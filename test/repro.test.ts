import { describe, test, expect, beforeEach, afterEach } from "bun:test"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"
import { MicrosoftTeamsPlugin } from "../src/index"

type Route = (url: string, init?: RequestInit) => Response | null
let routes: Route[]
let fetchCalls: { url: string; init?: RequestInit }[]
const realFetch = globalThis.fetch

const json = (o: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(o), { status, headers: { "Content-Type": "application/json", ...headers } })

let tmp: string, configDir: string, dataDir: string, tokenFile: string, logFile: string, configFile: string
let activeHooks: any[] = []
const CHAT_ID = "19:realchat@thread.v2"

function writeConfig() {
  const cfg = {
    azure: { tenantId: "tenant-test", clientId: "client-test", scopes: "Chat.Read offline_access" },
    polling: { intervalMs: 5, backoff1Ms: 600_000, backoff2Ms: 3_600_000 },
    log: { file: logFile, enabled: true },
    tokens: { file: tokenFile },
  }
  fs.mkdirSync(path.dirname(configFile), { recursive: true })
  fs.writeFileSync(configFile, JSON.stringify(cfg, null, 2))
}

beforeEach(() => {
  routes = []
  fetchCalls = []
  globalThis.fetch = (async (input: unknown, init?: RequestInit) => {
    const url = String(input)
    fetchCalls.push({ url, init })
    for (const r of routes) {
      const res = r(url, init)
      if (res) return res
    }
    throw new Error(`unrouted fetch: ${url}`)
  }) as typeof fetch
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), "mst-"))
  configDir = path.join(tmp, "opencode")
  dataDir = path.join(configDir, "opencode-ms-teams")
  tokenFile = path.join(dataDir, "tokens.json")
  logFile = path.join(dataDir, "opencode-ms-teams.log")
  configFile = path.join(dataDir, "opencode-ms-teams.json")
  process.env.MICROSOFT_TEAMS_CONFIG_DIR = configDir
  writeConfig()
})

afterEach(async () => {
  for (const h of activeHooks) { try { await h.dispose?.() } catch {} }
  activeHooks = []
  globalThis.fetch = realFetch
  fs.rmSync(tmp, { recursive: true, force: true })
  delete process.env.MICROSOFT_TEAMS_CONFIG_DIR
})

const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms))

function writeTokens() {
  fs.mkdirSync(path.dirname(tokenFile), { recursive: true })
  fs.writeFileSync(tokenFile,
    JSON.stringify({ access_token: "AT", refresh_token: "RT", expires_at: Date.now() + 3_600_000 }),
  )
}

const asst = (text: string) => ({ info: { role: "assistant" }, parts: [{ type: "text", text }] })

const teamsMsg = (id: string, content: string, messageType = "message") => ({
  id, messageType, deletedDateTime: null, body: { contentType: "html", content },
})

function makeFakeClient(messagesRef: { current: object[] }) {
  const calls = { promptAsync: [] as any[], showToast: [] as any[] }
  const client = {
    session: {
      get: async (o: any) => ({ data: { id: o.path.id, title: "실기 세션" } }),
      messages: async () => ({ data: messagesRef.current }),
      promptAsync: async (o: any) => { calls.promptAsync.push(o); return { data: undefined, error: undefined } },
      prompt: async () => ({ data: undefined, error: undefined }),
    },
    tui: { showToast: async (o: any) => { calls.showToast.push(o); return { data: true } } },
  }
  return { client, calls }
}

async function initPlugin(fake: any, workspace = "/tmp") {
  const hooks = await MicrosoftTeamsPlugin({
    client: fake.client, project: {}, directory: workspace, worktree: workspace,
    serverUrl: new URL("http://localhost:1"), $: null,
  } as any)
  activeHooks.push(hooks)
  return hooks
}

async function fireIdle(hooks: any, sid: string) {
  await hooks.event({ event: { type: "session.idle", properties: { sessionID: sid } } })
}

async function fireChatMessage(hooks: any, sid: string, text: string) {
  const parts = [{ type: "text", text }]
  await hooks["chat.message"]!({ sessionID: sid } as never, { message: { role: "user" }, parts } as never)
  return parts
}

async function fireTextComplete(hooks: any, sid: string, mid: string, text: string) {
  const output = { text }
  await hooks["experimental.text.complete"]!({ sessionID: sid, messageID: mid, partID: mid + "_p1" }, output)
  return output.text
}

const CONNECT_MARKER = "<!-- microsoft-teams:connect-request v1 -->"

async function fireConnectTeams(hooks: any, sid: string) {
  await fireTextComplete(hooks, sid, `msg_connect_${sid}`, `${CONNECT_MARKER}\n**Microsoft Teams 연동을 시작합니다.**`)
  await sleep(30)
}

describe("BUG REPRO — 최초 Teams 연동 (사전 단계 포함)", () => {
  test("풀 시나리오: 1) 첫 프롬프트 (연동 시작) → 2) chatId paste → 3) 최초 idle. 과거 이력이 프롬프트로 주입되지 않아야 한다", async () => {
    writeTokens()

    // 채팅방에 10개의 과거 메시지가 이미 존재
    const pastMessages = Array.from({ length: 10 }, (_, i) =>
      teamsMsg(String(1_700_000_000_000 + i * 1000), `과거 메시지 ${i}`)
    )
    const latestPast = pastMessages[pastMessages.length - 1]

    // Graph 라우팅
    routes.push((url, init) => {
      // /me/chats — 채팅 목록
      if (url.includes("graph.microsoft.com/v1.0/me/chats")) {
        return json({
          value: [{ id: CHAT_ID, topic: "실기 채팅방", chatType: "group", lastUpdatedDateTime: "2026-08-01T00:00:00Z" }],
        })
      }
      // /chats/{id}/messages
      if (url.includes("graph.microsoft.com") && url.includes("/chats/") && url.includes("/messages")) {
        if (init?.method === "POST") return json({ id: "2000000000000" })
        if (url.includes("%24top=1")) return json({ value: [latestPast] })
        return json({ value: [...pastMessages].reverse() })
      }
      return null
    })

    const SID = "ses_real_session"
    const messagesRef = { current: [] as object[] }
    const fake = makeFakeClient(messagesRef)
    const hooks = await initPlugin(fake)

    // 사전: /connect-teams 스킬 호출로 lazy connect 활성화 + ensureChatDirectory fetch 완료 대기
    await fireConnectTeams(hooks, SID)
    await sleep(80)

    // ===== 사전 단계 1: 첫 프롬프트 (일반 텍스트, chatId 없음) =====
    await fireChatMessage(hooks, SID, "안녕, 뭐 좀 물어볼게")
    // Assistant 응답 → text.complete 훅에서 chat select 안내 append
    messagesRef.current = [asst("네 무엇을 도와드릴까요")]
    const withNotice = await fireTextComplete(hooks, SID, "msg_1", "네 무엇을 도와드릴까요")
    expect(withNotice).toContain("채팅방 선택 필요") // 사전 단계 정상 확인

    // 첫 idle → chatId 없어서 skip
    await fireIdle(hooks, SID)

    // ===== 사전 단계 2: 사용자가 chatId 붙여넣기 =====
    const parts = await fireChatMessage(hooks, SID, CHAT_ID)
    // 텍스트가 안내로 치환됨
    expect(parts[0].text).toContain("연동되었습니다")
    // Assistant 는 치환된 텍스트를 보고 응답
    messagesRef.current = [asst("연동 확인했습니다. 이어서 진행하겠습니다.")]

    // ===== 최초 유효 idle: 커서 seed → send → poll 시작 =====
    await fireIdle(hooks, SID)
    // 폴 루프가 최소 1회 돌 시간
    await sleep(60)

    // 검증: 과거 이력이 promptAsync 로 주입되지 않아야 한다
    if (fake.calls.promptAsync.length > 0) {
      console.error("[BUG] promptAsync payload:")
      for (const c of fake.calls.promptAsync) {
        console.error(JSON.stringify(c.body.parts, null, 2))
      }
    }
    expect(fake.calls.promptAsync.length).toBe(0)
  })
})
