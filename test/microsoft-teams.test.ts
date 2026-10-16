import { describe, test, expect, beforeEach, afterEach } from "bun:test"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"
import { MicrosoftTeamsPlugin } from "../src/index"

type Route = (url: string, init?: RequestInit) => Response | Promise<Response> | null
let routes: Route[]
let fetchCalls: { url: string; init?: RequestInit }[]
const realFetch = globalThis.fetch

const json = (o: unknown, status = 200, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(o), { status, headers: { "Content-Type": "application/json", ...headers } })

let tmp: string
let configDir: string
let dataDir: string
let tokenFile: string
let logFile: string
let configFile: string
let activeHooks: any[] = []

const CHAT_ID = "19:test@thread.v2"

function writeConfig(overrides: Record<string, unknown> = {}) {
  const cfg = {
    azure: {
      tenantId: "tenant-test",
      clientId: "client-test",
      scopes: "Chat.Read offline_access",
    },
    polling: {
      intervalMs: 10,
      backoff1Ms: 600_000,
      backoff2Ms: 3_600_000,
    },
    log: {
      file: logFile,
      enabled: true,
    },
    tokens: {
      file: tokenFile,
    },
    ...overrides,
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
  for (const h of activeHooks) {
    try {
      await h.dispose?.()
    } catch {}
  }
  activeHooks = []
  globalThis.fetch = realFetch
  fs.rmSync(tmp, { recursive: true, force: true })
  delete process.env.MICROSOFT_TEAMS_CONFIG_DIR
})

const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms))

function writeTokens(over: Record<string, unknown> = {}) {
  fs.mkdirSync(path.dirname(tokenFile), { recursive: true })
  fs.writeFileSync(
    tokenFile,
    JSON.stringify({ access_token: "AT", refresh_token: "RT", expires_at: Date.now() + 3_600_000, ...over }),
  )
}

const asst = (text: string) => ({ info: { role: "assistant" }, parts: [{ type: "text", text }] })
const userMsg = (text: string) => ({ info: { role: "user" }, parts: [{ type: "text", text }] })

const teamsMsg = (id: string, content: string, messageType = "message") => ({
  id,
  messageType,
  deletedDateTime: null,
  body: { contentType: "html", content },
})

function makeFakeClient(opts: { session?: object | null; messages?: object[] } = {}) {
  const calls = {
    get: [] as any[],
    messages: [] as any[],
    promptAsync: [] as any[],
    prompt: [] as any[],
    showToast: [] as any[],
    permissionRespond: [] as any[],
  }
  // permissionRespondResult 는 테스트에서 강제로 실패 (return null) 시킬 수 있도록 mutable 로 노출.
  const state = { permissionRespondResult: { data: true, error: undefined } as any }
  const client: any = {
    session: {
      get: async (o: any) => {
        calls.get.push(o)
        return { data: opts.session === undefined ? { id: o.path.id, title: "테스트 세션" } : opts.session }
      },
      messages: async (o: any) => {
        calls.messages.push(o)
        return { data: opts.messages ?? [] }
      },
      promptAsync: async (o: any) => {
        calls.promptAsync.push(o)
        return { data: undefined, error: undefined }
      },
      prompt: async (o: any) => {
        calls.prompt.push(o)
        return { data: undefined, error: undefined }
      },
    },
    // v1 SDK: permission 응답은 OpencodeClient top-level 메서드 (session.* 아님).
    // 근거: node_modules/@opencode-ai/sdk/dist/gen/sdk.gen.d.ts L381
    postSessionIdPermissionsPermissionId: async (o: any) => {
      calls.permissionRespond.push(o)
      return state.permissionRespondResult
    },
    tui: {
      showToast: async (o: any) => {
        calls.showToast.push(o)
        return { data: true }
      },
    },
  }
  return { client, calls, state }
}

async function initPlugin(fake: ReturnType<typeof makeFakeClient>, workspace = "/tmp") {
  const hooks = await MicrosoftTeamsPlugin({
    client: fake.client,
    project: {},
    directory: workspace,
    worktree: workspace,
    serverUrl: new URL("http://localhost:1"),
    $: null,
  } as any)
  activeHooks.push(hooks)
  return hooks
}

async function fireIdle(hooks: any, sid = "ses_abc12345") {
  await hooks.event({ event: { type: "session.idle", properties: { sessionID: sid } } })
}

const CONNECT_MARKER = "<!-- microsoft-teams:connect-request v1 -->"

// /connect-teams 스킬을 호출한 것과 동일한 효과 (assistant 응답 marker 감지 → activate).
async function fireConnectTeams(hooks: any, sid = "ses_abc12345") {
  const output = { text: `${CONNECT_MARKER}\n**Microsoft Teams 연동을 시작합니다.**` }
  await hooks["experimental.text.complete"]!({ sessionID: sid, messageID: `msg_connect_${sid}`, partID: "p_connect" }, output)
  await sleep(30)
  return output
}

async function selectChat(hooks: any, sid = "ses_abc12345", chatId = CHAT_ID) {
  await fireConnectTeams(hooks, sid)
  const parts = [{ type: "text", text: chatId }]
  await hooks["chat.message"]!({ sessionID: sid } as never, { message: { role: "user" }, parts } as never)
}

function routeChatList(chats: Array<{ id: string; topic?: string; chatType?: string; lastActivity?: string }>) {
  routes.push((url, init) =>
    url.includes("graph.microsoft.com/v1.0/me/chats") && (init?.method ?? "GET") === "GET"
      ? json({
          value: chats.map((c) => ({
            id: c.id,
            topic: c.topic ?? null,
            chatType: c.chatType ?? "group",
            lastUpdatedDateTime: c.lastActivity ?? "2026-08-01T00:00:00Z",
          })),
        })
      : null,
  )
}

function routeTokenRefresh(resp: object = { access_token: "AT2", refresh_token: "RT2", expires_in: 3600 }, status = 200) {
  routes.push((url, init) =>
    url.includes("/oauth2/v2.0/token") && String(init?.body).includes("grant_type=refresh_token")
      ? json(resp, status)
      : null,
  )
}

function routeDeviceCode() {
  routes.push((url) =>
    url.includes("/oauth2/v2.0/devicecode")
      ? json({ device_code: "DC", user_code: "USERCODE1", verification_uri: "https://login.microsoft.com/device", interval: 3600, expires_in: 900 })
      : null,
  )
}

function routeGraphSend(id = "2000") {
  routes.push((url, init) =>
    url.includes("graph.microsoft.com") && url.includes("/messages") && init?.method === "POST" ? json({ id }) : null,
  )
}

function routeGraphList(msgs: object[]) {
  routes.push((url, init) =>
    url.includes("graph.microsoft.com") && url.includes("%24top") && (init?.method ?? "GET") === "GET"
      ? json({ value: msgs })
      : null,
  )
}

describe("config file", () => {
  test("설정 파일이 없으면 default 값으로 자동 생성된다", async () => {
    fs.rmSync(configFile, { force: true })
    writeTokens()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    await initPlugin(fake)
    expect(fs.existsSync(configFile)).toBe(true)
    const c = JSON.parse(fs.readFileSync(configFile, "utf8"))
    expect(c.azure.tenantId).toBeDefined()
    expect(c.log.enabled).toBe(false)
    expect(c.teams).toBeUndefined()
  })

  test("설정 파일은 데이터 디렉토리(opencode-ms-teams/) 하위에 생성된다", async () => {
    fs.rmSync(configFile, { force: true })
    fs.rmSync(dataDir, { recursive: true, force: true })
    writeTokens()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    await initPlugin(fake)
    expect(configFile).toBe(path.join(dataDir, "opencode-ms-teams.json"))
    expect(fs.existsSync(configFile)).toBe(true)
    expect(fs.existsSync(path.join(configDir, "opencode-ms-teams.json"))).toBe(false)
  })

  test("설정 파일 grouped 스키마의 azure/polling/log/tokens 필드가 읽힌다 (teams.chatId 는 스키마 밖)", async () => {
    writeConfig({
      polling: { intervalMs: 20, backoff1Ms: 100_000, backoff2Ms: 200_000 },
    })
    writeTokens()
    routeGraphList([])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_abc12345", "19:custom@thread.v2")
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    expect(send!.url).toContain("19%3Acustom%40thread.v2")
  })

  test("환경변수 (OPENCODE_TEAMS_*) 는 완전히 무시된다", async () => {
    process.env.OPENCODE_TEAMS_CHAT_ID = "19:env-ignored@thread.v2"
    writeConfig()
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    expect(send!.url).not.toContain("env-ignored")
    expect(send!.url).toContain("19%3Atest%40thread.v2")
    delete process.env.OPENCODE_TEAMS_CHAT_ID
  })
})

describe("log", () => {
  test("log.enabled=false 이면 로그 파일이 생성되지 않는다", async () => {
    writeConfig({ log: { file: logFile, enabled: false } })
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    expect(fs.existsSync(logFile)).toBe(false)
  })

  test("log.enabled=true 이면 플러그인 로드 시 로그 파일이 truncate 된다", async () => {
    fs.mkdirSync(path.dirname(logFile), { recursive: true })
    fs.writeFileSync(logFile, "OLD-CONTENT-BEFORE-INIT\n")
    writeConfig({ log: { file: logFile, enabled: true } })
    writeTokens()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    await initPlugin(fake)
    const content = fs.readFileSync(logFile, "utf8")
    expect(content).not.toContain("OLD-CONTENT-BEFORE-INIT")
    expect(content).toContain("plugin.loaded")
  })
})

describe("auth", () => {
  test("만료 임박 토큰은 refresh되고 회전된 refresh_token이 저장된다", async () => {
    writeTokens({ expires_at: Date.now() + 10_000 })
    routeTokenRefresh()
    routeChatList([{ id: CHAT_ID, topic: "test" }])
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("hello")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const refreshCall = fetchCalls.find((c) => c.url.includes("/oauth2/v2.0/token"))
    expect(refreshCall).toBeDefined()
    expect(refreshCall!.url).toContain("tenant-test")
    expect(String(refreshCall!.init?.body)).toContain("client_id=client-test")
    const saved = JSON.parse(fs.readFileSync(tokenFile, "utf8"))
    expect(saved.refresh_token).toBe("RT2")
    expect(saved.access_token).toBe("AT2")
    expect(saved.expires_at).toBeGreaterThan(Date.now())
  })

  test("expires_at 없는 토큰파일도 refresh 경로로 처리된다", async () => {
    writeTokens({ expires_at: undefined })
    routeTokenRefresh()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    expect(fetchCalls.some((c) => c.url.includes("/oauth2/v2.0/token"))).toBe(true)
  })

  test("토큰파일이 없으면 device flow가 시작되고 코드가 프롬프트+로그로 안내된다", async () => {
    routeDeviceCode()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    const connectOutput = await fireConnectTeams(hooks)
    await fireIdle(hooks)
    await sleep(30)
    expect(fetchCalls.some((c) => c.url.includes("/devicecode"))).toBe(true)
    expect(connectOutput.text).toContain("USERCODE1")
    expect(connectOutput.text).toContain("microsoft-teams:auth-notice")
    expect(fake.calls.showToast.some((t: any) => JSON.stringify(t).includes("USERCODE1"))).toBe(false)
    const logText = fs.readFileSync(logFile, "utf8")
    expect(logText).toContain("USERCODE1")
  })

  test("로그인 완료 시 Toast 가 'Microsoft Teams' 타이틀과 success variant 로 발생한다", async () => {
    routes.push((url) =>
      url.includes("/oauth2/v2.0/devicecode")
        ? json({ device_code: "DC", user_code: "UC2", verification_uri: "https://login.microsoft.com/device", interval: 0, expires_in: 900 })
        : null,
    )
    routes.push((url, init) => {
      if (!url.includes("/oauth2/v2.0/token")) return null
      if (!String(init?.body).includes("device_code")) return null
      return json({ access_token: "AT_OK", refresh_token: "RT_OK", expires_in: 3600 })
    })
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks)
    await sleep(100)
    const successToasts = fake.calls.showToast.filter((t: any) => t.body.variant === "success")
    expect(successToasts.length).toBeGreaterThan(0)
    expect(successToasts[0].body.title).toBe("Microsoft Teams")
    expect(successToasts[0].body.message).toContain("로그인 완료")
  })

  test("refresh가 invalid_grant면 device flow로 전환된다", async () => {
    writeTokens({ expires_at: 0 })
    routeTokenRefresh({ error: "invalid_grant" }, 400)
    routeDeviceCode()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks)
    await fireIdle(hooks)
    await sleep(30)
    expect(fetchCalls.some((c) => c.url.includes("/devicecode"))).toBe(true)
  })

  test("서브세션(parentID 존재)의 idle은 auth조차 시도하지 않는다", async () => {
    writeTokens()
    routeChatList([])
    const fake = makeFakeClient({ session: { id: "ses_child", title: "child", parentID: "ses_parent" } })
    const hooks = await initPlugin(fake)
    await fireIdle(hooks, "ses_child")
    const sessionFetches = fetchCalls.filter(
      (c) => c.url.includes("/oauth2/") || (c.url.includes("graph.microsoft.com") && c.url.includes("/chats/")),
    )
    expect(sessionFetches.length).toBe(0)
    expect(fake.calls.messages.length).toBe(0)
  })

  test("assistant 텍스트가 없으면(도구 호출만) 아무것도 하지 않는다", async () => {
    writeTokens()
    routeChatList([])
    const fake = makeFakeClient({ messages: [userMsg("질문"), { info: { role: "assistant" }, parts: [{ type: "tool" }] }] })
    const hooks = await initPlugin(fake)
    await fireIdle(hooks)
    const sessionFetches = fetchCalls.filter(
      (c) => c.url.includes("/oauth2/") || (c.url.includes("graph.microsoft.com") && c.url.includes("/chats/")),
    )
    expect(sessionFetches.length).toBe(0)
  })

  test("유효한(만료 임박 아님) 토큰은 재사용되고 refresh를 호출하지 않는다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("hello")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    expect(fetchCalls.some((c) => c.url.includes("/oauth2/v2.0/token"))).toBe(false)
  })

  test("중복 idle 스킵: 같은 세션·같은 assistant 텍스트의 idle은 두 번째부터 추가 fetch 없음", async () => {
    writeTokens({ expires_at: Date.now() + 10_000 })
    routeTokenRefresh()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("hello")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const countAfterFirst = fetchCalls.length
    await fireIdle(hooks)
    expect(fetchCalls.length).toBe(countAfterFirst)
  })

  test("refresh로 저장된 토큰 파일은 0600 권한을 갖는다", async () => {
    writeTokens({ expires_at: Date.now() + 10_000 })
    routeTokenRefresh()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("hello")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const mode = fs.statSync(tokenFile).mode & 0o777
    expect(mode).toBe(0o600)
  })
})

describe("background token refresh (요구사항 1)", () => {
  test("유효 토큰이 있으면 activate 후 refresh 예약이 로그에 기록된다", async () => {
    writeTokens({ expires_at: Date.now() + 3_600_000 })
    const fake = makeFakeClient({ messages: [asst("hello")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks)
    expect(fetchCalls.some((c) => c.url.includes("/oauth2/v2.0/token"))).toBe(false)
    const logText = fs.readFileSync(logFile, "utf8")
    expect(logText).toContain("token.refresh.scheduled")
  })

  test("dispose 호출 시 예약된 refresh timer 가 취소되어 이후 fetch 가 발생하지 않는다", async () => {
    writeTokens({ expires_at: Date.now() + 3_600_000 })
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks)
    await (hooks as any).dispose?.()
    const before = fetchCalls.length
    await sleep(50)
    expect(fetchCalls.length).toBe(before)
  })

  test("device flow 성공 후에도 다음 refresh 가 자동 예약된다 (self-scheduling)", async () => {
    routes.push((url) =>
      url.includes("/oauth2/v2.0/devicecode")
        ? json({
            device_code: "DC",
            user_code: "UC",
            verification_uri: "https://login.microsoft.com/device",
            interval: 0,
            expires_in: 900,
          })
        : null,
    )
    routes.push((url, init) => {
      if (!url.includes("/oauth2/v2.0/token")) return null
      if (!String(init?.body).includes("device_code")) return null
      return json({ access_token: "AT_NEW", refresh_token: "RT_NEW", expires_in: 3600 })
    })
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks)
    await sleep(200)
    const logText = fs.readFileSync(logFile, "utf8")
    const scheduledCount = (logText.match(/token\.refresh\.scheduled/g) ?? []).length
    expect(scheduledCount).toBeGreaterThanOrEqual(1)
  })
})

describe("outbound", () => {
  test("idle → 커서 seed(top=1) 후 HTML 발신, sentIds 기록", async () => {
    writeTokens()
    routeGraphList([teamsMsg("1000", "<p>기존 메시지</p>")])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("**결과** 완료\n`code`")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const seed = fetchCalls.find((c) => c.url.includes("%24top=1"))
    expect(seed).toBeDefined()
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("19%3Atest%40thread.v2"))
    expect(send).toBeDefined()
    const body = JSON.parse(String(send!.init!.body))
    expect(body.body.contentType).toBe("html")
    expect(body.body.content).toContain("<b>결과</b>")
    expect(body.body.content).toMatch(/<span style="[^"]*monospace[^"]*">code<\/span>/)
    expect(body.body.content).toContain("테스트 세션")
  })

  test("같은 내용이면 재발신하지 않는다 (fingerprint 가드)", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("동일 응답")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await fireIdle(hooks)
    const sends = fetchCalls.filter((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    expect(sends.length).toBe(1)
  })

  test("코드펜스는 <pre><code> 로 렌더되고 HTML 특수문자는 이스케이프된다 (line-number 회피: class 없음)", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("설명 <중요>\n```ts\nconst a = 1 < 2\nconst b = 3\n```")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    expect(content).toContain("&lt;중요&gt;")
    expect(content).toMatch(/<pre style="[^"]*monospace[^"]*"><code>const a = 1 &lt; 2\nconst b = 3<\/code><\/pre>/)
    expect(content).not.toMatch(/class="language-/)
    expect(content).not.toContain(">ts<")
  })

  test("들여쓰기 있는 코드가 <pre><code> 안에서 원본 그대로 보존된다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const md = "```python\nprint('hi')\nfor i in range(3):\n    print(i)\n```"
    const fake = makeFakeClient({ messages: [asst(md)] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    expect(content).toContain("<pre style=")
    expect(content).toContain("<code>print('hi')")
    expect(content).toContain("    print(i)")
    expect(content).toContain("</code></pre>")
  })

  test("여러 코드펜스가 연속으로 있어도 각각 코드블록으로 파싱된다 (평문 뭉침 회귀 방지)", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const md = "설명1\n\n```js\nconst a = 1\n```\n\n중간\n\n```py\nprint(2)\n```\n\n끝"
    const fake = makeFakeClient({ messages: [asst(md)] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    const codeBlocks = content.match(/<pre style="[^"]*"><code>/g) ?? []
    expect(codeBlocks.length).toBe(2)
    expect(content).toContain("const a = 1")
    expect(content).toContain("print(2)")
    expect(content).toContain("설명1")
    expect(content).toContain("중간")
    expect(content).toContain("끝")
    expect(content).not.toContain("```")
  })

  test("코드펜스 안에 인라인 백틱이 있어도 fence 로 잘못 인식되지 않는다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const md = "```ts\nconst s = `hello ${x}`\n```"
    const fake = makeFakeClient({ messages: [asst(md)] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    const codeBlocks = content.match(/<pre style="[^"]*"><code>/g) ?? []
    expect(codeBlocks.length).toBe(1)
    expect(content).toContain("hello ${x}")
  })

  test("25KB 초과 본문은 잘리고 '(생략)'이 붙는다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const fake = makeFakeClient({ messages: [asst("가".repeat(30_000))] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    expect(Buffer.byteLength(content, "utf8")).toBeLessThanOrEqual(25_000)
    expect(content).toContain("…(생략)")
  })

  test("발신 실패(null id) 시 waiting 미등록 → 다음 idle에서 재시도된다", async () => {
    writeTokens()
    routeGraphList([])
    routes.push((url, init) =>
      url.includes("graph.microsoft.com") && url.includes("/chats/") && init?.method === "POST"
        ? json({ error: { code: "x" } }, 500)
        : null,
    )
    const fake = makeFakeClient({ messages: [asst("실패할 발신")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    routes.length = 0
    routeGraphList([])
    routeGraphSend()
    await fireIdle(hooks)
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(2)
  })

  test("최초 연동 시 latestMessageId(top=1) 실패 → 발신 skip + 과거 이력 미주입 (cursor=0 오염 차단)", async () => {
    writeTokens()
    routes.push((url, init) =>
      url.includes("graph.microsoft.com") &&
      url.includes("/chats/") &&
      url.includes("%24top=1") &&
      (init?.method ?? "GET") === "GET"
        ? json({ error: { code: "throttle" } }, 500)
        : null,
    )
    routeGraphSend("2000")
    routeGraphList([
      teamsMsg("990", "과거 대화 1"),
      teamsMsg("991", "과거 대화 2"),
      teamsMsg("992", "과거 대화 3"),
    ])
    const fake = makeFakeClient({ messages: [asst("첫 응답")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(80)
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(0)
    expect(fake.calls.promptAsync.length).toBe(0)
  })

  test("최초 연동 후 latestMessageId 성공 시 이후 도착 답신만 주입된다 (과거 이력 필터)", async () => {
    writeTokens()
    routeGraphSend("2000")
    let listCalls = 0
    routes.push((url, init) => {
      if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
      if ((init?.method ?? "GET") !== "GET") return null
      if (url.includes("%24top=1")) return json({ value: [teamsMsg("1500", "seed 시점 최신")] })
      listCalls++
      if (listCalls === 1) {
        return json({
          value: [
            teamsMsg("1500", "seed 시점 최신"),
            teamsMsg("1400", "과거 대화 A"),
            teamsMsg("1300", "과거 대화 B"),
          ],
        })
      }
      return json({
        value: [
          teamsMsg("1600", "#ask 연동 이후 신규 답신"),
          teamsMsg("1500", "seed 시점 최신"),
          teamsMsg("1400", "과거 대화 A"),
        ],
      })
    })
    const fake = makeFakeClient({ messages: [asst("첫 응답")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(120)
    expect(fake.calls.promptAsync.length).toBe(1)
    const injected = fake.calls.promptAsync[0].body.parts[0].text as string
    expect(injected).toContain("연동 이후 신규 답신")
    expect(injected).not.toContain("과거 대화 A")
    expect(injected).not.toContain("과거 대화 B")
    expect(injected).not.toContain("seed 시점 최신")
  })
})

describe("inbound", () => {
  test("답신 도착 → promptAsync로 주입되고 마커가 붙는다", async () => {
    writeTokens()
    routeGraphSend("2000")
    let listCalls = 0
    routes.push((url, init) => {
      if (!url.includes("%24top") || (init?.method ?? "GET") !== "GET") return null
      if (!url.includes("/chats/")) return null
      listCalls++
      if (listCalls <= 2) return json({ value: [] })
      return json({ value: [teamsMsg("3000", "<p>#ask 답신입니다 &amp; 진행해줘</p>"), teamsMsg("2000", "내가 보낸 것")] })
    })
    const fake = makeFakeClient({ messages: [asst("작업 완료")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(120)
    expect(fake.calls.promptAsync.length).toBe(1)
    const arg = fake.calls.promptAsync[0]
    expect(arg.path.id).toBe("ses_abc12345")
    const text = arg.body.parts[0].text as string
    expect(text).toContain("<!-- microsoft-teams:injected -->")
    expect(text).toContain("답신입니다 & 진행해줘")
    expect(text).not.toContain("내가 보낸 것")
  })

  test("systemEventMessage 등 비일반 메시지는 주입되지 않는다", async () => {
    writeTokens()
    routeGraphSend("2000")
    routeGraphList([teamsMsg("3000", "<systemEventMessage/>", "systemEventMessage")])
    const fake = makeFakeClient({ messages: [asst("결과")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(60)
    expect(fake.calls.promptAsync.length).toBe(0)
  })

  test("여러 답신은 id 오름차순으로 병합 주입된다", async () => {
    writeTokens()
    routeGraphSend("2000")
    routes.push((url, init) => {
      if (!url.includes("%24top") || (init?.method ?? "GET") !== "GET") return null
      if (!url.includes("/chats/")) return null
      if (url.includes("%24top=1")) return json({ value: [] })
      return json({ value: [teamsMsg("3200", "#ask 두 번째"), teamsMsg("3100", "#ask 첫 번째")] })
    })
    const fake = makeFakeClient({ messages: [asst("결과")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(80)
    expect(fake.calls.promptAsync.length).toBe(1)
    const text = fake.calls.promptAsync[0].body.parts[0].text as string
    expect(text.indexOf("첫 번째")).toBeLessThan(text.indexOf("두 번째"))
  })

  test("발신 시 커서 미전진 — 발신 직전 도착한 답신(발신 id보다 작은 id)도 주입된다", async () => {
    writeTokens()
    routes.push((url, init) => {
      if (!url.includes("%24top") || (init?.method ?? "GET") !== "GET") return null
      if (!url.includes("/chats/")) return null
      if (url.includes("%24top=1")) return json({ value: [teamsMsg("1000", "기존")] })
      return json({ value: [teamsMsg("2000", "내 발신"), teamsMsg("1500", "#ask 발신 직전 도착한 답신")] })
    })
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("결과")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(80)
    expect(fake.calls.promptAsync.length).toBe(1)
    const text = fake.calls.promptAsync[0].body.parts[0].text as string
    expect(text).toContain("발신 직전 도착한 답신")
    expect(text).not.toContain("내 발신")
  })

  test("busy 중 큐잉 실제 시나리오 — 새 응답으로 idle 시 커서 이전 미처리 답신이 즉시 주입된다", async () => {
    writeTokens()
    let sendCount = 0
    routes.push((url, init) => {
      if (url.includes("graph.microsoft.com") && url.includes("/chats/") && init?.method === "POST") {
        sendCount++
        return json({ id: sendCount === 1 ? "2000" : "4000" })
      }
      return null
    })
    let phase = 0
    routes.push((url, init) => {
      if (!url.includes("%24top") || (init?.method ?? "GET") !== "GET") return null
      if (!url.includes("/chats/")) return null
      if (url.includes("%24top=1")) return json({ value: [] })
      if (phase === 0) return json({ value: [teamsMsg("3000", "#ask 1차 답신")] })
      return json({ value: [teamsMsg("4000", "내 2차 발신"), teamsMsg("3500", "#ask busy 중 답신"), teamsMsg("3000", "#ask 1차 답신")] })
    })
    const fakeMessages: object[] = [asst("1차 결과")]
    const fake = makeFakeClient({ messages: fakeMessages })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(80)
    expect(fake.calls.promptAsync.length).toBe(1)
    phase = 1
    fakeMessages.length = 0
    fakeMessages.push(asst("2차 결과 — 답신 반영"))
    await fireIdle(hooks)
    await sleep(80)
    expect(sendCount).toBe(2)
    expect(fake.calls.promptAsync.length).toBe(2)
    const text = fake.calls.promptAsync[1].body.parts[0].text as string
    expect(text).toContain("busy 중 답신")
    expect(text).not.toContain("내 2차 발신")
  })
})

describe("rendering", () => {
  test("마크다운 표는 <table>/<thead>/<tbody>/<th>/<td> 로 렌더된다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const md = "결과 요약:\n\n| 이름 | 값 |\n|:-----|---:|\n| a    | 1  |\n| b    | 2  |\n\n끝."
    const fake = makeFakeClient({ messages: [asst(md)] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    expect(content).toContain("<table")
    expect(content).toContain("<thead>")
    expect(content).toContain("<tbody>")
    expect(content).toContain('<th align="left">이름</th>')
    expect(content).toContain('<th align="right">값</th>')
    expect(content).toContain('<td align="left">a</td>')
    expect(content).toContain('<td align="right">1</td>')
    expect(content).toContain("결과 요약")
    expect(content).toContain("끝.")
  })

  test("코드 블록은 <pre><code> 로 렌더되며 syntax highlight class 는 부착되지 않는다 (Teams line-number 회피)", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const md = "```python\nprint('hi')\n```"
    const fake = makeFakeClient({ messages: [asst(md)] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    expect(content).toMatch(/<pre style="[^"]*"><code>print\('hi'\)<\/code><\/pre>/)
    expect(content).not.toMatch(/class="language-/)
  })
})

describe("auth notice append (experimental.text.complete)", () => {
  test("activate (connect trigger) 이후 토큰이 없으면 device flow 가 즉시 시작된다", async () => {
    routeDeviceCode()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks)
    await sleep(30)
    expect(fetchCalls.some((c) => c.url.includes("/devicecode"))).toBe(true)
  })

  test("device code 가 준비되면 text.complete 훅이 assistant 응답 끝에 안내를 append 한다", async () => {
    routeDeviceCode()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks, "ses_x")
    await sleep(80)
    const output = { text: "안녕하세요! 무엇을 도와드릴까요?" }
    await hooks["experimental.text.complete"]!({ sessionID: "ses_x", messageID: "msg_1", partID: "prt_1" }, output)
    expect(output.text.startsWith("안녕하세요! 무엇을 도와드릴까요?")).toBe(true)
    expect(output.text).toContain("<!-- microsoft-teams:auth-notice -->")
    expect(output.text).toContain("USERCODE1")
    expect(output.text).toContain("Microsoft Teams 로그인 필요")
  })

  test("같은 messageID 로 훅이 여러 번 호출되어도 안내는 1회만 append 된다", async () => {
    routeDeviceCode()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks, "s")
    await sleep(80)
    const o1 = { text: "part 1" }
    const o2 = { text: "part 2" }
    await hooks["experimental.text.complete"]!({ sessionID: "s", messageID: "msg_dup", partID: "p1" }, o1)
    await hooks["experimental.text.complete"]!({ sessionID: "s", messageID: "msg_dup", partID: "p2" }, o2)
    expect(o1.text).toContain("USERCODE1")
    expect(o2.text).toBe("part 2")
  })

  test("Microsoft Teams 발신 시 auth-notice 마커 이후는 제거되어 원본만 전송된다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend()
    const original = "실제 응답 본문입니다."
    const withNotice = `${original}\n\n<!-- microsoft-teams:auth-notice -->\n🔐 로그인 필요: 이 부분은 안 감`
    const fake = makeFakeClient({ messages: [asst(withNotice)] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    const send = fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com"))
    const content = JSON.parse(String(send!.init!.body)).body.content as string
    expect(content).toContain("실제 응답 본문입니다")
    expect(content).not.toContain("로그인 필요")
    expect(content).not.toContain("auth-notice")
  })
})

describe("errors", () => {
  test("Graph 401 → forceRefresh 후 1회 재시도로 성공한다", async () => {
    writeTokens()
    routeTokenRefresh({ access_token: "AT2", refresh_token: "RT2", expires_in: 3600 })
    let graphCalls = 0
    routes.push((url, init) => {
      if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
      graphCalls++
      if (graphCalls <= 1) return json({ error: { code: "InvalidAuthenticationToken" } }, 401)
      if (init?.method === "POST") return json({ id: "2000" })
      return json({ value: [] })
    })
    const fake = makeFakeClient({ messages: [asst("401 테스트")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    expect(fetchCalls.some((c) => String(c.init?.body ?? "").includes("grant_type=refresh_token"))).toBe(true)
    expect(
      fetchCalls.some(
        (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
      ),
    ).toBe(true)
  })

  test("Graph 429 → Retry-After 후 재시도한다", async () => {
    writeTokens()
    let graphCalls = 0
    routes.push((url, init) => {
      if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
      graphCalls++
      if (graphCalls === 1) return json({ error: { code: "TooManyRequests" } }, 429, { "Retry-After": "0" })
      if (init?.method === "POST") return json({ id: "2000" })
      return json({ value: [] })
    })
    const fake = makeFakeClient({ messages: [asst("429 테스트")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    expect(graphCalls).toBeGreaterThanOrEqual(2)
  })

  test("폴링 중 네트워크 오류가 나도 루프가 죽지 않고 다음 답신을 잡는다", async () => {
    writeTokens()
    routeGraphSend("2000")
    let listCalls = 0
    routes.push((url, init) => {
      if (!url.includes("%24top") || (init?.method ?? "GET") !== "GET") return null
      if (!url.includes("/chats/")) return null
      listCalls++
      if (url.includes("%24top=1")) return json({ value: [] })
      if (listCalls === 2) throw new Error("ECONNRESET")
      return json({ value: [teamsMsg("3000", "#ask 복구 후 답신")] })
    })
    const fake = makeFakeClient({ messages: [asst("네트워크 테스트")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(150)
    expect(fake.calls.promptAsync.length).toBe(1)
  })
})

describe("session model/agent injection", () => {
  test("assistant 메시지의 model/mode 를 답신 주입 시 promptAsync body 에 전달한다", async () => {
    writeTokens()
    const asstWithModel = {
      info: {
        role: "assistant",
        modelID: "claude-opus-4-7",
        providerID: "anthropic",
        mode: "build",
      },
      parts: [{ type: "text", text: "assistant 응답" }],
    }
    routes.push((url, init) => {
      if (url.includes("graph.microsoft.com") && url.includes("/chats/") && init?.method === "POST")
        return json({ id: "9999" })
      return null
    })
    routes.push((url, init) => {
      if (!url.includes("%24top") || (init?.method ?? "GET") !== "GET") return null
      if (!url.includes("/chats/")) return null
      if (url.includes("%24top=1")) return json({ value: [] })
      return json({ value: [teamsMsg("11000", "#ask 답신 메시지")] })
    })
    const fake = makeFakeClient({ messages: [asstWithModel] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(100)
    expect(fake.calls.promptAsync.length).toBe(1)
    const arg = fake.calls.promptAsync[0]
    expect(arg.body.model).toEqual({ providerID: "anthropic", modelID: "claude-opus-4-7" })
    expect(arg.body.agent).toBe("build")
  })

  test("assistant 메시지에 model/mode 가 없으면 promptAsync body 에 model/agent 필드가 빠진다", async () => {
    writeTokens()
    routes.push((url, init) => {
      if (url.includes("graph.microsoft.com") && url.includes("/chats/") && init?.method === "POST")
        return json({ id: "9999" })
      return null
    })
    routes.push((url, init) => {
      if (!url.includes("%24top") || (init?.method ?? "GET") !== "GET") return null
      if (!url.includes("/chats/")) return null
      if (url.includes("%24top=1")) return json({ value: [] })
      return json({ value: [teamsMsg("11000", "#ask 답신 메시지")] })
    })
    const fake = makeFakeClient({ messages: [asst("모델 정보 없는 응답")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(100)
    expect(fake.calls.promptAsync.length).toBe(1)
    const arg = fake.calls.promptAsync[0]
    expect(arg.body.model).toBeUndefined()
    expect(arg.body.agent).toBeUndefined()
    expect(arg.body.parts[0].type).toBe("text")
  })
})

// permission.ask plugin hook 은 opencode 1.4.17 core 에 trigger 호출부가 없어 dead API 다
// (omo-opencode-hook-analysis.md §2, §6.5 참고). 실제 발화하는 이벤트는 permission.asked 다.
// 따라서 알림은 event hook 의 permission.asked 분기로 옮겼고, 아래 케이스도 event 로 검증한다.
describe("permission.asked event (요구사항 1)", () => {
  const firePerm = async (
    hooks: any,
    props: {
      id: string
      sessionID: string
      permission?: string
      type?: string
      title?: string
      metadata?: Record<string, unknown>
      pattern?: string | string[]
    },
  ) => {
    await hooks.event({ event: { type: "permission.asked", properties: props } })
  }

  test("workspace 외부 경로 접근은 Microsoft Teams 로 알림을 발신한다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    const outside = "/etc/passwd"
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_x")
    await firePerm(hooks, {
      id: "p2",
      type: "read",
      sessionID: "ses_x",
      pattern: outside,
      title: "read /etc/passwd",
    })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    expect(body).toContain("권한 요청")
    expect(body).toContain("/etc/passwd")
    expect(body).toContain(workspace)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("workspace 내부 경로도 알림이 발신되지만 outside 리스트는 비어있다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    const inside = path.join(workspace, "src/foo.ts")
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_x")
    await firePerm(hooks, { id: "p1", type: "read", sessionID: "ses_x", pattern: inside, title: "read" })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    expect(body).toContain("권한 요청")
    expect(body).not.toContain("Workspace 외부 경로")
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("같은 permission.id 로 여러 번 발화해도 알림은 1회만 발신된다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "s")
    for (let i = 0; i < 3; i++) {
      await firePerm(hooks, { id: "same-id", type: "read", sessionID: "s", pattern: "/etc/passwd" })
    }
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("metadata.cwd 필드에서도 경로를 추출해 알림 본문에 포함한다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "s")
    await firePerm(hooks, {
      id: "p3",
      type: "bash",
      sessionID: "s",
      metadata: { cwd: "/opt/other" },
      title: "bash",
    })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    expect(body).toContain("/opt/other")
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("chatId 미선택 세션의 permission 은 Teams 알림을 발신하지 않는다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await fireConnectTeams(hooks, "ses_no_chat")
    await firePerm(hooks, {
      id: "p-nochat",
      type: "read",
      sessionID: "ses_no_chat",
      pattern: "/etc/passwd",
    })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(0)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("connectRequested=false 상태에서는 permission.asked 이벤트가 알림을 발신하지 않는다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await firePerm(hooks, { id: "p-off", type: "read", sessionID: "ses_off", pattern: "/etc/passwd" })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(0)
    expect(fetchCalls.some((c) => c.url.includes("/oauth2/"))).toBe(false)
    fs.rmSync(workspace, { recursive: true, force: true })
  })
})

describe("question.asked event (Interview Mode / 요구사항 1)", () => {
  const fireQ = async (
    hooks: any,
    props: {
      id: string
      sessionID: string
      questions: Array<{ question?: string; header?: string; options?: unknown; multiple?: boolean }>
    },
  ) => {
    await hooks.event({ event: { type: "question.asked", properties: props } })
  }

  test("connectRequested + chatId 확정된 세션의 question.asked → Teams 알림 발신", async () => {
    writeTokens()
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_q")
    await fireQ(hooks, {
      id: "q1",
      sessionID: "ses_q",
      questions: [
        {
          header: "배포 승인",
          question: "production 에 배포할까요?",
          options: [
            { label: "yes", description: "지금 배포" },
            { label: "no", description: "취소" },
          ],
          multiple: false,
        },
      ],
    })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    expect(body).toContain("세션 질문")
    expect(body).toContain("배포 승인")
    expect(body).toContain("production 에 배포할까요")
    expect(body).toContain("yes")
    expect(body).toContain("no")
  })

  test("connectRequested=false 상태에서는 question.asked 가 무시된다", async () => {
    writeTokens()
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await fireQ(hooks, { id: "q-off", sessionID: "ses_off", questions: [{ header: "H", question: "Q" }] })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(0)
  })

  test("같은 question.id 로 여러 번 발화해도 알림은 1회만 발신된다", async () => {
    writeTokens()
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_q")
    for (let i = 0; i < 3; i++) {
      await fireQ(hooks, { id: "same-q", sessionID: "ses_q", questions: [{ header: "H", question: "Q" }] })
    }
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
  })

  test("sub session 의 question.asked 는 root chat 으로 bubble-up 되고 본문에 sub-session 배너가 포함된다 (invariant 10)", async () => {
    writeTokens()
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    ;(fake.client.session as any).get = async (o: any) => ({
      data:
        o.path.id === "ses_child_q"
          ? { id: "ses_child_q", title: "sub", parentID: "ses_parent" }
          : { id: o.path.id, title: "parent" },
    })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_parent")
    await fireQ(hooks, {
      id: "q-sub",
      sessionID: "ses_child_q",
      questions: [{ header: "H", question: "Q" }],
    })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    expect(sends[0].url).toContain(encodeURIComponent(CHAT_ID))
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    expect(body).toContain("sub-session")
    expect(body).toContain("🌿")
  })

  test("multiple 플래그가 true 이면 알림 본문에 (복수 선택) 표기가 붙는다", async () => {
    writeTokens()
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_qm")
    await fireQ(hooks, {
      id: "qm",
      sessionID: "ses_qm",
      questions: [{ header: "복수", question: "여러 개 고르세요", options: ["a", "b"], multiple: true }],
    })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    expect(body).toContain("복수 선택")
  })
})

describe("lazy connect (요구사항 2)", () => {
  test("플러그인 로드만 하고 아무 hook 도 fire 하지 않으면 auth / me/chats fetch 가 발생하지 않는다", async () => {
    writeTokens()
    const fake = makeFakeClient({ messages: [] })
    await initPlugin(fake)
    await sleep(50)
    expect(fetchCalls.some((c) => c.url.includes("/oauth2/"))).toBe(false)
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))).toBe(false)
  })

  test("connect trigger 이전 text.complete 는 output.text 를 절대 변형하지 않는다", async () => {
    writeTokens()
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    const output = { text: "assistant 응답 원본" }
    await hooks["experimental.text.complete"]!(
      { sessionID: "ses_off", messageID: "msg_off", partID: "p1" },
      output,
    )
    expect(output.text).toBe("assistant 응답 원본")
    expect(output.text).not.toContain("microsoft-teams")
  })

  test("connect trigger 이전 chat.message 는 me/chats fetch 를 트리거하지 않는다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await hooks["chat.message"]!(
      { sessionID: "ses_off" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: "일반 프롬프트" }] } as never,
    )
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))).toBe(false)
  })

  test("connect trigger 이전 session.idle 은 Teams 발신을 시도하지 않는다", async () => {
    writeTokens()
    const fake = makeFakeClient({ messages: [asst("hi")] })
    const hooks = await initPlugin(fake)
    await fireIdle(hooks, "ses_off")
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"))).toBe(false)
  })

  test("text.complete 로 CONNECT_TRIGGER_MARKER 감지 → auth.bootstrap + me/chats fetch 활성화", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    const output = { text: `${CONNECT_MARKER}\n**Microsoft Teams 연동을 시작합니다.**` }
    await hooks["experimental.text.complete"]!(
      { sessionID: "ses_on", messageID: "msg_on", partID: "p1" },
      output,
    )
    await sleep(50)
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))).toBe(true)
  })

  test("chat.message 의 /connect-teams slash 로도 활성화된다 (marker 없이)", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await hooks["chat.message"]!(
      { sessionID: "ses_slash" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: "/connect-teams" }] } as never,
    )
    await sleep(50)
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))).toBe(true)
  })

  // opencode TUI 는 skill 을 slash command 로 등록하고, /connect-teams 입력 시
  // 리터럴 "/connect-teams" 대신 SKILL.md 본문 (frontmatter 제외) 을 user message 로 보낸다.
  function expandedConnectSkillBody(): string {
    const raw = fs.readFileSync(path.join(configDir, "skills", "connect-teams", "SKILL.md"), "utf8")
    return raw.replace(/^---\n[\s\S]*?\n---\n/, "").trim()
  }

  test("opencode 가 /connect-teams 를 SKILL.md 본문으로 확장해 보내도 chat.message 에서 활성화된다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await hooks["chat.message"]!(
      { sessionID: "ses_skill" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: expandedConnectSkillBody() }] } as never,
    )
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))).toBe(true)
  })

  test("skill 본문 확장 경로는 Graph 응답이 2초를 넘어도 스킬 응답에 채팅방 목록을 첨부한다", async () => {
    writeTokens()
    routes.push((url, init) =>
      url.includes("graph.microsoft.com/v1.0/me/chats") && (init?.method ?? "GET") === "GET"
        ? sleep(2_300).then(() =>
            json({ value: [{ id: "19:slow-chat@thread.v2", topic: "느린방", chatType: "group", lastUpdatedDateTime: "2026-08-01T00:00:00Z" }] }),
          )
        : null,
    )
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await hooks["chat.message"]!(
      { sessionID: "ses_slow" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: expandedConnectSkillBody() }] } as never,
    )
    const output = { text: `${CONNECT_MARKER}\n**Microsoft Teams 연동을 시작합니다.**` }
    await hooks["experimental.text.complete"]!({ sessionID: "ses_slow", messageID: "msg_slow", partID: "p1" }, output)
    expect(output.text).toContain("19:slow-chat@thread.v2")
  }, 10_000)

  test("marker 만 포함되고 skill heading 줄이 없는 메시지는 연동을 활성화하지 않는다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await hooks["chat.message"]!(
      { sessionID: "ses_marker_only" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: `이 주석 뭐야? ${CONNECT_MARKER} 그리고 # Skill: connect-teams 도 설명해줘` }] } as never,
    )
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))).toBe(false)
  })

  test("skill heading 만 언급한 일반 메시지는 연동을 활성화하지 않는다 (marker 동반 필요)", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await hooks["chat.message"]!(
      { sessionID: "ses_mention" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: "# Skill: connect-teams 문서 설명해줘" }] } as never,
    )
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))).toBe(false)
  })

  test("chat.message 로 chatId 를 직접 붙여넣으면 implicit connect 로 활성화된다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await hooks["chat.message"]!(
      { sessionID: "ses_impl" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: "19:implicit@thread.v2" }] } as never,
    )
    await sleep(50)
    ;(fake.client.session as any).get = async (o: any) => ({ data: { id: o.path.id, title: "impl" } })
    await fireIdle(hooks, "ses_impl")
    expect(fetchCalls.some((c) => c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"))).toBe(true)
  })

  test("활성화 후에는 permission.ask hook 이 등록되어 있지 않다 (dead API 정리)", async () => {
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    expect((hooks as any)["permission.ask"]).toBeUndefined()
  })
})

describe("chat selection (요구사항 1, 2)", () => {
  test("chatId 없는 세션의 첫 idle 은 Teams 발신을 skip 한다", async () => {
    writeTokens()
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await fireIdle(hooks)
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(0)
  })

  test("activate 이후 chatId 없는 세션의 chat.message 는 me/chats?$top=10 을 호출해 목록을 캐시한다", async () => {
    writeTokens()
    routeChatList([
      { id: "19:chat-a@thread.v2", topic: "Alice" },
      { id: "19:chat-b@thread.v2", topic: "Bob" },
    ])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks, "ses_abc12345")
    await sleep(50)
    const parts = [{ type: "text", text: "무언가 일반적인 질문입니다" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts } as never)
    const chatListCall = fetchCalls.find((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))
    expect(chatListCall).toBeDefined()
    expect(chatListCall!.url).toContain("%24top=10")
  })

  test("activate 이후 text.complete 훅이 캐시된 채팅 목록을 chatId + 이름 + 날짜 blockquote 로 append 한다", async () => {
    writeTokens()
    routeChatList([
      { id: "19:chat-a@thread.v2", topic: "Alice - 안녕", lastActivity: "2026-08-01T10:23:45Z" },
      { id: "19:chat-b@thread.v2", topic: "Bob - 프로젝트", lastActivity: "2026-07-20T05:00:00Z" },
    ])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks, "ses_abc12345")
    await sleep(80)
    const parts = [{ type: "text", text: "일반 프롬프트" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts } as never)
    const output = { text: "assistant 답변 본문" }
    await hooks["experimental.text.complete"]!(
      { sessionID: "ses_abc12345", messageID: "msg_chat_sel", partID: "p1" },
      output,
    )
    expect(output.text).toContain("<!-- microsoft-teams:chat-select -->")
    expect(output.text).toContain("Microsoft Teams 채팅방 선택 필요")
    expect(output.text.split("\n").filter((l) => l.startsWith(">")).length).toBeGreaterThanOrEqual(6)
    expect(output.text).toContain("Alice - 안녕")
    expect(output.text).toContain("Bob - 프로젝트")
    expect(output.text).toContain("2026-08-01 10:23")
    expect(output.text).toContain("2026-07-20 05:00")
    expect(output.text).toContain("`19:chat-a@thread.v2`")
    expect(output.text).toContain("`19:chat-b@thread.v2`")
    expect(output.text).toContain("toast")
  })

  test("activate 후 chat.message → text.complete 순으로 이어질 때 첫 응답에 이미 안내가 append 된다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks, "ses_first")
    await sleep(80)
    const parts = [{ type: "text", text: "첫 사용자 프롬프트" }]
    await hooks["chat.message"]!({ sessionID: "ses_first" } as never, { message: { role: "user" }, parts } as never)
    const output = { text: "assistant 첫 응답" }
    await hooks["experimental.text.complete"]!(
      { sessionID: "ses_first", messageID: "msg_first", partID: "p1" },
      output,
    )
    expect(output.text).toContain("Microsoft Teams 채팅방 선택 필요")
    expect(output.text).toContain("Alice")
  })

  test("bare chatId 만 입력해도 확정된다 (fetch 없이 즉시 처리)", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    const parts = [{ type: "text", text: "19:chat-b@thread.v2" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts } as never)
    ;(fake.client.session as any).messages = async () => ({ data: [asst("응답 갱신됨")] })
    await fireIdle(hooks)
    const send = fetchCalls.find(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(send).toBeDefined()
    expect(send!.url).toContain("19%3Achat-b%40thread.v2")
  })

  test("사용자가 목록 없이 chatId 를 직접 알고 있어서 바로 입력해도 확정된다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    const parts = [{ type: "text", text: "19:direct-chat@thread.v2" }]
    await hooks["chat.message"]!({ sessionID: "ses_direct" } as never, { message: { role: "user" }, parts } as never)
    ;(fake.client.session as any).get = async (o: any) => ({
      data: { id: o.path.id, title: "직접 세션" },
    })
    await fireIdle(hooks, "ses_direct")
    const send = fetchCalls.find(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(send).toBeDefined()
    expect(send!.url).toContain("19%3Adirect-chat%40thread.v2")
  })

  test("chatId 확정 시 toast 알림 (variant=success) 이 발생한다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    const introParts = [{ type: "text", text: "안녕" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts: introParts } as never)
    fake.calls.showToast.length = 0
    const parts = [{ type: "text", text: "19:chat-a@thread.v2" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts } as never)
    const successToasts = fake.calls.showToast.filter((t) => t.body.variant === "success")
    expect(successToasts.length).toBeGreaterThanOrEqual(1)
    expect(JSON.stringify(successToasts)).toContain("Microsoft Teams 채팅방 연동 완료")
    expect(JSON.stringify(successToasts)).toContain("Alice")
  })

  test("chat.message 훅이 chatId 부분을 안내 텍스트로 치환한다 (LLM 이 명령을 그대로 보지 않도록)", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    const intro = [{ type: "text", text: "안녕" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts: intro } as never)
    const parts = [{ type: "text", text: "19:chat-a@thread.v2" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts } as never)
    expect(parts[0].text).not.toContain("19:chat-a@thread.v2")
    expect(parts[0].text).toContain("Alice")
    expect(parts[0].text).toContain("연동되었습니다")
  })

  test("이미 확정된 세션에서는 chat.message 훅이 chatId 를 재파싱하지 않는다 (chatId 잠금)", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_locked", "19:first@thread.v2")
    const parts = [{ type: "text", text: "19:second@thread.v2" }]
    await hooks["chat.message"]!({ sessionID: "ses_locked" } as never, { message: { role: "user" }, parts } as never)
    expect(parts[0].text).toBe("19:second@thread.v2")
    await fireIdle(hooks, "ses_locked")
    const send = fetchCalls.find(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(send!.url).toContain("19%3Afirst%40thread.v2")
    expect(send!.url).not.toContain("second")
  })

  test("chatId 패턴이 없는 일반 프롬프트는 chat.message 훅이 chatId 저장을 하지 않는다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    const parts = [{ type: "text", text: "안녕하세요, 오늘 날씨 어떤가요?" }]
    await hooks["chat.message"]!({ sessionID: "ses_abc12345" } as never, { message: { role: "user" }, parts } as never)
    expect(parts[0].text).toBe("안녕하세요, 오늘 날씨 어떤가요?")
    await fireIdle(hooks)
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(0)
  })

  test("세션마다 chatId 는 독립적으로 관리된다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_1", "19:chat-1@thread.v2")
    await selectChat(hooks, "ses_2", "19:chat-2@thread.v2")
    ;(fake.client.session as any).get = async (o: any) => ({
      data: { id: o.path.id, title: `sess ${o.path.id}` },
    })
    await fireIdle(hooks, "ses_1")
    await fireIdle(hooks, "ses_2")
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(2)
    expect(sends[0].url).toContain("chat-1")
    expect(sends[1].url).toContain("chat-2")
  })

  test("1:1 채팅(`@unq.gbl.spaces` suffix) 도 정상 감지된다", async () => {
    writeTokens()
    routeGraphList([])
    routeGraphSend("2000")
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    const parts = [{ type: "text", text: "19:abc123_xyz@unq.gbl.spaces" }]
    await hooks["chat.message"]!({ sessionID: "ses_1to1" } as never, { message: { role: "user" }, parts } as never)
    ;(fake.client.session as any).get = async (o: any) => ({ data: { id: o.path.id, title: "1:1" } })
    await fireIdle(hooks, "ses_1to1")
    const send = fetchCalls.find(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(send).toBeDefined()
    expect(send!.url).toContain("unq.gbl.spaces")
  })

  test("activate 이후 chat.message 에서 fetch 는 최초 1회만 실행되고 이후는 캐시 재사용 (plugin-wide 캐시)", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    const hooks = await initPlugin(fake)
    await fireConnectTeams(hooks, "ses_a")
    await sleep(80)
    for (const sid of ["ses_a", "ses_b", "ses_c"]) {
      const parts = [{ type: "text", text: "일반 프롬프트" }]
      await hooks["chat.message"]!({ sessionID: sid } as never, { message: { role: "user" }, parts } as never)
    }
    const chatListCalls = fetchCalls.filter((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats"))
    expect(chatListCalls.length).toBe(1)
  })

  test("sub session (parentID 존재) 의 text.complete 는 chat-select 안내를 append 하지 않는다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    ;(fake.client.session as any).get = async (o: any) => ({
      data:
        o.path.id === "ses_child"
          ? { id: "ses_child", title: "sub", parentID: "ses_parent" }
          : { id: o.path.id, title: "parent" },
    })
    const hooks = await initPlugin(fake)
    // chatDirectory 캐시가 로드되도록 parent 세션에서 일반 프롬프트 처리
    await hooks["chat.message"]!(
      { sessionID: "ses_parent" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: "일반 프롬프트" }] } as never,
    )
    const output = { text: "sub 응답 본문" }
    await hooks["experimental.text.complete"]!(
      { sessionID: "ses_child", messageID: "msg_child_1", partID: "p1" },
      output,
    )
    expect(output.text).toBe("sub 응답 본문")
    expect(output.text).not.toContain("<!-- microsoft-teams:chat-select -->")
    expect(output.text).not.toContain("채팅방 선택 필요")
  })

  test("parent 세션이 chat 선택을 완료한 뒤 spawn 된 sub session 의 text.complete 에도 안내가 뜨지 않는다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    ;(fake.client.session as any).get = async (o: any) => ({
      data:
        o.path.id === "ses_child"
          ? { id: "ses_child", title: "sub", parentID: "ses_parent" }
          : { id: o.path.id, title: "parent" },
    })
    const hooks = await initPlugin(fake)
    // 실제 시나리오: parent 가 먼저 일반 프롬프트로 chat directory 캐시를 채운 뒤 chatId 를 확정한다
    await hooks["chat.message"]!(
      { sessionID: "ses_parent" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: "안녕" }] } as never,
    )
    await selectChat(hooks, "ses_parent", "19:chat-a@thread.v2")
    const output = { text: "sub 응답 본문" }
    await hooks["experimental.text.complete"]!(
      { sessionID: "ses_child", messageID: "msg_child_2", partID: "p1" },
      output,
    )
    expect(output.text).toBe("sub 응답 본문")
    expect(output.text).not.toContain("<!-- microsoft-teams:chat-select -->")
    expect(output.text).not.toContain("채팅방 선택 필요")
  })

  test("sub session 의 chat.message 는 me/chats fetch 를 새로 트리거하지 않는다", async () => {
    writeTokens()
    routeChatList([{ id: "19:chat-a@thread.v2", topic: "Alice" }])
    const fake = makeFakeClient({ messages: [asst("응답")] })
    ;(fake.client.session as any).get = async (o: any) => ({
      data:
        o.path.id === "ses_sub_only"
          ? { id: "ses_sub_only", title: "sub", parentID: "ses_parent" }
          : { id: o.path.id, title: "parent" },
    })
    const hooks = await initPlugin(fake)
    // plugin load 시 eager fetch 가 완료되도록 microtask flush
    await sleep(30)
    const before = fetchCalls.filter((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats")).length
    await hooks["chat.message"]!(
      { sessionID: "ses_sub_only" } as never,
      { message: { role: "user" }, parts: [{ type: "text", text: "sub 세션 프롬프트" }] } as never,
    )
    const after = fetchCalls.filter((c) => c.url.includes("graph.microsoft.com/v1.0/me/chats")).length
    expect(after).toBe(before)
    // sub session 에서 사용자가 chatId 를 붙여넣더라도 저장/응답 치환은 하지 않는다
    const parts = [{ type: "text", text: "19:chat-a@thread.v2" }]
    await hooks["chat.message"]!(
      { sessionID: "ses_sub_only" } as never,
      { message: { role: "user" }, parts } as never,
    )
    expect(parts[0].text).toBe("19:chat-a@thread.v2")
  })
})

describe("connect-teams 스킬 자동 설치", () => {
  test("플러그인 로드 시 connect-teams SKILL.md 가 skills 디렉토리에 생성된다", async () => {
    const fake = makeFakeClient()
    await initPlugin(fake)
    const skillPath = path.join(configDir, "skills", "connect-teams", "SKILL.md")
    expect(fs.existsSync(skillPath)).toBe(true)
    const content = fs.readFileSync(skillPath, "utf8")
    expect(content).toContain("name: connect-teams")
    expect(content).toContain("Microsoft Teams 연동을 시작합니다")
  })

  test("이미 동일 버전 스킬이 있으면 파일을 덮어쓰지 않는다", async () => {
    const skillDir = path.join(configDir, "skills", "connect-teams")
    fs.mkdirSync(skillDir, { recursive: true })
    const skillPath = path.join(skillDir, "SKILL.md")
    const sentinel = `<!-- skill-version: 2 -->\n기존_sentinel_내용`
    fs.writeFileSync(skillPath, sentinel)
    const fake = makeFakeClient()
    await initPlugin(fake)
    expect(fs.readFileSync(skillPath, "utf8")).toBe(sentinel)
  })

  test("버전이 다른 스킬이 있으면 최신 내용으로 덮어쓴다", async () => {
    const skillDir = path.join(configDir, "skills", "connect-teams")
    fs.mkdirSync(skillDir, { recursive: true })
    const skillPath = path.join(skillDir, "SKILL.md")
    fs.writeFileSync(skillPath, `<!-- skill-version: 0 -->\n구버전 내용`)
    const fake = makeFakeClient()
    await initPlugin(fake)
    const content = fs.readFileSync(skillPath, "utf8")
    expect(content).toContain("name: connect-teams")
    expect(content).not.toContain("구버전 내용")
  })

  test("스킬 설치 실패해도 플러그인 초기화는 정상적으로 완료된다", async () => {
    const skillsPath = path.join(configDir, "skills")
    fs.mkdirSync(path.dirname(skillsPath), { recursive: true })
    fs.writeFileSync(skillsPath, "not-a-directory")
    const fake = makeFakeClient()
    await expect(initPlugin(fake)).resolves.toBeDefined()
  })
})

// Prefix Keyword Trigger — Teams 답신 주입 직전 결정론적 키워드 게이트 (명세 v1).
// contentType="text" 로 보내 htmlToText 변환 우회 → detectTrigger 에 원문 도달.
async function runTriggerCase(replyText: string): Promise<{
  promptAsyncCalls: number
  injectedPayload: string | null
}> {
  writeTokens()
  routeGraphSend("2000")
  let listCalls = 0
  routes.push((url, init) => {
    if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
    if ((init?.method ?? "GET") !== "GET") return null
    if (url.includes("%24top=1")) return json({ value: [] })
    listCalls++
    if (listCalls <= 1) return json({ value: [] })
    return json({
      value: [
        {
          id: "3000",
          messageType: "message",
          deletedDateTime: null,
          body: { contentType: "text", content: replyText },
        },
      ],
    })
  })
  const fake = makeFakeClient({ messages: [asst("작업 완료")] })
  const hooks = await initPlugin(fake)
  await selectChat(hooks)
  await fireIdle(hooks)
  await sleep(120)
  const promptAsyncCalls = fake.calls.promptAsync.length
  const injectedPayload =
    promptAsyncCalls > 0 ? (fake.calls.promptAsync[0].body.parts[0].text as string) : null
  return { promptAsyncCalls, injectedPayload }
}

/** 주입된 텍스트에서 INJECT_MARKER 를 제거하고 실제 payload 부분만 반환. */
function stripInjectMarker(text: string): string {
  return text.replace(/^<!-- microsoft-teams:injected -->\n/, "")
}

describe("trigger.detection", () => {
  test("#1 기본 ASCII 매칭: '#ask 질문 있어' → payload='질문 있어'", async () => {
    const r = await runTriggerCase("#ask 질문 있어")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("질문 있어")
  })

  test("#2 대소문자 통일: '#Agent 리뷰해줘' → payload='리뷰해줘'", async () => {
    const r = await runTriggerCase("#Agent 리뷰해줘")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("리뷰해줘")
  })

  test("#3 대소문자 통일 (전체 대문자): '#AI 코드 봐줘' → payload='코드 봐줘'", async () => {
    const r = await runTriggerCase("#AI 코드 봐줘")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("코드 봐줘")
  })

  test("#4 NFKC 전각 #: '＃ask 질문' → 탐지", async () => {
    const r = await runTriggerCase("＃ask 질문")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("질문")
  })

  test("#5 선두 공백 trim: '  #opencode 설정' → payload='설정'", async () => {
    const r = await runTriggerCase("  #opencode 설정")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("설정")
  })

  test("#6 한글 + 쉼표 구분자: '막둥아, 이거 봐' → payload='이거 봐'", async () => {
    const r = await runTriggerCase("막둥아, 이거 봐")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("이거 봐")
  })

  test("#7 연속 구분자 제거: '가재야!! 배포?' → payload='배포?'", async () => {
    const r = await runTriggerCase("가재야!! 배포?")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("배포?")
  })

  test("#8 후행 경계 미탐지 (한글 어절 연속): '막둥아니야' → 미탐지", async () => {
    const r = await runTriggerCase("막둥아니야")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#9 후행 경계 미탐지 (한글 어절 연속): '가재야구장 가자' → 미탐지", async () => {
    const r = await runTriggerCase("가재야구장 가자")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#10 후행 경계 미탐지 (ASCII 어절 연속): '#asking about this' → 미탐지", async () => {
    const r = await runTriggerCase("#asking about this")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#11 offset 0 제약: '질문인데 #ask 이거' → 미탐지", async () => {
    const r = await runTriggerCase("질문인데 #ask 이거")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#12 이스케이프: '\\#ask 질문' → 미탐지", async () => {
    const r = await runTriggerCase("\\#ask 질문")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#13 선두 백틱: '`#ask` 는 명령어야' → 미탐지", async () => {
    const r = await runTriggerCase("`#ask` 는 명령어야")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#14 선두 인용문 '>': '> #ask 인용문' → 미탐지", async () => {
    const r = await runTriggerCase("> #ask 인용문")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#15 재스캔 없음: '#ask #agent 질문' → payload='#agent 질문'", async () => {
    const r = await runTriggerCase("#ask #agent 질문")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("#agent 질문")
  })

  test("#16 빈 payload: '#agent' 단독 → 액션 미실행", async () => {
    const r = await runTriggerCase("#agent")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("#17 개행 구분자: '#AI\\n다음 줄 본문' → payload='다음 줄 본문'", async () => {
    const r = await runTriggerCase("#AI\n다음 줄 본문")
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("다음 줄 본문")
  })

  test("#18 쿨다운 없음: 동일 키워드 메시지 5회 연속 모두 탐지 (per-message 누적)", async () => {
    writeTokens()
    routeGraphSend("2000")
    // Teams 답신 배치: [3001, 3002, 3003, 3004, 3005] 모두 '#ask X' 형태
    let listCalls = 0
    routes.push((url, init) => {
      if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
      if ((init?.method ?? "GET") !== "GET") return null
      if (url.includes("%24top=1")) return json({ value: [] })
      listCalls++
      if (listCalls <= 1) return json({ value: [] })
      return json({
        value: [1, 2, 3, 4, 5].map((n) => ({
          id: String(3000 + n),
          messageType: "message",
          deletedDateTime: null,
          body: { contentType: "text", content: `#ask 질문${n}` },
        })),
      })
    })
    const fake = makeFakeClient({ messages: [asst("작업 완료")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(120)
    expect(fake.calls.promptAsync.length).toBe(1)
    const text = fake.calls.promptAsync[0].body.parts[0].text as string
    for (const n of [1, 2, 3, 4, 5]) {
      expect(text).toContain(`질문${n}`)
    }
    // 키워드 원문은 세션에 전달되지 않아야 한다
    expect(text).not.toContain("#ask")
  })

  test("#19 NFKC 조합형 자모: '막둥아'.normalize('NFD') 형태로 입력 → 탐지", async () => {
    const decomposed = "막둥아".normalize("NFD") + ", 이거 봐"
    const r = await runTriggerCase(decomposed)
    expect(r.promptAsyncCalls).toBe(1)
    expect(stripInjectMarker(r.injectedPayload!)).toBe("이거 봐")
  })

  test("#20 코드블록 내부 키워드: '```\\n#opencode 설정\\n```' → 미탐지 (offset 0 이 백틱)", async () => {
    const r = await runTriggerCase("```\n#opencode 설정\n```")
    expect(r.promptAsyncCalls).toBe(0)
  })

  test("혼합 배치: 매칭/미매칭이 섞이면 매칭 것만 payload 로 주입된다", async () => {
    writeTokens()
    routeGraphSend("2000")
    let listCalls = 0
    routes.push((url, init) => {
      if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
      if ((init?.method ?? "GET") !== "GET") return null
      if (url.includes("%24top=1")) return json({ value: [] })
      listCalls++
      if (listCalls <= 1) return json({ value: [] })
      return json({
        value: [
          {
            id: "3001",
            messageType: "message",
            deletedDateTime: null,
            body: { contentType: "text", content: "그냥 잡담 메시지" },
          },
          {
            id: "3002",
            messageType: "message",
            deletedDateTime: null,
            body: { contentType: "text", content: "#ask 실제 질문" },
          },
          {
            id: "3003",
            messageType: "message",
            deletedDateTime: null,
            body: { contentType: "text", content: "이건 미매칭" },
          },
        ],
      })
    })
    const fake = makeFakeClient({ messages: [asst("작업 완료")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(120)
    expect(fake.calls.promptAsync.length).toBe(1)
    const text = fake.calls.promptAsync[0].body.parts[0].text as string
    expect(text).toContain("실제 질문")
    expect(text).not.toContain("잡담")
    expect(text).not.toContain("미매칭")
    expect(text).not.toContain("#ask")
  })

  test("전체 미매칭 배치: promptAsync 호출 없이 커서만 전진 (다음 폴에서 과거 이력 재조회 안 됨)", async () => {
    writeTokens()
    routeGraphSend("2000")
    routes.push((url, init) =>
      url.includes("%24top=1") && (init?.method ?? "GET") === "GET" && url.includes("/chats/")
        ? json({ value: [] })
        : null,
    )
    let listCalls = 0
    routes.push((url, init) => {
      if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
      if ((init?.method ?? "GET") !== "GET") return null
      if (url.includes("%24top=1")) return null
      listCalls++
      if (listCalls === 1) {
        return json({
          value: [
            {
              id: "3002",
              messageType: "message",
              deletedDateTime: null,
              body: { contentType: "text", content: "미매칭 A" },
            },
            {
              id: "3001",
              messageType: "message",
              deletedDateTime: null,
              body: { contentType: "text", content: "미매칭 B" },
            },
          ],
        })
      }
      return json({ value: [] })
    })
    const fake = makeFakeClient({ messages: [asst("작업 완료")] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks)
    await fireIdle(hooks)
    await sleep(120)
    expect(fake.calls.promptAsync.length).toBe(0)
  })
})

describe("Teams 답장 → permission/question 대체 응답 (요구사항 신규)", () => {
  const firePerm = async (
    hooks: any,
    props: {
      id: string
      sessionID: string
      permission?: string
      type?: string
      title?: string
      metadata?: Record<string, unknown>
      pattern?: string | string[]
    },
  ) => {
    await hooks.event({ event: { type: "permission.asked", properties: props } })
  }
  const fireQ = async (
    hooks: any,
    props: {
      id: string
      sessionID: string
      questions: Array<{ question?: string; header?: string; options?: unknown; multiple?: boolean }>
    },
  ) => {
    await hooks.event({ event: { type: "question.asked", properties: props } })
  }

  function extractShortCode(html: string, kind: "perm" | "quest"): string {
    const prefix = kind === "perm" ? "p" : "q"
    const m = new RegExp(`<code>(${prefix}[0-9a-z]{${5}})</code>`).exec(html)
    if (!m) throw new Error(`shortCode not found in alert body:\n${html}`)
    return m[1]
  }

  function routeGraphReplyList(replyText: string, msgId = "9500"): void {
    let called = 0
    routes.push((url, init) => {
      if (!url.includes("graph.microsoft.com") || !url.includes("/chats/")) return null
      if ((init?.method ?? "GET") !== "GET") return null
      if (url.includes("%24top=1")) return json({ value: [] })
      called++
      if (called === 1) {
        return json({
          value: [
            { id: msgId, messageType: "message", deletedDateTime: null, body: { contentType: "text", content: replyText } },
          ],
        })
      }
      return json({ value: [] })
    })
  }

  test("permission 알림 본문에 shortCode 와 답장 사용법이 포함된다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_1", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace }, title: "run" })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    const code = extractShortCode(body, "perm")
    expect(code).toMatch(/^p[0-9a-z]{5}$/)
    expect(body).toContain(`#allow ${code}`)
    expect(body).toContain(`#allow-always ${code}`)
    expect(body).toContain(`#reject ${code}`)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("question 알림 본문에 shortCode 와 label 안내가 포함된다", async () => {
    writeTokens()
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_qq")
    await fireQ(hooks, {
      id: "que_1",
      sessionID: "ses_qq",
      questions: [
        {
          header: "배포 승인",
          question: "production 에 배포할까요?",
          options: [
            { label: "yes", description: "지금 배포" },
            { label: "no", description: "취소" },
          ],
          multiple: false,
        },
      ],
    })
    const sends = fetchCalls.filter(
      (c) => c.init?.method === "POST" && c.url.includes("graph.microsoft.com") && c.url.includes("/chats/"),
    )
    expect(sends.length).toBe(1)
    const body = JSON.parse(String(sends[0].init!.body)).body.content as string
    const code = extractShortCode(body, "quest")
    expect(code).toMatch(/^q[0-9a-z]{5}$/)
    expect(body).toContain(`answer ${code}`)
    expect(body).toContain("yes")
  })

  test("Teams 답장 '#allow <code>' → SDK 로 permission 'once' 응답 호출", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_1", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")

    routeGraphReplyList(`#allow ${code}`, "9500")
    await sleep(200)

    expect(fake.calls.permissionRespond.length).toBe(1)
    const call = fake.calls.permissionRespond[0]
    expect(call.path.id).toBe("ses_a")
    expect(call.path.permissionID).toBe("perm_1")
    expect(call.body.response).toBe("once")
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("Teams 답장 '#reject <code>' → SDK 로 permission 'reject' 응답 호출", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_r", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")
    routeGraphReplyList(`#reject ${code}`, "9501")
    await sleep(200)
    expect(fake.calls.permissionRespond.length).toBe(1)
    expect(fake.calls.permissionRespond[0].body.response).toBe("reject")
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("Teams 답장 '#allow-always <code>' → SDK 로 permission 'always' 응답 호출 + 반영 확인 발신", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_aa", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")
    routeGraphReplyList(`#allow-always ${code}`, "9502")
    await sleep(200)
    expect(fake.calls.permissionRespond.length).toBe(1)
    const call = fake.calls.permissionRespond[0]
    expect(call.path.id).toBe("ses_a")
    expect(call.path.permissionID).toBe("perm_aa")
    expect(call.body.response).toBe("always")
    expect(fake.calls.promptAsync.length).toBe(0)
    const confirmations = fetchCalls
      .filter((c) => c.init?.method === "POST" && c.url.includes("/chats/"))
      .map((c) => JSON.parse(String(c.init!.body)).body.content as string)
      .filter((html) => html.includes("반영됨"))
    expect(confirmations.length).toBe(1)
    expect(confirmations[0]).toContain(`<code>${code}</code> allow-always`)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("Teams 답장 '#ALLOW-ALWAYS <code>' 대문자도 'always' 로 매칭된다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_au", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")
    routeGraphReplyList(`#ALLOW-ALWAYS ${code.toUpperCase()}`, "9504")
    await sleep(200)
    expect(fake.calls.permissionRespond.length).toBe(1)
    expect(fake.calls.permissionRespond[0].body.response).toBe("always")
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("'#allow-always' 파싱 실패 (sub-code / 여분 인자) 는 SDK 호출도 promptAsync 주입도 없다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_ab", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")
    routeGraphReplyList(`#allow-always ${code}-1 extra`, "9505")
    await sleep(200)
    expect(fake.calls.permissionRespond.length).toBe(0)
    expect(fake.calls.promptAsync.length).toBe(0)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("legacy '#opencode allow <code>' 문법은 SDK 호출 없이 promptAsync 로 fallthrough", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    routes.push((url, init) =>
      url.includes("%24top=1") && (init?.method ?? "GET") === "GET" && url.includes("/chats/")
        ? json({ value: [] })
        : null,
    )
    const fake = makeFakeClient({ messages: [asst("작업 완료")] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await fireIdle(hooks, "ses_a")
    await sleep(50)
    await firePerm(hooks, { id: "perm_leg", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/") && String(c.init?.body).includes("권한"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")
    routeGraphReplyList(`#opencode allow ${code}`, "9503")
    await sleep(300)
    expect(fake.calls.permissionRespond.length).toBe(0)
    expect(fake.calls.promptAsync.length).toBeGreaterThanOrEqual(1)
    const legacyCall = fake.calls.promptAsync.find((c: any) => (c.body.parts[0].text as string).includes(`allow ${code}`))
    expect(legacyCall).toBeTruthy()
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("Teams 답장 '#answer <code> yes' → raw fetch 로 question reply 호출 (answers=[['yes']])", async () => {
    writeTokens()
    routeGraphSend("9999")
    routes.push((url, init) => {
      if (init?.method !== "POST") return null
      if (!url.startsWith("http://localhost:1/question/") || !url.endsWith("/reply")) return null
      return json(true)
    })
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_q1")
    await fireQ(hooks, {
      id: "que_x",
      sessionID: "ses_q1",
      questions: [{ header: "H", question: "Q", options: [{ label: "yes", description: "" }, { label: "no", description: "" }] }],
    })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "quest")
    routeGraphReplyList(`#answer ${code} yes`, "9503")
    await sleep(200)
    const replyCall = fetchCalls.find((c) => c.init?.method === "POST" && c.url === `http://localhost:1/question/que_x/reply`)
    expect(replyCall).toBeTruthy()
    const body = JSON.parse(String(replyCall!.init!.body))
    expect(body).toEqual({ answers: [["yes"]] })
  })

  test("multi-question: 각 질문에 sub-code 부여, 모두 답변되면 한번에 assemble 되어 POST", async () => {
    writeTokens()
    routeGraphSend("9999")
    routes.push((url, init) => {
      if (init?.method !== "POST") return null
      if (!url.startsWith("http://localhost:1/question/") || !url.endsWith("/reply")) return null
      return json(true)
    })
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_q2")
    await fireQ(hooks, {
      id: "que_multi",
      sessionID: "ses_q2",
      questions: [
        { header: "A", question: "Q1", options: [{ label: "yes", description: "" }, { label: "no", description: "" }] },
        { header: "B", question: "Q2", options: [{ label: "prod", description: "" }, { label: "stage", description: "" }] },
      ],
    })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const rootCode = extractShortCode(alertBody, "quest")
    expect(alertBody).toContain(`${rootCode}-1`)
    expect(alertBody).toContain(`${rootCode}-2`)

    routeGraphReplyList(`#answer ${rootCode}-1 yes`, "9600")
    await sleep(200)
    expect(fetchCalls.filter((c) => c.url === "http://localhost:1/question/que_multi/reply").length).toBe(0)

    routes.length = 0
    routeGraphSend("9999")
    routes.push((url, init) => {
      if (init?.method !== "POST") return null
      if (!url.startsWith("http://localhost:1/question/") || !url.endsWith("/reply")) return null
      return json(true)
    })
    routeGraphReplyList(`#answer ${rootCode}-2 prod`, "9601")
    await hooks.event!({ event: { type: "permission.asked", properties: { id: "perm_kick", type: "bash", sessionID: "ses_q2", metadata: { cwd: "/tmp" } } } as any })
    await sleep(300)

    const replyCall = fetchCalls.find((c) => c.init?.method === "POST" && c.url === "http://localhost:1/question/que_multi/reply")
    expect(replyCall).toBeTruthy()
    const body = JSON.parse(String(replyCall!.init!.body))
    expect(body).toEqual({ answers: [["yes"], ["prod"]] })
  })

  test("multiple=true 인 단일 질문에 comma 로 여러 label 응답 → answers=[[label1, label2]]", async () => {
    writeTokens()
    routeGraphSend("9999")
    routes.push((url, init) => {
      if (init?.method !== "POST") return null
      if (!url.startsWith("http://localhost:1/question/") || !url.endsWith("/reply")) return null
      return json(true)
    })
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_qmulti")
    await fireQ(hooks, {
      id: "que_mm",
      sessionID: "ses_qmulti",
      questions: [{
        header: "복수", question: "여러 개",
        options: [
          { label: "opt-a", description: "" },
          { label: "opt-b", description: "" },
          { label: "opt-c", description: "" },
        ],
        multiple: true,
      }],
    })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "quest")
    routeGraphReplyList(`#answer ${code} opt-a, opt-c`, "9700")
    await sleep(200)
    const replyCall = fetchCalls.find((c) => c.init?.method === "POST" && c.url === "http://localhost:1/question/que_mm/reply")
    expect(replyCall).toBeTruthy()
    const body = JSON.parse(String(replyCall!.init!.body))
    expect(body).toEqual({ answers: [["opt-a", "opt-c"]] })
  })

  test("Teams 답장 '#skip <code>' → raw fetch 로 question reject 호출", async () => {
    writeTokens()
    routeGraphSend("9999")
    routes.push((url, init) => {
      if (init?.method !== "POST") return null
      if (!url.startsWith("http://localhost:1/question/") || !url.endsWith("/reject")) return null
      return json(true)
    })
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake)
    await selectChat(hooks, "ses_qsk")
    await fireQ(hooks, {
      id: "que_sk",
      sessionID: "ses_qsk",
      questions: [{ header: "H", question: "Q", options: [{ label: "yes", description: "" }] }],
    })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "quest")
    routeGraphReplyList(`#skip ${code}`, "9701")
    await sleep(200)
    const rejectCall = fetchCalls.find((c) => c.init?.method === "POST" && c.url === "http://localhost:1/question/que_sk/reject")
    expect(rejectCall).toBeTruthy()
  })

  test("알 수 없는 code 의 답장은 조용히 무시 (SDK/raw 어떤 호출도 없음)", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_1", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    routeGraphReplyList(`#allow pXXXXX`, "9900")
    await sleep(200)
    expect(fake.calls.permissionRespond.length).toBe(0)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("답장 명령은 promptAsync 세션 주입을 발생시키지 않는다 (트리거 payload 로 흘러들지 않음)", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_np", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")
    routeGraphReplyList(`#allow ${code}`, "9910")
    await sleep(200)
    expect(fake.calls.promptAsync.length).toBe(0)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("permission.replied 이벤트 발화 시 pending 이 정리되어 후속 답장은 무시된다", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_a")
    await firePerm(hooks, { id: "perm_ev", type: "bash", sessionID: "ses_a", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")

    await hooks.event!({ event: { type: "permission.replied", properties: { sessionID: "ses_a", permissionID: "perm_ev", response: "once" } } as any })

    routeGraphReplyList(`#allow ${code}`, "9920")
    await sleep(200)
    expect(fake.calls.permissionRespond.length).toBe(0)
    fs.rmSync(workspace, { recursive: true, force: true })
  })

  test("sub-session 의 permission 답장은 원본 sub-session sid 로 SDK 호출 (root sid 아님)", async () => {
    writeTokens()
    const workspace = fs.mkdtempSync(path.join(os.tmpdir(), "ws-"))
    routeGraphSend("9999")
    const fake = makeFakeClient({ messages: [] })
    ;(fake.client.session as any).get = async (o: any) => ({
      data:
        o.path.id === "ses_child"
          ? { id: "ses_child", title: "sub", parentID: "ses_parent" }
          : { id: o.path.id, title: "parent" },
    })
    const hooks = await initPlugin(fake, workspace)
    await selectChat(hooks, "ses_parent")
    await firePerm(hooks, { id: "perm_sub", type: "bash", sessionID: "ses_child", metadata: { cwd: workspace } })
    const alertBody = JSON.parse(String(fetchCalls.find((c) => c.init?.method === "POST" && c.url.includes("/chats/"))!.init!.body)).body.content as string
    const code = extractShortCode(alertBody, "perm")
    routeGraphReplyList(`#allow ${code}`, "9930")
    await sleep(200)
    expect(fake.calls.permissionRespond.length).toBe(1)
    expect(fake.calls.permissionRespond[0].path.id).toBe("ses_child")
    expect(fake.calls.permissionRespond[0].path.permissionID).toBe("perm_sub")
    fs.rmSync(workspace, { recursive: true, force: true })
  })
})
