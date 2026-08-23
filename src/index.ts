// opencode-ms-teams — OpenCode ↔ Microsoft Teams 양방향 플러그인.
// 설계: docs/2026-08-02-teams-opencode-bridge-design.md
// 구현 계획: docs/2026-08-03-teams-bridge-implementation.md
// 대상 opencode 버전: v1.4.17+ (WSL Rocky Linux 8.10, glibc 2.28)
// 주의: opencode 플러그인 로더는 모든 named export 를 Plugin 함수로 호출한다.
//       배포 파일의 유일한 export 는 MicrosoftTeamsPlugin 하나만 유지할 것.
import type { Plugin } from "@opencode-ai/plugin"
import { createHash } from "node:crypto"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"

// ---------- 상수 ----------

const DATA_DIR_NAME = "opencode-ms-teams"
const CONFIG_FILE_NAME = "opencode-ms-teams.json"
const LOG_FILE_NAME = "opencode-ms-teams.log"
const TOKENS_FILE_NAME = "tokens.json"

const INJECT_MARKER = "<!-- microsoft-teams:injected -->"
// assistant 응답 텍스트에 append 되는 인증 안내의 시작 마커.
// experimental.text.complete 훅이 붙이고, Microsoft Teams 발신 전에는 이 마커 이후를 제거하여 원본만 발신한다.
const AUTH_NOTICE_MARKER = "\n\n<!-- microsoft-teams:auth-notice -->\n"
// 채팅 선택 안내의 시작 마커. auth notice 와 동일한 규칙으로 stripAuthNotice() 가 제거한다.
// (실제로 Teams 발신은 채팅이 선택된 후에만 일어나므로 마커가 새어나갈 일은 거의 없지만
//  방어적으로 동일한 스트립 함수로 처리.)
const CHAT_SELECT_MARKER = "\n\n<!-- microsoft-teams:chat-select -->\n"

// Microsoft Teams chat ID 패턴 (group `19:...@thread.v2` / 1:1 `19:...@unq.gbl.spaces`).
// pending 상태에서 사용자가 chatId 만 붙여넣기로 입력하면 이 패턴으로 감지한다 (prefix 없음).
const CHAT_ID_RE = /(19:[A-Za-z0-9._~%!$&'()*+,;=:@-]+@(?:thread\.v2|unq\.gbl\.spaces))/

// Prefix Keyword Trigger — Teams 답신 주입 게이트 (명세 v1).
// 후행 경계 문자: ASCII 공백류 + 마침표/느낌표/물음표/틸드/콜론/세미콜론/쉼표
// + 유니코드 ellipsis(U+2026) + 전각 문장부호(U+FF0C 쉼표, U+3002 마침표, U+FF01 느낌표, U+FF1F 물음표)
// + 전각 공백(U+3000). 후행 경계는 lookahead 로 검사하되 매칭 후 연속 구분자는 소비한다.
const TRIGGER_DELIMS = " \t\n,.!?~:;\u2026\uFF0C\u3002\uFF01\uFF1F\u3000"
// alternation 은 longest-match 안전을 위해 길이 내림차순으로 정렬.
// 답장 verb keyword (#allow/#reject/#answer/#skip) 는 permission/question 응답 전용이며,
// 매칭 시 pollLoop 이 promptAsync 로 흘리지 않고 executeReplyCommand 로 라우팅한다 (불변식 16).
const TRIGGER_RE =
  /^(?:#(?:opencode|question|answer|reject|agent|allow|skip|ask|ai)|가재야|막둥아)(?=[ \t\n,.!?~:;\u2026\uFF0C\u3002\uFF01\uFF1F\u3000]|$)/iu

const REPLY_VERB_KEYWORDS = new Set(["#allow", "#reject", "#answer", "#skip"])

type TriggerMatch =
  | { matched: false }
  | { matched: true; keyword: string; matchedSpan: string; payload: string }

const REPLY_SHORTCODE_LEN = 5
const REPLY_TTL_MS = 24 * 60 * 60 * 1000
const REPLY_SHORTCODE_MAX_ATTEMPTS = 8

type ReplyCommand =
  | { verb: "allow" | "reject"; code: string }
  | { verb: "answer"; code: string; labels: string[]; text: string | null }
  | { verb: "skip"; code: string }

const REPLY_CODE_RE = /^([pq][0-9a-z]{5})(?:-([1-9][0-9]?))?$/i

// keyword 는 detectTrigger 가 반환한 소문자 형태 (#allow/#reject/#answer/#skip 중 하나).
// verb 는 keyword 로 이미 확정되었으므로 payload 는 code 부터 시작한다.
// text: 는 free-text answer 의 greedy tail terminator (Oracle 조정 #6).
function parseReplyCommand(keyword: string, payload: string): (ReplyCommand & { subIndex: number | null }) | null {
  if (!REPLY_VERB_KEYWORDS.has(keyword)) return null
  const verb = keyword.slice(1) as "allow" | "reject" | "answer" | "skip"
  const trimmed = payload.trim()
  if (!trimmed) return null
  const spaceIdx = trimmed.search(/\s/)
  const codePart = (spaceIdx < 0 ? trimmed : trimmed.slice(0, spaceIdx)).trim()
  const rest = spaceIdx < 0 ? "" : trimmed.slice(spaceIdx + 1).trim()
  const parseCode = (s: string): { code: string; subIndex: number | null } | null => {
    const m = REPLY_CODE_RE.exec(s)
    if (!m) return null
    return { code: m[1].toLowerCase(), subIndex: m[2] ? Number(m[2]) : null }
  }
  const parsed = parseCode(codePart)
  if (!parsed) return null
  if (verb === "allow" || verb === "reject") {
    if (parsed.subIndex != null || rest) return null
    return { verb, code: parsed.code, subIndex: null }
  }
  if (verb === "skip") {
    if (parsed.subIndex != null || rest) return null
    return { verb: "skip", code: parsed.code, subIndex: null }
  }
  if (!rest) return null
  const textMatch = /^text\s*:\s*(.*)$/is.exec(rest)
  if (textMatch) {
    const text = textMatch[1].trim()
    if (!text) return null
    return { verb: "answer", code: parsed.code, labels: [], text, subIndex: parsed.subIndex }
  }
  const labels = rest.split(",").map((s) => s.trim()).filter(Boolean)
  if (!labels.length) return null
  return { verb: "answer", code: parsed.code, labels, text: null, subIndex: parsed.subIndex }
}

function generateShortCode(prefix: "p" | "q", existing: Set<string>): string | null {
  const chars = "0123456789abcdefghijklmnopqrstuvwxyz"
  for (let attempt = 0; attempt < REPLY_SHORTCODE_MAX_ATTEMPTS; attempt++) {
    let code = prefix
    for (let i = 0; i < REPLY_SHORTCODE_LEN; i++) code += chars[Math.floor(Math.random() * chars.length)]
    if (!existing.has(code)) return code
  }
  return null
}

// 정규화: 제어문자 제거 → NFKC → CRLF 통일 → NBSP/전각공백을 공백으로 → trim.
// NFKC 는 전각 '＃'(U+FF03) → '#', 조합형 자모 → 완성형으로 흡수한다.
function normalizeForTrigger(raw: string): string {
  let s = raw.replace(/[\u200B\uFEFF\u200E\u200F]/g, "")
  s = s.normalize("NFKC")
  s = s.replace(/\r\n?/g, "\n")
  s = s.replace(/[\u00A0\u3000]/g, " ")
  return s.trim()
}

// offset 0 에서만 매칭. payload 는 키워드 + 연속 후행 구분자 소비 후의 나머지 (좌측 whitespace trim).
function detectTrigger(raw: string): TriggerMatch {
  const normalized = normalizeForTrigger(raw)
  const m = TRIGGER_RE.exec(normalized)
  if (!m) return { matched: false }
  let end = m[0].length
  while (end < normalized.length && TRIGGER_DELIMS.includes(normalized[end])) end++
  return {
    matched: true,
    keyword: m[0].toLowerCase(),
    matchedSpan: normalized.slice(0, end),
    payload: normalized.slice(end).replace(/^[ \t\n]+/, ""),
  }
}

// Azure 값은 하드코딩하지 않는다 — 사용자마다 tenant/client/scopes 가 다르며,
// 잘못된 default 로 device flow 를 시도하면 Azure 측 에러 메시지가 애매해져 원인 추적이 어렵다.
// opencode-ms-teams.json 최초 생성 시 빈 문자열로 기록하여 사용자에게 명시적 설정을 유도한다.
const DEFAULT_TENANT_ID = ""
const DEFAULT_CLIENT_ID = ""
const DEFAULT_SCOPES = ""
const DEFAULT_POLL_MS = 5000
const DEFAULT_BACKOFF1_MS = 10 * 60 * 1000
const DEFAULT_BACKOFF2_MS = 60 * 60 * 1000
const MAX_HTML_BYTES = 25_000
const CHAT_LIST_TOP = 10

const REFRESH_LEAD_MS = 5 * 60 * 1000
const REFRESH_RETRY_MS = 30 * 1000
const REFRESH_MIN_MS = 15 * 1000

// ---------- connect-teams 스킬 ----------

// 이 숫자를 올리면 설치된 스킬 파일을 최신 내용으로 덮어쓴다.
const CONNECT_TEAMS_SKILL_VERSION = "2"

// 스킬 실행 시 assistant 가 응답 최상단에 그대로 출력해야 하는 HTML 주석 marker.
// 플러그인이 experimental.text.complete 훅에서 이 marker 를 감지해 연동 활성화 (lazy connect) 트리거를 발동한다.
// 사용자가 스킬 없이 임의로 "Microsoft Teams 연동을 시작합니다." 문구만 언급하는 경우와 구분하기 위해 marker 병행.
const CONNECT_TRIGGER_MARKER = "<!-- microsoft-teams:connect-request v1 -->"
// 스킬 없이도 assistant 가 안내 문구를 그대로 출력한 경우를 대비한 secondary 감지 문구 (Oracle 권고 defense-in-depth).
const CONNECT_TRIGGER_PHRASE = "**Microsoft Teams 연동을 시작합니다.**"

const CONNECT_TEAMS_SKILL_CONTENT = `---
name: connect-teams
description: Microsoft Teams 연동 시작 스킬. 추가 프롬프트 없이 이 스킬만 호출하면 된다.
---

<!-- skill-version: ${CONNECT_TEAMS_SKILL_VERSION} -->

# Skill: connect-teams
이 스킬을 실행하면 현재 세션과 Microsoft Teams를 연동한다.
플러그인이 인증 상태와 채팅방 목록을 자동으로 처리하므로 추가 입력 없이 안내를 따르기만 하면 된다.

## 수행 절차
### Step 1 — 연동 시작 응답
아래 **두 줄을 그대로** 응답 최상단에 출력한다. 첫 줄의 HTML 주석은 플러그인이 연동 활성화 트리거로 사용하므로 절대 생략·변형하지 말 것.

첫 줄:
\`\`\`
${CONNECT_TRIGGER_MARKER}
\`\`\`
두 줄:
> **Microsoft Teams 연동을 시작합니다.**

이 두 줄을 출력하면 플러그인이 인증 상태를 확인하고 안내를 자동으로 하단에 첨부한다.

### Step 2 — 첨부 내용에 따른 행동
응답 하단에 자동으로 첨부된 블록을 확인한다.
인증 안내가 첨부된 경우 플러그인이 \`Microsoft Teams 로그인 필요\` 블록을 첨부한다.

사용자에게 안내한다:
1. 첨부된 URL을 브라우저에서 열고 코드를 입력하여 회사 계정으로 로그인
2. 로그인 완료 후 아무 메시지나 보내면 채팅방 목록이 표시됨
로그인 완료 시 다음 응답에 \`✅ Microsoft Teams 로그인 완료\` 안내가 자동 표시된다.
채팅방 목록이 첨부된 경우: 플러그인이 \`Microsoft Teams 채팅방 선택 필요\` 블록과 채팅방 목록을 첨부한다.

사용자에게 안내한다:
- 목록에서 사용할 채팅방을 찾아 \`19:...\` 형식 ID를 **다음 메시지에 그대로 붙여넣는다**
- ID 외의 다른 내용은 입력하지 않아도 된다

#### 아무것도 첨부되지 않은 경우
인증과 채팅방 선택이 이미 완료된 상태다.
사용자에게 알린다: "Microsoft Teams 연동이 이미 활성화되어 있습니다."

### Step 3 — 채팅방 연동 확인
사용자가 채팅방 ID를 붙여넣으면:
- 플러그인이 자동으로 채팅방을 확정하고 toast 알림 "Microsoft Teams 채팅방 연동 완료: {채팅방 이름}"을 표시한다
- 이후 세션 idle 시 응답이 Teams 채팅으로 자동 전송된다
- Teams에서 답장하면 세션에 자동으로 주입된다

toast 알림이 표시되면 연동 성공이다. 별도 확인 요청 없이 사용자에게 완료를 알린다.

## 제약
- 채팅방 ID를 직접 물어보지 않는다. 플러그인이 첨부한 목록에서 선택하도록 안내한다.
- 이 스킬 실행 중 다른 작업을 병행하지 않는다.
- 연동이 이미 완료된 경우, 간단히 알리고 종료한다.
`

function installConnectTeamsSkill(): void {
  try {
    const skillDir = path.join(configDir(), "skills", "connect-teams")
    const skillPath = path.join(skillDir, "SKILL.md")
    try {
      const existing = fs.readFileSync(skillPath, "utf8")
      if (existing.includes(`skill-version: ${CONNECT_TEAMS_SKILL_VERSION}`)) return
    } catch {
      // 파일 없음 → 설치 진행
    }
    fs.mkdirSync(skillDir, { recursive: true })
    fs.writeFileSync(skillPath, CONNECT_TEAMS_SKILL_CONTENT, "utf8")
  } catch {
    // 스킬 설치 실패는 무시
  }
}

// ---------- 설정 로더 ----------

// 채팅 ID 는 config 에 저장하지 않는다. 세션마다 최초 idle 에서 사용자에게 선택받아
// process memory 에만 보관한다 (요구사항: session 별로 Teams 채팅방을 지정하여 연동).
export type Config = {
  azure: {
    tenantId: string
    clientId: string
    scopes: string
  }
  polling: {
    intervalMs: number
    backoff1Ms: number
    backoff2Ms: number
  }
  log: {
    file: string
    enabled: boolean
  }
  tokens: {
    file: string
  }
}

// HOME 이 없고 실행 uid 가 /etc/passwd 에 없으면 os.homedir() 는 libuv ENOENT 로 throw 한다
// (컨테이너·systemd 유닛에서 실제로 발생하는 리눅스 전용 경로). 그대로 두면 플러그인
// 팩토리의 loadConfig() 에서 터져 opencode 의 플러그인 로드가 통째로 실패하고,
// 사용자에게는 "uv_os_homedir returned ENOENT" 라는 원인 불명 메시지만 남는다.
function homeDir(): string {
  try {
    return os.homedir()
  } catch {
    throw new Error(
      "홈 디렉토리를 확인할 수 없습니다 (HOME 미설정 + /etc/passwd 에 uid 없음). " +
        "MICROSOFT_TEAMS_CONFIG_DIR / OPENCODE_CONFIG_DIR / XDG_CONFIG_HOME 중 하나를 설정하세요.",
    )
  }
}

function configDir(): string {
  if (process.env.MICROSOFT_TEAMS_CONFIG_DIR) return process.env.MICROSOFT_TEAMS_CONFIG_DIR
  if (process.env.OPENCODE_CONFIG_DIR) return process.env.OPENCODE_CONFIG_DIR
  const xdg = process.env.XDG_CONFIG_HOME
  if (xdg) return path.join(xdg, "opencode")
  return path.join(homeDir(), ".config", "opencode")
}

function dataDir(): string {
  return path.join(configDir(), DATA_DIR_NAME)
}

// 설정 파일도 데이터 디렉토리 아래에 둔다 (요구사항 3: opencode-ms-teams.json 을
// ~/.config/opencode/opencode-ms-teams/ 로 이동). 로그/토큰과 동일 폴더로 통합.
function configFilePath(): string {
  return path.join(dataDir(), CONFIG_FILE_NAME)
}

function expandPath(p: string): string {
  if (!p) return p
  if (p === "~") return homeDir()
  if (p.startsWith("~/")) return path.join(homeDir(), p.slice(2))
  return p
}

function defaultConfig(): Config {
  const dd = dataDir()
  return {
    azure: {
      tenantId: DEFAULT_TENANT_ID,
      clientId: DEFAULT_CLIENT_ID,
      scopes: DEFAULT_SCOPES,
    },
    polling: {
      intervalMs: DEFAULT_POLL_MS,
      backoff1Ms: DEFAULT_BACKOFF1_MS,
      backoff2Ms: DEFAULT_BACKOFF2_MS,
    },
    log: {
      file: path.join(dd, LOG_FILE_NAME),
      enabled: false,
    },
    tokens: {
      file: path.join(dd, TOKENS_FILE_NAME),
    },
  }
}

function mergeConfig(base: Config, partial: unknown): Config {
  if (!partial || typeof partial !== "object") return base
  const p = partial as Record<string, unknown>
  const pick = <T extends object>(k: keyof Config, def: T): T => {
    const v = p[k as string]
    return v && typeof v === "object" ? ({ ...def, ...(v as object) } as T) : def
  }
  return {
    azure: pick("azure", base.azure),
    polling: pick("polling", base.polling),
    log: pick("log", base.log),
    tokens: pick("tokens", base.tokens),
  }
}

function loadConfig(): Config {
  const base = defaultConfig()
  const cfgPath = configFilePath()
  let loaded: Config
  try {
    const raw = fs.readFileSync(cfgPath, "utf8")
    loaded = mergeConfig(base, JSON.parse(raw))
  } catch (err: unknown) {
    // 파일 없음 → default 로 생성. 파싱 실패 → default 로 계속 진행 (플러그인이 세션을 방해하면 안 됨).
    const code = (err as NodeJS.ErrnoException | undefined)?.code
    if (code === "ENOENT") {
      try {
        fs.mkdirSync(path.dirname(cfgPath), { recursive: true })
        fs.writeFileSync(cfgPath, JSON.stringify(base, null, 2) + "\n", "utf8")
      } catch {
        // 파일 생성 실패해도 default 로 동작
      }
    }
    loaded = base
  }
  loaded.log.file = expandPath(loaded.log.file)
  loaded.tokens.file = expandPath(loaded.tokens.file)
  return loaded
}

// ---------- 로거 ----------

type Log = { info: L; warn: L; error: L; debug: L }
type L = (msg: string, data?: object) => void
type ToastVariant = "info" | "success" | "warning" | "error"
type NotifyOpts = { sid?: string; variant?: ToastVariant; prefix?: string; title?: string }
type Notify = (message: string, opts?: NotifyOpts) => void

const NOOP_LOG: Log = { info: () => {}, warn: () => {}, error: () => {}, debug: () => {} }

function makeLogger(cfg: Config): Log {
  if (!cfg.log.enabled) return NOOP_LOG
  // 플러그인 로드 시점에 로그 파일을 truncate 한 뒤 이후는 append 로 기록 (요구사항 4).
  try {
    fs.mkdirSync(path.dirname(cfg.log.file), { recursive: true })
    fs.writeFileSync(cfg.log.file, "")
  } catch {
    // truncate 실패해도 append 시도는 이어감
  }
  const write = (level: string, msg: string, data?: object) => {
    try {
      fs.appendFileSync(
        cfg.log.file,
        `${new Date().toISOString()} [${level.toUpperCase()}] ${msg}${data ? " " + JSON.stringify(data) : ""}\n`,
      )
    } catch {
      // 로그 실패는 무시 (플러그인이 세션을 방해하면 안 됨)
    }
  }
  return {
    info: (m, d) => write("info", m, d),
    warn: (m, d) => write("warn", m, d),
    error: (m, d) => write("error", m, d),
    debug: (m, d) => write("debug", m, d),
  }
}

// ---------- 유틸 ----------

const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms))
const fingerprint = (s: string) => createHash("sha1").update(s).digest("hex")

function stripAuthNotice(text: string): string {
  let cut = text.length
  for (const marker of [AUTH_NOTICE_MARKER, CHAT_SELECT_MARKER]) {
    const i = text.indexOf(marker)
    if (i >= 0 && i < cut) cut = i
  }
  return text.slice(0, cut).trimEnd()
}

function delayFor(elapsedMs: number, cfg: Config): number {
  const p = cfg.polling
  if (elapsedMs < p.backoff1Ms) return p.intervalMs
  if (elapsedMs < p.backoff2Ms) return p.intervalMs * 3
  return p.intervalMs * 6
}

// ---------- 인증 ----------

type TokenSet = { access_token: string; refresh_token: string; expires_at?: number; expires_in?: number }

class Auth {
  private refreshing: Promise<string | null> | null = null
  private deviceFlowRunning = false
  private lastSid: string | undefined
  private deviceFlowSid: string | undefined
  private pendingDeviceCode: { verificationUri: string; userCode: string; expiresAt: number } | null = null
  // device flow Phase 1 (device code 발급) 완료를 알리는 Promise.
  // startDeviceFlow() 가 생성하고, code 수신 직후 resolve 한다.
  // activateConnect() 가 await 하여 text.complete 훅에서 즉시 auth notice 를 첨부할 수 있게 한다.
  private devicePhase1: Promise<void> | null = null
  private refreshTimer: ReturnType<typeof setTimeout> | null = null
  private disposed = false

  constructor(
    private cfg: Config,
    private log: Log,
    private notify: Notify,
  ) {}

  getAuthNotice(): string | null {
    if (this.pendingDeviceCode && Date.now() < this.pendingDeviceCode.expiresAt) {
      const { verificationUri, userCode } = this.pendingDeviceCode
      return (
        `---\n\n` +
        `> **🔐 Microsoft Teams 로그인 필요**\n` +
        `>\n` +
        `> 1. 브라우저에서 [${verificationUri}](${verificationUri}) 를 여세요\n` +
        `> 2. 코드 \`${userCode}\` 를 입력\n` +
        `> 3. 회사 계정으로 로그인 완료 시 우측 상단에 완료 알림이 표시됩니다`
      )
    }
    return null
  }

  awaitDeviceCode(): Promise<void> {
    return this.devicePhase1 ?? Promise.resolve()
  }

  private tokenUrl() {
    return `https://login.microsoftonline.com/${this.cfg.azure.tenantId}/oauth2/v2.0/token`
  }

  private load(): TokenSet | null {
    try {
      return JSON.parse(fs.readFileSync(this.cfg.tokens.file, "utf8")) as TokenSet
    } catch {
      return null
    }
  }

  private save(t: TokenSet) {
    fs.mkdirSync(path.dirname(this.cfg.tokens.file), { recursive: true })
    fs.writeFileSync(this.cfg.tokens.file, JSON.stringify(t, null, 2), { mode: 0o600 })
    try {
      fs.chmodSync(this.cfg.tokens.file, 0o600)
    } catch {}
  }

  dispose() {
    this.disposed = true
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer)
      this.refreshTimer = null
    }
  }

  private scheduleRefreshFromToken(t: TokenSet | null): void {
    if (this.disposed) return
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer)
      this.refreshTimer = null
    }
    if (!t?.refresh_token) return
    const expiresAt = t.expires_at ?? Date.now() + (t.expires_in ?? 3600) * 1000
    const delay = Math.max(REFRESH_MIN_MS, expiresAt - Date.now() - REFRESH_LEAD_MS)
    this.refreshTimer = setTimeout(() => {
      this.refreshTimer = null
      void this.backgroundRefresh()
    }, delay)
    this.log.debug("token.refresh.scheduled", { inMs: delay })
  }

  bootstrap(): void {
    const saved = this.load()
    if (saved?.refresh_token) {
      this.scheduleRefreshFromToken(saved)
    } else {
      void this.getAccessToken()
    }
  }

  private async backgroundRefresh(): Promise<void> {
    if (this.disposed) return
    const cur = this.load()
    if (!cur?.refresh_token) return
    const next = await this.refresh(cur)
    if (this.disposed) return
    if (next) {
      this.scheduleRefreshFromToken(this.load())
    } else {
      this.refreshTimer = setTimeout(() => {
        this.refreshTimer = null
        void this.backgroundRefresh()
      }, REFRESH_RETRY_MS)
      this.log.debug("token.refresh.retryScheduled", { inMs: REFRESH_RETRY_MS })
    }
  }

  async getAccessToken(sid?: string): Promise<string | null> {
    if (sid) this.lastSid = sid
    const t = this.load()
    if (!t?.refresh_token) {
      this.startDeviceFlow(sid ?? this.lastSid)
      return null
    }
    if (t.expires_at && Date.now() < t.expires_at - 60_000) return t.access_token
    const at = await this.refresh(t)
    if (at) this.scheduleRefreshFromToken(this.load())
    return at
  }

  async forceRefresh(): Promise<string | null> {
    const t = this.load()
    if (!t?.refresh_token) {
      this.startDeviceFlow(this.lastSid)
      return null
    }
    const at = await this.refresh(t)
    if (at) this.scheduleRefreshFromToken(this.load())
    return at
  }

  private refresh(t: TokenSet): Promise<string | null> {
    if (this.refreshing) return this.refreshing
    this.refreshing = (async () => {
      try {
        const res = await fetch(this.tokenUrl(), {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            grant_type: "refresh_token",
            client_id: this.cfg.azure.clientId,
            refresh_token: t.refresh_token,
            scope: this.cfg.azure.scopes,
          }),
        })
        const body = (await res.json()) as TokenSet & { error?: string }
        if (!res.ok) {
          if (body.error === "invalid_grant") {
            this.log.warn("token.revoked — device flow 재시작")
            this.startDeviceFlow(this.lastSid)
          } else {
            this.log.error("token.refresh.failed", { error: body.error })
          }
          return null
        }
        const nt: TokenSet = { ...body, expires_at: Date.now() + (body.expires_in ?? 3600) * 1000 }
        this.save(nt)
        this.log.debug("token.refreshed")
        return nt.access_token
      } catch (e) {
        this.log.error("token.refresh.network", { error: String(e) })
        return null
      } finally {
        this.refreshing = null
      }
    })()
    return this.refreshing
  }

  startDeviceFlow(sid?: string): void {
    if (this.deviceFlowRunning || this.disposed) return
    this.deviceFlowRunning = true
    this.deviceFlowSid = sid ?? this.lastSid

    let resolvePhase1!: () => void
    this.devicePhase1 = new Promise<void>((r) => { resolvePhase1 = r })

    void (async () => {
      try {
        const res = await fetch(
          `https://login.microsoftonline.com/${this.cfg.azure.tenantId}/oauth2/v2.0/devicecode`,
          {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({ client_id: this.cfg.azure.clientId, scope: this.cfg.azure.scopes }),
          },
        )
        const d = (await res.json()) as {
          device_code: string
          user_code: string
          verification_uri: string
          interval?: number
          expires_in?: number
          error?: string
        }
        if (!res.ok) {
          this.log.error("devicecode.request.failed", { error: d.error })
          this.notify(`Microsoft Teams 로그인 준비 실패: ${d.error ?? "unknown"}`, { variant: "error" })
          resolvePhase1()
          return
        }
        const expiresInMs = (d.expires_in ?? 900) * 1000
        this.pendingDeviceCode = {
          verificationUri: d.verification_uri,
          userCode: d.user_code,
          expiresAt: Date.now() + expiresInMs,
        }
        this.log.info("devicecode.obtained", { userCode: d.user_code, verificationUri: d.verification_uri })
        resolvePhase1()
        const interval = (d.interval ?? 5) * 1000
        const deadline = Date.now() + expiresInMs
        while (Date.now() < deadline && !this.disposed) {
          await sleep(interval)
          const tr = await fetch(this.tokenUrl(), {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams({
              grant_type: "urn:ietf:params:oauth:grant-type:device_code",
              client_id: this.cfg.azure.clientId,
              device_code: d.device_code,
            }),
          })
          const tb = (await tr.json()) as TokenSet & { error?: string }
          if (tr.ok) {
            this.save({ ...tb, expires_at: Date.now() + (tb.expires_in ?? 3600) * 1000 })
            this.pendingDeviceCode = null
            this.notify("Microsoft Teams 로그인 완료", { variant: "success" })
            this.scheduleRefreshFromToken(this.load())
            return
          }
          if (tb.error === "authorization_pending" || tb.error === "slow_down") continue
          this.log.error("devicecode.polling.failed", { error: tb.error })
          this.pendingDeviceCode = null
          this.notify(`Microsoft Teams 로그인 실패: ${tb.error}`, { variant: "error" })
          return
        }
        this.log.warn("devicecode.expired — 다음 auth 시도에서 재발급")
        this.pendingDeviceCode = null
      } catch (e) {
        this.log.error("devicecode.network", { error: String(e) })
        this.pendingDeviceCode = null
        resolvePhase1()
      } finally {
        this.deviceFlowRunning = false
        this.deviceFlowSid = undefined
      }
    })()
  }
}

// ---------- 포맷팅 ----------

const CODE_FONT_STYLE = "font-family:Consolas,'Courier New',monospace"
// Teams 채팅은 white-space:pre 만 붙은 <div> 에서 개행이 collapsed 되는 관찰 결과로,
// 표준 <pre><code> 요소를 사용한다. Teams 가 syntax highlight/line number 를 붙이려면
// class="language-*" 가 필요한 것으로 보이며, class 를 생략하면 line number 없이 렌더된다.
const CODE_BLOCK_STYLE =
  `${CODE_FONT_STYLE};background:#f4f4f4;padding:8px;border-left:3px solid #ccc`
const INLINE_CODE_STYLE = `${CODE_FONT_STYLE};background:#f4f4f4;padding:1px 4px;border-radius:3px`

function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
}

// GFM 표 감지: 헤더 라인(파이프 포함)과 구분자 라인(:-+:) 을 만나면 <table> 로 치환.
const TABLE_SEP_RE = /^\s*\|?\s*:?-{3,}:?(\s*\|\s*:?-{3,}:?)+\s*\|?\s*$/
const TABLE_ROW_RE = /^\s*\|.*\|\s*$/
type Align = "left" | "center" | "right"

function splitRow(row: string): string[] {
  return row.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim())
}

function parseAlign(sep: string): Align[] {
  return splitRow(sep).map((c) => {
    const left = c.startsWith(":")
    const right = c.endsWith(":")
    if (left && right) return "center"
    if (right) return "right"
    return "left"
  })
}

function inlineMd(t: string): string {
  return escapeHtml(t)
    .replace(/`([^`]+)`/g, `<span style="${INLINE_CODE_STYLE}">$1</span>`)
    .replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>")
}

function renderTable(header: string, sep: string, rows: string[]): string {
  const aligns = parseAlign(sep)
  const th = splitRow(header)
    .map((c, i) => `<th align="${aligns[i] ?? "left"}">${inlineMd(c)}</th>`)
    .join("")
  const trs = rows
    .map((r) => {
      const cells = splitRow(r)
      return `<tr>${cells
        .map((c, i) => `<td align="${aligns[i] ?? "left"}">${inlineMd(c)}</td>`)
        .join("")}</tr>`
    })
    .join("")
  return `<table border="1" style="border-collapse:collapse"><thead><tr>${th}</tr></thead><tbody>${trs}</tbody></table>`
}

function convertTables(text: string): string {
  const lines = text.split("\n")
  const out: string[] = []
  let i = 0
  while (i < lines.length) {
    if (
      i + 1 < lines.length &&
      TABLE_ROW_RE.test(lines[i]) &&
      TABLE_SEP_RE.test(lines[i + 1])
    ) {
      const header = lines[i]
      const sep = lines[i + 1]
      const rows: string[] = []
      let j = i + 2
      while (j < lines.length && TABLE_ROW_RE.test(lines[j])) {
        rows.push(lines[j])
        j++
      }
      out.push(renderTable(header, sep, rows))
      i = j
      continue
    }
    out.push(lines[i])
    i++
  }
  return out.join("\n")
}

function renderCodeBlock(rawWithLang: string, langHint: string): string {
  void langHint
  return `<pre style="${CODE_BLOCK_STYLE}"><code>${escapeHtml(rawWithLang)}</code></pre>`
}

// 코드펜스는 라인의 시작에서 등장할 때만 인식한다.
// 인라인 백틱(단일 `) 과 혼동을 피하려고 fence 는 반드시 세 개 이상의 백틱 + 라인 시작.
// 예:
//   ```ts        ← open fence, language=ts
//   const a = 1
//   ```          ← close fence
function convertFencesAndTables(md: string): string {
  const lines = md.split("\n")
  const out: string[] = []
  let i = 0
  const fenceRe = /^(\s*)(`{3,})\s*([A-Za-z0-9_+.-]*)\s*$/
  while (i < lines.length) {
    const m = lines[i].match(fenceRe)
    if (m) {
      const fence = m[2]
      const lang = m[3]
      const inner: string[] = []
      let j = i + 1
      while (j < lines.length) {
        const closeM = lines[j].match(/^(\s*)(`{3,})\s*$/)
        if (closeM && closeM[2].length >= fence.length) break
        inner.push(lines[j])
        j++
      }
      // close fence 를 못 찾은 채 파일이 끝나는 경우도 코드 블록으로 처리 (사용자 관점 자연스러움).
      out.push(renderCodeBlock(inner.join("\n"), lang))
      i = j < lines.length ? j + 1 : j
      continue
    }
    out.push(lines[i])
    i++
  }
  // 코드펜스 밖 영역에 대해 표 감지 → 표 아닌 영역은 inlineMd + <br>
  const joined = out.join("\n")
  const segs = joined.split(/(<pre style="[^"]*"><code>[\s\S]*?<\/code><\/pre>)/)
  return segs
    .map((seg) => {
      if (seg.startsWith("<pre")) return seg
      const withTables = convertTables(seg)
      const tableSegs = withTables.split(/(<table\b[\s\S]*?<\/table>)/)
      return tableSegs
        .map((s) => (s.startsWith("<table") ? s : inlineMd(s).replace(/\n/g, "<br>")))
        .join("")
    })
    .join("")
}

function mdToHtml(md: string): string {
  return convertFencesAndTables(md)
}

function renderOutbound(title: string, sid: string, text: string, maxBytes: number): string {
  const shortId = sid.replace(/^ses_/, "").slice(0, 8)
  const header = `<b>[${escapeHtml(title)} · ${escapeHtml(shortId)}]</b><br>`
  let body = text
  let html = header + mdToHtml(body)
  while (Buffer.byteLength(html, "utf8") > maxBytes && body.length > 100) {
    body = body.slice(0, Math.floor(body.length * 0.8))
    html = header + mdToHtml(body) + "<br>…(생략)"
  }
  return html
}

function htmlToText(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&") // 반드시 마지막 (이중 디코드 방지)
    .trim()
}

// ---------- Graph ----------

type ChatMsg = {
  id: string
  messageType: string
  deletedDateTime?: string | null
  body: { contentType: string; content: string }
}

export type ChatSummary = {
  id: string
  topic: string
  chatType: string
  lastActivity: string | null
}

class Graph {
  constructor(
    private auth: Auth,
    private log: Log,
  ) {}

  private chatBase(chatId: string) {
    return `https://graph.microsoft.com/v1.0/chats/${encodeURIComponent(chatId)}`
  }

  private async call(method: "GET" | "POST", url: string, body?: object, retried = false): Promise<any | null> {
    const token = await this.auth.getAccessToken()
    if (!token) return null
    let res: Response
    try {
      res = await fetch(url, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          ...(body ? { "Content-Type": "application/json" } : {}),
        },
        body: body ? JSON.stringify(body) : undefined,
      })
    } catch (e) {
      const msg = String(e)
      this.log.warn("graph.network", { url, error: msg })
      if (/certificat|UNABLE_TO_VERIFY|self.signed/i.test(msg)) {
        this.log.error(
          "graph.ssl.hint",
          { hint: "사내망 프록시 인증서 문제로 보임. NODE_EXTRA_CA_CERTS=<사내 CA 번들 경로> 설정 후 OpenCode 재시작" },
        )
      }
      return null
    }
    if (res.status === 401 && !retried) {
      this.log.debug("graph.401 — forceRefresh 후 재시도")
      const t2 = await this.auth.forceRefresh()
      if (!t2) return null
      return this.call(method, url, body, true)
    }
    if (res.status === 429) {
      const ra = Number(res.headers.get("Retry-After") ?? "10")
      this.log.warn("graph.throttled", { retryAfterSec: ra })
      if (retried) return null
      await sleep(ra * 1000)
      return this.call(method, url, body, true)
    }
    if (!res.ok) {
      this.log.error("graph.error", { url, status: res.status })
      return null
    }
    if (res.status === 204) return {}
    return res.json()
  }

  async sendMessage(chatId: string, html: string): Promise<string | null> {
    const r = await this.call("POST", `${this.chatBase(chatId)}/messages`, {
      body: { contentType: "html", content: html },
    })
    return (r?.id as string | undefined) ?? null
  }

  // 반환 규약:
  //   number  — 성공. 채팅방이 비어 있으면 0 (messagesSince 도 빈 결과라 안전).
  //   null    — Graph 호출 실패 (401 재시도 실패, 429, 5xx, 네트워크 등).
  // 실패와 "빈 채팅방" 을 동일하게 0 으로 축약하면, 활성 채팅방에서 seed 실패 시
  // cursor=0 으로 오염되어 messagesSince 가 최근 50건을 전부 통과시킨다 (과거 이력 폭발).
  // 호출자는 null 을 반드시 별도로 처리해야 한다.
  async latestMessageId(chatId: string): Promise<number | null> {
    const r = await this.call("GET", `${this.chatBase(chatId)}/messages?%24top=1&%24orderby=createdDateTime%20desc`)
    if (r == null) return null
    const id = r?.value?.[0]?.id as string | undefined
    return id ? Number(id) : 0
  }

  async messagesSince(chatId: string, cursor: number): Promise<ChatMsg[]> {
    const r = await this.call("GET", `${this.chatBase(chatId)}/messages?%24top=50&%24orderby=createdDateTime%20desc`)
    const value = (r?.value as ChatMsg[] | undefined) ?? []
    return value
      .filter(
        (m) =>
          m.messageType === "message" &&
          Number(m.id) > cursor &&
          !m.deletedDateTime &&
          Boolean(m.body?.content),
      )
      .sort((a, b) => Number(a.id) - Number(b.id))
  }

  async listChats(top: number): Promise<ChatSummary[] | null> {
    const url =
      `https://graph.microsoft.com/v1.0/me/chats?` +
      `%24top=${encodeURIComponent(String(top))}` +
      `&%24orderby=${encodeURIComponent("lastMessagePreview/createdDateTime desc")}` +
      `&%24expand=${encodeURIComponent("lastMessagePreview")}`
    const r = await this.call("GET", url)
    if (!r) return null
    const rows = (r.value as Array<Record<string, unknown>> | undefined) ?? []
    return rows.map((row) => ({
      id: String(row.id ?? ""),
      topic: typeof row.topic === "string" && row.topic ? row.topic : "",
      chatType: String(row.chatType ?? "unknown"),
      lastActivity:
        typeof row.lastUpdatedDateTime === "string"
          ? row.lastUpdatedDateTime
          : ((row.lastMessagePreview as { createdDateTime?: string } | undefined)?.createdDateTime ?? null),
    }))
  }
}

// ---------- Permission Check (요구사항 8) ----------

// workspace 안 경로인지 판정.
// symlink/`..`/절대경로를 정규화한 뒤 workspace prefix 로 시작하는지 확인.
function isInsideWorkspace(target: string, workspace: string): boolean {
  if (!target || !workspace) return false
  try {
    const abs = path.isAbsolute(target) ? path.resolve(target) : path.resolve(workspace, target)
    const ws = path.resolve(workspace)
    const rel = path.relative(ws, abs)
    return rel === "" || (!rel.startsWith("..") && !path.isAbsolute(rel))
  } catch {
    return false
  }
}

// Permission input 에서 대상 경로 후보를 추출.
// opencode Permission 스키마: { type, pattern?, metadata } — 파일/쉘 접근 계열 hook 은
// pattern 또는 metadata.path / metadata.cwd 등에 경로가 들어온다.
function extractPathsFromPermission(input: {
  type?: string
  pattern?: string | string[]
  metadata?: Record<string, unknown>
}): string[] {
  const paths: string[] = []
  const pat = input.pattern
  if (typeof pat === "string") paths.push(pat)
  else if (Array.isArray(pat)) paths.push(...pat)
  const md = input.metadata ?? {}
  for (const k of ["path", "filePath", "file", "cwd", "target", "dir"] as const) {
    const v = md[k]
    if (typeof v === "string") paths.push(v)
  }
  return paths.filter(Boolean)
}

// input.serverUrl (URL) 우선. 실측: v1 SDK 의 client 인스턴스는 top-level 에 getConfig 를 노출하지 않고
// 내부 `_client.getConfig().baseUrl` 로만 접근 가능하므로 두 번째 폴백 경로를 이 위치로 조정한다.
// 두 소스 모두 없으면 null → question reply/reject 는 raw fetch 불가로 조용히 skip.
function resolveServerBase(input: unknown): string | null {
  const anyInput = input as {
    serverUrl?: URL | { origin?: string } | string
    client?: { _client?: { getConfig?: () => { baseUrl?: string } } }
  }
  const src = anyInput.serverUrl
  if (src) {
    if (typeof src === "string") return src.replace(/\/+$/, "")
    const origin = (src as { origin?: string }).origin
    if (typeof origin === "string" && origin) return origin.replace(/\/+$/, "")
  }
  try {
    const b = anyInput.client?._client?.getConfig?.().baseUrl
    if (typeof b === "string" && b) return b.replace(/\/+$/, "")
  } catch {
    // getConfig 접근 실패는 무시하고 null 반환
  }
  return null
}

async function callQuestionReply(baseUrl: string, requestID: string, answers: string[][]): Promise<boolean> {
  try {
    const r = await fetch(`${baseUrl}/question/${encodeURIComponent(requestID)}/reply`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ answers }),
    })
    return r.ok
  } catch {
    return false
  }
}

async function callQuestionReject(baseUrl: string, requestID: string): Promise<boolean> {
  try {
    const r = await fetch(`${baseUrl}/question/${encodeURIComponent(requestID)}/reject`, { method: "POST" })
    return r.ok
  } catch {
    return false
  }
}

// ---------- 플러그인 ----------

export const MicrosoftTeamsPlugin: Plugin = async (input) => {
  const { client } = input
  const workspace: string =
    (input as unknown as { directory?: string; worktree?: string }).directory ||
    (input as unknown as { worktree?: string }).worktree ||
    process.cwd()
  const serverBase = resolveServerBase(input)

  const cfg = loadConfig()
  const log = makeLogger(cfg)
  const notify: Notify = (message, opts = {}) => {
    const variant = opts.variant ?? "info"
    const title = opts.title ?? "Microsoft Teams"
    log.info("notify", { title, message, variant })
    void client.tui
      .showToast({
        body: { title, message, variant, duration: 6_000 },
      })
      .catch(() => {})
  }
  const auth = new Auth(cfg, log, notify)
  const graph = new Graph(auth, log)

  type WaitingState = {
    since: number
    model?: { providerID: string; modelID: string }
    agent?: string
  }
  const waiting = new Map<string, WaitingState>()
  const cursors = new Map<string, number>()
  const lastFp = new Map<string, string>()
  const sentIds = new Set<string>()
  const pollGen = new Map<string, number>()
  const noticedPermissions = new Set<string>()
  const noticedQuestions = new Set<string>()

  type PendingPermission = {
    sid: string
    permissionID: string
    rootSid: string
    createdAt: number
    outside: string[]
    kind: string
  }
  type PendingSubQuestion = { labels: string[] | null; text: string | null }
  type PendingQuestion = {
    sid: string
    requestID: string
    rootSid: string
    createdAt: number
    questions: QuestionAskedProps["questions"]
    subCodes: string[]
    subAnswers: Array<PendingSubQuestion | null>
  }
  const pendingPermissions = new Map<string, PendingPermission>()
  const pendingQuestions = new Map<string, PendingQuestion>()
  // 답장을 받아야 하는 root sid 집합. pollLoop 이 waiting.has(sid) 만으로 종료되지 않도록 게이트 (Oracle 블로커 #1).
  const alertPending = new Map<string, Set<string>>()

  function registerAlertPending(rootSid: string, code: string): void {
    let set = alertPending.get(rootSid)
    if (!set) {
      set = new Set()
      alertPending.set(rootSid, set)
    }
    set.add(code)
  }

  function releaseAlertPending(rootSid: string, code: string): void {
    const set = alertPending.get(rootSid)
    if (!set) return
    set.delete(code)
    if (!set.size) alertPending.delete(rootSid)
  }

  function existingShortCodes(): Set<string> {
    const all = new Set<string>()
    for (const c of pendingPermissions.keys()) all.add(c)
    for (const c of pendingQuestions.keys()) all.add(c)
    return all
  }

  function sweepExpiredPending(): void {
    const now = Date.now()
    for (const [code, p] of pendingPermissions) {
      if (now - p.createdAt > REPLY_TTL_MS) {
        pendingPermissions.delete(code)
        releaseAlertPending(p.rootSid, code)
        log.info("pending.perm.expired", { code, sid: p.sid, permissionID: p.permissionID })
      }
    }
    for (const [code, q] of pendingQuestions) {
      if (now - q.createdAt > REPLY_TTL_MS) {
        pendingQuestions.delete(code)
        releaseAlertPending(q.rootSid, code)
        log.info("pending.quest.expired", { code, sid: q.sid, requestID: q.requestID })
      }
    }
  }

  function ensurePollLoop(rootSid: string): void {
    if (waiting.has(rootSid)) return
    const gen = (pollGen.get(rootSid) ?? 0) + 1
    pollGen.set(rootSid, gen)
    void pollLoop(rootSid, gen)
  }

  let connectRequested = false
  let connectActivating: Promise<void> | null = null

  function activateConnect(via: string, sid?: string): void {
    if (connectRequested) return
    connectRequested = true
    log.info("connect.requested", { via, sessionID: sid })
    connectActivating = (async () => {
      auth.bootstrap()
      await Promise.race([auth.awaitDeviceCode(), sleep(5_000)])
      await Promise.race([ensureChatDirectory(), sleep(2_000)])
    })()
  }

  // 세션별 확정 chatId 는 sessionChat 에, plugin-wide 채팅 목록 캐시는 chatDirectory 에.
  // 채팅 목록은 사용자 계정 단위(/me/chats)이므로 세션마다 재조회하지 않고 최초 1회만 fetch.
  // pendingNoticed 는 같은 messageID 에 안내 중복 append 방지용.
  // sessionParent 는 sessionID → parentID (또는 root 의 경우 null) 를 캐시한다.
  // sub session 은 onIdle 에서 skip 하므로 chat 선택 안내도 append 하지 않아야 한다 (요구사항).
  const sessionChat = new Map<string, string>()
  const pendingNoticed = new Set<string>()
  const sessionParent = new Map<string, string | null>()
  // sessionRoot 은 sub-session sid → root sid 를 캐싱한다. blocking user-input 이벤트
  // (permission.asked / question.asked) 를 root 의 chatId 로 bubble-up 시킬 때 사용.
  // 트리 구조상 root 는 immutable 이므로 무효화 필요 없음.
  const sessionRoot = new Map<string, string>()
  const chatDirectory: { chats: ChatSummary[] | null; fetching: Promise<ChatSummary[] | null> | null } = {
    chats: null,
    fetching: null,
  }

  async function isSubSession(sid: string): Promise<boolean> {
    if (sessionParent.has(sid)) return sessionParent.get(sid) != null
    try {
      const res = await client.session.get({ path: { id: sid } })
      const pid = (res.data as { parentID?: string } | undefined)?.parentID ?? null
      sessionParent.set(sid, pid)
      return pid != null
    } catch {
      // transient 실패 시 캐시에 null 을 남기면 이후 sub-session 이 영구적으로 root 로 오분류되어
      // resolveRoot 가 잘못된 sid 로 chatId 를 찾다 silent skip 되는 회귀가 발생한다.
      return false
    }
  }

  async function resolveRoot(sid: string): Promise<string> {
    const cached = sessionRoot.get(sid)
    if (cached) return cached
    const visited = new Set<string>()
    let current = sid
    while (!visited.has(current)) {
      visited.add(current)
      let parent = sessionParent.get(current)
      if (parent === undefined) {
        try {
          const res = await client.session.get({ path: { id: current } })
          parent = (res.data as { parentID?: string } | undefined)?.parentID ?? null
          sessionParent.set(current, parent)
        } catch {
          // 실패 시 캐시하지 않음 (isSubSession 과 동일 원칙). current 를 root 로 간주하고 종료.
          log.warn("resolveRoot.getFailed", { sid, at: current })
          return current
        }
      }
      if (!parent) {
        for (const v of visited) sessionRoot.set(v, current)
        return current
      }
      current = parent
    }
    log.warn("resolveRoot.cycle", { sid, visited: [...visited] })
    sessionRoot.set(sid, sid)
    return sid
  }

  async function ensureChatDirectory(): Promise<ChatSummary[] | null> {
    if (chatDirectory.chats) return chatDirectory.chats
    if (chatDirectory.fetching) return chatDirectory.fetching
    chatDirectory.fetching = (async () => {
      const res = await graph.listChats(CHAT_LIST_TOP)
      if (res) {
        chatDirectory.chats = res
        log.info("chatlist.fetched", { count: res.length })
        notify("Microsoft Teams 채팅방 목록 중 선택 필요", {
          title: "Microsoft Teams 연동 준비 중",
          variant: "info",
        })
      } else {
        log.warn("chatlist.failed")
      }
      chatDirectory.fetching = null
      return res
    })()
    return chatDirectory.fetching
  }

  function formatChatDate(iso: string | null): string {
    if (!iso) return "-"
    const d = new Date(iso)
    if (Number.isNaN(d.getTime())) return "-"
    return d.toISOString().replace("T", " ").slice(0, 16)
  }

  function renderChatSelectNotice(chats: ChatSummary[]): string {
    const rows = chats
      .map((c, i) => {
        const name = c.topic || `(${c.chatType})`
        const when = formatChatDate(c.lastActivity)
        return `> ${i + 1}. **${name}** · ${when}\n>    \`${c.id}\``
      })
      .join("\n>\n")
    return (
      `---\n\n` +
      `> **🔗 Microsoft Teams 채팅방 선택 필요**\n` +
      `>\n` +
      `> 이 세션에서 사용할 채팅방의 \`chatId\` 를 다음 프롬프트에 입력해 주세요.\n` +
      `> (아래 목록에서 원하는 채팅의 \`19:...\` 문자열을 그대로 복사·붙여넣기)\n` +
      `>\n` +
      `${rows}\n` +
      `>\n` +
      `> 선택 완료 시 우측 상단에 toast 알림이 표시됩니다.`
    )
  }

  async function pollLoop(sid: string, gen: number): Promise<void> {
    // waiting (idle 응답 답신 대기) 이 없어도 alertPending (permission/question 답장 대기) 이 있으면 계속 폴링한다.
    // 이 조건이 false → true 로 바뀌는 이유를 이해하는 것이 이 함수의 lifecycle 을 안전하게 유지하는 열쇠 (Oracle 블로커 #1).
    const shouldPoll = () => (waiting.has(sid) || (alertPending.get(sid)?.size ?? 0) > 0) && pollGen.get(sid) === gen
    while (shouldPoll()) {
      const st = waiting.get(sid)
      const since = st?.since ?? Date.now()
      await sleep(delayFor(Date.now() - since, cfg))
      if (!shouldPoll()) return
      const chatId = sessionChat.get(sid)
      if (!chatId) return
      sweepExpiredPending()
      let msgs: ChatMsg[]
      try {
        msgs = await graph.messagesSince(chatId, cursors.get(sid) ?? 0)
      } catch (e) {
        log.warn("poll.error", { sid, error: String(e) })
        continue
      }
      if (!msgs.length) continue
      const maxId = Math.max(...msgs.map((m) => Number(m.id)))
      const replies = msgs.filter((m) => !sentIds.has(m.id))
      cursors.set(sid, Math.max(cursors.get(sid) ?? 0, maxId))
      if (!replies.length) continue
      const payloads: string[] = []
      for (const m of replies) {
        const raw = m.body.contentType === "html" ? htmlToText(m.body.content) : m.body.content.trim()
        if (!raw) continue
        const trig = detectTrigger(raw)
        if (!trig.matched) {
          log.info("trigger.skip.nomatch", { sid, mid: m.id, len: raw.length })
          continue
        }
        if (!trig.payload) {
          log.info("trigger.skip.emptyPayload", { sid, mid: m.id, keyword: trig.keyword })
          continue
        }
        // Reply verb keyword (#allow/#reject/#answer/#skip) 는 executeReplyCommand 로만 라우팅한다.
        // 파싱 실패 시에도 promptAsync 로 fallthrough 하지 않는다 ("allow p2a7k9" 가 세션에 주입되는 회귀 방지).
        if (REPLY_VERB_KEYWORDS.has(trig.keyword)) {
          const cmd = parseReplyCommand(trig.keyword, trig.payload)
          if (cmd) await executeReplyCommand(cmd, sid, chatId)
          else log.info("reply.parse.failed", { sid, mid: m.id, keyword: trig.keyword, payloadLen: trig.payload.length })
          continue
        }
        log.info("trigger.match", {
          sid,
          mid: m.id,
          keyword: trig.keyword,
          matchedSpan: trig.matchedSpan,
          payloadLen: trig.payload.length,
        })
        payloads.push(trig.payload)
      }
      if (!payloads.length || !st) continue
      const text = payloads.join("\n---\n")
      waiting.delete(sid)
      const res = await client.session.promptAsync({
        path: { id: sid },
        body: {
          parts: [{ type: "text", text: `${INJECT_MARKER}\n${text}` }],
          ...(st.model ? { model: st.model } : {}),
          ...(st.agent ? { agent: st.agent } : {}),
        },
      })
      if (res.error) {
        log.error("inject.failed", { sid, error: JSON.stringify(res.error) })
        waiting.set(sid, st)
        continue
      }
      log.info("injected", { sid, count: replies.length })
      notify("Microsoft Teams 답신을 세션에 주입했습니다")
      if (!shouldPoll()) return
    }
  }

  async function sendReplyConfirmation(rootSid: string, chatId: string, html: string): Promise<void> {
    const mid = await graph.sendMessage(chatId, html)
    if (mid) sentIds.add(mid)
    void rootSid
  }

  async function executeReplyCommand(
    cmd: ReplyCommand & { subIndex: number | null },
    rootSid: string,
    chatId: string,
  ): Promise<void> {
    if (cmd.verb === "allow" || cmd.verb === "reject") {
      const p = pendingPermissions.get(cmd.code)
      if (!p) {
        log.info("reply.perm.unknown", { code: cmd.code, verb: cmd.verb })
        return
      }
      if (p.rootSid !== rootSid) {
        log.warn("reply.perm.wrongChat", { code: cmd.code, expected: p.rootSid, actual: rootSid })
        return
      }
      const response = cmd.verb === "allow" ? "once" : "reject"
      let ok = false
      try {
        const res = await (client as unknown as {
          postSessionIdPermissionsPermissionId: (o: unknown) => Promise<{ data?: unknown; error?: unknown }>
        }).postSessionIdPermissionsPermissionId({
          path: { id: p.sid, permissionID: p.permissionID },
          body: { response },
        })
        ok = !res.error
      } catch (e) {
        log.error("reply.perm.crash", { code: cmd.code, error: String(e) })
      }
      pendingPermissions.delete(cmd.code)
      releaseAlertPending(p.rootSid, cmd.code)
      const outsideNote = p.outside.length > 0 ? " · <b>⚠️ workspace 외부 경로 승인됨</b>" : ""
      const html = ok
        ? `<b>✅ 반영됨</b> · <code>${escapeHtml(cmd.code)}</code> ${escapeHtml(cmd.verb)}${outsideNote}`
        : `<b>❌ 적용 실패</b> · <code>${escapeHtml(cmd.code)}</code> ${escapeHtml(cmd.verb)} (opencode 세션이 이미 응답했거나 만료됨)`
      await sendReplyConfirmation(rootSid, chatId, html)
      log.info("reply.perm.done", { code: cmd.code, response, ok, outside: p.outside })
      return
    }
    if (cmd.verb === "skip") {
      const q = pendingQuestions.get(cmd.code)
      if (!q) {
        log.info("reply.quest.unknown", { code: cmd.code, verb: "skip" })
        return
      }
      if (q.rootSid !== rootSid) {
        log.warn("reply.quest.wrongChat", { code: cmd.code })
        return
      }
      let ok = false
      if (serverBase) ok = await callQuestionReject(serverBase, q.requestID)
      else log.warn("reply.quest.noBaseUrl", { code: cmd.code })
      pendingQuestions.delete(cmd.code)
      releaseAlertPending(q.rootSid, cmd.code)
      const html = ok
        ? `<b>✅ 반영됨</b> · <code>${escapeHtml(cmd.code)}</code> skip`
        : `<b>❌ 적용 실패</b> · <code>${escapeHtml(cmd.code)}</code> skip`
      await sendReplyConfirmation(rootSid, chatId, html)
      log.info("reply.quest.reject.done", { code: cmd.code, ok })
      return
    }
    if (cmd.verb !== "answer") return
    // parseReplyCommand 는 sub-code (예: q9rbvg-2) 를 code=q9rbvg + subIndex=2 로 분해해 반환한다.
    // 따라서 pendingQuestions.get(cmd.code) 은 항상 root 를 조회한다.
    const q = pendingQuestions.get(cmd.code)
    if (!q) {
      log.info("reply.quest.answer.unknown", { code: cmd.code, subIndex: cmd.subIndex })
      return
    }
    if (q.rootSid !== rootSid) {
      log.warn("reply.quest.answer.wrongChat", { code: cmd.code })
      return
    }
    const isSub = cmd.subIndex != null
    if (!isSub && q.questions.length !== 1) {
      log.info("reply.quest.answer.needSubCode", { code: cmd.code, questions: q.questions.length })
      return
    }
    const targetIdx = isSub ? cmd.subIndex! - 1 : 0
    if (targetIdx < 0 || targetIdx >= q.questions.length) {
      log.warn("reply.quest.answer.badIndex", { code: cmd.code, targetIdx })
      return
    }
    q.subAnswers[targetIdx] = cmd.text != null ? { labels: null, text: cmd.text } : { labels: cmd.labels, text: null }
    log.info("reply.quest.answer.recorded", { rootCode: cmd.code, subIndex: targetIdx + 1, hasText: cmd.text != null })
    if (q.subAnswers.some((a) => a == null)) {
      const remaining = q.subAnswers.filter((a) => a == null).length
      const displayed = isSub ? `${cmd.code}-${cmd.subIndex}` : cmd.code
      const html = `<b>📝 답변 기록됨</b> · <code>${escapeHtml(displayed)}</code> (남은 질문 ${remaining}개)`
      await sendReplyConfirmation(rootSid, chatId, html)
      return
    }
    // 모든 sub-answers 채워짐 → assemble & POST once. custom text 는 label 위치에 그대로 삽입 (SDK QuestionAnswer = Array<string>).
    const answers: string[][] = q.subAnswers.map((a) => (a!.text != null ? [a!.text] : a!.labels ?? []))
    let ok = false
    if (serverBase) ok = await callQuestionReply(serverBase, q.requestID, answers)
    else log.warn("reply.quest.answer.noBaseUrl", { rootCode: cmd.code })
    pendingQuestions.delete(cmd.code)
    releaseAlertPending(q.rootSid, cmd.code)
    const html = ok
      ? `<b>✅ 반영됨</b> · <code>${escapeHtml(cmd.code)}</code> answer`
      : `<b>❌ 적용 실패</b> · <code>${escapeHtml(cmd.code)}</code> answer (opencode 세션이 이미 응답했거나 만료됨)`
    await sendReplyConfirmation(rootSid, chatId, html)
    log.info("reply.quest.answer.done", { rootCode: cmd.code, ok })
  }

  type SessionContext = {
    text: string
    model?: { providerID: string; modelID: string }
    agent?: string
  }

  async function lastAssistantContext(sid: string): Promise<SessionContext | null> {
    const res = await client.session.messages({ path: { id: sid } })
    const items = res.data
    if (!items) return null
    let assistantIdx = -1
    let text = ""
    for (let i = items.length - 1; i >= 0; i--) {
      const { info, parts } = items[i]
      if (info.role !== "assistant") continue
      const raw = parts
        .filter((p): p is Extract<typeof p, { type: "text" }> => p.type === "text" && !(p as { synthetic?: boolean }).synthetic)
        .map((p) => p.text)
        .join("\n")
        .trim()
      const clean = stripAuthNotice(raw)
      if (clean) {
        assistantIdx = i
        text = clean
        break
      }
    }
    if (assistantIdx < 0) return null
    // 세션의 현재 model/agent 는 두 곳에서 얻을 수 있다:
    //   1) AssistantMessage.modelID/providerID/mode  (실제 응답에 쓰인 값)
    //   2) 그 직전 UserMessage.model/agent           (사용자가 명시 선택한 값)
    // 답신을 세션에 재주입할 때는 (1) 을 우선한다 —
    // 실제 응답 모델과 동일해야 사용자 기대치와 일치한다.
    const info = items[assistantIdx].info as {
      role: "assistant"
      modelID?: string
      providerID?: string
      mode?: string
    }
    const model =
      info.providerID && info.modelID
        ? { providerID: info.providerID, modelID: info.modelID }
        : undefined
    let agent = info.mode || undefined
    if (!model || !agent) {
      // 폴백: 가장 최근 user 메시지에서 model/agent 조회
      for (let i = assistantIdx - 1; i >= 0; i--) {
        const u = items[i].info as {
          role: string
          agent?: string
          model?: { providerID: string; modelID: string }
        }
        if (u.role !== "user") continue
        if (!agent && u.agent) agent = u.agent
        break
      }
    }
    return { text, model, agent }
  }

  async function onIdle(sid: string): Promise<void> {
    const sres = await client.session.get({ path: { id: sid } })
    const session = sres.data
    if (!session || session.parentID) return
    const ctx = await lastAssistantContext(sid)
    if (!ctx) return
    const fp = fingerprint(ctx.text)
    if (lastFp.get(sid) === fp) {
      log.debug("skip.unchanged", { sid })
      return
    }
    const token = await auth.getAccessToken(sid)
    if (!token) {
      log.warn("skip.noauth", { sid })
      return
    }
    const chatId = sessionChat.get(sid)
    if (!chatId) {
      void ensureChatDirectory()
      log.info("skip.nochat", { sid })
      return
    }
    if (!cursors.has(sid)) {
      const seed = await graph.latestMessageId(chatId)
      if (seed == null) {
        log.warn("skip.seed.failed", { sid, chatId })
        return
      }
      cursors.set(sid, seed)
    }
    const html = renderOutbound(session.title ?? "OpenCode", sid, ctx.text, MAX_HTML_BYTES)
    const mid = await graph.sendMessage(chatId, html)
    if (!mid) {
      log.error("send.failed", { sid })
      return
    }
    sentIds.add(mid)
    // 커서는 여기서 전진시키지 않는다. busy 중 도착해 아직 처리 못한 답신이
    // 커서~mid 사이에 있을 수 있고, 자기 메시지는 sentIds가 걸러낸다 (설계 결정 #11).
    lastFp.set(sid, fp)
    const already = waiting.has(sid)
    waiting.set(sid, { since: Date.now(), model: ctx.model, agent: ctx.agent })
    log.info("sent", { sid, mid, model: ctx.model, agent: ctx.agent })
    if (!already) {
      const gen = (pollGen.get(sid) ?? 0) + 1
      pollGen.set(sid, gen)
      void pollLoop(sid, gen)
    }
  }

  log.info("plugin.loaded")

  // Lazy connect: auth.bootstrap() / ensureChatDirectory() 는 /connect-teams 스킬 호출 시에만 실행한다.
  // SKILL.md 자체는 opencode 가 skill 등록을 위해 파일이 존재해야 하므로 로드 시 설치를 유지한다.
  installConnectTeamsSkill()

  // messageID 당 최초 완료 파트에만 auth 안내를 append 하기 위한 dedup 세트 (여러 텍스트 파트 방지).
  const noticedMessages = new Set<string>()

  return {
    dispose: async () => {
      auth.dispose()
      log.info("plugin.disposed")
    },
    event: async ({ event }) => {
      try {
        await onEvent(event as unknown as { type?: string; properties?: Record<string, unknown> })
      } catch (e) {
        log.error("event.crash", { error: String(e) })
      }
    },
    "chat.message": async (input, output) => {
      try {
        await onChatMessage(input, output)
      } catch (e) {
        log.error("chat.message.crash", { error: String(e) })
      }
    },
    "experimental.text.complete": async (input, output) => {
      try {
        if (!connectRequested && looksLikeConnectTrigger(output.text)) {
          activateConnect("text.complete.marker", input.sessionID)
        }
        if (!connectRequested) return
        if (connectActivating) {
          await connectActivating
          connectActivating = null
        }
        // sub-session 응답에는 auth notice / chat-select notice 를 모두 append 하지 않는다.
        // sub-session 의 response text 는 Task tool result 로 root 에 되돌아오므로,
        // notice 가 새어들어가면 사용자 대화가 오염된다.
        const isSub = await isSubSession(input.sessionID)
        const authNotice = !isSub && !noticedMessages.has(input.messageID) ? auth.getAuthNotice() : null
        if (authNotice) {
          noticedMessages.add(input.messageID)
          output.text = output.text.trimEnd() + AUTH_NOTICE_MARKER + authNotice
          log.debug("auth.notice.appended", { sessionID: input.sessionID, messageID: input.messageID })
        }
        if (!sessionChat.has(input.sessionID) && chatDirectory.chats && !pendingNoticed.has(input.messageID)) {
          if (isSub) {
            pendingNoticed.add(input.messageID)
          } else {
            pendingNoticed.add(input.messageID)
            output.text = output.text.trimEnd() + CHAT_SELECT_MARKER + renderChatSelectNotice(chatDirectory.chats)
            log.debug("chat.select.notice.appended", { sessionID: input.sessionID, messageID: input.messageID })
          }
        }
      } catch (e) {
        log.error("text.complete.crash", { error: String(e) })
      }
    },
  }

  function looksLikeConnectTrigger(text: string): boolean {
    if (!text) return false
    return text.includes(CONNECT_TRIGGER_MARKER) || text.includes(CONNECT_TRIGGER_PHRASE)
  }

  // 사용자 입력에서 /connect-teams slash command 를 감지 (assistant 응답 marker 와 무관하게 즉시 활성화).
  function detectConnectSlashCommand(parts: Array<{ type: string; text?: string }>): boolean {
    for (const p of parts) {
      if (p.type !== "text" || typeof p.text !== "string") continue
      const first = p.text.trimStart().split(/\r?\n/, 1)[0] ?? ""
      if (/^\/connect-teams(\b|$)/.test(first)) return true
    }
    return false
  }

  // SDK v1 Event union 은 question.* / permission.asked 를 포함하지 않는다 (ref: omo-opencode-hook-analysis.md §3.2).
  // 런타임에는 발화하므로 unknown → 최소 shape narrowing 후 소비한다.
  type UnknownEvent = { type?: string; properties?: Record<string, unknown> }
  type QuestionAskedProps = {
    id: string
    sessionID: string
    questions: Array<{ question?: string; header?: string; options?: unknown; multiple?: boolean }>
  }
  type PermissionAskedProps = {
    id: string
    sessionID: string
    permission?: string
    type?: string
    title?: string
    metadata?: Record<string, unknown>
    pattern?: string | string[]
    patterns?: string[]
  }

  function isQuestionAsked(e: UnknownEvent): e is { type: "question.asked"; properties: QuestionAskedProps } {
    const p = e.properties as QuestionAskedProps | undefined
    return e.type === "question.asked" && !!p && typeof p.id === "string" && typeof p.sessionID === "string" && Array.isArray(p.questions)
  }

  function isPermissionAsked(e: UnknownEvent): e is { type: "permission.asked"; properties: PermissionAskedProps } {
    const p = e.properties as PermissionAskedProps | undefined
    return (
      (e.type === "permission.asked" || e.type === "permission.updated") &&
      !!p && typeof p.id === "string" && typeof p.sessionID === "string"
    )
  }

  type PermissionRepliedProps = { sessionID: string; permissionID: string; response?: string }
  type QuestionRepliedProps = { sessionID: string; requestID: string }

  function isPermissionReplied(e: UnknownEvent): e is { type: "permission.replied"; properties: PermissionRepliedProps } {
    const p = e.properties as PermissionRepliedProps | undefined
    return e.type === "permission.replied" && !!p && typeof p.sessionID === "string" && typeof p.permissionID === "string"
  }

  function isQuestionSettled(e: UnknownEvent): e is { type: "question.replied" | "question.rejected"; properties: QuestionRepliedProps } {
    const p = e.properties as QuestionRepliedProps | undefined
    return (e.type === "question.replied" || e.type === "question.rejected") &&
      !!p && typeof p.sessionID === "string" && typeof p.requestID === "string"
  }

  // opencode 세션에서 이미 응답한 경우 pending 을 즉시 정리한다.
  // 이후 도착한 Teams 답장은 "알 수 없는 code" 로 조용히 skip 되므로 race 자체가 해소된다 (Oracle 블로커 #3).
  function dropPendingByPermissionID(permissionID: string): void {
    for (const [code, p] of pendingPermissions) {
      if (p.permissionID === permissionID) {
        pendingPermissions.delete(code)
        releaseAlertPending(p.rootSid, code)
        log.info("pending.perm.dropped", { code, permissionID })
      }
    }
  }

  function dropPendingByRequestID(requestID: string): void {
    for (const [code, q] of pendingQuestions) {
      if (q.requestID === requestID) {
        pendingQuestions.delete(code)
        releaseAlertPending(q.rootSid, code)
        log.info("pending.quest.dropped", { code, requestID })
      }
    }
  }

  async function onEvent(event: UnknownEvent): Promise<void> {
    const type = event?.type
    if (!type) return

    if (type === "session.idle") {
      if (!connectRequested) return
      const sid = (event.properties as { sessionID?: string } | undefined)?.sessionID
      if (!sid) return
      try {
        await onIdle(sid)
      } catch (e) {
        log.error("onIdle.crash", { sid, error: String(e) })
      }
      return
    }

    if (!connectRequested) return
    if (isQuestionAsked(event)) {
      await onQuestionAsked(event.properties)
      return
    }
    if (isPermissionAsked(event)) {
      await onPermissionAsked(event.properties)
      return
    }
    if (isPermissionReplied(event)) {
      dropPendingByPermissionID(event.properties.permissionID)
      return
    }
    if (isQuestionSettled(event)) {
      dropPendingByRequestID(event.properties.requestID)
      return
    }
  }

  async function alertGateCommon(sid: string): Promise<{ chatId: string; token: string } | null> {
    const chatId = sessionChat.get(sid)
    if (!chatId) return null
    const token = await auth.getAccessToken(sid)
    if (!token) return null
    return { chatId, token }
  }

  async function alertGateForEvents(
    sid: string,
  ): Promise<{ chatId: string; token: string; rootSid: string; isSub: boolean } | null> {
    const rootSid = await resolveRoot(sid)
    const common = await alertGateCommon(rootSid)
    if (!common) return null
    return { ...common, rootSid, isSub: rootSid !== sid }
  }

  async function onQuestionAsked(props: QuestionAskedProps): Promise<void> {
    if (noticedQuestions.has(props.id)) return
    // in-flight dedup: gate 진입 전에 예약해 race 를 차단, 실패 시 rollback 하여 재시도 가능하게 유지.
    noticedQuestions.add(props.id)
    const gate = await alertGateForEvents(props.sessionID)
    if (!gate) {
      noticedQuestions.delete(props.id)
      return
    }
    const code = generateShortCode("q", existingShortCodes())
    if (!code) {
      log.error("question.shortcode.exhausted", { id: props.id })
      noticedQuestions.delete(props.id)
      return
    }
    const subCodes = props.questions.length > 1 ? props.questions.map((_, i) => `${code}-${i + 1}`) : []
    pendingQuestions.set(code, {
      sid: props.sessionID,
      requestID: props.id,
      rootSid: gate.rootSid,
      createdAt: Date.now(),
      questions: props.questions,
      subCodes,
      subAnswers: props.questions.map(() => null),
    })
    registerAlertPending(gate.rootSid, code)
    const subInfo = gate.isSub ? { rootSid: gate.rootSid, subSid: props.sessionID } : null
    const html = renderQuestionAlert(props.sessionID, props.questions, subInfo, code, subCodes)
    const mid = await graph.sendMessage(gate.chatId, html)
    if (mid) sentIds.add(mid)
    ensurePollLoop(gate.rootSid)
    notify(`❓ Microsoft Teams: 세션이 질문을 대기 중`, { variant: "warning" })
    log.info("question.notified", { id: props.id, sid: props.sessionID, rootSid: gate.rootSid, isSub: gate.isSub, code, subCount: subCodes.length })
  }

  async function onPermissionAsked(props: PermissionAskedProps): Promise<void> {
    if (noticedPermissions.has(props.id)) return
    noticedPermissions.add(props.id)
    const gate = await alertGateForEvents(props.sessionID)
    if (!gate) {
      noticedPermissions.delete(props.id)
      return
    }
    const paths = extractPathsFromPermission({
      type: props.type,
      pattern: props.pattern ?? (props.patterns && props.patterns.length ? props.patterns : undefined),
      metadata: props.metadata,
    })
    const outside = paths.filter((p) => !isInsideWorkspace(p, workspace))
    const code = generateShortCode("p", existingShortCodes())
    if (!code) {
      log.error("permission.shortcode.exhausted", { id: props.id })
      noticedPermissions.delete(props.id)
      return
    }
    pendingPermissions.set(code, {
      sid: props.sessionID,
      permissionID: props.id,
      rootSid: gate.rootSid,
      createdAt: Date.now(),
      outside,
      kind: props.permission ?? props.type ?? "unknown",
    })
    registerAlertPending(gate.rootSid, code)
    const subInfo = gate.isSub ? { rootSid: gate.rootSid, subSid: props.sessionID } : null
    const html = renderPermissionAlert(props.sessionID, props, outside, subInfo, code)
    const mid = await graph.sendMessage(gate.chatId, html)
    if (mid) sentIds.add(mid)
    ensurePollLoop(gate.rootSid)
    const label = outside.length > 0 ? `workspace 외부 접근` : `permission`
    notify(`⚠️ Microsoft Teams: ${label} 요청 대기 중`, { variant: "warning" })
    log.info("permission.notified", {
      id: props.id,
      sid: props.sessionID,
      rootSid: gate.rootSid,
      isSub: gate.isSub,
      outside,
      code,
    })
  }

  type SubInfo = { rootSid: string; subSid: string } | null

  function renderSubSessionBanner(subInfo: SubInfo): string {
    if (!subInfo) return ""
    const subShort = escapeHtml(subInfo.subSid.replace(/^ses_/, "").slice(0, 8))
    const rootShort = escapeHtml(subInfo.rootSid.replace(/^ses_/, "").slice(0, 8))
    return (
      `<div style="font-size:12px;color:#6b6b6b">` +
      `🌿 sub-session <b>${subShort}</b> · root <b>${rootShort}</b>` +
      `</div>`
    )
  }

  function renderReplyHint(lines: string[]): string {
    const items = lines.map((l) => `<li><code>${escapeHtml(l)}</code></li>`).join("")
    return `<br><i>Teams 답장으로도 응답 가능:</i><ul>${items}</ul>`
  }

  function renderQuestionAlert(
    sid: string,
    questions: QuestionAskedProps["questions"],
    subInfo: SubInfo = null,
    code = "",
    subCodes: string[] = [],
  ): string {
    const shortId = sid.replace(/^ses_/, "").slice(0, 8)
    const rows = questions
      .map((q, i) => {
        const head = escapeHtml(q.header ?? `Q${i + 1}`)
        const body = escapeHtml(q.question ?? "")
        const opts = Array.isArray(q.options)
          ? (q.options as Array<{ label?: unknown; description?: unknown } | string>)
              .map((o) => {
                if (typeof o === "string") return `<li>${escapeHtml(o)}</li>`
                const label = escapeHtml(String((o as { label?: unknown }).label ?? ""))
                const desc = (o as { description?: unknown }).description
                return `<li><b>${label}</b>${desc ? ` — ${escapeHtml(String(desc))}` : ""}</li>`
              })
              .join("")
          : ""
        const multi = q.multiple ? " <i>(복수 선택)</i>" : ""
        const codeHint = code
          ? subCodes.length > 0
            ? ` · <code>${escapeHtml(subCodes[i] ?? "")}</code>`
            : ` · <code>${escapeHtml(code)}</code>`
          : ""
        return (
          `<div><b>${head}</b>${multi}${codeHint}<br>${body}` +
          (opts ? `<ul>${opts}</ul>` : "") +
          `</div>`
        )
      })
      .join("<hr>")
    const replyHint = code
      ? renderReplyHint(
          subCodes.length > 0
            ? [
                `#answer ${subCodes[0]} <label>`,
                `#answer ${subCodes[1] ?? code} <label>`,
                `#skip ${code}`,
              ]
            : [`#answer ${code} <label>[, <label2>]`, `#answer ${code} text: <자유 텍스트>`, `#skip ${code}`],
        )
      : ""
    return (
      renderSubSessionBanner(subInfo) +
      `<b>[❓ Microsoft Teams: 세션 질문 · ${escapeHtml(shortId)}]</b>` +
      (code ? ` · <code>${escapeHtml(code)}</code>` : "") +
      `<br>` +
      rows +
      replyHint
    )
  }

  function renderPermissionAlert(
    sid: string,
    props: PermissionAskedProps,
    outsidePaths: string[],
    subInfo: SubInfo = null,
    code = "",
  ): string {
    const shortId = sid.replace(/^ses_/, "").slice(0, 8)
    const kind = props.permission ?? props.type ?? "unknown"
    const outsideList = outsidePaths.length
      ? `<b>⚠️ Workspace 외부 경로 (Teams 승인 시 주의)</b>:<ul>${outsidePaths.map((p) => `<li>${escapeHtml(p)}</li>`).join("")}</ul>`
      : ""
    const replyHint = code ? renderReplyHint([`#allow ${code}`, `#reject ${code}`]) : ""
    return (
      renderSubSessionBanner(subInfo) +
      `<b>[⚠️ Microsoft Teams: 권한 요청 · ${escapeHtml(shortId)}]</b>` +
      (code ? ` · <code>${escapeHtml(code)}</code>` : "") +
      `<br>` +
      `<b>Type</b>: ${escapeHtml(String(kind))}<br>` +
      (props.title ? `<b>Title</b>: ${escapeHtml(props.title)}<br>` : "") +
      `<b>Workspace</b>: <span style="${INLINE_CODE_STYLE}">${escapeHtml(workspace)}</span><br>` +
      outsideList +
      replyHint
    )
  }

  async function onChatMessage(
    input: { sessionID: string },
    output: { parts: Array<{ type: string; text?: string }> },
  ): Promise<void> {
    const parts = output.parts ?? []
    // slash command 는 sessionChat 이 이미 확정된 세션에서도 재활성화 신호로 유효해야 하므로 gate 전에 검사한다.
    if (!connectRequested && detectConnectSlashCommand(parts)) {
      activateConnect("chat.message.slash", input.sessionID)
    }
    if (sessionChat.has(input.sessionID)) return
    // sub session 은 Teams 발신에 관여하지 않으므로 chatId 파싱/저장/디렉토리 fetch 를 모두 skip.
    if (await isSubSession(input.sessionID)) return
    let detectedPart: { type: string; text?: string } | null = null
    let detectedId: string | null = null
    for (const p of parts) {
      if (p.type !== "text" || typeof p.text !== "string") continue
      const m = CHAT_ID_RE.exec(p.text)
      if (m) {
        detectedPart = p
        detectedId = m[1]
        break
      }
    }
    if (!detectedId) {
      if (connectRequested && !chatDirectory.chats) await ensureChatDirectory()
      return
    }
    // chatId 를 붙여넣기만 하는 경우도 implicit connect 로 취급 — 사용자의 의도가 명확하다.
    if (!connectRequested) activateConnect("chat.message.chatId", input.sessionID)
    if (connectActivating) {
      await connectActivating
      connectActivating = null
    }
    const matched = chatDirectory.chats?.find((c) => c.id === detectedId)
    const label = matched?.topic || detectedId
    sessionChat.set(input.sessionID, detectedId)
    log.info("chat.selected", { sid: input.sessionID, chatId: detectedId, label, fromList: Boolean(matched) })
    notify(`Microsoft Teams 채팅방 연동 완료: ${label}`, { variant: "success" })
    if (detectedPart && typeof detectedPart.text === "string") {
      detectedPart.text = detectedPart.text.replace(
        CHAT_ID_RE,
        `(Microsoft Teams 채팅방 "${label}" 이(가) 이번 세션에 연동되었습니다. 짧게 확인만 해주고 이전 작업을 계속 이어가세요.)`,
      )
    }
  }
}
