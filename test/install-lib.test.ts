import { describe, test, expect, beforeEach, afterEach } from "bun:test"
import * as fs from "node:fs"
import * as os from "node:os"
import * as path from "node:path"
import { install, uninstall, doctor } from "../scripts/install-lib.mjs"

let tmp: string
let configDir: string
let dataDir: string
let configFile: string
let opencodeJson: string
let prevCacheHome: string | undefined

beforeEach(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), "mst-installlib-"))
  // uninstall() 은 opencode 캐시 링크를 지운다. 격리하지 않으면 개발자의 실제 캐시가 삭제된다.
  prevCacheHome = process.env.XDG_CACHE_HOME
  process.env.XDG_CACHE_HOME = path.join(tmp, "cache")
  configDir = path.join(tmp, "opencode")
  dataDir = path.join(configDir, "opencode-ms-teams")
  configFile = path.join(dataDir, "opencode-ms-teams.json")
  opencodeJson = path.join(configDir, "opencode.json")
  fs.mkdirSync(configDir, { recursive: true })
  fs.writeFileSync(opencodeJson, JSON.stringify({ $schema: "https://opencode.ai/config.json" }, null, 2))
})

afterEach(() => {
  if (prevCacheHome === undefined) delete process.env.XDG_CACHE_HOME
  else process.env.XDG_CACHE_HOME = prevCacheHome
  fs.rmSync(tmp, { recursive: true, force: true })
})

describe("install() config scaffolding", () => {
  test("데이터 디렉토리와 opencode-ms-teams.json 을 생성한다", async () => {
    expect(fs.existsSync(dataDir)).toBe(false)
    expect(fs.existsSync(configFile)).toBe(false)

    const result = await install({ configDir })

    expect(fs.existsSync(dataDir)).toBe(true)
    expect(fs.existsSync(configFile)).toBe(true)
    expect(result.settings).toBeDefined()
    expect(result.settings.path).toBe(configFile)
    expect(result.settings.created).toBe(true)
  })

  test("생성된 config 의 azure 값은 빈 문자열이다 (사용자가 직접 편집)", async () => {
    await install({ configDir })
    const cfg = JSON.parse(fs.readFileSync(configFile, "utf8"))
    expect(cfg.azure).toEqual({ tenantId: "", clientId: "", scopes: "" })
  })

  test("생성된 config 의 log.file / tokens.file 은 configDir 하위 경로다", async () => {
    await install({ configDir })
    const cfg = JSON.parse(fs.readFileSync(configFile, "utf8"))
    expect(cfg.log.file).toBe(path.join(dataDir, "opencode-ms-teams.log"))
    expect(cfg.log.enabled).toBe(false)
    expect(cfg.tokens.file).toBe(path.join(dataDir, "tokens.json"))
  })

  test("생성된 config 의 polling 은 src/index.ts 의 default 와 일치한다", async () => {
    await install({ configDir })
    const cfg = JSON.parse(fs.readFileSync(configFile, "utf8"))
    expect(cfg.polling).toEqual({ intervalMs: 5000, backoff1Ms: 600_000, backoff2Ms: 3_600_000 })
  })

  test("기존 config 파일이 있으면 절대 overwrite 하지 않는다", async () => {
    fs.mkdirSync(dataDir, { recursive: true })
    const existing = { azure: { tenantId: "existing-tenant", clientId: "cid", scopes: "s" }, custom: 42 }
    fs.writeFileSync(configFile, JSON.stringify(existing, null, 2))

    const result = await install({ configDir })

    const cfg = JSON.parse(fs.readFileSync(configFile, "utf8"))
    expect(cfg).toEqual(existing)
    expect(result.settings.created).toBe(false)
    expect(result.settings.reason).toBe("already-exists")
  })

  test("--force 플래그가 있어도 기존 config 는 overwrite 하지 않는다 (플러그인 재등록만 강제)", async () => {
    fs.mkdirSync(dataDir, { recursive: true })
    const existing = { azure: { tenantId: "keep-me", clientId: "keep", scopes: "keep" } }
    fs.writeFileSync(configFile, JSON.stringify(existing, null, 2))

    await install({ configDir, force: true })

    const cfg = JSON.parse(fs.readFileSync(configFile, "utf8"))
    expect(cfg).toEqual(existing)
  })

  test("tokens.json 은 install 이 만들지 않는다 (첫 device flow 시 생성)", async () => {
    await install({ configDir })
    expect(fs.existsSync(path.join(dataDir, "tokens.json"))).toBe(false)
  })

  test("uninstall 은 데이터 디렉토리와 config 파일을 남긴다", async () => {
    await install({ configDir })
    await uninstall({ configDir })
    expect(fs.existsSync(configFile)).toBe(true)
    expect(fs.existsSync(dataDir)).toBe(true)
  })

  test("doctor 는 install 직후 settings-present 를 ✓ 로 표시한다", async () => {
    await install({ configDir })
    const report = await doctor({ configDir })
    const check = report.checks.find((c: any) => c.name === "settings-present")
    expect(check?.ok).toBe(true)
    expect(check?.detail).toBe(configFile)
  })
})
