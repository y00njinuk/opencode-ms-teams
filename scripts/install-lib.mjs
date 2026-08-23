// scripts/install-lib.mjs — opencode-ms-teams 설치/제거/진단 로직.
//
// 순수 Node ESM 모듈. postinstall.mjs 와 bin/cli.mjs 가 공유한다.
//
// 핵심 동작:
//   - resolveConfigDir(): ~/.config/opencode 위치 결정
//     (환경변수 XDG_CONFIG_HOME / OPENCODE_CONFIG_DIR 오버라이드 지원)
//   - install(): opencode.json 의 plugin 배열에 우리 패키지 spec 을 추가.
//     기존 opencode.json 은 .bak.<timestamp> 로 백업. 이미 등록되어 있으면 skip.
//   - uninstall(): plugin 배열에서 우리 패키지 spec 을 제거.
//   - doctor(): 설치 상태와 tokens.json 존재 여부 진단.
//
// 규칙:
//   - PLUGIN_SPEC 은 항상 semver 없이 "opencode-ms-teams" 만 사용.
//     opencode 로더가 npm 이 이미 설치한 최신 버전을 자동으로 잡는다 (semver 고정은 install 시 하지 않음).
//   - opencode.json 이 없거나 JSON 파싱에 실패하면 새로 만들지 않고 에러 반환 —
//     사용자가 기존 config 를 손상시켰을 가능성을 알린다.
//   - 이 파일은 오직 로컬 파일시스템만 만진다. 네트워크 호출 없음.

import { promises as fs } from "node:fs"
import { homedir } from "node:os"
import { join } from "node:path"

export const PLUGIN_SPEC = "opencode-ms-teams"

// src/index.ts 의 defaultConfig() 와 스키마를 반드시 동일하게 유지할 것.
// install-lib 은 순수 Node ESM 이라 TS 소스를 import 할 수 없으므로 별도 선언한다.
// 값을 바꾼다면 src/index.ts 의 DEFAULT_* 상수와 defaultConfig() 도 함께 수정하고
// test/install-lib.test.ts, test/microsoft-teams.test.ts 도 동시 갱신할 것.
export const SETTINGS_FILE_NAME = "opencode-ms-teams.json"
export const DATA_DIR_NAME = "opencode-ms-teams"
export const TOKENS_FILE_NAME = "tokens.json"
export const LOG_FILE_NAME = "opencode-ms-teams.log"

function defaultSettings(dataDir) {
  return {
    azure: { tenantId: "", clientId: "", scopes: "" },
    polling: { intervalMs: 5000, backoff1Ms: 600_000, backoff2Ms: 3_600_000 },
    log: { file: join(dataDir, LOG_FILE_NAME), enabled: false },
    tokens: { file: join(dataDir, TOKENS_FILE_NAME) },
  }
}

// 데이터 디렉토리와 opencode-ms-teams.json 을 스캐폴딩한다.
// 요구사항: 기존 파일이 있으면 절대 overwrite 하지 않는다.
// tokens.json 은 여기서 만들지 않는다 (device flow 가 자동 발급하며,
// 빈 파일이 미리 있으면 인증 로직이 '토큰 있음' 으로 오판할 수 있다).
async function ensureSettingsFile(configDir) {
  const dd = join(configDir, DATA_DIR_NAME)
  const filePath = join(dd, SETTINGS_FILE_NAME)
  await fs.mkdir(dd, { recursive: true })
  try {
    await fs.access(filePath)
    return { path: filePath, created: false, reason: "already-exists" }
  } catch {
    await fs.writeFile(filePath, JSON.stringify(defaultSettings(dd), null, 2) + "\n", "utf8")
    return { path: filePath, created: true }
  }
}

// opencode 로더는 플러그인을 자체 npm 클라이언트로 registry 에서 받아 캐시에 보관한다.
// 방금 npm 으로 설치한 버전을 그대로 쓰게 하려면 캐시 위치에 심볼릭 링크를 심어야 한다
// (링크가 없으면 npm ls -g 는 새 버전을 보고하는데 opencode 는 캐시의 옛 버전을 계속 로드한다).
// XDG_CACHE_HOME 을 존중한다 (resolveConfigDir 의 XDG_CONFIG_HOME 처리와 대칭).
// uninstall() 은 캐시 링크를 무조건 지우므로, 이 오버라이드가 없으면 테스트가
// 개발자의 실제 opencode 캐시 링크를 삭제한다.
function cacheDirPath() {
  const xdg = process.env.XDG_CACHE_HOME
  const base = xdg && xdg.trim() ? xdg : join(resolveHome(), ".cache")
  return join(base, "opencode", "packages", `${PLUGIN_SPEC}@latest`)
}

// PLUGIN_SPEC 이 scope 를 갖든 안 갖든 동일하게 동작하도록 "/" 로 분해해 이어붙인다.
function cacheLinkPath() {
  return join(cacheDirPath(), "node_modules", ...PLUGIN_SPEC.split("/"))
}

async function linkToOpencodeCache(packageDir) {
  const link = cacheLinkPath()
  await fs.mkdir(join(link, ".."), { recursive: true })
  try {
    const st = await fs.lstat(link)
    if (st.isSymbolicLink() || st.isDirectory()) {
      await fs.rm(link, { recursive: true, force: true })
    }
  } catch {}
  await fs.symlink(packageDir, link, "dir")
  // opencode 는 캐시 폴더 최상위에 package.json 을 요구한다 (npm install 결과 모방).
  const outerPkg = join(cacheDirPath(), "package.json")
  try {
    await fs.access(outerPkg)
  } catch {
    await fs.writeFile(outerPkg, JSON.stringify({ dependencies: { [PLUGIN_SPEC]: "latest" } }, null, 2))
  }
  return link
}

async function unlinkFromOpencodeCache() {
  const link = cacheLinkPath()
  try {
    await fs.rm(link, { recursive: true, force: true })
    return true
  } catch {
    return false
  }
}

// HOME 이 없고 실행 uid 가 /etc/passwd 에 없으면 homedir() 는 libuv ENOENT 로 throw 한다
// (컨테이너·systemd 유닛에서 실제로 발생). 원인과 해결책이 드러나는 메시지로 바꾼다.
function resolveHome() {
  try {
    return homedir()
  } catch {
    throw new Error(
      "홈 디렉토리를 확인할 수 없습니다 (HOME 미설정 + /etc/passwd 에 uid 없음). " +
        "OPENCODE_CONFIG_DIR / XDG_CONFIG_HOME / XDG_CACHE_HOME 을 설정하거나 HOME 을 지정하세요.",
    )
  }
}

export function resolveConfigDir(explicit) {
  if (explicit) return explicit
  if (process.env.OPENCODE_CONFIG_DIR) return process.env.OPENCODE_CONFIG_DIR
  const xdg = process.env.XDG_CONFIG_HOME
  if (xdg) return join(xdg, "opencode")
  return join(resolveHome(), ".config", "opencode")
}

function configPath(configDir) {
  return join(configDir, "opencode.json")
}

async function readJson(path) {
  const raw = await fs.readFile(path, "utf8")
  return JSON.parse(raw)
}

async function writeJson(path, data) {
  await fs.writeFile(path, JSON.stringify(data, null, 2) + "\n", "utf8")
}

async function backupConfig(path) {
  const ts = new Date().toISOString().replace(/[:.]/g, "-")
  const backup = `${path}.bak.${ts}`
  await fs.copyFile(path, backup)
  return backup
}

function specMatches(entry, spec) {
  if (typeof entry === "string") return entry === spec || entry.startsWith(`${spec}@`)
  if (Array.isArray(entry) && typeof entry[0] === "string") return entry[0] === spec || entry[0].startsWith(`${spec}@`)
  return false
}

export async function install({ configDir, packageDir, force = false } = {}) {
  const dir = resolveConfigDir(configDir)
  const cfgPath = configPath(dir)

  let cfg
  try {
    cfg = await readJson(cfgPath)
  } catch (err) {
    if (err.code === "ENOENT") {
      throw new Error(
        `opencode.json 이 없습니다: ${cfgPath}\n` +
          `먼저 opencode 를 최소 1회 실행하여 기본 config 를 생성한 뒤 다시 설치하세요.`,
      )
    }
    throw new Error(`opencode.json 파싱 실패: ${cfgPath}\n원인: ${err.message}`)
  }

  const plugins = Array.isArray(cfg.plugin) ? cfg.plugin : []
  const already = plugins.findIndex((p) => specMatches(p, PLUGIN_SPEC))

  let cacheLink
  if (packageDir) {
    cacheLink = await linkToOpencodeCache(packageDir)
  }

  // 데이터 디렉토리와 opencode-ms-teams.json 은 opencode.json 재등록 여부와 무관하게
  // 항상 스캐폴딩한다 (사용자가 azure 값을 install 직후 편집할 수 있어야 함).
  const settings = await ensureSettingsFile(dir)

  if (already >= 0 && !force) {
    return { changed: false, reason: "already-registered", path: cfgPath, cacheLink, settings }
  }

  const backup = await backupConfig(cfgPath)
  const next = { ...cfg, plugin: [...plugins.filter((p) => !specMatches(p, PLUGIN_SPEC)), PLUGIN_SPEC] }
  await writeJson(cfgPath, next)
  return { changed: true, path: cfgPath, backup, cacheLink, settings }
}

export async function uninstall({ configDir } = {}) {
  const dir = resolveConfigDir(configDir)
  const cfgPath = configPath(dir)

  await unlinkFromOpencodeCache()

  let cfg
  try {
    cfg = await readJson(cfgPath)
  } catch (err) {
    if (err.code === "ENOENT") return { changed: false, reason: "no-config" }
    throw new Error(`opencode.json 파싱 실패: ${cfgPath}\n원인: ${err.message}`)
  }

  const plugins = Array.isArray(cfg.plugin) ? cfg.plugin : []
  const filtered = plugins.filter((p) => !specMatches(p, PLUGIN_SPEC))
  if (filtered.length === plugins.length) {
    return { changed: false, reason: "not-registered", path: cfgPath }
  }

  const backup = await backupConfig(cfgPath)
  await writeJson(cfgPath, { ...cfg, plugin: filtered })
  return { changed: true, path: cfgPath, backup }
}

export async function doctor({ configDir } = {}) {
  const dir = resolveConfigDir(configDir)
  const cfgPath = configPath(dir)
  const report = { configDir: dir, configPath: cfgPath, checks: [] }

  const push = (name, ok, detail) => report.checks.push({ name, ok, detail })

  try {
    const cfg = await readJson(cfgPath)
    push("config-readable", true, `${cfgPath}`)
    const plugins = Array.isArray(cfg.plugin) ? cfg.plugin : []
    const registered = plugins.some((p) => specMatches(p, PLUGIN_SPEC))
    push("plugin-registered", registered, registered ? PLUGIN_SPEC : "opencode-ms-teams install 로 등록하세요")
  } catch (err) {
    push("config-readable", false, err.message)
    push("plugin-registered", false, "config 를 먼저 복구하세요")
  }

  // opencode 로더 우회용 캐시 심볼릭 링크 확인
  const link = cacheLinkPath()
  try {
    const st = await fs.lstat(link)
    if (st.isSymbolicLink()) {
      const target = await fs.readlink(link)
      push("cache-link", true, `${link} -> ${target}`)
    } else {
      push("cache-link", true, `${link} (directory)`)
    }
  } catch {
    push("cache-link", false, `${link} 없음 — opencode-ms-teams install 재실행 필요`)
  }

  const settingsPath = join(dir, "opencode-ms-teams", "opencode-ms-teams.json")
  try {
    const st = await fs.stat(settingsPath)
    push("settings-present", st.size > 0, settingsPath)
  } catch {
    push("settings-present", false, `${settingsPath} — 플러그인 최초 로드 시 default 로 자동 생성됩니다`)
  }

  const tokenPath = join(dir, "opencode-ms-teams", "tokens.json")
  try {
    const st = await fs.stat(tokenPath)
    push("tokens-present", st.size > 0, tokenPath)
  } catch {
    push("tokens-present", false, `${tokenPath} — 첫 프롬프트 응답 시 device flow 로 자동 발급됩니다`)
  }

  const logPath = join(dir, "opencode-ms-teams", "opencode-ms-teams.log")
  try {
    const st = await fs.stat(logPath)
    push("log-present", true, `${logPath} (${st.size} bytes)`)
  } catch {
    push("log-present", false, `${logPath} — 로그가 비활성이거나 아직 생성되지 않음`)
  }

  return report
}
