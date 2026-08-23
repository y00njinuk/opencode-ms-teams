#!/usr/bin/env node
// bin/cli.mjs — opencode-ms-teams CLI.
//
// 명령:
//   opencode-ms-teams install [--force]   plugin 을 opencode.json 에 등록
//   opencode-ms-teams uninstall           opencode.json 에서 plugin 제거
//   opencode-ms-teams doctor              설치 상태·토큰·로그 진단
//   opencode-ms-teams help                도움말 출력

import { dirname, resolve } from "node:path"
import { fileURLToPath } from "node:url"
import { install, uninstall, doctor, PLUGIN_SPEC, resolveConfigDir } from "../scripts/install-lib.mjs"

const HERE = dirname(fileURLToPath(import.meta.url))
// bin/cli.mjs 는 패키지 루트/bin 안에 있으므로 부모 폴더가 패키지 루트.
const PACKAGE_DIR = resolve(HERE, "..")

const args = process.argv.slice(2)
const cmd = args[0] || "help"

const has = (flag) => args.includes(flag)

function printHelp() {
  console.log(`opencode-ms-teams CLI

Usage:
  opencode-ms-teams install [--force]   plugin 등록 (기본: 이미 등록되어 있으면 skip)
  opencode-ms-teams uninstall           plugin 제거
  opencode-ms-teams doctor              설치 상태 진단
  opencode-ms-teams help                이 도움말

Env:
  OPENCODE_CONFIG_DIR         opencode config 경로 (기본 ~/.config/opencode)
  MICROSOFT_TEAMS_CONFIG_DIR  플러그인 설정 파일 위치 (기본: OPENCODE_CONFIG_DIR)

Config file:
  \${OPENCODE_CONFIG_DIR}/opencode-ms-teams/opencode-ms-teams.json  (grouped schema: azure/polling/log/tokens; chatId 는 세션별 in-memory)
`)
}

async function main() {
  try {
    if (cmd === "install") {
      const result = await install({ packageDir: PACKAGE_DIR, force: has("--force") })
      if (result.changed) {
        console.log(`[install] ${PLUGIN_SPEC} 를 ${result.path} 에 등록했습니다`)
        console.log(`[install] backup: ${result.backup}`)
      } else {
        console.log(`[install] 이미 등록되어 있음 (${result.reason})`)
      }
      if (result.cacheLink) console.log(`[install] opencode 캐시 링크 → ${result.cacheLink}`)
      if (result.settings?.created) {
        console.log(`[install] settings 생성됨: ${result.settings.path}`)
        console.log(`[install] azure.tenantId / clientId / scopes 를 편집한 뒤 opencode 를 실행하세요`)
      } else if (result.settings) {
        console.log(`[install] settings 이미 존재 (유지): ${result.settings.path}`)
      }
      console.log(`[install] opencode TUI 를 재시작하세요`)
    } else if (cmd === "uninstall") {
      const result = await uninstall({})
      if (result.changed) {
        console.log(`[uninstall] ${PLUGIN_SPEC} 를 ${result.path} 에서 제거했습니다`)
        console.log(`[uninstall] backup: ${result.backup}`)
      } else {
        console.log(`[uninstall] 등록되어 있지 않음 (${result.reason})`)
      }
    } else if (cmd === "doctor") {
      const report = await doctor({})
      console.log(`configDir: ${report.configDir}`)
      console.log(`configPath: ${report.configPath}`)
      console.log()
      let bad = 0
      for (const c of report.checks) {
        const mark = c.ok ? "✓" : "✗"
        if (!c.ok) bad++
        console.log(`${mark} ${c.name.padEnd(20)} ${c.detail ?? ""}`)
      }
      process.exit(bad > 0 ? 1 : 0)
    } else {
      printHelp()
      if (cmd !== "help" && cmd !== "-h" && cmd !== "--help") process.exit(1)
    }
  } catch (err) {
    console.error(`[opencode-ms-teams] ${err.message}`)
    process.exit(1)
  }
}

main()
