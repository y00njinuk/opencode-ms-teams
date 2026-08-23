#!/usr/bin/env node
// postinstall.mjs — 글로벌 설치 시 opencode.json 에 자동 등록.
//
// 실행 조건:
//   - npm_config_global === "true"        (글로벌 설치),  또는
//   - OPENCODE_MS_TEAMS_AUTO_INSTALL=1    (로컬 의존성 케이스 opt-in)
//
// 건너뛰는 조건:
//   - CI 환경 (opt-in 없이는 CI 에서 파일시스템 손상 방지)
//   - 로컬 install (opt-in 없이는 개발자가 다른 경로로 등록할 수 있음)
//
// 실패 시: 경고만 남기고 exit 0 — npm install 자체는 성공하도록 한다.
//         사용자는 이후 `opencode-ms-teams install` 로 수동 재시도.

import { dirname } from "node:path"
import { fileURLToPath } from "node:url"
import { install, resolveConfigDir } from "./scripts/install-lib.mjs"

const HERE = dirname(fileURLToPath(import.meta.url))

function shouldRun() {
  const isGlobal = process.env.npm_config_global === "true"
  const optIn = process.env.OPENCODE_MS_TEAMS_AUTO_INSTALL === "1"
  const isCI = process.env.CI === "true" || process.env.CONTINUOUS_INTEGRATION === "true"
  if (isCI && !optIn) {
    console.log("[opencode-ms-teams] postinstall: CI 감지 — 건너뜀. 필요 시 `opencode-ms-teams install` 수동 실행.")
    return false
  }
  if (!isGlobal && !optIn) {
    console.log(
      "[opencode-ms-teams] postinstall: 로컬 install 감지 — 건너뜀. 배포하려면 `opencode-ms-teams install` 실행 또는 OPENCODE_MS_TEAMS_AUTO_INSTALL=1 설정.",
    )
    return false
  }
  return true
}

async function main() {
  if (!shouldRun()) return
  try {
    const configDir = resolveConfigDir(undefined)
    console.log(`[opencode-ms-teams] postinstall: ${configDir}/opencode.json 에 플러그인 등록 중`)
    const result = await install({ configDir, packageDir: HERE })
    if (result.changed) {
      console.log(`[opencode-ms-teams] postinstall: 등록 완료 (backup: ${result.backup})`)
    } else {
      console.log(`[opencode-ms-teams] postinstall: 이미 등록됨 (${result.reason})`)
    }
    if (result.cacheLink) {
      console.log(`[opencode-ms-teams] postinstall: opencode 캐시 링크 → ${result.cacheLink}`)
    }
    console.log("[opencode-ms-teams] postinstall: 확인은 `opencode-ms-teams doctor`")
  } catch (err) {
    console.warn(`[opencode-ms-teams] postinstall: ${err.message}`)
    console.warn("[opencode-ms-teams] postinstall: 설치 미완료 — `opencode-ms-teams install` 로 재시도")
    process.exit(0)
  }
}

main()
