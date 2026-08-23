# CLAUDE.md — opencode-ms-teams

이 저장소는 OpenCode ↔ Microsoft Teams 양방향 npm 플러그인의 개발/배포 워크스페이스다.
Claude Code 등 AI 에이전트가 이 저장소에서 작업할 때 지켜야 할 규칙과 진입점을 정리한다.

## Project Overview

- `ARCHITECTURE.md` 에는 시스템 아키텍처와 설계 원칙이 기술되어 있으며, `CLAUDE.md` 는 개발자들이 실제로 코드를 작성할 때 따라야 하는 구체적인 가이드라인과 규칙을 담고 있다.

## 저장소 구조

```
opencode-ms-teams/
├── src/
│   └── index.ts              ← 배포 대상 (단일 파일, 유일한 export = MicrosoftTeamsPlugin)
├── bin/
│   └── cli.mjs               ← opencode-ms-teams install/uninstall/doctor CLI
├── scripts/
│   ├── install-lib.mjs       ← opencode.json 자동 패치 + 설정 스캐폴딩 라이브러리
│   ├── release.sh            ← 두 승인 게이트 + 롤백 + push 자동 릴리스
│   └── test-linux.sh         ← Ubuntu/Rocky 컨테이너 배포본 스모크 테스트 드라이버
├── postinstall.mjs           ← 글로벌 npm 설치 시 자동 배포 진입점
├── dist/                     ← tsc 빌드 산출물 (.gitignore, npm tarball 포함)
├── test/
│   ├── microsoft-teams.test.ts  ← 플러그인 코어 격리 테스트 (bun test)
│   ├── install-lib.test.ts      ← 설치/제거/진단 라이브러리 테스트
│   ├── repro.test.ts            ← 회귀 재현 테스트
│   └── linux/                   ← 배포본 리눅스 검증 (Dockerfile 2종 + smoke.sh)
├── opencode.json.example     ← 사용자용 plugin 배열 예시
├── package.json              ← opencode-ms-teams (공개 npm), publishConfig, bin, exports, files
├── tsconfig.json             ← src+test 타입 체크
├── tsconfig.build.json       ← src → dist 빌드
├── ARCHITECTURE.md           ← 시스템 아키텍처 및 설계 원칙
├── LICENSE                   ← MIT
├── .gitignore
├── README.md
└── CLAUDE.md                 ← 본 파일
```

## 핵심 불변식

1. **단일 파일 플러그인 코어**: `src/index.ts` 는 런타임 의존성 제로 (Bun 내장 `fetch` + `node:` 빌트인만). 외부 패키지 `import` 추가 금지. 타입만 필요한 경우 `import type` 사용.
2. **export 는 `MicrosoftTeamsPlugin` 하나만**: opencode 플러그인 로더가 named export 를 모두 Plugin 함수로 호출한다. 헬퍼 함수/클래스를 `export` 하면 로딩이 침묵 실패한다. (`Config`/`ChatSummary` type 은 `export type` 이므로 안전.)
3. **버전 고정**: `@opencode-ai/plugin` 과 `@opencode-ai/sdk` 는 **1.4.17** 정확 핀 (`^` 금지). peerDependency 는 `>=1.4.17` 로 유연하게, devDependency 는 정확 핀.
4. **토큰 취급**: 실 Azure 토큰 파일 (`tokens.json` 등) 은 저장소 내부에 두지 않는다. 런타임 토큰은 `~/.config/opencode/opencode-ms-teams/tokens.json` 에만 저장되며, `.gitignore` 의 `tokens.json` 규칙은 실수 방지용 방어선이다.
5. **AUTH_NOTICE_MARKER / CHAT_SELECT_MARKER 계약**: `experimental.text.complete` 훅이 붙이는 안내와 `stripAuthNotice` 가 제거하는 마커는 반드시 같은 상수를 참조해야 한다. 분리하면 Microsoft Teams 발신에 안내가 새어나간다.
6. **설정 단일 소스**: 모든 설정은 `~/.config/opencode/opencode-ms-teams/opencode-ms-teams.json` 에서 관리. 환경변수 오버라이드 도입 금지. **chatId 는 config 스키마에 두지 말 것** — 세션별 in-memory 관리가 요구사항.
7. **로그 정책**: `log.enabled=false` (기본) 이면 파일 자체를 생성하지 않는다. `true` 이면 플러그인 로드 시 truncate 후 append.
8. **용어 통일**: 사용자 노출 문자열/심볼/파일 모두 "Microsoft Teams" 표현을 사용. "Teams", "bridge", "teams-bridge" 표현 신규 도입 금지 (도메인 이름은 `microsoft-teams` 로 hyphenated).
9. **세션별 채팅 격리**: `sessionChat` Map 은 sessionID 를 키로 격리한다. 세션 하나의 chatId 변경이 다른 세션에 영향 주지 않아야 함. 세션 종료/재시작 시 상태 소실은 허용 (요구사항).
10. **Sub session gate (2단 파이프라인)**: `parentID` 가 있는 세션은 두 파이프라인이 다르게 취급된다.
    - **Idle send pipeline** (`onIdle` — assistant idle 응답 발신, chat-select 안내 append, auth notice append, chat.message 의 chatId 파싱·저장): sub session 에서 **완전 차단**. `sessionParent` 캐시를 우회하면 sub session 마다 안내가 반복 append 되는 회귀가 발생한다.
    - **Blocking user-input events** (`permission.asked`, `question.asked`): `resolveRoot()` 로 root sessionID 를 찾아 root 의 chatId 로 **bubble-up 발신**한다. 원격 사용자가 opencode UI 로 돌아와 승인/응답해야 sub session 이 unblock 되므로 무음 처리 시 요구사항 위반. 알림 본문에는 sub-session 배너 (`🌿 sub-session {short} · root {short}`) 를 prepend 하여 컨텍스트 식별을 돕는다. `isSubSession()` 이 transient `session.get` 실패에 `sessionParent.set(sid, null)` 로 캐시를 오염시키면 이후 bubble-up 이 영구적으로 잘못된 sid 로 chatId 를 찾는 회귀가 발생하므로 catch 블록은 절대 캐시를 남기지 않는다.
11. **`defaultSettings()` / `defaultConfig()` 동기화**: `install-lib.mjs` 의 `defaultSettings()` 와 `src/index.ts` 의 `defaultConfig()` 는 스키마·기본값을 반드시 동일하게 유지할 것.
12. **Lazy connect 원칙**: 플러그인 로드만으로는 어떤 외부 통신 (Azure device flow / Graph `/me/chats` / notify toast / 응답 텍스트 mutation) 도 발생시키지 않는다. `activateConnect()` 는 오직 세 경로 (① `experimental.text.complete` 의 `CONNECT_TRIGGER_MARKER` / `CONNECT_TRIGGER_PHRASE`, ② `chat.message` 의 `/connect-teams` slash, ③ `chat.message` 의 chatId 붙여넣기) 로만 트리거된다. eager bootstrap 재도입 금지.
13. **`CONNECT_TRIGGER_MARKER` / SKILL.md 동기화**: `src/index.ts` 의 상수 `CONNECT_TRIGGER_MARKER` 문자열과 `CONNECT_TEAMS_SKILL_CONTENT` 의 marker 지시 문구가 동일해야 한다. marker 를 변경하려면 `CONNECT_TEAMS_SKILL_VERSION` 을 반드시 bump 하여 사용자 스킬 파일이 자동 갱신되도록 할 것.
14. **`permission.ask` plugin hook 등록 금지**: opencode 1.4.17 core 에 trigger 호출부가 없어 dead API 이다 (근거: `omo-opencode-hook-analysis.md` §2 항목 4). 권한 알림은 반드시 `event` hook 의 `permission.asked` bus event 분기로 처리한다. `permission.updated` 는 SDK 타입 오류일 뿐 실제 발화하는 이름은 `permission.asked`이다.
15. **question.asked / permission.asked runtime narrowing**: SDK v1 Event union 에 두 이벤트가 누락되어 있으므로 `event as unknown as {type?: string; properties?: Record<string, unknown>}` 로 캐스팅 후 `isQuestionAsked()` / `isPermissionAsked()` 로 최소 shape 검증 후 소비. `event as any` 로 축약하지 말 것 (후속 개발자가 타입 무시로 오해할 수 있음).
16. **Teams 답장 대체 응답 파이프라인 불변식**:
    - **답장 verb 는 top-level keyword**: `#allow` / `#reject` / `#answer` / `#skip` 은 `TRIGGER_RE` alternation 에 포함된 최상위 트리거 keyword 이다. `REPLY_VERB_KEYWORDS = new Set(["#allow","#reject","#answer","#skip"])` 로 관리. `TRIGGER_RE` 는 반드시 keyword 길이 내림차순 (`opencode|question|answer|reject|agent|allow|skip|ask|ai`) 정렬을 유지해 longest-match 를 보장한다.
    - **shortCode 생성**: `generateShortCode('p' | 'q', existing)` 로 base36 5자 무작위 + prefix. 충돌 시 최대 8회 재시도 후 실패는 알림 skip 으로 유예 (attack surface 를 줄이기 위해 hash 기반 결정성 부여 금지).
    - **alertPending 게이트**: `alertPending: Map<rootSid, Set<code>>` 를 `pollLoop` 종료 조건에 포함시켜, `waiting` 없이도 permission/question 답장을 계속 폴링한다. `ensurePollLoop(rootSid)` 는 `onPermissionAsked` / `onQuestionAsked` 성공 후 반드시 호출. 미호출 시 답장이 도달해도 소비하지 못하는 회귀 (Oracle 블로커 #1).
    - **Reply verb keyword 라우팅**: pollLoop 에서 `REPLY_VERB_KEYWORDS.has(trig.keyword)` 이면 `parseReplyCommand(keyword, payload)` 로 라우팅. 파싱 성공/실패 무관하게 `continue` 하여 절대 `promptAsync` 로 fallthrough 하지 않는다 ("allow p2a7k9" 가 세션 메시지로 주입되는 회귀 방지). `#opencode` / `#question` / `#agent` / `#ask` / `#ai` / `가재야` / `막둥아` 는 기존대로 promptAsync 주입 경로 유지.
    - **parseReplyCommand 시그니처**: `parseReplyCommand(keyword: string, payload: string)`. verb 는 keyword 로 이미 확정되므로 payload 는 `<code>` 부터 시작한다. `text:` prefix 는 free-text answer 의 greedy tail terminator (Oracle 조정 #6).
    - **`allow-always` 금지**: Teams 원격 답장으로 영구 권한 부여는 지원하지 않는다 (원격 계정 탈취 시 permanent broad grant 위험). `once` 만 지원. 영구 승인은 opencode TUI 에서 처리.
    - **Permission 응답 채널**: v1 SDK top-level 메서드 `client.postSessionIdPermissionsPermissionId({ path: { id, permissionID }, body: { response } })` 를 사용. `client.session.*` 아래에는 없다 (혼동 방지). 대상 sid 는 `pending.sid` (원본 sub-session sid) 이지 `rootSid` 아님.
    - **Question 응답 채널**: v1 SDK 에 client 메서드 없음 → `${input.serverUrl.origin}/question/{reqID}/reply` (또는 `/reject`) 를 raw fetch 로 POST. baseUrl 우선 순위: `input.serverUrl` (URL) → `client._client.getConfig().baseUrl` 폴백. **실측 주의**: TS 선언은 `client.getConfig()` 를 노출하나 런타임 인스턴스는 `_client.getConfig()` 로만 접근 가능하므로 top-level `getConfig` 를 그대로 호출하면 undefined 참조. 두 소스 모두 없으면 조용히 skip.
    - **Multi-question sub-code**: `questions.length > 1` 이면 각 질문에 `{rootCode}-{index+1}` 부여. 사용자는 각 sub-code 로 개별 답변. 모든 sub 채워지면 한번에 `Array<Array<string>>` assemble 후 POST. `parseReplyCommand` 가 `q9rbvg-2` 를 `code=q9rbvg, subIndex=2` 로 분해해 반환하므로 `pendingQuestions.get(cmd.code)` 는 항상 root 를 조회한다.
    - **완료 이벤트 정리**: `permission.replied`, `question.replied`, `question.rejected` 는 v1 SDK Event union 에 미포함이므로 `isPermissionReplied` / `isQuestionSettled` runtime narrower 로 소비. 매칭 시 `dropPendingByPermissionID` / `dropPendingByRequestID` 로 pending 즉시 정리 (TUI 가 먼저 응답한 경우 stale code 방지 — Oracle 블로커 #3).
    - **TTL**: `REPLY_TTL_MS = 24 * 60 * 60 * 1000` (24시간). `pollLoop` 각 tick 마다 `sweepExpiredPending()` 이 만료된 pending 을 정리 (SSE 이벤트 유실 대비 belt-and-suspenders). 정상 flow 에서는 `permission.replied` / `question.replied` 이벤트로 훨씬 빨리 정리되므로 24시간은 극단 상황 (SSE 유실·플러그인 재시작 등) 대비 상한.
    - **응답 확인 UX**: 성공 시 `<b>✅ 반영됨</b> · <code>{code}</code> {verb}`, 실패 시 `<b>❌ 적용 실패</b> ...` 를 Teams 로 후속 발신. Multi-question 부분 답변 시 `<b>📝 답변 기록됨</b>` 후속 발신. 모든 후속 메시지의 mid 는 반드시 `sentIds.add(mid)` 로 등록 (자기 발신 재소비 방지).

## 개발 워크플로

```bash
# 초기 설치
npm install --no-audit --no-fund

# 검증 (변경 후 반드시)
bun test              # 130 pass 기대 (3개 파일)
bunx tsc --noEmit     # 오류 0건 기대
npm run build         # dist 재빌드
npm publish --dry-run # tarball 구성 확인

# 배포 대상 리눅스 검증 (release 전 필수, docker 필요)
npm run test:linux    # Ubuntu 24.04 + Rocky Linux 9 에서 tarball 글로벌 설치 스모크
npm run test:ubuntu   # 한 배포판만 돌릴 때
npm run test:rocky

# 로컬 실기 검증 (opencode 캐시에 심볼릭 링크)
CACHE_DIR="$HOME/.cache/opencode/packages/opencode-ms-teams@latest/node_modules"
mkdir -p "$CACHE_DIR"
ln -sfn "$(pwd)" "$CACHE_DIR/opencode-ms-teams"
node bin/cli.mjs install
opencode serve --print-logs   # plugin.loaded 로그 확인
```

새 기능 추가 시:
1. `ARCHITECTURE.md` 의 설계 원칙과 상충하지 않는지 확인
2. `test/microsoft-teams.test.ts` 에 실패 테스트를 먼저 작성 (TDD)
3. `src/index.ts` 만 수정 (export 1개 원칙 준수)
4. `bun test && bunx tsc --noEmit && npm run build` 통과 확인
5. `ARCHITECTURE.md`, `README.md`, `CLAUDE.md` 영향 있으면 함께 갱신

## 자주 하는 실수

- `src/index.ts` 에 헬퍼를 `export` 로 노출 → 로더가 Plugin 함수로 호출하다 크래시 (`export type` 은 안전)
- 커서를 발신 시점에 전진 → busy 중 도착한 답신 유실 (ARCHITECTURE.md 세션 발신 파이프라인 참고)
- fingerprint 가드 제거 → 무한 재발신 루프
- refresh 시 scope 에서 `offline_access` 제외 → 다음 token refresh 실패
- 실 토큰 파일 (`tokens.json`) 을 저장소/배포에 포함 → 사용자 토큰 유출
- 코드 블록에 `class="language-*"` 부착 → Microsoft Teams 가 line number 를 자동으로 붙임. 반드시 `<pre style="..."><code>...</code></pre>` (class 없음) 로 감쌀 것.
- `<div style="white-space:pre">` 사용 → Teams UI 에서 개행이 collapsed 되어 코드가 한 줄로 뭉침.
- device flow 안내를 별도 프롬프트/toast 로만 표시 → `experimental.text.complete` 훅에서 응답 끝에 `AUTH_NOTICE_MARKER` 로 append 해야 함.
- Microsoft Teams 발신 전 `stripAuthNotice()` 미호출 → 안내 텍스트가 Teams 로 새어나감.
- eager auth 재도입 → `/connect-teams` 스킬을 호출하지 않은 세션에서도 device flow 가 시작되어 요구사항 위반. `auth.bootstrap()` 은 오직 `activateConnect()` 안에서만 호출한다.
- 답신 주입 시 `promptAsync` body 에 model/agent 미전달 → 다른 모델이 답신 처리에 개입. `lastAssistantContext()` 로 캡처한 값 재사용.
- background refresh timer 를 dispose hook 에서 취소하지 않음 → OpenCode 종료 후에도 fetch 잔여.
- 환경변수 (`OPENCODE_TEAMS_*`) 재도입 → 설정 소스 이원화 (요구사항 위반).
- `permission.ask` plugin hook 등록 → dead API (불변식 14). 알림은 `event: permission.asked` 로 처리.
- `Graph.latestMessageId` 반환값을 `number` 로 축약 (실패=0, 빈방=0 동일 취급) → 활성 채팅방 최초 연동 시 cursor=0 오염 → `messagesSince` 가 최근 50건 과거 이력 전부를 `promptAsync` 로 주입하는 회귀. 반드시 `Promise<number | null>` 유지하고 `onIdle` 은 null 시 발신 skip.
- `detectTrigger()` 를 우회해 raw reply text 를 그대로 `promptAsync` 로 주입 → 잡담 답신이 세션에 흘러들어옴 (불변식 12 위반). 신규 답신 처리 로직 추가 시 반드시 트리거 게이트를 먼저 통과시킬 것.
- 트리거 미매칭 답신에 대해 `cursors.set()` 를 건너뜀 → 다음 폴링에서 동일 메시지 재조회 무한 루프. 커서는 반드시 게이트와 무관하게 전진.
- `os.homedir()` 직접 호출 → HOME 미설정 + `/etc/passwd` 에 uid 없는 리눅스 환경 (컨테이너·systemd 유닛) 에서 libuv ENOENT 로 throw. `src/index.ts` 는 `homeDir()`, `install-lib.mjs` 는 `resolveHome()` 을 경유해 조치 가능한 메시지로 감쌀 것. 플러그인 팩토리는 `loadConfig()` 를 최상단에서 호출하므로 여기서 throw 하면 **플러그인 로드 전체가 실패**한다.
- 리눅스 검증 없이 릴리스 → Rocky 9 의 기본 `nodejs` 는 16 (engines `>=20` 미만) 이고, RHEL 계열은 `curl`/`coreutils` 를 그냥 install 하면 기본 탑재본과 충돌한다. `npm run test:linux` 로 확인할 것.
- 캐시 경로를 다루는 테스트에서 `XDG_CACHE_HOME` 미설정 → `uninstall()` 이 개발자의 **실제** opencode 캐시 링크를 삭제. `cacheDirPath()` 는 `XDG_CACHE_HOME` 을 존중하므로 테스트는 반드시 임시 경로로 오버라이드할 것.
- 트리거 정규화 단계 (NFKC / NBSP / CRLF / trim) 순서 임의 변경 → 명세 §11 수용 기준의 NFKC·조합형 자모 케이스가 오탐 · 미탐으로 회귀. 순서는 `normalizeForTrigger()` 주석과 명세 §4 를 기준으로 유지.

## Code Style

- 주석은 최대한 간결하게 작성한다. 코드 자체로 의도가 명확하지 않은 경우에만 사용한다. (예: 복잡한 알고리즘, 트릭, 외부 API 사용법 등)
  - 예시: `result 추출 우선순위: fileManualResult.classResult → IndicatorExtractor(evidences) → 기존값유지(None)`
  - 예시: `Finagle Backoff 는 "재시도 횟수" 단위이므로 maxAttempts 에서 1 을 뺀다 (초기 호출 포함).`
- 코드 구현 시 명확한 의도를 표현한다. 의미 있는 변수명과 메서드명을 사용하여 코드의 목적을 명확히 한다.
  - 예시: `fetchFileReputation(fileSeq)` → 메서드명만으로도 "fileSeq 에 해당하는 파일 평판을 조회한다"는 의도가 명확히 드러남

## Git Commit Guidelines

- 모두 한국말(Korean)로 반드시 작성. 제목과 본문은 빈 줄로 구분하며, 본문은 필요할 때만 작성
- 기능 단위로 적절히 분리하여 "1 changes = 1 commit" 원칙을 따른다.
- `ARCHITECTURE.md` 에는 현재 커밋에 반영된 내용으로 인해 변경되어야 하는 부분이 없는지 **반드시** 검토하고 수정한다.
- 1 commit 에 너무 많은 변경을 담지 않는다. 동일한 기능을 구현하더라도 범위가 크면 나누어서 작성 (최대 5개 파일의 변경까지만 허용)
- 제목 형식: `<type>: <subject>` (예: `Feat: 기능 구현`)
- type 종류: feat, fix, chore, refactor, docs, style, test, perf, ci, build, revert
- 제목은 50자 이내, 대문자로 시작, 마침표 금지, 명령조(imperative)로 작성
- 본문은 "어떻게"보다 "무엇을, 왜"에 초점 (변경 이유와 영향 설명)
- 제목 예시:
  ```
  Feat: 세션 idle 발신 파이프라인 추가
  Chore: dist 디렉토리 gitignore 추가
  ```
- 본문 예시:
  ```
  Test: 트리거 정규화 NFKC 케이스 테스트 추가

  전각 `#` 과 NBSP 가 섞인 답신이 트리거로 인식되는지 검증하는 테스트를 추가한다.

  - microsoft-teams.test.ts: 두 가지 시나리오 검증
      1) `＃agent` (전각) → `#agent` 로 정규화되어 매칭
      2) NBSP 후행 경계 케이스 매칭
  ```
  ```
  Refactor: Graph 응답 → 세션 주입 변환 헬퍼 메서드 수정

  - 메서드 이름에서 동작을 유추할 수 있도록 stripAuthNotice 이름으로 변경
  - 분리되어 있던 marker 제거 로직을 하나의 메서드로 통일하고 필요한 데이터를 인자로 직접 전달받는다.
  ```

## 인증 정보 접근

이 저장소는 다음 외부 시스템과 상호작용한다.

| 시스템 | 인증 | 용도 |
|---|---|---|
| Microsoft Graph | Azure AD delegated (device code) | Microsoft Teams 채팅 발신/폴링 |
| npm registry (registry.npmjs.org) | `~/.npmrc` 의 `_authToken` | 공개 npm publish |

Azure 관련 값 (`tenantId`, `clientId`, `scopes`) 은 `defaultConfig()` 에서 빈 문자열로 초기화되며,
설치 후 사용자가 `~/.config/opencode/opencode-ms-teams/opencode-ms-teams.json` 을 열어 직접 채워야 한다.
필요한 delegated scope: `Chat.Create Chat.Read Chat.ReadBasic Chat.ReadWrite ChatMessage.Read ChatMessage.Send offline_access User.Read`

Azure AD 앱 등록/권한 신청은 조직의 Azure 관리자 안내를 따른다.

## 관련 참고

- npm: `@opencode-ai/plugin@1.4.17`, `@opencode-ai/sdk@1.4.17`
