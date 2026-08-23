# ARCHITECTURE.md — opencode-ms-teams

## 시스템 개요

OpenCode 세션과 Microsoft Teams 채팅을 양방향으로 잇는 단일 파일 npm 플러그인.
루트 세션이 idle 이 되면 마지막 assistant 응답을 Teams 채팅으로 발신하고, Teams 답신을 동일 세션에 프롬프트로 재주입해 연속 대화를 만든다.

```
OpenCode 세션 idle
      │
      ▼
MicrosoftTeamsPlugin (session.idle 이벤트)
      │
      ├─ lastAssistantContext(): 마지막 assistant 텍스트 + model/agent 캡처
      ├─ fingerprint 가드: 동일 텍스트 재발신 방지
      ├─ stripAuthNotice(): auth/chat-select 안내 제거 → 순수 텍스트만 추출
      ├─ Graph POST /chats/{id}/messages       (아웃바운드)
      └─ pollLoop() → Graph GET /chats/{id}/messages  (커서 폴링)
             │
             └─ 답신 감지 → client.session.promptAsync(text, model, agent) 로 세션 재개
```

런타임 의존성 제로 — Bun 내장 `fetch` 와 `node:` 빌트인만 사용.

---

## 컴포넌트

### Auth

- **lazy bootstrap**: `auth.bootstrap()` 은 `/connect-teams` 스킬 호출 (`connectRequested = true`) 이후에만 실행.
  - 저장된 refresh_token 있으면 → `scheduleRefreshFromToken()` 으로 background timer 예약
  - 없으면 → `startDeviceFlow()` 로 device code 요청 시작
- **`getAuthNotice()`**: pending device code 또는 login 완료 알림 문자열 반환. 완료 알림은 1회만 반환 (`completionAt` 소비 후 0 리셋).
- **토큰 저장**: tokens.json (mode 0600) 에 저장. `expires_at` 은 `Date.now() + expires_in * 1000` 으로 계산해 함께 저장.

### Graph

| 메서드 | 엔드포인트 | 용도 |
|---|---|---|
| `sendMessage(chatId, html)` | POST `/chats/{id}/messages` | 아웃바운드 발신 |
| `latestMessageId(chatId)` | GET `/chats/{id}/messages?$top=1` | 커서 초기화 |
| `messagesSince(chatId, cursor)` | GET `/chats/{id}/messages?$top=50` | 답신 폴링 |
| `listChats(top)` | GET `/me/chats?$top=N&$orderby=...` | 채팅방 목록 |

- 401 → `auth.forceRefresh()` 후 1회 재시도
- 429 → `Retry-After` 헤더 대기 후 1회 재시도
- 사내망 SSL 오류 감지 → `NODE_EXTRA_CA_CERTS` 설정 힌트 로깅

### Plugin 훅

| 훅 | 핸들러 | 동작 |
|---|---|---|
| `event: session.idle` | `onIdle()` | fingerprint 체크 → Teams 발신 → `pollLoop()` 시작 |
| `event: question.asked` | `onQuestionAsked()` | assistant `question` tool 발동 시 Teams 로 질문/옵션 알림 (Interview Mode 포함) |
| `event: permission.asked` | `onPermissionAsked()` | opencode core 가 사용자 승인을 요청하면 Teams 로 알림 (알림만; 승인은 opencode UI 에서 결정) |
| `chat.message` | `onChatMessage()` | `/connect-teams` slash 감지 → activate / 입력 텍스트에서 chatId 감지 → 세션 등록 → toast |
| `experimental.text.complete` | (inline) | `CONNECT_TRIGGER_MARKER` 감지 → activate / auth 안내 / chat-select 안내 append |

> `permission.ask` plugin hook 은 opencode 1.4.17 core 에 trigger 호출부가 없어 dead API 다 (omo v3.17.15 코드 조사로 확인). 실제 발화하는 것은 bus event `permission.asked` 이며, 알림은 `event` hook 분기로 옮겼다.

---

## 활성화 (Lazy Connect)

플러그인 로드만으로는 어떤 외부 통신도 발생하지 않는다. 사용자가 명시적으로 `/connect-teams` 스킬을 호출한 경우에만 auth/chatDirectory bootstrap 이 시작된다.

```
activateConnect(via, sid)
  ├─ connectRequested 이미 true 면 즉시 return (멱등)
  ├─ connectRequested = true
  └─ connectActivating = Promise.race([
        (async () => { auth.bootstrap(); await ensureChatDirectory() })(),
        sleep(1800ms),  ← Oracle 권고 bounded wait; 첫 응답에 안내가 붙도록 UX guard
     ])
```

**감지 경로 3종 (defense-in-depth)**:
1. `experimental.text.complete` → `output.text` 에 `CONNECT_TRIGGER_MARKER` (`<!-- microsoft-teams:connect-request v1 -->`) 포함 시. 스킬이 assistant 에게 이 marker 를 응답 최상단에 출력하도록 지시함.
2. `experimental.text.complete` → `output.text` 에 `CONNECT_TRIGGER_PHRASE` (`**Microsoft Teams 연동을 시작합니다.**`) 포함 시 (secondary — 모델이 marker 를 누락한 경우 대비).
3. `chat.message` → `input.parts` 의 첫 줄이 `/connect-teams` slash command 로 시작 시.

또한 사용자가 처음부터 chatId (`19:...`) 를 붙여넣으면 **implicit connect** 으로 취급하여 activate.

**Why lazy**: eager bootstrap 은 사용자가 스킬을 호출하지 않은 상태에서도 device flow 요청·`/me/chats` fetch·toast 알림을 발생시켜 요구사항 위반 (첫 프롬프트에 자동 연동 시작 금지).

---

## 인증 흐름 (Azure AD Device Code Flow)

```
auth.bootstrap()   ← activateConnect() 안에서만 호출됨
  ├─ tokens.json 있음 → scheduleRefreshFromToken() → background timer 예약
  └─ 없음 → startDeviceFlow()
        │
        ├─ POST /oauth2/v2.0/devicecode → pendingDeviceCode 저장 → toast 알림
        └─ 폴링 (POST /oauth2/v2.0/token with device_code)
              ├─ authorization_pending / slow_down → 대기 후 재시도
              ├─ 성공 → tokens.json 저장 → completionAt 세트 → background refresh 예약
              └─ 실패 → pendingDeviceCode = null

experimental.text.complete 훅  (connectRequested = true 인 경우에만 mutate)
  ├─ auth.getAuthNotice() → pendingDeviceCode 있으면 로그인 안내 append (AUTH_NOTICE_MARKER)
  └─ completionAt > 0 이면 완료 알림 append (1회 후 리셋)
```

---

## 채팅방 선택 흐름

```
/connect-teams 스킬 호출 (activateConnect)
  └─ ensureChatDirectory() 호출 (백그라운드 /me/chats fetch, 1800ms bounded)

사용자 프롬프트 (connectRequested = true 상태에서만 fetch 발생)
  └─ chat.message 훅
        ├─ /connect-teams slash 감지 → activate (이미 활성이면 no-op)
        ├─ chatId 없음 → (activated 경우) ensureChatDirectory() blocking → return
        └─ chatId 감지 (CHAT_ID_RE)
              ├─ activated 아니었으면 implicit connect (chatId 붙여넣기 → 명확한 의도)
              ├─ sessionChat.set(sid, chatId)
              ├─ toast "Microsoft Teams 채팅방 연동 완료: {label}"
              └─ 입력 텍스트 내 chatId → 안내 문구로 치환 (LLM 혼동 방지)

experimental.text.complete 훅  (connectRequested = true 인 경우에만 mutate)
  └─ chatDirectory.chats 있고 sessionChat 없으면
        ├─ isSubSession() → true → skip
        └─ false → 채팅방 목록 안내 append (CHAT_SELECT_MARKER)
```

- `chatDirectory` 는 플러그인 범위 싱글톤 (계정 단위 채팅 목록은 세션마다 재조회하지 않음)
- `pendingNoticed` Set 으로 동일 messageID 에 안내 중복 append 방지
- chatId 패턴: `19:...@thread.v2` (group) 또는 `19:...@unq.gbl.spaces` (1:1)

---

## 세션 발신 파이프라인 (onIdle)

```
onIdle(sid)
  ├─ session.get → parentID 있으면 즉시 return  (sub session gate)
  ├─ lastAssistantContext()
  │     ├─ INJECT_MARKER 포함 파트 제외 (자기 주입 메시지 재발신 방지)
  │     └─ stripAuthNotice() 적용 → 순수 텍스트 추출
  ├─ fingerprint(text) === lastFp.get(sid) → skip.unchanged
  ├─ auth.getAccessToken() → null → skip.noauth
  ├─ sessionChat.get(sid) 없으면 → ensureChatDirectory() + skip.nochat
  ├─ cursors.has(sid) 없으면 → latestMessageId() 로 커서 초기화
  │     └─ null (Graph 호출 실패) → skip.seed.failed 후 return
  │        (cursor=0 저장 금지 — messagesSince 가 활성 채팅방의 과거 메시지 최대 50건을 무필터 통과시켜
  │         promptAsync 로 폭탄 주입되는 회귀 방지)
  ├─ renderOutbound() → markdown → HTML, 25KB 초과 시 절단 + "…(생략)"
  ├─ graph.sendMessage() → mid → sentIds.add(mid)
  ├─ 커서 미전진  ← 설계 원칙: busy 중 도착 답신 유실 방지 (sentIds 로 자기 메시지 제외)
  ├─ lastFp.set(sid, fp)
  └─ pollLoop() 시작 (이미 대기 중이면 세대 번호 갱신으로 기존 loop 는 자연 종료)
```

---

## 폴링 및 백오프

| 경과 시간 | 폴링 주기 |
|---|---|
| 0 ~ `backoff1Ms` (기본 10분) | `intervalMs` (기본 5s) |
| `backoff1Ms` ~ `backoff2Ms` (기본 1시간) | `intervalMs × 3` (기본 15s) |
| `backoff2Ms` 이상 | `intervalMs × 6` (기본 30s) |

- `pollGen` Map 으로 세대 관리 → 새 발신 시 세대 증가 → 이전 loop 자연 종료
- 답신 감지 → `waiting.delete(sid)` → `promptAsync` 주입 → 성공 시 return
- 주입 실패 (`res.error`) → `waiting.set(sid, st)` 복구 → 다음 폴링에서 재시도

---

## Prefix Keyword Trigger 게이트

Teams 답신을 세션에 주입하기 전 `detectTrigger()` 로 결정론적 키워드 매칭을 수행. 매칭 실패한 답신은 **주입 파이프라인에서 완전 제외**되어 잡담이 세션에 흘러들지 않는다.

```
messagesSince() → replies
  └─ per-reply:
        raw = html ? htmlToText(content) : content.trim()
        detectTrigger(raw):
          ├─ normalizeForTrigger():
          │    ├─ 제어문자 제거 (U+200B, U+FEFF, U+200E/F)
          │    ├─ NFKC        (전각 '＃' → '#', 조합형 자모 → 완성형)
          │    ├─ CRLF → LF
          │    ├─ NBSP·전각공백 → 공백
          │    └─ trim
          ├─ TRIGGER_RE.exec(normalized)  (offset 0, IGNORECASE, UNICODE)
          │    └─ 매칭 X → skip (log: trigger.skip.nomatch)
          ├─ 후행 구분자 소비 → matched_span 확장
          └─ payload = normalized.slice(end).replace(/^[ \t\n]+/, "")
               └─ payload 빈 문자열 → skip (log: trigger.skip.emptyPayload)
매칭 성공 payload 들 → "\n---\n" 로 join → promptAsync
```

**계약**:
- 키워드 집합: `#opencode`, `#question`, `#agent`, `#ask`, `#ai`, `가재야`, `막둥아` (§ Prefix Keyword Trigger 명세 v1)
- 후행 경계: ` \t\n,.!?~:;` + `…，。！？　` + EOS
- alternation 은 longest-match 안전을 위해 길이 내림차순 정렬 (미래에 `#agentic` 등 추가 대비)
- payload 만 세션 주입 (키워드 원문은 세션 컨텍스트에서 제거)
- 재스캔 없음: `#ask #agent 질문` → `#ask` 만 소비, payload = `#agent 질문`
- 커서(`cursors`)는 필터 여부와 무관하게 전진 — 미매칭 메시지도 소비된 것으로 간주 (다음 폴에서 과거 이력 재조회 안 됨)
- 로그는 `log.enabled=true` 일 때만 파일에 기록 (`trigger.match`/`trigger.skip.*` 이벤트)

**게이트 방향**:
- Teams → OpenCode 주입 경로 (본 게이트) — 미매칭 답신 skip
- OpenCode → Teams 발신 경로 (변경 없음) — 필터링 안 함

---

## 지속 토큰 Refresh

```
scheduleRefreshFromToken(t)
  └─ delay = max(REFRESH_MIN_MS, expiresAt - now - REFRESH_LEAD_MS)
        └─ setTimeout → backgroundRefresh()
              ├─ refresh 성공 → scheduleRefreshFromToken(newToken)
              ├─ refresh 실패 → 30s 후 재시도
              └─ invalid_grant → startDeviceFlow()

auth.dispose()
  └─ clearTimeout(refreshTimer)  ← OpenCode 종료 후 fetch 방지
```

상수: `REFRESH_LEAD_MS = 5min`, `REFRESH_RETRY_MS = 30s`, `REFRESH_MIN_MS = 15s`

---

## 답신 주입 (model/agent 재사용)

Teams 답신을 세션에 재주입할 때 `lastAssistantContext()` 로 캡처한 model/agent 를 `promptAsync` body 에 전달해 대화 일관성을 유지한다.

```
우선순위:
  1) AssistantMessage.modelID / providerID / mode  (실제 응답에 쓰인 값)
  2) 직전 UserMessage.agent                         (사용자 명시 선택)
  3) 두 정보 모두 없으면 body 에서 필드 생략 → opencode 기본값
```

세션 busy 중 도착한 답신은 `waiting` 이 아직 살아 있으므로 poll loop 가 계속 실행 → 다음 idle 에 순서대로 주입.

---

## 서브세션 게이트 (2단 파이프라인)

`parentID` 가 있는 세션 (Task tool 등에서 spawn) 은 두 파이프라인이 다르게 취급된다.

### Idle send pipeline — sub-session 완전 차단

- `onIdle`: `session.get` → `parentID` 존재 시 즉시 return (assistant 응답 발신 skip)
- `chat.message`: `isSubSession()` true 이면 chatId 파싱·저장 skip
- `text.complete`: `isSubSession()` true 이면 chat-select 안내 append skip, auth notice append skip

### Blocking user-input events — root 로 bubble-up

`permission.asked` / `question.asked` 는 sub-session 에서 발화해도 root sessionID 로 bubble-up 하여 root 의 chatId 로 발신한다. 이유: 두 이벤트는 사용자 응답 없이는 sub-session progress 가 blocking 되므로, 원격 사용자가 opencode UI 로 돌아와 승인/응답해야 unblock 된다. 무음 처리 시 사용자가 눈치채지 못하고 sub-session 이 무한 대기.

```
onPermissionAsked / onQuestionAsked
  ├─ noticedPermissions/Questions.has(id) → in-flight dedup (has → add 즉시, gate 실패 시 delete)
  ├─ alertGateForEvents(sid):
  │    ├─ resolveRoot(sid):
  │    │    ├─ sessionRoot 캐시 hit → 즉시 반환
  │    │    ├─ sessionParent 캐시 → 없으면 session.get 반복 walk
  │    │    ├─ visited-set cycle guard
  │    │    └─ 체인 전체를 sessionRoot 에 캐싱 (성공 시)
  │    └─ alertGateCommon(rootSid): sessionChat.get + auth.getAccessToken
  ├─ isSub = (rootSid !== sid)
  └─ renderPermissionAlert/QuestionAlert(props, subInfo=isSub ? {rootSid, subSid} : null)
        └─ subInfo 있으면 본문에 🌿 sub-session · root 배너 prepend
```

### 캐시 원칙

- `sessionParent: Map<sid, parentID | null>` — 반복 API 호출 방지
- `sessionRoot: Map<sid, rootSid>` — bubble-up 시 부모 체인 traversal 캐싱 (opencode 세션 트리는 immutable 이므로 무효화 없음)
- **catch 블록은 캐시를 남기지 않는다**: `isSubSession()` / `resolveRoot()` 가 transient `session.get` 실패 시 `sessionParent.set(sid, null)` 로 오염시키면 이후 sub-session 이 영구적으로 root 로 오분류되어 bubble-up 이 잘못된 sid 로 chatId 를 찾는다.

### 이벤트별 매트릭스

| 이벤트 (sub-session 발생) | Teams 발신 | 근거 |
|---|---|---|
| `session.idle` | ❌ | `onIdle` line 초입 `session.parentID` 체크 |
| `experimental.text.complete` | ❌ (auth notice · chat-select 모두 skip) | `text.complete` 훅 초입 `isSubSession()` 체크 |
| `chat.message` (chatId 붙여넣기) | ❌ (파싱·저장 skip) | `onChatMessage` 초입 `isSubSession()` 체크 |
| `permission.asked` | ✅ (root chatId, 배너 포함) | `alertGateForEvents` root bubble-up |
| `question.asked` | ✅ (root chatId, 배너 포함) | `alertGateForEvents` root bubble-up |

---

## Teams 답장 대체 응답 (permission / question)

opencode 세션이 사용자 승인 (`permission.asked`) 또는 질문 응답 (`question.asked`) 을 기다리며 blocking 상태일 때, 사용자가 opencode TUI 로 돌아가지 않고 Teams 채팅 답장만으로 응답할 수 있다.

```
event: permission.asked / question.asked
      │
      ▼
onPermissionAsked / onQuestionAsked
      ├─ alertGateForEvents(sid) → root sid + chatId 결정
      ├─ generateShortCode('p'|'q', existing)  ← base36 5자 + prefix, 최대 8회 충돌 재시도
      ├─ pendingPermissions.set(code, {sid, permissionID, rootSid, outside, ...})
      ├─ registerAlertPending(rootSid, code)   ← pollLoop 종료 게이트
      ├─ graph.sendMessage(chatId, html)       ← alert 본문에 shortCode + 답장 명령 예시 포함
      └─ ensurePollLoop(rootSid)               ← waiting 없어도 계속 폴링 (Oracle 블로커 #1)

pollLoop (waiting.has(sid) OR alertPending 채워짐 동안 계속)
      └─ per-reply detectTrigger:
            ├─ TRIGGER_RE 매칭 (offset 0, longest-first alternation)
            ├─ keyword ∈ REPLY_VERB_KEYWORDS ({#allow, #reject, #answer, #skip}) 이면:
            │     parseReplyCommand(keyword, payload):
            │       ├─ #allow  <code>                                → verb=allow
            │       ├─ #reject <code>                                → verb=reject
            │       ├─ #answer <code>[-<sub>] <label>[, <label2>]    → verb=answer, labels=[...]
            │       ├─ #answer <code>[-<sub>] text: <자유텍스트>     → verb=answer, text=...
            │       └─ #skip   <code>                                → verb=skip
            │     매칭 시 executeReplyCommand → continue (promptAsync fallthrough 금지)
            │       ├─ allow/reject → client.postSessionIdPermissionsPermissionId({path, body:{response}})
            │       ├─ answer       → callQuestionReply(baseUrl, reqID, [[labels...]])
            │       └─ skip         → callQuestionReject(baseUrl, reqID)
            │     성공/실패 후 <b>✅ 반영됨</b> / <b>❌ 적용 실패</b> Teams 로 후속 발신
            └─ keyword ∈ {#opencode, #question, #agent, #ask, #ai, 가재야, 막둥아} 이면:
                  기존 pipeline (payload 를 promptAsync 로 주입)
```

### 답장 문법 스펙 (v1.3)

| 명령 | 대상 | 매핑 |
|---|---|---|
| `#allow <code>` | permission | `response: "once"` |
| `#reject <code>` | permission | `response: "reject"` |
| `#answer <code> <label>[, <label2>...]` | question (single) | `answers: [[labels...]]` |
| `#answer <code>-<n> <label>` | question (multi, per sub-code) | 모두 채워진 후 assemble → POST |
| `#answer <code> text: <free text>` | question (custom text) | `answers: [[<free text>]]` |
| `#skip <code>` | question | reject |

- `<code>` = `p` 또는 `q` prefix + base36 5자 (예: `pa7k9m`, `q3x9m2`)
- Multi-question 은 `{rootCode}-1`, `{rootCode}-2`, ... 형태의 sub-code 로 답변
- 모든 sub-code 답변이 채워지면 한 번에 서버로 POST
- **`#allow-always` 는 지원하지 않는다** — `#allow` 뒤 `-` 는 후행 경계가 아니므로 detectTrigger 미매칭. 원격 계정 탈취 시 permanent grant 위험 (Oracle 블로커 #5). 영구 승인은 opencode TUI 에서 처리
- **legacy `#opencode <verb>` 문법 폐기 (v1.3.0)** — `#opencode` 는 기존 세션 주입 keyword 로 유지되므로 `#opencode allow ...` 는 assistant 세션에 일반 텍스트로 주입될 뿐 SDK 호출 없음

### API 채널

| 대상 | 채널 | 근거 |
|---|---|---|
| permission | `client.postSessionIdPermissionsPermissionId` (v1 SDK top-level 메서드) | v1 SDK dist/gen/sdk.gen.d.ts L381 |
| question reply | raw fetch `POST {baseUrl}/question/{reqID}/reply` body `{answers}` | v1 SDK 미포함, 실측 (127.0.0.1:4096, curl 200) |
| question reject | raw fetch `POST {baseUrl}/question/{reqID}/reject` | v1 SDK 미포함, 실측 |

baseUrl 결정 순서: `input.serverUrl.origin` (Plugin input 필드, 항상 존재) → `client.getConfig().baseUrl` (폴백) → null 이면 조용히 skip.

### 안전장치 · Race 처리

- **workspace 외부 경로**: alert 본문에 `<b>⚠️ Workspace 외부 경로 (Teams 승인 시 주의)</b>` 강조. 별도 verb 는 두지 않음 (Oracle 조정: 두 verb 는 friction 만 늘리고 실질 안전 효과 없음).
- **알 수 없는 code**: pending map miss → 조용히 skip (attack surface 축소).
- **chatId 불일치**: `pending.rootSid !== 답장이 온 rootSid` → 조용히 skip (다른 채팅방에서 온 code 우연 매칭 방지).
- **`permission.replied` / `question.replied` / `question.rejected` 이벤트**: opencode TUI 가 먼저 응답한 경우 즉시 pending 정리 → stale code 방지 (Oracle 블로커 #3).
- **TTL**: `REPLY_TTL_MS = 24시간`. `pollLoop` tick 마다 `sweepExpiredPending()` (SSE 유실 대비 belt-and-suspenders). 정상 flow 에서는 `permission.replied` / `question.replied` 이벤트로 훨씬 빨리 정리되므로 24시간은 극단 상황 대비 상한.
- **HTTP 4xx/5xx**: `<b>❌ 적용 실패</b>` Teams 로 후속 발신. Pending 은 유지 (재시도 여지). TTL 도달 시 자동 정리.

### 이벤트별 매트릭스 (재확인, 답장 파이프라인 포함)

| 이벤트 | 답장 지원 | 채널 |
|---|---|---|
| `permission.asked` (root/sub) | ✅ | v1 SDK `client.postSessionIdPermissionsPermissionId` |
| `question.asked` (root/sub) | ✅ | raw fetch `POST /question/{reqID}/reply` |
| `permission.replied` / `question.replied` / `question.rejected` | (pending 정리 트리거) | `dropPendingByPermissionID` / `dropPendingByRequestID` |

---

## Interview Mode / Question / Permission 알림

opencode core 가 사용자에게 무언가 결정을 요구할 때마다 Teams 채팅으로 알림을 보낸다. 결정 자체는 opencode UI 에서 처리하며 (Teams 응답 UX 는 hook 시그니처상 불가), Teams 는 대기 상태 통지 전용이다.

### 감지 이벤트

| 이벤트 | 소스 | 알림 조건 |
|---|---|---|
| `question.asked` | assistant 의 `question` tool 발동, Interview Mode (Prometheus 등) 세션의 질문 | `connectRequested` + `resolveRoot(sid)` 성공 + `sessionChat.has(rootSid)` + auth OK. sub-session 은 root 로 bubble-up 되며 본문에 배너 prepend. |
| `permission.asked` | opencode core 의 `Permission.ask` (bash, edit, read, webfetch 등) | 위와 동일. workspace 외부 경로가 있으면 별도 강조 |

두 이벤트 모두 SDK v1 Event union 에는 포함되어 있지 않다 (`permission.asked` 는 SDK 가 `permission.updated` 로 잘못 타이핑; `question.*` 은 아예 누락). 런타임에는 정상 발화하며, `src/index.ts` 의 `isQuestionAsked` / `isPermissionAsked` runtime narrower 로 unknown → 최소 shape 검증 후 소비한다. 근거: `omo-opencode-hook-analysis.md` §3.2, §5.3, §6.1.

### omo v3.17.15 와 공존

- omo 는 `event` hook 을 `void hook.event?.(...)` 로 병렬 호출하므로 순서 의존성 없음.
- omo 는 `permission.ask` plugin hook 자체를 등록하지 않으므로 `output.status` 충돌 없음.
- omo 는 `experimental.text.complete` 를 사용하지 않으므로 marker append 충돌 없음.
- omo 의 `sessionNotification` 이 동일 이벤트로 자체 알림을 발생시키지만, 본 플러그인의 알림 채널 (Teams) 과 omo 의 채널 (openclaw/system notification) 은 독립적이다.

### 중복 알림 방지

- `noticedPermissions: Set<permissionID>` — 동일 permission 재발화 시 무시
- `noticedQuestions: Set<questionID>` — 동일 question 재발화 시 무시
- 프로세스 수명 동안 unbounded 로 유지 (opencode 프로세스는 일반적으로 24h 미만; ceiling 관측 시 재검토)

### 죽은 API 정리

`permission.ask` plugin hook 은 opencode 1.4.17 core 에 trigger 호출부가 없다 (`omo-opencode-hook-analysis.md` §2 항목 4). 이 플러그인은 해당 hook 등록을 하지 않는다.

---

## Microsoft Teams 렌더링 규칙

| 마크다운 | 변환 |
|---|---|
| ` ```lang\n...\n``` ` | `<pre style="font-family:Consolas,...;background:#f4f4f4;padding:8px;border-left:3px solid #ccc"><code>...</code></pre>` |
| `` `inline` `` | `<span style="font-family:Consolas,...;background:#f4f4f4;padding:1px 4px;border-radius:3px">` |
| `**bold**` | `<b>` |
| GFM 표 (`\| col \|` + `\|---\|`) | `<table border="1" style="border-collapse:collapse">` + `<thead>`/`<tbody>` |
| 정렬 `:---` / `---:` / `:---:` | `align="left/right/center"` |
| HTML 특수문자 | 항상 escape (`&`, `<`, `>`, `"`) |

**코드 블록 두 가지 금지 패턴**:
- `class="language-*"` 부착 → Teams 가 line number 자동 부착
- `<div style="white-space:pre">` 사용 → Teams UI 에서 개행 collapsed

---

## 설정 스키마

파일 위치: `~/.config/opencode/opencode-ms-teams/opencode-ms-teams.json`

```jsonc
{
  "azure": {
    "tenantId": "",    // 빈 문자열로 초기화 — 사용자가 반드시 채워야 함
    "clientId": "",
    "scopes": ""       // offline_access 반드시 포함 (refresh token 발급)
  },
  "polling": {
    "intervalMs": 5000,
    "backoff1Ms": 600000,
    "backoff2Ms": 3600000
  },
  "log": {
    "file": "~/.config/opencode/opencode-ms-teams/opencode-ms-teams.log",
    "enabled": false   // true 이면 플러그인 로드 시 truncate 후 append
  },
  "tokens": {
    "file": "~/.config/opencode/opencode-ms-teams/tokens.json"
  }
}
```

- `azure` 값 하드코딩 금지: 잘못된 tenant 로 device flow 를 시도하면 Azure 에러가 모호해짐 → 명시적 실패 + 사용자 안내 흐름
- `chatId` 는 config 에 저장하지 않음 — 세션마다 in-memory 관리 (요구사항)
- ENOENT 일 때만 default 로 생성. 기존 파일은 `--force` 로도 overwrite 불가.
- `install-lib.mjs` 의 `defaultSettings()` 와 `src/index.ts` 의 `defaultConfig()` 는 스키마·기본값을 항상 동일하게 유지할 것.

---

## 설치 메커니즘 — opencode 캐시 링크

opencode 는 플러그인을 자체 npm 클라이언트로 registry 에서 받아 `~/.cache/opencode/packages/` 에 보관한다.
캐시가 이미 차 있으면 재-fetch 하지 않으므로, 링크가 없으면 `npm ls -g` 는 새 버전을 보고하는데
opencode 는 캐시에 남은 옛 버전을 계속 로드할 수 있다.

**장치**: npm 글로벌 설치 위치를 opencode 캐시 경로에 심볼릭 링크로 연결해 방금 설치한 버전으로 고정.

```
~/.cache/opencode/packages/opencode-ms-teams@latest/
  package.json                            ←  { dependencies: { "opencode-ms-teams": "latest" } } 스텁
  node_modules/opencode-ms-teams  →  (symlink)  →  npm global install 위치
```

- `install-lib.mjs` 의 `cacheDirPath()` / `cacheLinkPath()` / `linkToOpencodeCache()` 에 격리
- 경로는 `PLUGIN_SPEC` 에서 파생되므로 패키지 이름이 바뀌어도 상수 한 곳만 수정하면 된다

---

## 릴리스 파이프라인

`scripts/release.sh` 실행 단계:

1. Pre-flight 체크 (git 상태·브랜치·원격 동기화)
2. `bun test` + `tsc --noEmit` + `npm run build`
3. **승인 게이트 #1** (버전 bump 확인)
4. `npm version <bump>` (커밋 + 태그)
5. `npm publish --dry-run` 검증
6. **승인 게이트 #2** (공개 npm 배포 확인, 돌이킬 수 없음)
7. `npm publish` → 공개 npm registry (registry.npmjs.org)
8. `git push --follow-tags`

`--yes` 옵션은 CI 전용. 대화형 세션에서 사용 금지.
