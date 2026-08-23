#!/usr/bin/env bash
# release.sh - opencode-ms-teams npm 패키지 릴리스 자동화.
#
# 사용법:
#   scripts/release.sh <patch|minor|major> [--yes]
#
# 워크플로우:
#   1. Pre-flight 체크 (git 상태, 브랜치, 원격 동기화)
#   2. 테스트 실행 (bun test + typecheck + build)
#   3. 버전 bump 미리보기
#   4. **사용자 승인 게이트 #1** (버전 bump 확인)
#   5. npm version <bump> 실행 (커밋 + 태그 생성)
#   6. Publish 미리보기 (npm publish --dry-run)
#   7. **사용자 승인 게이트 #2** (공개 npm 배포 확인)
#   8. npm publish → 공개 npm registry (registry.npmjs.org)
#   9. git push --follow-tags
#
# 옵션:
#   --yes: 모든 확인 프롬프트 자동 승인 (CI 전용, 대화형에서는 사용 금지)

set -euo pipefail

RED=$'\033[0;31m'
GREEN=$'\033[0;32m'
YELLOW=$'\033[0;33m'
BLUE=$'\033[0;34m'
CYAN=$'\033[0;36m'
BOLD=$'\033[1m'
RESET=$'\033[0m'

log()    { printf "%s[release]%s %s\n" "$CYAN" "$RESET" "$*"; }
info()   { printf "%s[info]%s    %s\n" "$BLUE" "$RESET" "$*"; }
ok()     { printf "%s[ok]%s      %s\n" "$GREEN" "$RESET" "$*"; }
warn()   { printf "%s[warn]%s    %s\n" "$YELLOW" "$RESET" "$*"; }
err()    { printf "%s[error]%s   %s\n" "$RED" "$RESET" "$*" >&2; }
prompt() { printf "%s[confirm]%s %s" "$BOLD" "$RESET" "$*"; }

# ---------- 인자 파싱 ----------
BUMP="${1:-}"
AUTO_YES=false
for arg in "$@"; do
  case "$arg" in
    --yes|-y) AUTO_YES=true ;;
  esac
done

if [[ ! "$BUMP" =~ ^(patch|minor|major)$ ]]; then
  err "usage: $0 <patch|minor|major> [--yes]"
  err "  patch: 0.1.0 → 0.1.1 (버그 수정)"
  err "  minor: 0.1.0 → 0.2.0 (기능 추가, backward-compatible)"
  err "  major: 0.1.0 → 1.0.0 (breaking change)"
  exit 1
fi

# ---------- 프로젝트 루트 이동 ----------
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_ROOT"
log "프로젝트 루트: $PROJECT_ROOT"
log "Bump 유형:    $BUMP"

# ---------- 사용자 확인 헬퍼 ----------
confirm() {
  local message="$1"
  if [ "$AUTO_YES" = true ]; then
    info "[--yes] 자동 승인: $message"
    return 0
  fi
  local reply
  prompt "$message [y/N]: "
  read -r reply </dev/tty
  case "$reply" in
    [yY]|[yY][eE][sS]) return 0 ;;
    *) return 1 ;;
  esac
}

# ---------- STEP 1: Pre-flight ----------
log "=========================================="
log "STEP 1: Pre-flight 체크"
log "=========================================="

if [ -n "$(git status --porcelain)" ]; then
  err "작업 디렉토리에 커밋되지 않은 변경사항이 있음:"
  git status --short
  err "먼저 변경사항을 커밋하거나 stash 하세요."
  exit 1
fi
ok "Working tree clean"

CURRENT_BRANCH="$(git rev-parse --abbrev-ref HEAD)"
if [ "$CURRENT_BRANCH" != "master" ] && [ "$CURRENT_BRANCH" != "main" ]; then
  warn "현재 브랜치가 master/main이 아님: $CURRENT_BRANCH"
  confirm "이 브랜치에서 릴리스를 진행하시겠습니까?" || { err "사용자가 취소함"; exit 1; }
fi
ok "브랜치: $CURRENT_BRANCH"

log "원격 저장소 fetch..."
git fetch origin --quiet || warn "git fetch 실패 (원격 없음일 수 있음) - 계속 진행"
UPSTREAM="origin/$CURRENT_BRANCH"
if git rev-parse --verify "$UPSTREAM" >/dev/null 2>&1; then
  LOCAL="$(git rev-parse HEAD)"
  REMOTE="$(git rev-parse "$UPSTREAM")"
  BASE="$(git merge-base HEAD "$UPSTREAM")"
  if [ "$LOCAL" = "$REMOTE" ]; then
    ok "원격과 동기화됨"
  elif [ "$LOCAL" = "$BASE" ]; then
    err "로컬이 원격보다 뒤처져 있음. 먼저 pull 하세요."
    exit 1
  elif [ "$REMOTE" = "$BASE" ]; then
    info "로컬에 push되지 않은 커밋 있음 (release 후 함께 push됨)"
    git log --oneline "$UPSTREAM..HEAD"
  else
    err "로컬과 원격이 diverge됨. 먼저 정리하세요."
    exit 1
  fi
else
  warn "원격 브랜치 $UPSTREAM 없음 (새 브랜치)"
fi

# npm 로그인 확인 (publish 직전에 실패하면 버전 bump 롤백이 필요해지므로 먼저 검사)
if ! NPM_USER="$(npm whoami 2>/dev/null)"; then
  err "npm 로그인 상태가 아님. 'npm login' 후 다시 실행하세요."
  exit 1
fi
ok "npm 사용자: $NPM_USER"

# ---------- STEP 2: 테스트 ----------
log "=========================================="
log "STEP 2: 테스트 실행 (bun test + typecheck + build)"
log "=========================================="

if ! bun test; then err "bun test 실패. 릴리스 중단."; exit 1; fi
if ! bunx tsc --noEmit; then err "typecheck 실패. 릴리스 중단."; exit 1; fi
if ! npm run build; then err "build 실패. 릴리스 중단."; exit 1; fi
ok "모든 테스트/빌드 통과"

# ---------- STEP 3: 버전 미리보기 ----------
log "=========================================="
log "STEP 3: 버전 bump 미리보기"
log "=========================================="

CURRENT_VERSION="$(node -p "require('./package.json').version")"
NEXT_VERSION="$(node -e "
  const semver = '$CURRENT_VERSION'.split('.').map(Number);
  const bump = '$BUMP';
  if (bump === 'patch') semver[2]++;
  else if (bump === 'minor') { semver[1]++; semver[2] = 0; }
  else if (bump === 'major') { semver[0]++; semver[1] = 0; semver[2] = 0; }
  console.log(semver.join('.'));
")"

info "현재 버전: $CURRENT_VERSION"
info "다음 버전: $NEXT_VERSION"

log "이번 릴리스에 포함될 커밋:"
LAST_TAG="$(git describe --tags --abbrev=0 2>/dev/null || echo '')"
if [ -n "$LAST_TAG" ]; then
  git log --oneline "${LAST_TAG}..HEAD" | head -20
else
  git log --oneline -10
fi

# ---------- STEP 4: 승인 게이트 #1 ----------
log "=========================================="
log "STEP 4: 버전 bump 승인 (게이트 #1)"
log "=========================================="

confirm "${BOLD}${YELLOW}버전을 $CURRENT_VERSION → $NEXT_VERSION 로 올리시겠습니까?${RESET}" || {
  err "사용자가 버전 bump를 거부함. 릴리스 중단."
  exit 1
}

# ---------- STEP 5: 버전 bump ----------
log "=========================================="
log "STEP 5: npm version $BUMP 실행"
log "=========================================="

VERSION_OUTPUT="$(npm version "$BUMP")"
NEW_TAG="$VERSION_OUTPUT"
ok "버전 bump 완료: $NEW_TAG"

rollback_version() {
  err "롤백 중..."
  git tag -d "$NEW_TAG" 2>/dev/null || true
  git reset --hard HEAD~1
  warn "버전 bump 커밋 및 태그 롤백 완료. 현재 버전: $CURRENT_VERSION"
}

# ---------- STEP 6: dry-run ----------
log "=========================================="
log "STEP 6: Publish 미리보기 (dry-run)"
log "=========================================="

info "npm publish --dry-run 실행..."
if ! npm publish --dry-run 2>&1 | tail -40; then
  err "Dry-run 실패. 롤백 필요."
  rollback_version
  exit 1
fi
ok "Dry-run 통과 - 패키지 구조 이상 없음"

info "패키지 정보:"
node -e "
  const pkg = require('./package.json');
  console.log('  name:    ' + pkg.name);
  console.log('  version: ' + pkg.version);
  console.log('  target:  ' + (pkg.publishConfig?.registry || 'registry.npmjs.org (public)'));
"

# ---------- STEP 7: 승인 게이트 #2 ----------
log "=========================================="
log "STEP 7: 공개 npm 배포 승인 (게이트 #2)"
log "=========================================="

warn "이제 공개 npm registry (registry.npmjs.org) 에 $NEW_TAG 를 배포합니다."
warn "공개 배포이므로 누구나 설치할 수 있게 됩니다."
warn "배포 후에는 동일 버전 재-publish 가 registry에 의해 거부됩니다."
warn "돌이킬 수 없는 작업이므로 신중히 확인하세요."

confirm "${BOLD}${RED}공개 npm registry에 $NEW_TAG 를 배포하시겠습니까?${RESET}" || {
  warn "사용자가 배포를 거부함. 버전 bump는 유지됩니다."
  info "나중에 배포하려면: npm publish && git push origin HEAD --follow-tags"
  info "취소하려면: git tag -d $NEW_TAG && git reset --hard HEAD~1"
  exit 0
}

# ---------- STEP 8: publish ----------
log "=========================================="
log "STEP 8: npm publish (공개 npm registry 배포)"
log "=========================================="

if ! npm publish; then
  err "npm publish 실패."
  err "수동 확인 필요:"
  err "  - npm 로그인 상태 (npm whoami)"
  err "  - npm registry 접근성 (~/.npmrc)"
  err "  - 동일 버전 중복 배포 여부"
  err ""
  err "  롤백:   git tag -d $NEW_TAG && git reset --hard HEAD~1"
  err "  재시도: npm publish"
  exit 1
fi
ok "npm 배포 완료: $NEW_TAG"

# ---------- STEP 9: push ----------
log "=========================================="
log "STEP 9: git push --follow-tags"
log "=========================================="

if ! git push origin "$CURRENT_BRANCH" --follow-tags; then
  err "git push 실패. 수동으로 push 하세요:"
  err "  git push origin $CURRENT_BRANCH --follow-tags"
  exit 1
fi
ok "커밋 및 태그 원격 push 완료"

log "=========================================="
ok "${BOLD}${GREEN}릴리스 완료: $NEW_TAG${RESET}"
log "=========================================="
info ""
info "다음 단계 (사용자 환경 업데이트):"
info "  npm install -g opencode-ms-teams@${NEXT_VERSION}"
info "  opencode-ms-teams doctor"
info ""
