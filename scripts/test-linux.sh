#!/usr/bin/env bash
# test-linux.sh — Ubuntu / Rocky Linux 컨테이너에서 배포본 스모크 테스트를 돌린다.
#
# 사용법:
#   scripts/test-linux.sh            # 두 배포판 모두
#   scripts/test-linux.sh ubuntu     # Ubuntu 24.04 만
#   scripts/test-linux.sh rocky      # Rocky Linux 9 만
#
# 왜 필요한가:
#   개발은 macOS 에서 하지만 배포 대상에는 Ubuntu·Rocky 가 포함된다. 두 환경은
#   심볼릭 링크·실행 비트·기본 Node 버전·os.homedir() 폴백에서 갈리며, 실제로
#   Rocky 의 기본 nodejs 는 16 (engines 미만) 이고 HOME 없는 uid 에서는
#   os.homedir() 가 throw 한다. 이 스크립트는 그 차이를 어느 호스트에서든 재현한다.
#
# 검증 방식:
#   소스 트리가 아니라 `npm pack` 산출물(tarball)을 컨테이너에 넣고 글로벌 설치한다.
#   즉 사용자가 `npm install -g opencode-ms-teams` 했을 때와 같은 경로를 밟는다.
#
# 정리:
#   컨테이너는 --rm 으로 즉시 제거된다. docker 데몬이 꺼져 있어 이 스크립트가
#   직접 기동한 경우에만 종료 시 다시 끈다 (MS_TEAMS_KEEP_DOCKER=1 로 끄기 방지).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

RED=$'\033[0;31m'; GREEN=$'\033[0;32m'; YELLOW=$'\033[0;33m'; CYAN=$'\033[0;36m'; RESET=$'\033[0m'
log()  { printf "%s[test-linux]%s %s\n" "$CYAN" "$RESET" "$*"; }
ok()   { printf "%s[ok]%s          %s\n" "$GREEN" "$RESET" "$*"; }
warn() { printf "%s[warn]%s        %s\n" "$YELLOW" "$RESET" "$*"; }
err()  { printf "%s[error]%s       %s\n" "$RED" "$RESET" "$*" >&2; }

TARGET="${1:-all}"
case "$TARGET" in
  all|ubuntu|rocky) ;;
  *) err "usage: $0 [all|ubuntu|rocky]"; exit 1 ;;
esac

if ! command -v docker >/dev/null 2>&1; then
  err "docker 를 찾을 수 없다. Docker / OrbStack / colima 중 하나가 필요하다."
  exit 1
fi

# ---------- docker 데몬 준비 ----------
STARTED_DAEMON=0
PKGDIR=""
cleanup() {
  local status=$?
  [ -n "$PKGDIR" ] && rm -rf "$PKGDIR"
  if [ "$STARTED_DAEMON" = "1" ] && [ "${MS_TEAMS_KEEP_DOCKER:-0}" != "1" ]; then
    log "docker 데몬 종료 (이 스크립트가 기동했으므로 원상복구)"
    orb stop >/dev/null 2>&1 || warn "orb stop 실패 — 수동으로 'orb stop'"
  fi
  return "$status"
}
trap cleanup EXIT

if ! docker info >/dev/null 2>&1; then
  if command -v orb >/dev/null 2>&1; then
    log "docker 데몬이 꺼져 있음 — OrbStack 기동"
    orb start >/dev/null 2>&1 || true
    for _ in $(seq 1 30); do docker info >/dev/null 2>&1 && break; sleep 2; done
    docker info >/dev/null 2>&1 && STARTED_DAEMON=1
  fi
fi
docker info >/dev/null 2>&1 || { err "docker 데몬에 연결할 수 없다."; exit 1; }

# ---------- 배포본 준비 ----------
PKGDIR="$(mktemp -d)"
log "npm pack 실행 (prepack: clean + build 포함)"
npm pack --pack-destination "$PKGDIR" >/dev/null
TARBALL="$(ls "$PKGDIR"/opencode-ms-teams-*.tgz | head -1)"
ok "$(basename "$TARBALL") ($(du -h "$TARBALL" | cut -f1))"

run_one() {
  local name="$1" dockerfile="$2" tag="$3"
  log "=========================================="
  log "$name"
  log "=========================================="
  docker build -q -f "test/linux/$dockerfile" -t "$tag" test/linux >/dev/null
  docker run --rm \
    -v "$PKGDIR:/pkg:ro" \
    -v "$ROOT/test/linux:/smoke:ro" \
    "$tag" bash /smoke/smoke.sh
}

FAILED=()
if [ "$TARGET" = "all" ] || [ "$TARGET" = "ubuntu" ]; then
  run_one "Ubuntu 24.04" Dockerfile.ubuntu opencode-ms-teams-test:ubuntu-24.04 || FAILED+=("ubuntu")
fi
if [ "$TARGET" = "all" ] || [ "$TARGET" = "rocky" ]; then
  run_one "Rocky Linux 9" Dockerfile.rocky opencode-ms-teams-test:rocky-9 || FAILED+=("rocky")
fi

if [ "${#FAILED[@]}" -gt 0 ]; then
  err "실패한 배포판: ${FAILED[*]}"
  exit 1
fi
ok "모든 배포판 통과"
