#!/usr/bin/env bash
# test/linux/smoke.sh — Linux 배포본 스모크 테스트 (컨테이너 안에서 실행된다).
#
# 검증 대상은 "npm 으로 배포된 tarball 이 리눅스에서 설치·동작하는가" 이다.
# 소스 트리가 아니라 /pkg 에 마운트된 tarball 을 글로벌 설치해서 확인한다.
#
# 각 단계는 실패 시 즉시 종료하며, 무엇이 깨졌는지 stdout 에 남긴다.
set -euo pipefail

TARBALL="$(ls /pkg/opencode-ms-teams-*.tgz | head -1)"
PASS=0
step() { printf '\n[%s] %s\n' "$(( ++PASS ))" "$*"; }

step "환경"
echo "  distro : $(. /etc/os-release && echo "$PRETTY_NAME")"
echo "  node   : $(node -v)"
echo "  npm    : $(npm -v)"
echo "  tarball: $(basename "$TARBALL")"

step "글로벌 설치 (postinstall 포함)"
# opencode.json 이 아직 없으므로 postinstall 은 경고만 남기고 exit 0 이어야 한다.
# 여기서 npm install 이 실패하면 사용자의 npm install -g 가 통째로 깨진다.
npm install -g "$TARBALL"

step "CLI 실행 (PATH · shebang · 실행 비트)"
opencode-ms-teams help > /dev/null
echo "  opencode-ms-teams help → OK"

step "opencode.json 부재 시 install 이 명확히 실패하는가"
if opencode-ms-teams install > /tmp/out 2>&1; then
  echo "  FAIL: opencode.json 없이 install 이 성공했다"; cat /tmp/out; exit 1
fi
grep -q "opencode.json 이 없습니다" /tmp/out || { echo "  FAIL: 예상 메시지 없음"; cat /tmp/out; exit 1; }
echo "  기대한 안내 메시지 확인"

step "install → opencode.json 등록 + 캐시 링크 + 설정 스캐폴딩"
mkdir -p "$HOME/.config/opencode"
printf '{\n  "$schema": "https://opencode.ai/config.json"\n}\n' > "$HOME/.config/opencode/opencode.json"
opencode-ms-teams install
node -e '
const fs = require("node:fs");
const p = process.env.HOME + "/.config/opencode/opencode.json";
const cfg = JSON.parse(fs.readFileSync(p, "utf8"));
if (!Array.isArray(cfg.plugin) || !cfg.plugin.includes("opencode-ms-teams")) {
  console.error("  FAIL: plugin 배열 미등록 →", JSON.stringify(cfg.plugin)); process.exit(1);
}
console.log("  plugin 배열:", JSON.stringify(cfg.plugin));
'

step "캐시 심볼릭 링크가 실제 설치본을 가리키는가"
LINK="$HOME/.cache/opencode/packages/opencode-ms-teams@latest/node_modules/opencode-ms-teams"
[ -L "$LINK" ] || { echo "  FAIL: 심볼릭 링크 아님: $LINK"; ls -la "$(dirname "$LINK")" || true; exit 1; }
TARGET="$(readlink -f "$LINK")"
[ -f "$TARGET/dist/index.js" ] || { echo "  FAIL: 링크 대상에 dist/index.js 없음: $TARGET"; exit 1; }
echo "  $LINK"
echo "    → $TARGET"
STUB="$HOME/.cache/opencode/packages/opencode-ms-teams@latest/package.json"
grep -q '"opencode-ms-teams"' "$STUB" || { echo "  FAIL: 스텁 package.json 이상"; cat "$STUB"; exit 1; }
echo "  스텁 package.json OK"

step "빌드 산출물 ESM 로드 + export 계약 (export 는 MicrosoftTeamsPlugin 하나여야 한다)"
export NPM_ROOT_G="$(npm root -g)"
node --input-type=module -e '
const root = process.env.NPM_ROOT_G;
const m = await import(`file://${root}/opencode-ms-teams/dist/index.js`);
const keys = Object.keys(m);
console.log("  exports:", JSON.stringify(keys));
if (keys.length !== 1 || keys[0] !== "MicrosoftTeamsPlugin") {
  console.error("  FAIL: export 계약 위반 (불변식 2)"); process.exit(1);
}
if (typeof m.MicrosoftTeamsPlugin !== "function") {
  console.error("  FAIL: MicrosoftTeamsPlugin 이 함수가 아님"); process.exit(1);
}
'

step "doctor 진단"
# doctor 는 tokens/log 미생성에도 exit 1 을 반환한다 (fresh install 에서는 정상 상태).
# 따라서 exit code 가 아니라 설치 관련 4개 체크가 통과했는지로 판정한다.
set +e
DOC="$(opencode-ms-teams doctor 2>&1)"; DOC_RC=$?
set -e
echo "$DOC" | sed 's/^/  | /'
for check in config-readable plugin-registered cache-link settings-present; do
  echo "$DOC" | grep -q "✓ $check" || { echo "  FAIL: $check 미통과"; exit 1; }
done
echo "  설치 관련 4개 체크 통과 (doctor exit=$DOC_RC — tokens/log 미생성으로 인한 값, fresh install 에서 정상)"

step "XDG_CONFIG_HOME 오버라이드 준수"
XDG_DIR=/tmp/xdg-config
mkdir -p "$XDG_DIR/opencode"
printf '{}\n' > "$XDG_DIR/opencode/opencode.json"
XDG_CONFIG_HOME="$XDG_DIR" opencode-ms-teams install > /dev/null
grep -q "opencode-ms-teams" "$XDG_DIR/opencode/opencode.json" || { echo "  FAIL: XDG_CONFIG_HOME 무시됨"; exit 1; }
echo "  XDG_CONFIG_HOME 경로에 등록 확인"

step "HOME 미설정 + /etc/passwd 항목 없는 uid 에서 크래시하지 않는가"
# 리눅스 전용 실패 모드: HOME 이 없으면 Node 는 getpwuid 로 폴백하는데,
# 컨테이너/systemd 처럼 uid 가 passwd 에 없으면 os.homedir() 가 ENOENT 로 throw 한다.
set +e
OUT="$(env -u HOME -u XDG_CONFIG_HOME setpriv --reuid=61234 --regid=61234 --clear-groups opencode-ms-teams doctor 2>&1)"
RC=$?
set -e
echo "$OUT" | sed 's/^/  | /'
# 원인 불명의 libuv 오류가 그대로 새어나오면 실패로 본다.
if echo "$OUT" | grep -qiE "uv_os_homedir|uv_os_get_passwd"; then
  echo "  FAIL: libuv 원본 오류가 그대로 노출됐다 — 조치 가능한 메시지로 감싸야 한다"; exit 1
fi
# 대신 어떤 환경변수를 설정하면 되는지 알려주는 메시지가 나와야 한다.
echo "$OUT" | grep -q "홈 디렉토리를 확인할 수 없습니다" || {
  echo "  FAIL: 조치 가능한 안내 메시지가 없다"; exit 1
}
echo "$OUT" | grep -q "OPENCODE_CONFIG_DIR" || { echo "  FAIL: 해결 방법(환경변수) 안내 누락"; exit 1; }
echo "  (exit=$RC — 진단 실패는 정상, 조치 가능한 메시지를 확인)"

step "uninstall 로 원복"
opencode-ms-teams uninstall
[ -L "$LINK" ] && { echo "  FAIL: uninstall 후에도 캐시 링크 잔존"; exit 1; }
echo "  캐시 링크 제거 확인"

printf '\n=== 모든 단계 통과 (%s) ===\n' "$(. /etc/os-release && echo "$PRETTY_NAME")"
