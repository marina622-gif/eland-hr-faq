#!/bin/bash
# update.sh — docs 변경사항 감지 → FAQ 시드 생성 → GitHub push → Render 재배포 → 로컬 재시작
set -e

cd "$(dirname "$0")"

REMOTE="origin"
RENDER_HOOK="https://api.render.com/deploy/srv-d8bgndm7r5hc73fql280?key=mCyIJu2YMbI"

echo "======================================"
echo "  이랜드건설 HR FAQ 업데이트 스크립트"
echo "======================================"

# 1. docs 변경사항 확인
echo ""
echo "→ [1/4] docs 변경사항 확인 중..."
git add docs/
if git diff --cached --quiet -- docs/; then
    echo "  변경된 파일 없음. 종료합니다."
    git restore --staged docs/ 2>/dev/null || true
    exit 0
fi
git diff --cached --name-status -- docs/
git restore --staged docs/ 2>/dev/null || true

# 2. FAQ 시드 자동 생성
echo ""
echo "→ [2/4] FAQ 시드 자동 생성 중..."
python make_seed.py

# 3. GitHub push
echo ""
echo "→ [3/4] GitHub push 중..."
git add docs/
if git diff --cached --quiet; then
    echo "  커밋할 변경사항 없음."
else
    TIMESTAMP=$(date '+%Y-%m-%d %H:%M')
    git commit -m "docs 업데이트 ($TIMESTAMP)"
    git push "$REMOTE" master
fi

# 4. Render 재배포 트리거
echo ""
echo "→ [4/4] Render 재배포 트리거 중..."
curl -s -X POST "$RENDER_HOOK" > /dev/null
echo "  Render 재배포 요청 완료 (1~2분 후 반영)"

# 5. 로컬 앱 재시작
echo ""
echo "→ 로컬 앱 재시작 중..."
taskkill //F //IM python.exe 2>/dev/null || true
sleep 2
python app.py &
sleep 3
echo "  로컬 서버 재시작 완료 → http://localhost:5000"

echo ""
echo "======================================"
echo "  ✅ 모든 업데이트 완료!"
echo "  로컬: http://localhost:5000"
echo "  외부: https://eland-hr-faq.onrender.com"
echo "======================================"
