#!/bin/bash
# update.sh — docs 폴더 변경사항을 감지해서 자동으로 처리 후 GitHub push
set -e

cd "$(dirname "$0")"

REMOTE="origin"

echo "======================================"
echo "  이랜드건설 HR FAQ 업데이트 스크립트"
echo "======================================"

# 1. docs 변경사항 확인
echo ""
echo "→ [1/3] docs 변경사항 확인 중..."
git add docs/
if git diff --cached --quiet -- docs/; then
    echo "  변경된 파일 없음. 종료합니다."
    git restore --staged docs/ 2>/dev/null || true
    exit 0
fi
git diff --cached --name-status -- docs/
git restore --staged docs/ 2>/dev/null || true

# 2. 새 파일에 대해 FAQ 시드 자동 생성
echo ""
echo "→ [2/3] FAQ 시드 자동 생성 중..."
python make_seed.py

# 3. git add → commit → push
echo ""
echo "→ [3/3] GitHub push 중..."
git add docs/

if git diff --cached --quiet; then
    echo "  커밋할 변경사항 없음."
    exit 0
fi

TIMESTAMP=$(date '+%Y-%m-%d %H:%M')
git commit -m "docs 업데이트 ($TIMESTAMP)"
git push "$REMOTE" master

echo ""
echo "======================================"
echo "  ✅ 완료! Render 자동 재배포 중..."
echo "  배포 완료까지 약 1~2분 소요됩니다."
echo "======================================"
