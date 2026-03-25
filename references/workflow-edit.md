# 기존 문서 편집 워크플로우 (unpack → Edit → pack)

```bash
source "$VENV"

# 1. HWPX → 디렉토리 (XML pretty-print)
python3 "$SKILL_DIR/scripts/office/unpack.py" document.hwpx ./unpacked/

# 2. XML 직접 편집 (Claude가 Read/Edit 도구로)
#    본문: ./unpacked/Contents/section0.xml
#    스타일: ./unpacked/Contents/header.xml

# 3. 다시 HWPX로 패키징
python3 "$SKILL_DIR/scripts/office/pack.py" ./unpacked/ edited.hwpx

# 4. 기본 구조 검증
python3 "$SKILL_DIR/scripts/validate.py" edited.hwpx

# 5. 심층 검증 (itemCnt, IDRef, mixed-content, secPr)
python3 "$SKILL_DIR/scripts/validate.py" edited.hwpx --deep

# 6. 편집 전후 무결성 비교 (필수 — 구조 변경·텍스트 손상 탐지)
python3 "$SKILL_DIR/scripts/edit_guard.py" --before document.hwpx --after edited.hwpx
```

## 주의사항

- `edit_guard.py`는 편집 전후 비교를 통해 의도하지 않은 구조 변경(문단/표 추가삭제), 텍스트 중복/소실, mixed-content 손상, ID 참조 불일치를 사전에 탐지한다.
- `edit_guard.py` FAIL 시 결과를 완료로 제출하지 않고, 원인을 수정 후 재패키징한다.
- `--before` 없이 `--after`만 지정하면 단독 구조 정합성 검증만 수행한다.
