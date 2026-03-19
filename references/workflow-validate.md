# 검증 워크플로우

```bash
source "$VENV"
python3 "$SKILL_DIR/scripts/validate.py" document.hwpx
```

검증 항목: ZIP 유효성, 필수 파일 존재, mimetype 내용/위치/압축방식, XML well-formedness
