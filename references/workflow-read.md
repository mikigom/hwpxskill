# 읽기/텍스트 추출 워크플로우

```bash
source "$VENV"

# 순수 텍스트
python3 "$SKILL_DIR/scripts/text_extract.py" document.hwpx

# 테이블 포함
python3 "$SKILL_DIR/scripts/text_extract.py" document.hwpx --include-tables

# 마크다운 형식
python3 "$SKILL_DIR/scripts/text_extract.py" document.hwpx --format markdown
```

## Python API

```python
from hwpx import TextExtractor
with TextExtractor("document.hwpx") as ext:
    text = ext.extract_text(include_nested=True, object_behavior="nested")
    print(text)
```
