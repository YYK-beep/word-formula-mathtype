# word-formula-mathtype

A Windows Codex skill for converting formula sources into genuine **MathType
OLE objects** inside Microsoft Word.

## Supported inputs

- Existing Word **OMML** equations
- **LaTeX text** in a DOCX: `$...$`, `$$...$$`, `\(...\)`, `\[...\]`
- `.tex`, Markdown, HTML, and MathML files
- Formula **screenshots or photos** embedded in a DOCX
- A standalone formula image file

Image formulas are recognized to LaTeX, logged for review, and then converted to
MathType. The skill does not silently fall back to Word-native OMML equations.

## Requirements

- Windows with Microsoft Word desktop
- MathType 7 with the Word add-in loaded
- Python 3.10+
- `pypandoc-binary` or Pandoc
- Vision OCR credentials, Mathpix credentials, or a manual OCR map for images

Check the environment:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

## Quick start

### Any formula source inside a DOCX

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -Mode Auto `
  -ImagePolicy Auto
```

### Existing OMML only

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_omml_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx"
```

### Formula photo

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\formula.jpg" `
  -TargetPath "D:\docs\formula_mathtype.docx"
```

### TeX file

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\paper.tex" `
  -TargetPath "D:\docs\paper_mathtype.docx"
```

### Dry-run and review

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -Mode Images `
  -ImagePolicy Auto `
  -DryRun
```

The dry-run JSON records image numbers, recognized LaTeX, confidence, and notes.
Correct or approve the map, then use:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -Mode Images `
  -OcrBackend Manual `
  -ManualOcrMapPath "D:\docs\reviewed_ocr.json"
```

### Verification

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify_mathtype.ps1 `
  -Path "D:\docs\input_mathtype.docx"
```

Expected results:

- `omathsInRange = 0`
- `mathTypeShapes > 0`
- `remainingFormulaDelimiters = 0`
- `remainingRawLatexParagraphs = 0`

Currency such as `$5` is reported separately from formula delimiters.

## OCR configuration

Default OpenAI-compatible VLM backend:

```powershell
$env:FORMULA_OCR_API_KEY = "..."
$env:FORMULA_OCR_BASE_URL = "https://token.sensenova.cn/v1"
$env:FORMULA_OCR_MODEL = "sensenova-6.8-flash-lite"
```

Aliases from the SenseNova environment are also accepted:

- `SN_VISION_API_KEY` / `SN_CHAT_API_KEY` / `SN_API_KEY`
- `SN_VISION_BASE_URL` / `SN_CHAT_BASE_URL` / `SN_BASE_URL`
- `SN_VISION_MODEL` / `SN_CHAT_MODEL`

Mathpix:

```powershell
$env:MATHPIX_APP_ID = "..."
$env:MATHPIX_APP_KEY = "..."
```

Use `-OcrBackend Mathpix` or `-OcrBackend Vlm`.

## Design notes

- The source document is always copied or imported to a new output.
- OMML extraction uses Pandoc's document order.
- Image OCR classifies first and only replaces images accepted as formulas.
- Pandoc cases output is normalized to a MathType-friendly `cases` environment.
- Final delimiter cleanup runs after all OLE objects have stabilized.
- Logs include OCR decisions, normalized LaTeX, timing, and failures.

## Scripts

| Script | Purpose |
| --- | --- |
| `scripts/setup_environment.ps1` | Environment validation |
| `scripts/convert_any_formula_to_mathtype.ps1` | Unified converter |
| `scripts/ocr_formula_image.py` | Image formula OCR |
| `scripts/convert_omml_to_mathtype.ps1` | Legacy OMML-only converter |
| `scripts/extract_math_map.py` | OMML-to-LaTeX mapping |
| `scripts/verify_mathtype.ps1` | Output verification |

## Limitations

- Microsoft Word desktop and MathType 7 are mandatory.
- Handwritten/ambiguous images require review.
- PDF pages and legacy Equation Editor 3.0 objects are not automatically
  decomposed into individual formulas.
- Complex malformed formulas may require manual adjustment in MathType.

## License

MIT

