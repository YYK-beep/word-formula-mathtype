---
name: word-formula-mathtype
description: "Convert formula sources into genuine MathType OLE objects in Word on Windows. Supports existing OMML equations, LaTeX text in DOCX, TeX/Markdown/HTML/MathML files, and formula screenshots or photos via OCR, with section scoping, dry-run mapping, delimiter cleanup, and verification. 将 Word OMML、TeX 文本、TeX 文件、公式截图/照片统一转换为 Word 中的 MathType 公式对象。"
---

# Word Formula to MathType

## What this skill does

This skill converts formulas from multiple representations into genuine
**MathType OLE objects** (`Equation.DSMT4`) inside Microsoft Word:

- Existing Word equations in **OMML** (`m:oMath`)
- **LaTeX text** embedded in a DOCX: `$...$`, `$$...$$`, `\(...\)`, `\[...\]`
- Standalone **`.tex`**, Markdown, HTML, or MathML files
- **Formula screenshots or photos** embedded in a DOCX
- A standalone **image file** containing one formula

Formula images are recognized to LaTeX first, producing a reviewable OCR map,
then inserted through MathType so the final result remains a real editable
MathType object, not a picture or a Word-native equation.

The skill never silently falls back to OMML when MathType is requested.

## Requirements

- Windows
- Microsoft Word desktop
- MathType 7 with `MathType Commands 2016.dotm` loaded in Word
- Python 3.10+
- `pypandoc-binary` or pandoc on `PATH`
- For image OCR, one of:
  - An OpenAI-compatible vision endpoint
  - A Mathpix account
  - A reviewed manual OCR map

Run the environment check first:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

### Vision OCR configuration

For the default VLM backend, set any compatible provider using:

```powershell
$env:FORMULA_OCR_API_KEY = "..."
$env:FORMULA_OCR_BASE_URL = "https://token.sensenova.cn/v1"
$env:FORMULA_OCR_MODEL = "sensenova-6.8-flash-lite"
```

These aliases are also accepted:

- Key: `SN_VISION_API_KEY`, `SN_CHAT_API_KEY`, `SN_API_KEY`
- URL: `SN_VISION_BASE_URL`, `SN_CHAT_BASE_URL`, `SN_BASE_URL`
- Model: `SN_VISION_MODEL`, `SN_CHAT_MODEL`

For Mathpix:

```powershell
$env:MATHPIX_APP_ID = "..."
$env:MATHPIX_APP_KEY = "..."
```

## Recommended workflow

1. Work on a copy; never overwrite the source DOCX.
2. For a document with many image formulas, run `-DryRun` first and review the OCR map.
3. Run the conversion.
4. Run verification.
5. Open representative inline, display, matrix, cases, and image-derived formulas in Word.

## Quick start

### Convert everything found in a Word document

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -Mode Auto `
  -ImagePolicy Auto
```

`Auto` processes OMML and delimited LaTeX, and examines likely formula images.
Ordinary charts are sent to the OCR classifier but are not replaced when the
classifier says they are not formulas.

### Inspect image OCR before changing the document

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -Mode Images `
  -ImagePolicy Auto `
  -DryRun
```

The `.dryrun.json` log contains each image number, recognized LaTeX,
confidence, and notes.

### Convert a section

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\section_mathtype.docx" `
  -SectionHeading "四、课题研究内容和工作方案" `
  -EndHeading "五、基础条件和优势"
```

### Convert a standalone formula photo

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\formula_photo.png" `
  -TargetPath "D:\docs\formula_photo_mathtype.docx" `
  -OcrBackend Vlm
```

### Convert a TeX file

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\paper.tex" `
  -TargetPath "D:\docs\paper_mathtype.docx"
```

The TeX file is imported to DOCX with Pandoc first; resulting OMML equations
are then converted to MathType.

### Review images manually, then convert

`-DryRun` produces a JSON map. Convert or correct it into this shape, keyed by
Word image number:

```json
{
  "1": {
    "is_formula": true,
    "latex": "\\frac{\\partial u}{\\partial t}=\\alpha\\nabla^2u+f(x,t)",
    "display": true,
    "confidence": 0.99,
    "notes": "reviewed"
  },
  "2": {
    "is_formula": false,
    "latex": "",
    "display": false,
    "confidence": 1.0,
    "notes": "ordinary chart"
  }
}
```

Then run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_any_formula_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -Mode Images `
  -OcrBackend Manual `
  -ManualOcrMapPath "D:\docs\reviewed_ocr.json"
```

### Verify

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify_mathtype.ps1 `
  -Path "D:\docs\input_mathtype.docx"
```

Expected values for a completed formula conversion:

- `omathsInRange = 0`
- `mathTypeShapes > 0`
- `remainingFormulaDelimiters = 0`
- `remainingRawLatexParagraphs = 0`

`remainingDollarCharacters` may be non-zero when the document intentionally
contains currency such as `$5`. Verification therefore distinguishes raw dollar
characters from formula delimiters.

## Important parameters

| Parameter | Meaning |
| --- | --- |
| `-Mode Auto` | OMML + delimited LaTeX + formula images |
| `-Mode Omml` | Existing Word OMML only |
| `-Mode Latex` | LaTeX text only |
| `-Mode Images` | Formula images only |
| `-Mode All` | All supported sources |
| `-LatexMode Delimited` | Convert only `$...$`, `$$...$$`, `\(...\)`, `\[...\]` |
| `-LatexMode Heuristic` | Also convert high-confidence standalone raw LaTeX paragraphs |
| `-ImagePolicy Auto` | Formula-alt-text, formula-labelled, or single-image paragraphs; OCR classifies before replacement |
| `-ImagePolicy All` | OCR every non-MathType image |
| `-ImagePolicy AltText` | Only image alt text/title that names a formula |
| `-ImageNumbers "1,3"` | Explicit Word inline-shape numbers |
| `-DryRun` | Recognize and plan without modifying the document |
| `-KeepWorkFiles` | Keep extracted images and OCR JSON for inspection |

## Supported source files

| Type | Handling |
| --- | --- |
| `.docx` | OMML, LaTeX text, and embedded images |
| `.tex` | Pandoc import, then OMML conversion |
| `.md`, `.markdown`, `.txt` | Pandoc Markdown import, then OMML conversion |
| `.html`, `.htm` | Pandoc HTML import, then OMML conversion |
| `.mml`, `.xml` | Pandoc MathML import, then OMML conversion |
| `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif`, `.tiff`, `.gif`, `.webp` | OCR, then MathType insertion |

## How it works

1. The source is copied or imported into a DOCX; the original is not edited.
2. OMML nodes are paired with Pandoc's ordered LaTeX extraction.
3. Delimited LaTeX is found directly in Word paragraph text.
4. Formula images are exported from Word, classified and transcribed by OCR.
5. Each accepted formula is replaced by a LaTeX wrapper and converted through
   MathType's `MTCommand_OnTexToggle`.
6. Pandoc cases/matrix output is normalized to a MathType-friendly `cases`
   environment, and dangerous trailing control characters are removed.
7. Delimiter cleanup runs only after all OLE objects have stabilized.
8. A JSON log records every conversion and OCR decision.

## Main scripts

| Script | Purpose |
| --- | --- |
| `scripts/setup_environment.ps1` | Check Word, MathType, Python, Pandoc, and OCR configuration |
| `scripts/convert_any_formula_to_mathtype.ps1` | Unified OMML/LaTeX/image/file converter |
| `scripts/ocr_formula_image.py` | VLM or Mathpix formula OCR to reviewed JSON |
| `scripts/convert_omml_to_mathtype.ps1` | Backward-compatible OMML-only converter |
| `scripts/extract_math_map.py` | Extract ordered LaTeX from OMML using Pandoc |
| `scripts/verify_mathtype.ps1` | Verify no formula sources remain and MathType objects exist |

## Constraints and known limits

- Windows, Word desktop, and MathType 7 are required for real MathType objects.
- The target document must be closed before overwriting.
- Two conversions must not target the same file at the same time.
- OCR accuracy depends on image quality and the selected backend. Always review
  `-DryRun` output for handwritten or ambiguous formulas.
- Formula detection in images is intentionally conservative in `Auto` policy.
  Use `-ImagePolicy All` for unlabelled images embedded in mixed prose.
- PDF pages, legacy Equation Editor 3.0 objects, and vector-only formulas in
  scanned pages are not automatically decomposed. Crop each formula to an image
  or supply a reviewed LaTeX map first.
- `\begin{aligned}` inside an outer cases expression is not reliably supported
  by MathType's TeX toggle. The converter normalizes the common Pandoc cases
  form automatically.
- Complex malformed/nested formulas may still require manual review.

