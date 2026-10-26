---
name: word-formula-mathtype
description: "Convert OMML, LaTeX text, TeX/Markdown/HTML/MathML files, and formula screenshots or photos into genuine MathType OLE objects in Windows Word for mathematical modeling papers, theses, reviews, course designs, and other Word writing. Includes mandatory preflight, reviewable OCR mapping, delimiter cleanup, and verification. 将数学建模、毕业论文、综述论文、课程设计等 Word 写作中的多种公式形式统一转换为 MathType，并在开始前自动检查依赖环境。
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

## Use cases

Use this skill for Word-based mathematical modeling papers, undergraduate and
graduate theses, review articles, journal/conference manuscripts, course
designs, experiment reports, and project proposals where formulas come from a
mixture of Word equations, LaTeX, TeX files, screenshots, or photos.

## Support conditions

| Component | Requirement | Tested |
| --- | --- | --- |
| Windows | Windows 10/11 | Windows 11, 10.0.22631 |
| PowerShell | 5.1+ | 5.1.22621.6133 |
| Microsoft Word | Desktop Word 2016+ | Office 2021, 16.0.20527.20088 |
| MathType | 7.x; `Equation.DSMT4` registered | MathType 7.12.3 |
| MathType Word add-in | `MathType Commands 2016.dotm` loaded | Present |
| Python | 3.10+ | 3.12.3 |
| Pandoc | 3.x, installed directly or via `pypandoc-binary` | 3.9 |
| Image OCR | Optional: VLM API, Mathpix, or Manual OCR map | SenseNova 6.8 Flash-Lite |

## Mandatory preflight before work

Before any conversion, run and verify:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1 -Json
```

The unified converter and the legacy OMML converter run this preflight
automatically before editing or copying the output. The preflight checks
PowerShell, Word, MathType registration/version, the MathType Word add-in,
Python, Pandoc, and OCR configuration. A failed required check exits non-zero
and the conversion stops.

If the preflight fails, report the missing dependency and stop. Do not silently
fall back to Word-native equations. `-SkipPreflight` is only for users who have
already completed and recorded a manual environment check.

For an image workflow that requires an API, preflight must also pass:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1 `
  -RequireOcr -OcrBackend Vlm
```

Use `-OcrBackend Mathpix` for Mathpix or `-OcrBackend Manual` with a reviewed
OCR map.

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

0. Run the mandatory preflight and stop if any required check fails.
1. Work on a copy; never overwrite the source DOCX.
2. For a document with many image formulas, run `-DryRun` first and review the OCR map.
3. Run the conversion.
4. Run verification.
5. Open representative inline, display, matrix, cases, and image-derived formulas in Word.
6. Review the conversion log and report any failed or low-confidence items.

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
| `-SkipPreflight` | Skip automatic preflight only after a manual environment check |

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

## Test evidence

A reproducible 12-formula test is included under
`examples/formula-speed-test/`: input DOCX, source text, output MathType DOCX,
output PDF, conversion log, benchmark report, and page previews. The tested run
converted 12/12 formulas with 12 MathType objects, zero remaining OMML, and no
remaining formula delimiters. End-to-end time was 23.881 seconds including mandatory
preflight on the documented test machine; the per-formula average was 0.700 s.

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

