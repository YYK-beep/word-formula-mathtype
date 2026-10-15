---
name: word-omml-mathtype
description: "Convert existing Word OMML equations in .docx files into MathType OLE objects on Windows. Use when a Word document already contains Word 2007+ equations and needs genuine MathType objects, including display and inline formulas, section-scoped conversion, cleanup of residual dollar delimiters, and verification. 将 Word .docx 中已有的 OMML 公式转换为 MathType 对象。"
---

# Word OMML to MathType

## What this skill does

This skill converts existing **Word OMML equations** (`m:oMath`) inside a `.docx`
into **MathType OLE objects** (`Equation.DSMT4`). It preserves the surrounding
document, works on display and inline equations, and can limit conversion to a
specific section.

Use it when the document already contains Word equations and the requirement is
specifically MathType, not Word’s native equation format. For plain LaTeX text
that has not yet become an equation, use `word-latex-mathtype` instead.

## Requirements

- Windows
- Microsoft Word desktop
- MathType 7 with the Word add-in `MathType Commands 2016.dotm`
- Python 3
- `pypandoc-binary` or `pandoc` on `PATH`

Run the environment check first:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

## Workflow

1. Work on a copy; never overwrite the source in place.
2. Run the conversion script.
3. Run the verification script.
4. Open the result in Word and inspect representative display and inline formulas.

### Convert a section

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_omml_to_mathtype.ps1 `
  -SourcePath "C:\path\input.docx" `
  -TargetPath "C:\path\input_mathtype.docx" `
  -SectionHeading "四、课题研究内容和工作方案" `
  -EndHeading "五、基础条件和优势"
```

### Convert the whole document

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_omml_to_mathtype.ps1 `
  -SourcePath "C:\path\input.docx" `
  -TargetPath "C:\path\input_mathtype.docx"
```

### Test a small number first

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_omml_to_mathtype.ps1 `
  -SourcePath "C:\path\input.docx" `
  -TargetPath "C:\path\probe.docx" `
  -SectionHeading "四、课题研究内容和工作方案" `
  -EndHeading "五、基础条件和优势" `
  -MaxFormulas 2
```

### Verify the result

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify_mathtype.ps1 `
  -Path "C:\path\input_mathtype.docx" `
  -SectionHeading "四、课题研究内容和工作方案" `
  -EndHeading "五、基础条件和优势"
```

Expected verification values for a completed section:

- `omathsInRange = 0`
- `mathTypeShapes > 0`
- `remainingDollarCharacters = 0`

## How it works

1. `extract_math_map.py` asks pandoc for the document AST and markdown.
2. Math nodes are extracted in document order and paired with Word `OMaths`.
3. Each OMath is deleted from the target copy and replaced with its LaTeX form.
4. MathType's `MTCommand_OnTexToggle` converts the LaTeX text into an
   `Equation.DSMT4` OLE object.
5. Display formulas are centered with the document’s existing tab stops.
6. Residual `$` delimiters around MathType objects are removed.
7. Single alphabetic tokens that fail conversion, such as `Ber` or `KL`, are
   retried as `\mathrm{...}`.
8. Progress is saved every five converted formulas, so a long document can be
   inspected if interrupted.

## Important constraints

- Keep the target document closed in Word before overwriting it.
- Do not run two conversions against the same target file at once.
- Conversion order is from the last OMath to the first; this keeps earlier
  ranges stable while editing.
- The script requires the document’s OMath order and pandoc’s math-node order
  to match. This is true for normal Word documents but should be checked on
  unusually complex documents.
- MathType is required for actual MathType objects. If MathType is missing, the
  script stops rather than silently producing Word equations.
- For very large documents, run a small `-MaxFormulas` probe first.
