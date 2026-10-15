# word-omml-mathtype

A Codex skill for converting **existing Word OMML equations** in a `.docx`
into genuine **MathType OLE objects** on Windows.

It handles display equations, inline formulas, section-scoped conversion,
residual `$` cleanup, and verification. It was developed and tested against a
document containing 71 OMML equations.

## Install as a Codex skill

Clone this repository into a skills folder, for example:

```powershell
git clone https://github.com/YYK-beep/word-omml-mathtype.git `
  "$env:USERPROFILE\.codex\skills\word-omml-mathtype"
```

The folder must contain `SKILL.md` at its root. Codex can then discover the
skill by name or description.

## Requirements

- Windows
- Microsoft Word desktop
- MathType 7 with `MathType Commands 2016.dotm` loaded in Word
- Python 3
- `pypandoc-binary` or `pandoc`

Install or verify the runtime dependencies:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

## Quick start

Convert only one section of an application or thesis:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_omml_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\input_mathtype.docx" `
  -SectionHeading "四、课题研究内容和工作方案" `
  -EndHeading "五、基础条件和优势"
```

Run a small probe first:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_omml_to_mathtype.ps1 `
  -SourcePath "D:\docs\input.docx" `
  -TargetPath "D:\docs\probe.docx" `
  -SectionHeading "四、课题研究内容和工作方案" `
  -EndHeading "五、基础条件和优势" `
  -MaxFormulas 2
```

Verify:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify_mathtype.ps1 `
  -Path "D:\docs\input_mathtype.docx" `
  -SectionHeading "四、课题研究内容和工作方案" `
  -EndHeading "五、基础条件和优势"
```

Expected result:

- `omathsInRange : 0`
- `mathTypeShapes : <number of converted formulas>`
- `remainingDollarCharacters : 0`

## Scripts

| Script | Purpose |
| --- | --- |
| `scripts/setup_environment.ps1` | Check Word, MathType, Python and pandoc/pypandoc |
| `scripts/extract_math_map.py` | Extract ordered LaTeX math nodes from a DOCX using pandoc |
| `scripts/convert_omml_to_mathtype.ps1` | Delete OMaths, convert their LaTeX to MathType, clean delimiters, log results |
| `scripts/verify_mathtype.ps1` | Verify that the selected range has no OMML, no `$`, and contains MathType objects |

## Performance

- 100 ms polling instead of 300 ms intervals.
- Standalone alphabetic formulas are pre-wrapped as `\mathrm{...}` instead of
  being attempted once and retried after a timeout.
- The document is saved every 20 formulas, not every 5.
- Delimiter `$` cleanup runs once at the end instead of after every formula.

In a 5-formula probe, total runtime including Word startup was about 12.6 s;
after the first formula, each formula typically completes in 0.4-0.5 s.

## Design notes

- The source is copied first; the original document is never edited in place.
- Conversion processes OMaths from last to first so earlier ranges remain stable.
- Progress is saved every five formulas.
- Single alphabetic formulas such as `Ber` or `KL` are retried with
  `\mathrm{}` because MathType’s TeX toggle may otherwise refuse to create an
  object for a standalone word.
- Display equation numbers remain as normal text; formulas use the existing
  center and right tab stops.

## Limitations

- Windows and Word desktop only.
- MathType 7 is required for real MathType objects.
- Pandoc must be able to read the document’s OMML equations.
- Complex nested or unusually malformed equations may still require manual
  review after conversion.
- The script cannot convert formulas in floating text boxes unless they are
  part of the document story processed by Word.

## License

MIT
