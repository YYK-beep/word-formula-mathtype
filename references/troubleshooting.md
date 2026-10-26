# Troubleshooting

## MathType add-in is not available

Confirm that `Equation.DSMT4` is registered and that
`MathType Commands 2016.dotm` exists in Word's `STARTUP` folder. Restart Word.
The converter stops rather than silently falling back to OMML.

## `MTCommand_OnTexToggle` times out

- Keep the target document closed before overwriting it.
- Do not run two conversions against the same target at once.
- Run a small probe with `-MaxFormulas 2`.
- Check the `.log.json` for the formula and error message.
- Some malformed LaTeX may need manual correction in MathType.

## OCR model returns HTTP 404

The provider accepted the API key but the configured model name does not exist.
Set a valid model explicitly:

```powershell
$env:FORMULA_OCR_MODEL = "sensenova-6.8-flash-lite"
```

or pass `-OcrModel`. At the time of testing, SenseNova exposed
`sensenova-6.8-flash-lite`, not the older `6.7` default.

## OCR says an ordinary chart is a formula

Use `-DryRun`, inspect the map, and set `is_formula` to `false` for that image.
Then rerun with `-OcrBackend Manual` and `-ManualOcrMapPath`.
`-ImagePolicy AltText` is the most conservative policy.

## An unlabelled formula image is skipped

Use one of:

```powershell
-ImagePolicy All
-ImageNumbers "3,7"
```

The image will still be OCR-classified before replacement.

## `remainingDollarCharacters` is non-zero

Raw dollar signs may be legitimate currency. Check
`remainingFormulaDelimiters`. If it is zero and the document intentionally
contains prices, the extraction is complete.

## The final formula still has a stray delimiter

The converter waits for all MathType objects to stabilize, then performs two
document-wide delimiter-cleanup passes. If a custom document still shows a
delimiter:

1. Inspect the `.log.json` for that formula.
2. Confirm the formula is a real `Equation.DSMT4` object.
3. Check that no dollar is inside the MathType equation itself.
4. Re-run with `-Force`.

## Cases or matrices look wrong

Pandoc sometimes emits:

```latex
\left\{ \begin{matrix} ... \end{matrix} \right.\
```

The converter normalizes common cases output to `\begin{cases}...\end{cases}`
and removes the trailing control backslash. Unsupported `aligned` inside an
outer cases expression may still need manual conversion.

## Word is locked or reports an invalid filename

Close the source and target documents, or use a different `-TargetPath`.

## PDF or legacy Equation Editor input

Automated PDF decomposition and legacy Equation Editor 3.0 conversion are not
provided. Crop each PDF formula to an image or supply reviewed LaTeX, then use
the image/manual workflow.

## Formula contains Chinese text

MathType's TeX toggle can be unstable with CJK inside `\text{...}`. For
publication-critical CJK formulas, review the result in Word; split CJK labels
from the mathematical expression when necessary.
