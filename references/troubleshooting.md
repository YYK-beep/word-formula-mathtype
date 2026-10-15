# Troubleshooting

## `MTCommand_OnTexToggle` is not available

MathType’s Word add-in is not loaded. Confirm that
`MathType Commands 2016.dotm` exists in Word’s `STARTUP` folder and restart
Word. The skill stops instead of falling back to OMML because the requested
result is MathType objects.

## Word is locked or reports an invalid file name

Close the target document in Word before running the converter. The converter
always writes to `TargetPath`; do not pass the same path for source and target.

## Conversion is slow

Each formula is converted through MathType’s OLE/TeX path. A document with
dozens of formulas can take several minutes. Progress is saved every five
formulas, and the script logs `elapsedMs` for each formula.

## A standalone formula such as `Ber` or `KL` fails

MathType may not create an equation object for a lone word. The converter
automatically retries single alphabetic tokens as `\mathrm{...}`. If a custom
formula still fails, inspect the `.log.json` file and convert that formula
manually in Word.

## Residual `$` characters remain

The converter removes `$` delimiters adjacent to MathType shapes and converts
failed text back to plain LaTeX. If `$` still appears in the verification
report, open the document and inspect the listed section. Do not use a global
find-and-replace if the document intentionally contains dollar amounts.

## Pandoc cannot read OMML

Install `pypandoc-binary` with the setup script. If pandoc still cannot extract
the equations, use the `word-latex-mathtype` skill when the formulas are plain
LaTeX text rather than Word OMath objects.

## Tab stops after conversion

Display equations normally use the paragraph’s existing center and right tab
stops: the first tab centers the formula, and the second tab right-aligns the
equation number. This is expected and does not indicate a conversion failure.
