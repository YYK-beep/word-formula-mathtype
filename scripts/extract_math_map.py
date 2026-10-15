#!/usr/bin/env python3
"""Extract ordered LaTeX math nodes from a DOCX using pandoc.

The output is a JSON list:
[
  {"index": 1, "latex": "...", "display": true},
  ...
]

`index` is one-based and follows the order of `OMaths` in Word.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import pypandoc
except ImportError:  # pypandoc is optional when pandoc is already on PATH
    pypandoc = None


def resolve_pandoc(explicit: str | None) -> str:
    if explicit:
        return explicit
    if pypandoc is not None:
        try:
            return pypandoc.get_pandoc_path()
        except Exception:
            pass
    found = shutil.which("pandoc")
    if found:
        return found
    raise FileNotFoundError(
        "pandoc was not found. Run scripts/setup_environment.ps1 or install pypandoc-binary."
    )


def run_pandoc(pandoc: str, args: list[str]) -> str:
    proc = subprocess.run(
        [pandoc, *args],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "pandoc conversion failed")
    return proc.stdout


def walk_math_nodes(node, out: list[dict]) -> None:
    if isinstance(node, dict):
        if node.get("t") == "Math":
            payload = node.get("c")
            if isinstance(payload, list) and len(payload) == 2:
                kind = payload[0].get("t")
                latex = payload[1]
                out.append({"kind": kind, "latex": latex})
        for value in node.values():
            walk_math_nodes(value, out)
    elif isinstance(node, list):
        for value in node:
            walk_math_nodes(value, out)


def extract_display_latex(markdown: str) -> list[str]:
    display: list[str] = []
    # A display equation from Word OMML is emitted by pandoc as a standalone
    # line containing one math expression, optionally followed by an equation
    # number such as （4-11） or (4-11).
    line_re = re.compile(
        r"^\s*(\$\$.+?\$\$|\$.+?\$)\s*(?:[（(][^（）()]{0,40}[）)])?\s*$",
        re.S,
    )
    for raw in markdown.splitlines():
        line = raw.strip()
        if not line:
            continue
        match = line_re.match(line)
        if not match:
            continue
        math = match.group(1)
        if math.startswith("$$") and math.endswith("$$"):
            display.append(math[2:-2].strip())
        elif math.startswith("$") and math.endswith("$"):
            display.append(math[1:-1].strip())
    return display


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract ordered LaTeX math nodes from DOCX")
    parser.add_argument("--source", required=True, help="Source .docx path")
    parser.add_argument("--output", required=True, help="Output JSON path")
    parser.add_argument("--pandoc", default=None, help="Explicit pandoc executable path")
    args = parser.parse_args()

    source = Path(args.source).resolve()
    output = Path(args.output).resolve()
    if not source.exists():
        raise FileNotFoundError(f"Source DOCX not found: {source}")

    pandoc = resolve_pandoc(args.pandoc)
    ast_text = run_pandoc(pandoc, ["-f", "docx", "-t", "json", str(source)])
    markdown = run_pandoc(
        pandoc,
        ["-f", "docx", "-t", "markdown", "--wrap=none", str(source)],
    )

    ast = json.loads(ast_text)
    math_nodes: list[dict] = []
    walk_math_nodes(ast, math_nodes)

    display_queue = extract_display_latex(markdown)
    display_ptr = 0
    items: list[dict] = []
    for index, node in enumerate(math_nodes, start=1):
        latex = node["latex"]
        is_display = node["kind"] == "DisplayMath"
        if not is_display and display_ptr < len(display_queue):
            if latex.strip() == display_queue[display_ptr].strip():
                is_display = True
                display_ptr += 1
        items.append({"index": index, "latex": latex, "display": is_display})

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "source": str(source),
        "output": str(output),
        "mathNodes": len(items),
        "displayMarked": sum(1 for item in items if item["display"]),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
