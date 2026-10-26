#!/usr/bin/env python3
"""Recognize a formula image and return a reviewable JSON result.

Backends:
  vlm     OpenAI-compatible vision endpoint (default; SenseNova-compatible)
  mathpix Mathpix /v3/text (specialized formula OCR)
  manual  Copy a reviewed result supplied by --manual-json
"""
from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path

SYSTEM_PROMPT = """You are a mathematical formula OCR engine.
Transcribe visible mathematical notation exactly. Never solve, simplify, or
rewrite the formula. Preserve every subscript, superscript, fraction,
integral, sum, matrix, cases block, Greek letter, accent, and delimiter.
Use standard LaTeX commands that MathType accepts. If several related
equations appear together, join them with standard LaTeX line breaks inside
an aligned or array environment. Return JSON only, with no Markdown fences.
"""

USER_PROMPT = """Inspect the image. Decide whether its primary content is a
mathematical formula/equation rather than prose, a chart, a table, a diagram,
or a photograph of an object. If it is a formula, transcribe it to LaTeX.
Return exactly this JSON schema:
{
  "is_formula": true,
  "latex": "standard LaTeX without surrounding dollar signs",
  "display": true,
  "confidence": 0.0,
  "notes": "short uncertainty notes, or empty string"
}
If it is not a formula, return:
{
  "is_formula": false,
  "latex": "",
  "display": false,
  "confidence": 0.0,
  "notes": "why it was rejected"
}
Do not include any other keys or text.
"""


def env_first(*names: str, default: str = "") -> str:
    for name in names:
        value = os.environ.get(name, "").strip()
        if value:
            return value
    return default


def image_data_uri(path: Path) -> str:
    mime, _ = mimetypes.guess_type(path.name)
    if not mime:
        mime = "image/png"
    payload = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{payload}"


def chat_endpoint(base_url: str) -> str:
    base = base_url.rstrip("/")
    if base.endswith("/v1"):
        return base + "/chat/completions"
    return base + "/v1/chat/completions"


def post_json(url: str, payload: dict, headers: dict[str, str], timeout: int) -> dict:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            text = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach {url}: {exc.reason}") from exc
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"Endpoint returned non-JSON data: {text[:500]}") from exc


def extract_json_object(text: str) -> dict:
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.I)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        value = json.loads(cleaned)
        if isinstance(value, dict):
            return value
    except json.JSONDecodeError:
        pass
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start >= 0 and end > start:
        value = json.loads(cleaned[start : end + 1])
        if isinstance(value, dict):
            return value
    raise RuntimeError(f"Could not parse JSON from model output: {text[:800]}")


def normalize_result(raw: dict, backend: str) -> dict:
    latex = str(raw.get("latex") or raw.get("latex_styled") or "").strip()
    latex = re.sub(r"^\s*\$\$?", "", latex)
    latex = re.sub(r"\$\$?\s*$", "", latex).strip()
    is_formula = raw.get("is_formula")
    if is_formula is None:
        is_formula = bool(latex)
    try:
        confidence = float(raw.get("confidence", 0.0))
    except (TypeError, ValueError):
        confidence = 0.0
    confidence = max(0.0, min(1.0, confidence))
    if is_formula and not latex:
        is_formula = False
    return {
        "status": "ok",
        "backend": backend,
        "is_formula": bool(is_formula),
        "latex": latex,
        "display": bool(raw.get("display", True)),
        "confidence": confidence,
        "notes": str(raw.get("notes", "")),
    }


def run_vlm(path: Path, args: argparse.Namespace) -> dict:
    api_key = args.api_key or env_first(
        "FORMULA_OCR_API_KEY", "SN_VISION_API_KEY", "SN_CHAT_API_KEY", "SN_API_KEY"
    )
    base_url = args.base_url or env_first(
        "FORMULA_OCR_BASE_URL",
        "SN_VISION_BASE_URL",
        "SN_CHAT_BASE_URL",
        "SN_BASE_URL",
        default="https://token.sensenova.cn/v1",
    )
    model = args.model or env_first(
        "FORMULA_OCR_MODEL",
        "SN_VISION_MODEL",
        "SN_CHAT_MODEL",
        default="sensenova-6.8-flash-lite",
    )
    if not api_key:
        raise RuntimeError(
            "No vision API key. Set FORMULA_OCR_API_KEY or SN_API_KEY, "
            "or use --backend mathpix / --manual-json."
        )
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": USER_PROMPT},
                    {"type": "image_url", "image_url": {"url": image_data_uri(path)}},
                ],
            },
        ],
    }
    response = post_json(
        chat_endpoint(base_url),
        payload,
        {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        args.timeout,
    )
    choices = response.get("choices") or []
    if not choices:
        raise RuntimeError(f"Vision response has no choices: {json.dumps(response)[:800]}")
    content = choices[0].get("message", {}).get("content", "")
    if isinstance(content, list):
        content = "".join(part.get("text", "") for part in content if isinstance(part, dict))
    result = normalize_result(extract_json_object(str(content)), "vlm")
    result["model"] = model
    result["base_url"] = base_url
    return result


def run_mathpix(path: Path, args: argparse.Namespace) -> dict:
    app_id = args.mathpix_app_id or env_first("MATHPIX_APP_ID")
    app_key = args.mathpix_app_key or env_first("MATHPIX_APP_KEY")
    if not app_id or not app_key:
        raise RuntimeError("Mathpix requires MATHPIX_APP_ID and MATHPIX_APP_KEY.")
    payload = {
        "src": image_data_uri(path),
        "formats": ["latex_styled", "text"],
        "data_options": {"include_latex": True},
    }
    response = post_json(
        "https://api.mathpix.com/v3/text",
        payload,
        {"app_id": app_id, "app_key": app_key, "Content-Type": "application/json"},
        args.timeout,
    )
    latex = response.get("latex_styled") or response.get("latex") or ""
    raw = {
        "is_formula": bool(latex),
        "latex": latex,
        "display": True,
        "confidence": response.get("confidence", 0.9),
        "notes": str(response.get("error") or ""),
    }
    result = normalize_result(raw, "mathpix")
    result["model"] = "mathpix-v3"
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Recognize a formula image to LaTeX JSON")
    parser.add_argument("--image", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--backend", choices=["vlm", "mathpix"], default="vlm")
    parser.add_argument("--api-key", default="")
    parser.add_argument("--base-url", default="")
    parser.add_argument("--model", default="")
    parser.add_argument("--mathpix-app-id", default="")
    parser.add_argument("--mathpix-app-key", default="")
    parser.add_argument("--manual-json", default="", help="Copy this reviewed result instead of OCR")
    parser.add_argument("--timeout", type=int, default=120)
    args = parser.parse_args()

    image = Path(args.image).resolve()
    output = Path(args.output).resolve()
    if not image.exists():
        raise FileNotFoundError(f"Image not found: {image}")
    output.parent.mkdir(parents=True, exist_ok=True)

    if args.manual_json:
        result = normalize_result(extract_json_object(Path(args.manual_json).read_text(encoding="utf-8-sig")), "manual")
    elif args.backend == "mathpix":
        result = run_mathpix(image, args)
    else:
        result = run_vlm(image, args)
    result["image"] = str(image)

    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        sys.exit(1)
