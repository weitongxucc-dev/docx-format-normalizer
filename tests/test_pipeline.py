#!/usr/bin/env python3
"""
Automated test pipeline for docx-format-normalizer.
Runs formatter + validator on all templates and reports PASS/FAIL summary.

Usage:
    python3 tests/test_pipeline.py [--input sample.docx]
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.parent
SCRIPTS_DIR = SCRIPT_DIR / "scripts"
TEMPLATES_DIR = SCRIPT_DIR / "templates"

TEMPLATES = [
    "government_document.json",
    "bidding_dark_jiangsu.json",
    "bidding_dark_guizhou.json",
    "bidding_dark_jian.json",
    "testing_report.json",
    "court_document.json",
    "academic_paper.json",
    "construction_plan.json",
]

REQUIRED_TEMPLATE_FIELDS = [
    "industry_name", "standard", "page_setup"
]

OPTIONAL_FIELDS = ["body_text", "headings", "header_footer", "page_zones"]

TEMPLATES_WITH_PAGE_ZONES = ["testing_report.json"]

REQUIRED_COVER_FIELDS = ["allowed_elements", "disallowed_elements"]


def find_python():
    venv_python = SCRIPT_DIR / ".venv" / "bin" / "python"
    if venv_python.exists():
        return str(venv_python)
    return sys.executable


def test_template_schema(template_path):
    """Validate template JSON has all required fields."""
    errors = []
    with open(template_path, "r", encoding="utf-8") as f:
        tpl = json.load(f)

    for field in REQUIRED_TEMPLATE_FIELDS:
        if field not in tpl:
            errors.append(f"Missing required field: {field}")

    cover_rules = tpl.get("cover_page_rules", {})
    if cover_rules:
        for field in REQUIRED_COVER_FIELDS:
            if field not in cover_rules:
                errors.append(f"cover_page_rules missing: {field}")
        layout = cover_rules.get("layout", {})
        if cover_rules.get("layout") and not layout.get("elements"):
            errors.append("cover_page_rules.layout.elements is empty")

    font_fallback = tpl.get("font_fallback")
    if font_fallback is not None and not isinstance(font_fallback, dict):
        errors.append("font_fallback must be a dict")

    if os.path.basename(template_path) in TEMPLATES_WITH_PAGE_ZONES:
        if not tpl.get("page_zones"):
            errors.append("testing_report template must have page_zones")

    return errors


def test_formatter(python_path, input_path, template_path, output_path):
    """Run formatter and check it produces valid output."""
    result = subprocess.run(
        [python_path, str(SCRIPTS_DIR / "docx_formatter.py"),
         "--input", input_path,
         "--template", template_path,
         "--output", output_path],
        capture_output=True, text=True, timeout=60
    )
    if result.returncode != 0:
        return False, f"Formatter failed: {result.stderr[:200]}"
    if not os.path.exists(output_path):
        return False, "Output file not created"
    return True, "OK"


def test_validator(python_path, input_path, template_path):
    """Run validator; judge by return code + stable JSON, not terminal text."""
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
        json_path = tmp.name
    result = subprocess.run(
        [python_path, str(SCRIPTS_DIR / "validate.py"),
         "--input", input_path,
         "--template", template_path,
         "--json", json_path],
        capture_output=True, text=True, timeout=30
    )
    summary = None
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            summary = json.load(f)
    except Exception:
        summary = None
    finally:
        if os.path.exists(json_path):
            os.unlink(json_path)

    # 以退出码 + JSON.failed 为准（修复旧版解析 "[FAIL]" 字符串导致的假阳性）
    failed = summary.get("failed", 0) if summary else -1
    if result.returncode != 0 or failed != 0:
        return False, f"FAIL: {failed} check(s) (exit={result.returncode})"
    if summary:
        return True, f"PASS: {summary.get('passed')} FAIL: 0 WARN: {summary.get('warnings')}"
    return True, "OK"


def run_tests(input_docx=None):
    python_path = find_python()
    print(f"Python: {python_path}")
    print(f"Templates: {len(TEMPLATES)}")
    print()

    schema_results = []
    formatter_results = []
    validator_results = []

    for tpl_name in TEMPLATES:
        tpl_path = TEMPLATES_DIR / tpl_name
        print(f"--- {tpl_name} ---")

        errors = test_template_schema(tpl_path)
        schema_results.append((tpl_name, len(errors) == 0, errors))
        if errors:
            for e in errors:
                print(f"  SCHEMA FAIL: {e}")
        else:
            print(f"  SCHEMA: PASS")

        if input_docx:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                output_path = tmp.name

            ok, msg = test_formatter(
                python_path, input_docx, str(tpl_path), output_path)
            formatter_results.append((tpl_name, ok, msg))
            print(f"  FORMATTER: {'PASS' if ok else 'FAIL'} - {msg[:80]}")

            if ok:
                ok, msg = test_validator(
                    python_path, output_path, str(tpl_path))
                validator_results.append((tpl_name, ok, msg))
                print(f"  VALIDATOR: {'PASS' if ok else 'FAIL'} - {msg[:80]}")

            if os.path.exists(output_path):
                os.unlink(output_path)
        print()

    print("=" * 60)
    schema_pass = sum(1 for _, ok, _ in schema_results if ok)
    print(f"Schema: {schema_pass}/{len(schema_results)} PASS")
    if input_docx:
        fmt_pass = sum(1 for _, ok, _ in formatter_results if ok)
        val_pass = sum(1 for _, ok, _ in validator_results if ok)
        print(f"Formatter: {fmt_pass}/{len(formatter_results)} PASS")
        print(f"Validator: {val_pass}/{len(validator_results)} PASS")

    total_pass = schema_pass
    total = len(schema_results)
    if input_docx:
        total_pass += sum(1 for _, ok, _ in formatter_results if ok)
        total_pass += sum(1 for _, ok, _ in validator_results if ok)
        total += len(formatter_results) + len(validator_results)

    print(f"\nTotal: {total_pass}/{total} PASS")
    return total_pass == total


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Test pipeline for docx-format-normalizer")
    parser.add_argument("--input", help="Sample .docx file to test with")
    args = parser.parse_args()

    success = run_tests(args.input)
    sys.exit(0 if success else 1)
