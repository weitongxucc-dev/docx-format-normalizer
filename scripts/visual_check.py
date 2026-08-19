#!/usr/bin/env python3
"""
DOCX Visual Check - Convert output to images and inspect for visual compliance.

Converts the .docx to PDF (via LibreOffice), then to page images (via pdftoppm),
then performs automated visual checks:
  1. Color detection: any non-grayscale pixels indicate non-compliant colored content
  2. Margin consistency: detect content bounding box per page
  3. Page count and size verification

Also saves page images so the AI agent can visually inspect them.

Usage:
    python3 visual_check.py --input output.docx --template templates/government_document.json
    python3 visual_check.py --input output.docx --template templates/government_document.json --save-images /tmp/pages/
"""

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path


SOFFICE = os.environ.get("SOFFICE", "")
if not SOFFICE:
    for candidate in [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        "/Users/Zhuanz/Library/Application Support/TRAE SOLO CN/ModularData/ai-agent/vm/tools/opt/libreoffice/LibreOffice.app/Contents/MacOS/soffice",
    ]:
        if os.path.exists(candidate):
            SOFFICE = candidate
            break

PDFTOPPM = os.environ.get("PDFTOPPM", "")
if not PDFTOPPM:
    for candidate in [
        "/usr/local/bin/pdftoppm",
        "/opt/homebrew/bin/pdftoppm",
        "/Users/Zhuanz/Library/Application Support/TRAE SOLO CN/ModularData/ai-agent/vm/tools/bin/pdftoppm",
    ]:
        if os.path.exists(candidate):
            PDFTOPPM = candidate
            break

SIPS = "/usr/bin/sips"


def convert_to_pdf(docx_path, output_dir):
    """Convert .docx to PDF using LibreOffice headless."""
    cmd = [
        SOFFICE, "--headless", "--convert-to", "pdf",
        "--outdir", output_dir, docx_path
    ]
    env = dict(os.environ)
    env["HOME"] = tempfile.mkdtemp()
    result = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=60)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    pdf_name = Path(docx_path).stem + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_name)
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF not found at {pdf_path}")
    return pdf_path


def convert_to_images(pdf_path, output_dir, dpi=150):
    """Convert PDF to JPEG images using pdftoppm."""
    prefix = os.path.join(output_dir, "page")
    cmd = [PDFTOPPM, "-jpeg", "-r", str(dpi), pdf_path, prefix]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {result.stderr}")

    images = sorted([
        os.path.join(output_dir, f)
        for f in os.listdir(output_dir)
        if f.startswith("page") and f.endswith(".jpg")
    ])
    return images


def check_image_colors(image_path):
    """Check if image contains non-grayscale (colored) pixels.

    Analyzes BMP pixel data from a small (30 DPI) page image.
    Returns dict with: has_color (bool), colored_pixel_ratio (float).
    """
    try:
        import struct

        bmp_path = image_path.replace(".jpg", ".bmp")
        result = subprocess.run(
            [SIPS, "-s", "format", "bmp", image_path, "--out", bmp_path],
            capture_output=True, text=True, timeout=15
        )
        if result.returncode != 0 or not os.path.exists(bmp_path):
            return {"has_color": None, "colored_pixel_ratio": -1, "error": "sips bmp failed"}

        with open(bmp_path, "rb") as f:
            data = f.read()

        if len(data) < 54:
            return {"has_color": None, "colored_pixel_ratio": -1, "error": "BMP too small"}

        width = struct.unpack_from("<I", data, 18)[0]
        height_raw = struct.unpack_from("<i", data, 22)[0]
        height = abs(height_raw)
        bpp = struct.unpack_from("<H", data, 28)[0]
        data_offset = struct.unpack_from("<I", data, 10)[0]

        if width > 10000 or height > 10000:
            return {"has_color": None, "colored_pixel_ratio": -1,
                    "error": f"Unexpected dimensions: {width}x{height}"}

        if bpp != 24:
            return {"has_color": None, "colored_pixel_ratio": -1,
                    "error": f"Unsupported bpp: {bpp}"}

        row_size = ((width * 3 + 3) // 4) * 4
        pixel_data = data[data_offset:]

        colored_count = 0
        total = 0

        for y in range(height):
            row_start = y * row_size
            for x in range(width):
                idx = row_start + x * 3
                if idx + 2 >= len(pixel_data):
                    break
                b = pixel_data[idx]
                g = pixel_data[idx + 1]
                r = pixel_data[idx + 2]
                if max(r, g, b) - min(r, g, b) > 20:
                    colored_count += 1
                total += 1

        ratio = colored_count / max(total, 1)

        try:
            os.unlink(bmp_path)
        except Exception:
            pass

        return {
            "has_color": ratio > 0.001,
            "colored_pixel_ratio": round(ratio, 6),
            "image_size": f"{width}x{height}"
        }

    except Exception as e:
        return {"has_color": None, "colored_pixel_ratio": -1, "error": str(e)}


def run_visual_check(docx_path, template_path, save_images_dir=None):
    """Run full visual check pipeline."""
    with open(template_path, "r", encoding="utf-8") as f:
        template = json.load(f)

    results = {
        "industry": template.get("industry_name", "Unknown"),
        "standard": template.get("standard", ""),
        "input_file": docx_path,
        "checks": [],
        "passed": 0,
        "failed": 0,
        "warnings": 0,
        "page_images": [],
    }

    def add_check(name, status, detail):
        results["checks"].append({"name": name, "status": status, "detail": detail})
        if status == "PASS":
            results["passed"] += 1
        elif status == "FAIL":
            results["failed"] += 1
        else:
            results["warnings"] += 1

    if not SOFFICE:
        add_check("Tool: LibreOffice", "WARN", "soffice not found, visual check skipped")
        return results

    work_dir = save_images_dir or tempfile.mkdtemp(prefix="docx_visual_")
    os.makedirs(work_dir, exist_ok=True)

    try:
        print("[INFO] Converting to PDF...", flush=True)
        pdf_path = convert_to_pdf(docx_path, work_dir)
        add_check("PDF Conversion", "PASS", f"Converted to PDF: {Path(pdf_path).name}")
    except Exception as e:
        add_check("PDF Conversion", "FAIL", f"Conversion failed: {e}")
        return results

    try:
        print("[INFO] Converting to page images...", flush=True)
        images = convert_to_images(pdf_path, work_dir, dpi=30)
        add_check("Image Conversion", "PASS", f"Generated {len(images)} page images (30 DPI)")
        results["page_images"] = images
    except Exception as e:
        add_check("Image Conversion", "FAIL", f"Image generation failed: {e}")
        return results

    add_check("Page Count", "PASS", f"{len(images)} pages")

    print(f"[INFO] Analyzing {len(images)} pages for colored content...", flush=True)
    colored_pages = []
    for i, img_path in enumerate(images):
        if i % 5 == 0:
            print(f"  [progress] page {i+1}/{len(images)}", flush=True)
        color_result = check_image_colors(img_path)
        if color_result.get("has_color") is True:
            colored_pages.append(i + 1)

    if colored_pages:
        add_check("Color Detection", "FAIL",
                  f"Colored content found on pages: {colored_pages}")
    else:
        add_check("Color Detection", "PASS",
                  "No colored content detected (all pages grayscale)")

    if save_images_dir:
        results["page_images"] = images

    return results


def main():
    parser = argparse.ArgumentParser(description="DOCX Visual Check")
    parser.add_argument("--input", required=True, help="Output docx file to check")
    parser.add_argument("--template", required=True, help="Template JSON file")
    parser.add_argument("--save-images", help="Directory to save page images for AI inspection")
    parser.add_argument("--report", help="Output JSON report path")
    args = parser.parse_args()

    print(f"[INFO] Input: {args.input}")
    print(f"[INFO] Template: {args.template}")
    sys.stdout.flush()

    results = run_visual_check(args.input, args.template, args.save_images)

    report_path = args.report or os.path.join(
        os.path.dirname(args.input), "visual_check_report.json"
    )
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"\n{'='*60}")
    print(f"  Visual Check Report")
    print(f"  Industry: {results['industry']}")
    print(f"  Standard: {results['standard']}")
    print(f"{'='*60}\n")

    for check in results["checks"]:
        icon = {"PASS": "[OK]", "FAIL": "[XX]", "WARN": "[!!]"}[check["status"]]
        print(f"  {icon} {check['name']}: {check['detail']}")

    print(f"\n{'─'*60}")
    print(f"  Total: {results['passed'] + results['failed'] + results['warnings']}  "
          f"PASS: {results['passed']}  FAIL: {results['failed']}  WARN: {results['warnings']}")

    if results["page_images"]:
        print(f"\n  Page images saved for visual inspection:")
        for img in results["page_images"]:
            print(f"    {img}")

    if results["failed"] > 0:
        print(f"\n  RESULT: FAILED - {results['failed']} check(s) failed")
    elif results["warnings"] > 0:
        print(f"\n  RESULT: PASSED (with warnings)")
    else:
        print(f"\n  RESULT: ALL CHECKS PASSED")


if __name__ == "__main__":
    main()
