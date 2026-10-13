#!/usr/bin/env python3
"""Render .docx files to PDF with LibreOffice and to page images with pdftoppm, for visual checks.

Usage:
  python3 scripts/render-docx.py FILE.docx [FILE.docx ...] --out DIR [--dpi 60] [--format png|jpeg]

For each FILE.docx it writes DIR/FILE.pdf and DIR/FILE-<page>.png, and prints one line per file with the page count.
LibreOffice substitutes fonts (Candara and Aptos are not installed in most sandboxes), so line breaks differ a little
from Word; colours, spacing, tables, panels, pictures, headers, footers and page numbers are faithful.
"""
import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile


def to_pdf(docx, outdir, timeout=240):
    profile = tempfile.mkdtemp(prefix="bwb_lo_")
    try:
        subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf",
                        "--outdir", outdir, docx], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                       timeout=timeout)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError(f"LibreOffice did not produce a PDF for {docx}")
    return pdf


def page_count(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    m = re.search(r"Pages:\s+(\d+)", out)
    return int(m.group(1)) if m else 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--out", required=True, help="folder for the PDF and page images")
    ap.add_argument("--dpi", type=int, default=60)
    ap.add_argument("--format", choices=("png", "jpeg"), default="png")
    a = ap.parse_args()
    for tool in ("soffice", "pdftoppm", "pdfinfo"):
        if not shutil.which(tool):
            sys.exit(f"{tool} is not installed")
    os.makedirs(a.out, exist_ok=True)
    for docx in a.files:
        stem = os.path.splitext(os.path.basename(docx))[0]
        for old in glob.glob(os.path.join(a.out, glob.escape(stem) + "-*." + ("png" if a.format == "png" else "jpg"))):
            os.remove(old)
        pdf = to_pdf(os.path.abspath(docx), os.path.abspath(a.out))
        subprocess.run(["pdftoppm", "-png" if a.format == "png" else "-jpeg", "-r", str(a.dpi), pdf,
                        os.path.join(a.out, stem)], check=True)
        print(f"{docx}: {page_count(pdf)} pages -> {pdf}")


if __name__ == "__main__":
    main()
