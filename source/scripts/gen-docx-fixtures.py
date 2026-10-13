#!/usr/bin/env python3
"""Build the differential test fixtures with the Python BwB builder (the reference the TypeScript port must match).

Usage:
  python3 scripts/gen-docx-fixtures.py [--skill SKILL_DIR] [NAME ...]

For each test/fixtures/docx/<NAME>.json (all of them when no NAME is given) the script builds the document the way
build_document.build does with rendering and text fixes switched off: Doc on public/templates/bwb-base.docx, then
compose, then save. The date is fixed at 5 October 2026 (the TypeScript test passes the same date as assets.now), and
the brand assets folder holds the base template's brand images plus public/templates/bwb-cover.jpg as
cover_default.jpg, so both builders embed the same cover photograph.

It writes test/fixtures/docx/expected/<NAME>/:
  word/document.xml, word/numbering.xml, word/footnotes.xml, word/_rels/document.xml.rels, [Content_Types].xml,
  docProps/core.xml (with its build time fixed at 2026-10-05T09:00:00Z)
  manifest.json   every part of the package with its SHA-256 and size, the media names, the icons used and the
                  builder's warnings
and test/fixtures/docx/expected/icon-policy.json with Python's answers for the icon colour rules, name resolution and
suggestions (iconPolicy.test.ts compares them).

Requirements: lxml, Pillow and cairosvg (pip install cairosvg --break-system-packages) for the icon fixtures.
SKILL_DIR is the bwb-document-format skill folder (or set $BWB_SKILL_DIR).
"""
import argparse
import datetime
import glob
import hashlib
import json
import os
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIXTURES = os.path.join(ROOT, "test", "fixtures", "docx")
BASE = os.path.join(ROOT, "public", "templates", "bwb-base.docx")
COVER = os.path.join(ROOT, "public", "templates", "bwb-cover.jpg")
FIXED_DATE = datetime.date(2026, 10, 5)
SAVED_PARTS = ["word/document.xml", "word/numbering.xml", "word/footnotes.xml", "word/_rels/document.xml.rels",
               "[Content_Types].xml", "docProps/core.xml"]
FIXED_UTC = datetime.datetime(2026, 10, 5, 9, 0, 0)


def load_builder(skill):
    sys.path.insert(0, os.path.join(skill, "scripts"))
    import build_document  # noqa: E402  (the skill's compose)
    from bwb_lib import builder, icons, package  # noqa: E402
    original = builder.fmt_date

    class FixedDatetime(datetime.datetime):
        @classmethod
        def utcnow(cls):
            return FIXED_UTC

    # core.xml carries the build time; a fixed one keeps the manifests stable when the fixtures are regenerated
    package.datetime = type("datetime_module", (), {"datetime": FixedDatetime})

    def fixed_fmt_date(d=None):
        return original(d or FIXED_DATE)

    builder.fmt_date = fixed_fmt_date            # used by Doc.cover and Doc.title_block
    build_document.fmt_date = fixed_fmt_date     # used by compose for the default date
    return build_document, builder, icons


def brand_assets(tmp):
    with zipfile.ZipFile(BASE) as z:
        for n in z.namelist():
            if n.startswith("word/media/"):
                with open(os.path.join(tmp, n.rsplit("/", 1)[1]), "wb") as f:
                    f.write(z.read(n))
    shutil.copy(COVER, os.path.join(tmp, "cover_default.jpg"))


def build_one(name, build_document, builder, assets, work):
    src = os.path.join(FIXTURES, name + ".json")
    with open(src, encoding="utf-8") as f:
        content = json.load(f)
    meta = content.get("meta", {})
    doc = builder.Doc(BASE, assets_dir=assets)
    build_document.compose(doc, content, FIXTURES)
    out = os.path.join(work, name + ".docx")
    doc.save(out, title=meta.get("title", ""), subject=meta.get("subject", ""), keywords=meta.get("keywords", ""),
             description=meta.get("description", ""))
    dest = os.path.join(FIXTURES, "expected", name)
    shutil.rmtree(dest, ignore_errors=True)
    os.makedirs(dest)
    parts = {}
    with zipfile.ZipFile(out) as z:
        for n in sorted(z.namelist()):
            data = z.read(n)
            parts[n] = {"sha256": hashlib.sha256(data).hexdigest(), "size": len(data)}
            if n in SAVED_PARTS:
                p = os.path.join(dest, n)
                os.makedirs(os.path.dirname(p), exist_ok=True)
                with open(p, "wb") as f:
                    f.write(data)
    manifest = {
        "source": f"{name}.json",
        "parts": parts,
        "media": sorted(n.rsplit("/", 1)[1] for n in parts if n.startswith("word/media/")),
        "icons": sorted(doc.icons_used),
        "warnings": doc.warnings,
    }
    with open(os.path.join(dest, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=1, ensure_ascii=False)
        f.write("\n")
    return manifest


def policy_reference(icons):
    """Python's answers for the icon rules (colour checks, name resolution, suggestions) for iconPolicy.test.ts."""
    colours = ["", "green", "navy", "charcoal", "white", "lime", "purple", "Grey", "#328840", "BED62F", "black", "#ffffff",
               "  Lime  "]
    checks = [[c, nav, fb] + list(icons.check_colour(c, nav, fb))
              for c in colours for nav in (False, True) for fb in (None, "navy", "white", "lime", "purple")]
    names = ["Leaf", "lucide-leaf", "Lucide:Leaf", "leaf-icon", "bar_chart", "Bar Chart", "alert-triangle", "nope", "",
             "  Droplets  ", "x-octagon", "LUCIDE-ICON"]
    queries = ["no-such-icon", "not-an-icon", "leef", "chart", "alert", "dolar-sign", "x", "user", "arrow-right-circle",
               "barchart", "Bar Chart Icon"]
    ref = {
        "check_colour": checks,
        "resolve": {n: icons.resolve(n) for n in names},
        "suggest3": {q: icons.suggest(q, 3) for q in queries},
        "suggest5": {q: icons.suggest(q) for q in queries},
        "px_for_cm": {str(cm): icons.px_for_cm(cm) for cm in (0.1, 0.4, 0.6, 0.8, 1.0, 1.2, 2.54)},
    }
    os.makedirs(os.path.join(FIXTURES, "expected"), exist_ok=True)
    with open(os.path.join(FIXTURES, "expected", "icon-policy.json"), "w", encoding="utf-8") as f:
        json.dump(ref, f, indent=1, ensure_ascii=False)
        f.write("\n")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("names", nargs="*", help="fixture names (default: every JSON file in test/fixtures/docx)")
    ap.add_argument("--skill", default=os.environ.get("BWB_SKILL_DIR"), required="BWB_SKILL_DIR" not in os.environ,
                    help="the bwb-document-format skill folder")
    a = ap.parse_args()
    build_document, builder, icons = load_builder(a.skill)
    if icons.renderer() != "cairosvg":
        print(f"warning: icons are drawn with {icons.renderer()}, not cairosvg (only icon names are compared)",
              file=sys.stderr)
    names = a.names or sorted(os.path.splitext(os.path.basename(p))[0] for p in glob.glob(os.path.join(FIXTURES, "*.json")))
    work = tempfile.mkdtemp(prefix="bwb_fixtures_")
    assets = os.path.join(work, "assets")
    os.makedirs(assets)
    try:
        brand_assets(assets)
        policy_reference(icons)
        for name in names:
            m = build_one(name, build_document, builder, assets, work)
            print(f"{name}: {len(m['parts'])} parts, {len(m['media'])} media, {len(m['warnings'])} warnings")
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
