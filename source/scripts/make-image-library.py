"""Builds the BwB image library the add-in offers in Word and PowerPoint.

  python3 scripts/make-image-library.py --website PATH/TO/bwb-website/public

Reads src/brand/imageCatalogue.json, converts each photo from the website (AVIF or WebP) to a JPEG that Word and
PowerPoint can insert (long edge at most 2000 px, camera metadata removed) plus a thumbnail, and writes
public/assets/images/index.json with the sizes. "@cover" stands for the report cover photo in public/templates.
"""
import argparse
import json
import os
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "public" / "assets" / "images"
MAX_EDGE = 2000
THUMB_WIDTH = 360


def save_jpeg(image: Image.Image, path: Path, quality: int) -> None:
    rgb = image.convert("RGB")
    rgb.save(path, "JPEG", quality=quality, optimize=True, progressive=True)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--website", default=os.environ.get("BWB_WEBSITE_DIR"), required="BWB_WEBSITE_DIR" not in os.environ,
                    help="the BwB website's public folder")
    args = ap.parse_args()
    website = Path(args.website)
    catalogue = json.loads((ROOT / "src" / "brand" / "imageCatalogue.json").read_text(encoding="utf-8"))

    (OUT / "thumbs").mkdir(parents=True, exist_ok=True)
    entries = []
    for item in catalogue["images"]:
        if item["file"] == "@cover":
            source = ROOT / "public" / "templates" / "bwb-cover.jpg"
            full_path = "templates/bwb-cover.jpg"
        else:
            source = website / item["file"]
            full_path = f"assets/images/{item['id']}.jpg"
        with Image.open(source) as image:
            image.load()
            if item["file"] != "@cover":
                full = image.copy()
                full.thumbnail((MAX_EDGE, MAX_EDGE), Image.LANCZOS)
                save_jpeg(full, OUT / f"{item['id']}.jpg", 85)
                width, height = full.size
            else:
                width, height = image.size
            thumb = image.copy()
            thumb.thumbnail((THUMB_WIDTH, THUMB_WIDTH * 2), Image.LANCZOS)
            save_jpeg(thumb, OUT / "thumbs" / f"{item['id']}.jpg", 74)
        entries.append({
            "id": item["id"], "title": item["title"], "category": item["category"], "tags": item["tags"],
            "file": full_path, "thumb": f"assets/images/thumbs/{item['id']}.jpg", "width": width, "height": height,
            "bytes": (ROOT / "public" / full_path).stat().st_size,
        })

    index = {"note": catalogue["note"], "categories": catalogue["categories"], "images": entries}
    (OUT / "index.json").write_text(json.dumps(index, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    total = sum(e["bytes"] for e in entries)
    print(f"{len(entries)} images, {total / 1e6:.1f} MB full size, index at {OUT / 'index.json'}")


if __name__ == "__main__":
    main()
