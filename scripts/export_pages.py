"""Split a checked vector PDF into individual paper figures; optional explicit crop rectangles."""
import argparse
import json
from pathlib import Path

import pymupdf


def export_pages(source, folder, crops=None):
    source, folder = Path(source).resolve(), Path(folder)
    crops = {} if crops is None else crops
    if not isinstance(crops, dict):
        raise ValueError("Crop specification must be an object keyed by one-based page number.")
    folder.mkdir(parents=True, exist_ok=True)
    with pymupdf.open(source) as original:
        pages = {str(i+1) for i in range(len(original))}
        if not crops.keys() <= pages:
            raise ValueError(f"Unknown pages in crop specification: {sorted(crops.keys()-pages)}")
        bounds = []
        for i, page in enumerate(original, 1):
            if page.get_image_info():
                raise ValueError(f"Page {i} contains a raster image. Review the publication source.")
            crop = None
            if str(i) in crops:
                crop = pymupdf.Rect(crops[str(i)])
                if page.rotation or page.cropbox != page.mediabox or page.rect.top_left != pymupdf.Point(0, 0):
                    raise ValueError(f"Page {i}: specify crops on the unrotated, uncropped slide export.")
                if crop.is_empty or not page.rect.contains(crop):
                    raise ValueError(f"Page {i}: crop must lie within the original page.")
            bounds.append(crop)
        for i, crop in enumerate(bounds):
            path = folder / f"figure-{i+1:02}.pdf"
            if path.resolve() == source:
                raise ValueError("Output would overwrite the input PDF.")
            with pymupdf.open() as single:
                single.insert_pdf(original, from_page=i, to_page=i)
                if crop is not None:
                    single[0].set_cropbox(crop)
                single.save(path, garbage=4, deflate=True)
    return len(bounds)


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("pdf")
    p.add_argument("folder")
    p.add_argument("--crops", type=Path, help='JSON such as {"1":[x0,y0,x1,y1]} in PDF points from top left')
    args = p.parse_args()
    crops = json.loads(args.crops.read_text(encoding="utf-8")) if args.crops else None
    print(f"Exported {export_pages(args.pdf, args.folder, crops)} vector pages.")
