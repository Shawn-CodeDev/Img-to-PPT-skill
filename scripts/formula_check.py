"""Measure horizontal/vertical scale of printed equation glyphs in isolated reference boxes."""
import argparse
import json
import math
from pathlib import Path
import re

from lxml import etree
import pymupdf

NUMBER = re.compile(r"[-+]?(?:\d*\.\d+|\d+\.?\d*)(?:[eE][-+]?\d+)?")
DELIMITERS = frozenset("()[]{}|‖⌊⌋⌈⌉⎡⎢⎣⎤⎥⎦⎛⎜⎝⎞⎟⎠√∫∑")


def glyph_runs(page):
    xml = etree.fromstring(page.get_svg_image(text_as_path=False).encode())
    for text in xml.xpath('//*[local-name()="text"]'):
        transform = text.get("transform", "matrix(1,0,0,1,0,0)")
        if not transform.startswith("matrix("):
            raise ValueError(f"Unsupported PDF text transform {transform}")
        a,b,c,d,e,f = map(float, NUMBER.findall(transform))
        sx = math.hypot(a,b)
        sy = abs(a*d-b*c) / sx if sx else 0
        if not sy:
            continue
        for span in text.xpath('.//*[local-name()="tspan"]'):
            x = NUMBER.findall(span.get("x", ""))
            y = NUMBER.findall(span.get("y", ""))
            chars = "".join(span.itertext()).strip()
            if not x or not y or not chars:
                continue
            x0,y0 = float(x[0]),float(y[0])
            yield a*x0+c*y0+e, b*x0+d*y0+f, chars, sx/sy


def check(pdf, records, tolerance_percent):
    if not records or not math.isfinite(tolerance_percent) or tolerance_percent < 0:
        raise ValueError("Supply at least one equation and a non-negative finite limit.")
    results, pages = [], {}
    with pymupdf.open(pdf) as doc:
        for record in records:
            no, name = record["page"], record["name"]
            if not isinstance(no, int) or not 1 <= no <= len(doc):
                raise ValueError(f"{name}: page number out of range")
            area = pymupdf.Rect(record["box_pt"])
            if area.is_empty or not doc[no-1].rect.contains(area):
                raise ValueError(f"{name}: invalid or out-of-page box")
            if no not in pages:
                pages[no] = list(glyph_runs(doc[no-1]))
            found = [(text, ratio) for x,y,text,ratio in pages[no]
                     if area.contains(pymupdf.Point(x,y))
                     and any(not (c in DELIMITERS or c.isspace()) for c in text)]
            delta = max((abs(ratio-1)*100 for _,ratio in found), default=None)
            results.append({"page": no, "name": name, "glyph_runs": len(found),
                            "printed_text_samples": [text for text,_ in found],
                            "max_deviation_percent": delta,
                            "passed": delta is not None and delta <= tolerance_percent})
    return {"passed": all(r["passed"] for r in results), "tolerance_percent": tolerance_percent, "equations": results}


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("pdf")
    p.add_argument("boxes", type=Path, help="JSON list of {page, name, box_pt:[x0,y0,x1,y1]} in top-left PDF points.")
    p.add_argument("--limit-percent", type=float, required=True)
    args = p.parse_args()
    report = check(args.pdf, json.loads(args.boxes.read_text(encoding="utf-8")), args.limit_percent)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(0 if report["passed"] else 1)
