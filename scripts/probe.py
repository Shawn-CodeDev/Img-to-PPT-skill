#!/usr/bin/env python3
"""Make a non-destructive, generic editability probe copy of a PPTX.

This changes one native text run, one OMML subscript, and translates one named
shape/group. Relationships stay intact because all package parts except the
selected slide XML are copied byte-for-byte.
"""
from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from lxml import etree as ET
from slide_parts import slide_parts

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"p": P, "a": A, "m": M}


def tag(ns, name): return f"{{{ns}}}{name}"


def named_nodes(root, name):
    return root.xpath(".//p:sp[p:nvSpPr/p:cNvPr[@name=$name]] | "
                      ".//p:grpSp[p:nvGrpSpPr/p:cNvPr[@name=$name]]",
                      namespaces=NS, name=name)


def exactly_one(items, label):
    if len(items) != 1: raise ValueError(f"expected one {label}; found {len(items)}")
    return items[0]


def prepare(source: Path, output: Path, slide_no: int, text_shape: str, old_text: str,
            new_text: str, math_shape: str, old_sub: str, new_sub: str,
            move_shape: str, delta_px: float, reference_width_px: float) -> Path:
    source, output = source.resolve(), output.resolve()
    if source == output: raise ValueError("output must differ from source; source is never modified")
    with zipfile.ZipFile(source) as zin:
        parts = slide_parts(zin)
        if not 1 <= slide_no <= len(parts):
            raise ValueError(f"Display page out of range: {slide_no}")
        slide_part = parts[slide_no-1]
        root = ET.fromstring(zin.read(slide_part))
        presentation = ET.fromstring(zin.read("ppt/presentation.xml"))
        size = presentation.find("p:sldSz", NS)
        if size is None: raise ValueError("presentation.xml has no slide size")
        dx = round(float(delta_px) * int(size.get("cx")) / float(reference_width_px))

        shape = exactly_one(named_nodes(root, text_shape), f"shape named {text_shape!r}")
        runs = [t for t in shape.iter(tag(A, "t")) if t.text == old_text]
        exactly_one(runs, f"text run {old_text!r} inside {text_shape!r}").text = new_text

        math_obj = exactly_one(named_nodes(root, math_shape), f"math host shape named {math_shape!r}")
        subs = [s for s in math_obj.iter(tag(M, "sSub"))
                if "".join(t.text or "" for t in s.findall("./m:sub//m:t", NS)) == old_sub]
        sub = exactly_one(subs, f"OMML subscript {old_sub!r} inside {math_shape!r}")
        sub_runs = sub.findall("./m:sub//m:t", NS)
        exactly_one(sub_runs, f"single subscript run {old_sub!r}").text = new_sub

        moving = exactly_one(named_nodes(root, move_shape), f"shape/group named {move_shape!r}")
        if moving.getparent().tag != tag(P, "spTree"):
            raise ValueError("Select a top-level shape or group for a page-coordinate move.")
        xfrm = moving.find("./p:spPr/a:xfrm", NS) if moving.tag == tag(P, "sp") else moving.find("./p:grpSpPr/a:xfrm", NS)
        off = xfrm.find("./a:off", NS) if xfrm is not None else None
        if off is None: raise ValueError(f"{move_shape!r} has no x offset")
        off.set("x", str(int(off.get("x")) + dx))
        slide_data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
            for info in zin.infolist():
                zout.writestr(info, slide_data if info.filename == slide_part else zin.read(info.filename))
    return output


def compare_rendered(baseline: Path, edited: Path, regions_path: Path, zoom: float = 2.0) -> dict:
    """Compare PDF raster output; region rectangles are [x0,y0,x1,y1] points.

    JSON is {"1": [{"name":"label", "rect":[...], "must_change":true}]}.
    Only pixels inside listed regions may change. Every region must explicitly
    set must_change true or false; true requires at least one differing RGB pixel.
    Unlisted portions of every page, and all pages without regions, must be identical.
    """
    import numpy as np
    import pymupdf

    if zoom <= 0: raise ValueError("zoom must be greater than zero")
    spec = json.loads(regions_path.read_text(encoding="utf-8"))
    if not isinstance(spec, dict): raise ValueError("regions JSON must be an object keyed by 1-based page number")
    # Normalize/validate declarations before rendering.
    regions = {}
    for key, entries in spec.items():
        page_no = int(key)
        if page_no < 1 or not isinstance(entries, list): raise ValueError(f"invalid page entry {key!r}")
        checked = []
        for item in entries:
            if not isinstance(item, dict) or "name" not in item or "rect" not in item or not isinstance(item.get("must_change"), bool):
                raise ValueError(f"each region needs name, rect, and explicit boolean must_change: {item!r}")
            rect = item["rect"]
            if not isinstance(rect, list) or len(rect) != 4 or not all(isinstance(v, (int, float)) for v in rect):
                raise ValueError(f"invalid rect for {item['name']!r}; expected [x0,y0,x1,y1] in PDF points")
            if rect[2] <= rect[0] or rect[3] <= rect[1]: raise ValueError(f"empty/reversed rect: {rect}")
            checked.append({"name": str(item["name"]), "rect": [float(v) for v in rect], "must_change": item["must_change"]})
        regions[page_no] = checked

    report_pages = []
    failures = []
    with pymupdf.open(baseline) as before, pymupdf.open(edited) as after:
        if before.page_count != after.page_count:
            raise ValueError(f"page count differs: baseline={before.page_count}, edited={after.page_count}")
        invalid_pages = sorted(page for page in regions if page > before.page_count)
        if invalid_pages: raise ValueError(f"regions refer to pages outside the PDFs: {invalid_pages}")
        for index, (bp, ep) in enumerate(zip(before, after), 1):
            if abs(bp.rect.width - ep.rect.width) > 0.01 or abs(bp.rect.height - ep.rect.height) > 0.01:
                raise ValueError(f"page {index} dimensions differ: baseline={bp.rect}, edited={ep.rect}")
            b = bp.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csRGB, alpha=False)
            e = ep.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csRGB, alpha=False)
            if (b.width, b.height, b.n) != (e.width, e.height, e.n): raise ValueError(f"render dimensions differ on page {index}")
            ba = np.frombuffer(b.samples, dtype=np.uint8).reshape(b.height, b.width, 3)
            ea = np.frombuffer(e.samples, dtype=np.uint8).reshape(e.height, e.width, 3)
            changed = np.any(ba != ea, axis=2)
            allowed = np.zeros(changed.shape, dtype=bool)
            page_regions = []
            for region in regions.get(index, []):
                x0, y0, x1, y1 = region["rect"]
                # Clamp to page raster; PDF point coordinates use top-left origin in PyMuPDF.
                ix0, iy0 = max(0, int(np.floor(x0 * zoom))), max(0, int(np.floor(y0 * zoom)))
                ix1, iy1 = min(b.width, int(np.ceil(x1 * zoom))), min(b.height, int(np.ceil(y1 * zoom)))
                if ix0 >= ix1 or iy0 >= iy1: raise ValueError(f"region {region['name']!r} falls outside page {index}")
                allowed[iy0:iy1, ix0:ix1] = True
                count = int(np.count_nonzero(changed[iy0:iy1, ix0:ix1]))
                status = "PASS"
                if region["must_change"] and count == 0:
                    status = "FAIL"; failures.append(f"page {index} region {region['name']!r} was required to change but did not")
                page_regions.append({"name": region["name"], "rect": region["rect"], "must_change": region["must_change"], "changed_pixels": count, "status": status})
            outside = int(np.count_nonzero(changed & ~allowed))
            if outside:
                failures.append(f"page {index}: {outside} changed pixels outside allowed regions")
            report_pages.append({"page": index, "changed_pixels": int(np.count_nonzero(changed)),
                                 "outside_allowed_pixels": outside, "regions": page_regions,
                                 "status": "PASS" if outside == 0 and all(r["status"] == "PASS" for r in page_regions) else "FAIL"})
    return {"status": "PASS" if not failures else "FAIL", "zoom": zoom,
            "pixel_rule": "exact RGB inequality; no tolerance", "pages": report_pages, "errors": failures}


def diff_main(argv):
    p = argparse.ArgumentParser(description="Compare baseline/edited PDF pixels with explicitly allowed regions.")
    p.add_argument("baseline", type=Path); p.add_argument("edited", type=Path); p.add_argument("--regions", required=True, type=Path)
    p.add_argument("--zoom", type=float, default=2.0); p.add_argument("--output", type=Path, help="write JSON report (also prints it)")
    a = p.parse_args(argv)
    try:
        result = compare_rendered(a.baseline, a.edited, a.regions, a.zoom)
    except Exception as e:
        p.exit(2, f"render diff failed: {e}\n")
    encoded = json.dumps(result, indent=2)
    if a.output:
        a.output.parent.mkdir(parents=True, exist_ok=True)
        a.output.write_text(encoded + "\n", encoding="utf-8")
    print(encoded)
    return 0 if result["status"] == "PASS" else 1


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("source", type=Path); p.add_argument("output", type=Path)
    p.add_argument("--slide", type=int, required=True, help="1-based slide number")
    p.add_argument("--text-shape", required=True); p.add_argument("--old-text", required=True); p.add_argument("--new-text", required=True)
    p.add_argument("--math-shape", required=True); p.add_argument("--old-subscript", required=True); p.add_argument("--new-subscript", required=True)
    p.add_argument("--move-shape", required=True, help="native shape or group name")
    p.add_argument("--delta-px", type=float, default=8.0); p.add_argument("--reference-width-px", type=float, default=1300.0)
    a = p.parse_args(argv)
    if a.slide < 1 or a.reference_width_px <= 0: p.error("slide must be positive and reference width must be > 0")
    try:
        print(prepare(a.source, a.output, a.slide, a.text_shape, a.old_text, a.new_text,
                      a.math_shape, a.old_subscript, a.new_subscript, a.move_shape,
                      a.delta_px, a.reference_width_px))
    except Exception as e:
        p.exit(2, f"probe failed: {e}\n")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "diff":
        raise SystemExit(diff_main(sys.argv[2:]))
    main()
