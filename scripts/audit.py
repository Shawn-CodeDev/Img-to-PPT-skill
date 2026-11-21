#!/usr/bin/env python3
"""Portable structural and PDF audit for editable PowerPoint deliveries.

Run: python audit.py deck.pptx deck.pdf [--expect-math]
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from slide_parts import slide_parts

try:
    import pymupdf
except ImportError as exc:  # pragma: no cover - actionable CLI diagnostic
    raise SystemExit("Install PyMuPDF (import name pymupdf) to audit PDF files") from exc

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"p": P, "a": A, "m": M}
BAD = re.compile("[\ufffd\x00-\x08\x0b\x0c\x0e-\x1f]")


def q(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def _visible(node: ET.Element) -> bool:
    for e in node.iter():
        for key, value in e.attrib.items():
            local = key.rsplit("}", 1)[-1].lower()
            v = value.lower().strip()
            if local == "hidden" and v in {"1", "true"}:
                return False
            if local in {"alpha", "alpha_mod"} and v in {"0", "0.0"}:
                return False
    return True


def _hidden_count(root: ET.Element) -> int:
    return sum(1 for e in root.iter() if not _visible(e))


def audit_slide(data: bytes) -> dict:
    root = ET.fromstring(data)
    ids = [e.get("id") for e in root.iter(q(P, "cNvPr"))]
    duplicate_ids = len(ids) != len(set(ids))
    pics = sum(1 for _ in root.iter(q(P, "pic")))
    blips = sum(1 for _ in root.iter(q(A, "blip")))
    texts = []
    for shape in root.iter(q(P, "sp")):
        body = shape.find("p:txBody", NS)
        if body is None or not _visible(shape):
            continue
        text = "".join(t.text or "" for t in body.iter(q(A, "t")))
        if text.strip():
            texts.append(text)
    math_count = sum(1 for e in root.iter(q(M, "oMath")))
    visible_math_count = sum(1 for shape in root.iter(q(P, "sp")) if _visible(shape)
                             for _ in shape.iter(q(M, "oMath")))
    hidden_nodes = _hidden_count(root)
    return {"hidden_slides": int(root.get("show", "1").lower() in {"0", "false"}),
            "duplicate_ids": duplicate_ids, "visible_text_shapes": len(texts),
            "visible_text_chars": sum(map(len, texts)), "omml_count": math_count,
            "visible_omml_count": visible_math_count,
            "picture_shapes": pics, "picture_blips": blips,
            "hidden_elements": hidden_nodes, "bad_glyphs": sum(len(BAD.findall(t)) for t in texts)}


def audit_pptx(path: Path) -> list[dict]:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        records = [audit_slide(z.read(part)) for part in slide_parts(z)]
        media = [n for n in names if n.startswith("ppt/media/") and not n.endswith("/")]
        raster_media = [n for n in media if Path(n).suffix.lower() in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}]
        template_parts = [n for n in names if re.fullmatch(r"ppt/(slideLayouts|slideMasters)/[^/]+\.xml", n)]
        hidden_templates = 0
        for part in template_parts:
            hidden_templates += _hidden_count(ET.fromstring(z.read(part)))
    for rec in records:
        rec["raster_media_parts"] = len(raster_media)
        rec["hidden_template_elements"] = hidden_templates
    return records


def audit_pdf(path: Path) -> list[dict]:
    pages = []
    with pymupdf.open(path) as doc:
        for page in doc:
            text = page.get_text("text") or ""
            fonts = {}
            for row in page.get_fonts(full=True):
                xref = int(row[0]); name = str(row[3] if len(row) > 3 else xref)
                embedded = False
                if xref > 0:
                    try:
                        item = doc.extract_font(xref)
                        embedded = bool(item and len(item) > 3 and item[3])
                    except Exception:
                        pass
                fonts[xref] = {"name": name, "embedded": embedded}
            # get_image_info reports images actually painted on this page, including inline images.
            image_count = len(page.get_image_info(xrefs=True))
            invisible_text = sum(1 for span in page.get_texttrace()
                                 if span.get("type") == 3 or span.get("opacity", 1) <= 0)
            invisible_drawings = 0
            for drawing in page.get_drawings():
                fill, stroke = drawing.get("fill_opacity"), drawing.get("stroke_opacity")
                if (fill is None or fill <= 0) and (stroke is None or stroke <= 0):
                    invisible_drawings += 1
            ocgs = doc.get_ocgs() or {}
            hidden_ocgs = sum(1 for value in ocgs.values() if not value.get("on", True))
            pages.append({"text_chars": len(text.strip()), "bad_glyphs": len(BAD.findall(text)),
                          "raster_images": image_count, "font_count": len(fonts),
                          "embedded_font_count": sum(v["embedded"] for v in fonts.values()),
                          "fonts_embedded": bool(fonts) and all(v["embedded"] for v in fonts.values()),
                          "hidden_content": invisible_text + invisible_drawings + hidden_ocgs,
                          "font_names": sorted({v["name"] for v in fonts.values()}),
                          "width_pt": round(page.rect.width, 2), "height_pt": round(page.rect.height, 2)})
    return pages


def run(pptx: Path, pdf: Path, expect_math: bool = False, allow_empty_text: set[int] | None = None) -> dict:
    slides = audit_pptx(pptx)
    pages = audit_pdf(pdf)
    allow_empty_text = allow_empty_text or set()
    if any(i < 1 or i > len(slides) for i in allow_empty_text):
        raise ValueError("An --allow-empty-text page lies outside the deck.")
    errors = []
    if len(slides) != len(pages):
        errors.append(f"slide/page count mismatch: {len(slides)} slides, {len(pages)} PDF pages")
    if expect_math and not sum(s["visible_omml_count"] for s in slides):
        errors.append("--expect-math specified but no visible OMML equations were found")
    for i, s in enumerate(slides, 1):
        start = len(errors)
        if s["duplicate_ids"]: errors.append(f"slide {i}: duplicate shape IDs")
        if s["visible_text_shapes"] < 1 and s["visible_omml_count"] < 1 and i not in allow_empty_text:
            errors.append(f"slide {i}: no visible native PowerPoint text or equation")
        if s["picture_shapes"] or s["picture_blips"] or s["raster_media_parts"]: errors.append(f"slide {i}: raster picture/flattened content found in PPTX")
        if s["hidden_elements"] or s["hidden_slides"] or s["hidden_template_elements"]: errors.append(f"slide {i}: hidden elements/slides found on slide or template")
        if s["bad_glyphs"]: errors.append(f"slide {i}: suspicious glyphs in native text")
        s["status"] = "PASS" if len(errors) == start else "FAIL"
    for i, p in enumerate(pages, 1):
        start = len(errors)
        if p["text_chars"] < 1 and i not in allow_empty_text: errors.append(f"PDF page {i}: no extractable text")
        if p["raster_images"]: errors.append(f"PDF page {i}: {p['raster_images']} raster image(s) drawn")
        if (p["font_count"] and not p["fonts_embedded"]) or (not p["font_count"] and i not in allow_empty_text):
            errors.append(f"PDF page {i}: no embedded fonts or one or more fonts are not embedded")
        if p["hidden_content"]: errors.append(f"PDF page {i}: hidden/transparent text, drawings, or layers found")
        if p["bad_glyphs"]: errors.append(f"PDF page {i}: suspicious glyphs")
        p["status"] = "PASS" if len(errors) == start else "FAIL"
    return {"status": "PASS" if not errors else "FAIL", "allow_empty_text_pages": sorted(allow_empty_text),
            "slides": slides, "pdf_pages": pages,
            "errors": errors, "omml_total": sum(s["omml_count"] for s in slides)}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx", type=Path); ap.add_argument("pdf", type=Path)
    ap.add_argument("--expect-math", action="store_true", help="fail if no Office Math (OMML) is present")
    ap.add_argument("--allow-empty-text", type=int, nargs="*", default=[], metavar="PAGE",
                    help="1-based pages intentionally containing only art; skips text/font-resource requirements")
    args = ap.parse_args(argv)
    try:
        if any(page < 1 for page in args.allow_empty_text): ap.error("--allow-empty-text page numbers must be positive")
        report = run(args.pptx, args.pdf, args.expect_math, set(args.allow_empty_text))
    except Exception as e:
        print(json.dumps({"status": "ERROR", "error": str(e)}, indent=2), file=sys.stderr)
        return 2
    print(json.dumps(report, indent=2))
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
