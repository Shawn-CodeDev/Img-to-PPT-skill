"""Measure editable Office Math dimensions and write a portable metrics manifest."""
import argparse
import json
from pathlib import Path
import re
import tempfile

import pymupdf
from lxml import etree as ET
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from render import render
from skill_portable_native import Canvas
from skill_portable_native.equations import math_run


REFERENCE_WIDTH = 1300
REFERENCE_HEIGHT = 743
REFERENCE_MM = 183
HOST_WIDTH_PX = 240
HOST_HEIGHT_PX = 50
FONT_PX = 20


def _formula_box(canvas, tex, font_px):
    x, y, w, h = 25, 25, HOST_WIDTH_PX, HOST_HEIGHT_PX
    shape = canvas._textbox(x, y, w, h, "Formula calibration specimen")
    paragraph = shape.text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    font_pt = font_px * canvas.pt_per_px
    paragraph.font.name = "Cambria Math"
    paragraph.font.size = Pt(font_pt)
    paragraph._p.append(math_run(tex, font_pt))
    return font_pt


def _main_glyph_scale(page, tex):
    svg = ET.fromstring(page.get_svg_image(text_as_path=False).encode())
    sizes, letters = [], []
    for text in svg.xpath('//*[local-name()="text"]'):
        content = "".join(text.itertext()).strip()
        if not content:
            continue
        transform = text.get("transform", "matrix(1 0 0 1 0 0)")
        matrix = [float(value) for value in re.findall(r"[-.\d]+(?:e[-+]?\d+)?", transform)]
        if len(matrix) < 4:
            continue
        font_size = float(text.get("font-size"))
        measured = (font_size * abs(matrix[0]), font_size * abs(matrix[3]))
        sizes.append(measured)
        if any(char.isalpha() or char.isdigit() for char in content):
            letters.append(measured)
    if not sizes:
        raise ValueError(f"LibreOffice PDF contains no extractable glyphs for {tex!r}")
    return max(letters or sizes, key=lambda pair: pair[0] * pair[1])


def calibrate(terms_path, metrics_path, soffice=None):
    terms_path, metrics_path = Path(terms_path), Path(metrics_path)
    terms = json.loads(terms_path.read_text(encoding="utf-8"))
    if not isinstance(terms, list) or not terms or any(not isinstance(t, str) or not t.strip() for t in terms):
        raise ValueError("terms.json must contain a non-empty list of non-empty LaTeX strings")
    if len(set(terms)) != len(terms):
        raise ValueError("terms.json contains duplicate formulas")

    from pptx import Presentation
    presentation = Presentation()
    width_pt = HOST_WIDTH_PX * (REFERENCE_MM * 36000 / REFERENCE_WIDTH) / 12700
    height_pt = HOST_HEIGHT_PX * (REFERENCE_MM * 36000 / REFERENCE_WIDTH) / 12700
    presentation.slide_width = round(REFERENCE_MM * 36000)
    presentation.slide_height = round(REFERENCE_MM * REFERENCE_HEIGHT / REFERENCE_WIDTH * 36000)
    specs = []
    for tex in terms:
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        canvas = Canvas(slide, REFERENCE_WIDTH, REFERENCE_HEIGHT, REFERENCE_MM)
        font_pt = _formula_box(canvas, tex, FONT_PX)
        specs.append((tex, font_pt))

    metrics_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="calibrate-math-", dir=metrics_path.parent) as tmp:
        tmp = Path(tmp)
        pptx_path, pdf_path = tmp / "calibration.pptx", tmp / "calibration.pdf"
        presentation.save(pptx_path)
        render(pptx_path, pdf_path, soffice)
        measured = {}
        with pymupdf.open(pdf_path) as pdf:
            if len(pdf) != len(specs):
                raise ValueError(f"Expected {len(specs)} rendered pages; received {len(pdf)}")
            for page, (tex, font_pt) in zip(pdf, specs, strict=True):
                sx, sy = _main_glyph_scale(page, tex)
                width_em = width_pt / sx
                height_em = height_pt / sy
                if not all(value > 0 for value in (width_em, height_em)):
                    raise ValueError(f"Invalid measured dimensions for formula {tex!r}")
                measured[tex] = [round(width_em, 4), round(height_em, 4)]

    metrics_path.write_text(json.dumps(measured, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return measured


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("terms", type=Path, help="JSON list of LaTeX formulas")
    parser.add_argument("metrics", type=Path, help="Output formula metrics JSON")
    parser.add_argument("--soffice", help="Path to LibreOffice executable")
    args = parser.parse_args()
    measured = calibrate(args.terms, args.metrics, args.soffice)
    print(f"Calibrated {len(measured)} formulas into {args.metrics}")


if __name__ == "__main__":
    main()
