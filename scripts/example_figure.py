"""A small reproducible diagram with a rounded route, editable subscript, and grouped artwork."""
import argparse
import json
from pathlib import Path

from skill_portable_native import new_deck


def example(directory, metrics):
    directory.mkdir(parents=True, exist_ok=True)
    ref_width, physical_width = 720, 120
    pres, canvas = new_deck(ref_width, 360, physical_width, metrics)
    font = "Liberation Sans"
    canvas.rect(35, 38, 200, 65, "#E6F0F7", "#45718F", 2, 10, name="Source block")
    canvas.text(47, 48, 176, 45, "Input", 28, font=font, name="Input label")
    canvas.rect(425, 145, 255, 72, "#E3EEE7", "#537965", 2, 10, name="Target block")
    canvas.text(443, 151, 219, 56, "Projection", 27, font=font, name="Target label")
    canvas.arrow([(235,70), (323,70), (323,181), (425,181)], width=3, radius=18,
                 name="Input to projection")
    with canvas.group("Basis symbol"):
        canvas.poly([(70,267), (107,222), (144,267)], "#6F92AD", name="Back face")
        canvas.poly([(107,222), (145,267), (121,258)], "#45637F", name="Shaded face")
    canvas.eq(468, 236, 165, 65, "y_w", size=36, name="Projected coordinate")
    pres.save(directory / "example.pptx")

    # Top-left coordinates in uncropped PDF points; include both positions of moved objects.
    pt = physical_width * 72 / 25.4 / ref_width
    def points(bounds):
        return [round(x * pt, 3) for x in bounds]
    regions = [
        {"name": "Input label", "rect": points([42,42,227,99]), "must_change": True},
        {"name": "Projected coordinate", "rect": points([461,230,640,307]), "must_change": True},
        {"name": "Basis symbol", "rect": points([65,217,160,273]), "must_change": True},
    ]
    (directory / "edit-regions.json").write_text(json.dumps({"1": regions}, indent=2), encoding="utf-8")
    eq = [{"page": 1, "name": "Projected coordinate", "box_pt": points([468,236,633,301])}]
    (directory / "formula-boxes.json").write_text(json.dumps(eq, indent=2), encoding="utf-8")


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("directory", type=Path)
    p.add_argument("--metrics", type=Path, help="Measured equation dimensions from calibrate_math.py")
    args = p.parse_args()
    example(args.directory, args.metrics)
