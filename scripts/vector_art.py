"""Turn an opaque raster artwork into editable PowerPoint DrawingML paths.

Requires Pillow, vtracer, lxml, fontTools and python-pptx. The source image is
center-cropped to the requested box aspect ratio, then traced and inserted as
an editable group. Transparent pixels are rejected because VTracer's stacked
color trace does not preserve arbitrary per-pixel alpha as editable fills.
"""

from __future__ import annotations

from pathlib import Path
import re
import tempfile
from xml.sax.saxutils import escape

from fontTools.pens.recordingPen import RecordingPen
from fontTools.svgLib.path import parse_path
from lxml import etree as ET
from PIL import Image
from pptx.oxml import parse_xml
import vtracer


_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_SVG = "http://www.w3.org/2000/svg"


def _drawing_path(path: ET._Element, svg_width: float, svg_height: float):
    """Convert one SVG path to a local-coordinate DrawingML custom path."""
    pen = RecordingPen()
    parse_path(path.get("d", ""), pen)
    transform = path.get("transform", "")
    match = re.search(r"translate\(\s*([-+\d.eE]+)(?:[, ]+([-+\d.eE]+))?\s*\)", transform)
    dx, dy = (float(match.group(1)), float(match.group(2) or 0)) if match else (0.0, 0.0)

    commands = []
    for command, points in pen.value:
        converted = [(min(svg_width, max(0.0, x + dx)), min(svg_height, max(0.0, y + dy))) for x, y in points]
        commands.append((command, converted))

    all_points = [point for _, points in commands for point in points]
    if not all_points:
        return None
    left = round(min(x for x, _ in all_points) * 100)
    top = round(min(y for _, y in all_points) * 100)
    right = round(max(x for x, _ in all_points) * 100)
    bottom = round(max(y for _, y in all_points) * 100)
    width, height = max(1, right - left), max(1, bottom - top)
    tags = {"moveTo": "moveTo", "lineTo": "lnTo", "curveTo": "cubicBezTo", "qCurveTo": "quadBezTo", "closePath": "close"}
    parts = []
    for command, points in commands:
        tag = tags.get(command)
        if tag is None:
            raise ValueError(f"Unsupported traced SVG command: {command}")
        if command == "closePath":
            parts.append("<a:close/>")
        else:
            if not points:
                continue
            pts = "".join(f'<a:pt x="{round(x * 100) - left}" y="{round(y * 100) - top}"/>' for x, y in points)
            parts.append(f"<a:{tag}>{pts}</a:{tag}>")
    return left, top, width, height, "".join(parts)


def _shape_xml(name: str, shape_id: int, left: int, top: int, width: int, height: int, path_xml: str, color: str) -> str:
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(name, {chr(34): "&quot;"})}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>'
        f'<a:pathLst><a:path w="{width}" h="{height}" stroke="0">{path_xml}</a:path></a:pathLst></a:custGeom>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
    )


def place_vector(
    slide,
    image_path: str | Path,
    x_emu: int,
    y_emu: int,
    width_emu: int,
    height_emu: int,
    *,
    name: str | None = None,
    max_paths: int = 1200,
):
    """Trace a local raster and add one editable group to ``slide``.

    Coordinates and dimensions are EMU. The image is center-cropped to the box
    aspect ratio (like ``ImageOps.fit``), preserving proportions without
    stretching. Returns python-pptx's GroupShape proxy. Transparent images and
    traces above ``max_paths`` raise ``ValueError`` with an actionable message.
    """
    if width_emu <= 0 or height_emu <= 0:
        raise ValueError("width_emu and height_emu must be positive")
    if max_paths < 1:
        raise ValueError("max_paths must be at least 1")
    source = Path(image_path)
    if not source.is_file():
        raise FileNotFoundError(source)

    with Image.open(source) as opened:
        rgba = opened.convert("RGBA")
        alpha = rgba.getchannel("A")
        amin, amax = alpha.getextrema()
        if amin != 255 or amax != 255:
            raise ValueError("Transparent artwork is not supported: provide an opaque raster image (all pixels alpha=255).")
        aspect = width_emu / height_emu
        sw, sh = rgba.size
        if sw / sh > aspect:
            new_width = max(1, round(sh * aspect))
            left = (sw - new_width) // 2
            rgba = rgba.crop((left, 0, left + new_width, sh))
        else:
            new_height = max(1, round(sw / aspect))
            top = (sh - new_height) // 2
            rgba = rgba.crop((0, top, sw, top + new_height))
        rgba.thumbnail((512, 512), Image.Resampling.LANCZOS)
        raster_width, raster_height = rgba.size
        rgb = Image.new("RGB", rgba.size, "white")
        rgb.paste(rgba, mask=rgba.getchannel("A"))

    with tempfile.TemporaryDirectory(prefix="ppt-vector-") as temp_dir:
        png_path = Path(temp_dir) / "source.png"
        svg_path = Path(temp_dir) / "trace.svg"
        rgb.save(png_path)
        vtracer.convert_image_to_svg_py(
            str(png_path), str(svg_path), colormode="color", hierarchical="stacked", mode="spline",
            filter_speckle=3, color_precision=6, layer_difference=12, corner_threshold=60,
            length_threshold=4.0, max_iterations=10, splice_threshold=45, path_precision=3,
        )
        root = ET.parse(str(svg_path)).getroot()
        raw_paths = list(root.iter(f"{{{_SVG}}}path"))
        if len(raw_paths) > max_paths:
            raise ValueError(
                f"Trace produced {len(raw_paths)} editable regions, above max_paths={max_paths}. "
                "Use simpler artwork, reduce image detail/colors, or explicitly raise max_paths."
            )
        viewbox = root.get("viewBox", f"0 0 {raster_width} {raster_height}").split()
        if len(viewbox) != 4:
            raise ValueError("VTracer returned an invalid SVG viewBox")
        svg_width, svg_height = float(viewbox[2]), float(viewbox[3])
        shapes = []
        for index, svg_path_node in enumerate(raw_paths, start=1):
            color = (svg_path_node.get("fill") or "").lstrip("#")
            if not re.fullmatch(r"[0-9a-fA-F]{6}", color):
                raise ValueError("VTracer returned an unsupported fill; expected opaque six-digit RGB colors")
            geometry = _drawing_path(svg_path_node, svg_width, svg_height)
            if geometry is None:
                continue
            left, top, path_width, path_height, path_xml = geometry
            shapes.append((left, top, _shape_xml(f"{name or source.stem} • region {index}", 0, left, top, path_width, path_height, path_xml, color.upper())))

    if not shapes:
        raise ValueError("No editable paths were produced; try a nonblank raster image")
    # Inserting each region as a group child keeps the artwork grouped and gives
    # PowerPoint valid unique non-visual IDs within the slide's shape tree.
    cx, cy = raster_width * 100, raster_height * 100
    group_name = name or source.stem
    children = "".join(xml for _, _, xml in shapes)
    group = (
        f'<p:grpSp xmlns:p="{_P}" xmlns:a="{_A}"><p:nvGrpSpPr><p:cNvPr id="0" name="{escape(group_name, {chr(34): "&quot;"})}"/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
        f'<a:off x="{int(x_emu)}" y="{int(y_emu)}"/><a:ext cx="{int(width_emu)}" cy="{int(height_emu)}"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></p:grpSpPr>{children}</p:grpSp>'
    )
    tree = parse_xml(group.encode("utf-8"))
    sp_tree = slide.shapes._spTree
    sp_tree.insert_element_before(tree, "p:extLst")
    # Allocate after insertion so python-pptx also sees the new IDs when turbo
    # shape-ID caching is enabled on the slide.
    for node in tree.iter(f"{{{_P}}}cNvPr"):
        node.set("id", str(slide.shapes._next_shape_id))
    return slide.shapes[-1]
