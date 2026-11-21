"""Resolve slide part paths in the order displayed in PowerPoint."""
import posixpath
from xml.etree import ElementTree as ET

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def slide_parts(archive):
    presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
    rels = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels if rel.get("Type") == f"{R}/slide"}
    parts = []
    for slide in presentation.findall(f"./{{{P}}}sldIdLst/{{{P}}}sldId"):
        rid = slide.get(f"{{{R}}}id")
        if rid not in targets:
            raise ValueError(f"Missing presentation slide relationship {rid!r}")
        target = targets[rid]
        part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join("ppt", target))
        archive.getinfo(part)
        parts.append(part)
    if not parts:
        raise ValueError("Presentation contains no slides.")
    return parts
