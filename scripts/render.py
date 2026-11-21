"""Export one editable deck with a dedicated LibreOffice profile."""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import tempfile

import pymupdf
from pptx import Presentation


def render(source, target, soffice=None):
    source, target = Path(source).resolve(), Path(target).resolve()
    if source.suffix.lower() != ".pptx" or target.suffix.lower() != ".pdf":
        raise ValueError("Expected a .pptx input and a .pdf output.")
    pages = len(Presentation(source).slides)
    if not pages:
        raise ValueError("The input presentation is empty.")
    program = soffice or os.environ.get("SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    if not program:
        raise FileNotFoundError("Install LibreOffice with Math, or set SOFFICE to its executable.")
    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ppt-proof-", dir=target.parent) as tmp:
        tmp = Path(tmp)
        result = subprocess.run(
            [str(program), f"-env:UserInstallation={(tmp / 'profile').as_uri()}", "--headless",
             "--convert-to", "pdf:impress_pdf_Export", "--outdir", str(tmp), str(source)],
            capture_output=True, text=True, timeout=300,
        )
        pdf = tmp / source.with_suffix(".pdf").name
        if result.returncode or not pdf.is_file():
            raise RuntimeError(f"LibreOffice export failed. {result.stdout} {result.stderr}")
        with pymupdf.open(pdf) as doc:
            if len(doc) != pages:
                raise ValueError(f"Export has {len(doc)} pages for {pages} slides.")
        pdf.replace(target)
    return target


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("pptx", type=Path)
    p.add_argument("pdf", type=Path)
    p.add_argument("--soffice", help="Path to the LibreOffice executable")
    args = p.parse_args()
    print(render(args.pptx, args.pdf, args.soffice))
