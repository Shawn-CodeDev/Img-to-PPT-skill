"""Rebuild, export and check one or all eight figures from the bundled historical case."""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from zipfile import ZipFile

from render import render


def font_dir(choice):
    if choice:
        path = Path(choice)
    elif os.environ.get("NATIVE_FONTS"):
        path = Path(os.environ["NATIVE_FONTS"])
    elif shutil.which("fc-match"):
        path = Path(subprocess.check_output(["fc-match", "-f", "%{file}", "Liberation Serif"], text=True)).parent
    else:
        raise ValueError("Pass --fonts with the directory containing Liberation Serif and Sans TTF files.")
    missing = [f"Liberation{family}-{style}.ttf" for family in ("Serif", "Sans")
               for style in ("Regular", "Italic", "Bold", "BoldItalic")
               if not (path / f"Liberation{family}-{style}.ttf").is_file()]
    if missing:
        raise ValueError(f"Missing measurement font in {path}: {', '.join(missing)}")
    return str(path.resolve())


def replay(args):
    os.environ["NATIVE_FONTS"] = font_dir(args.fonts)
    target = args.directory.resolve()
    target.mkdir(parents=True, exist_ok=True)
    if any(target.iterdir()):
        raise ValueError("Choose an empty work directory so existing figures cannot be overwritten.")
    archive = Path(__file__).resolve().parents[1] / "assets" / "diffusion-redraw-source.zip"
    with ZipFile(archive) as z:
        z.extractall(target)
    # The historical modules use local unqualified imports; give them their own directory first.
    sys.path.insert(0, str(target / "scripts/diffusion_native"))
    from build import FIGURES, OUT, WORK, draw, new
    from compare import compare

    WORK.mkdir(parents=True, exist_ok=True)
    singles = OUT / "Single_Figure_PPTX"
    singles.mkdir(exist_ok=True)
    ids = range(len(FIGURES)) if args.all else [args.figure - 1]
    deck, ledger, ordered_refs = new(), [], []
    with ZipFile(target / "注意力增强：扩散生成语义引导.zip") as z:
        refs = [n for n in z.namelist() if n.lower().endswith(".png") and not n.startswith("__MACOSX")]
        for i in ids:
            label = f"{i+1:02}_{FIGURES[i][0]}"
            ledger.append(draw(deck, i))
            if args.all:
                single = new()
                draw(single, i)
                single.save(singles / f"{label}.pptx")
            ref = WORK / f"original-{i+1:02}.png"
            ref.write_bytes(z.read(refs[i]))
            ordered_refs.append(ref)
    stem = "Diffusion_Methods_Native" if args.all else label
    native, pdf = OUT / f"{stem}.pptx", OUT / f"{stem}.pdf"
    deck.save(native)
    (WORK / "elements.json").write_text(json.dumps(ledger, ensure_ascii=False, indent=1), encoding="utf-8")
    render(native, pdf, args.soffice)
    from audit import run
    check = run(native, pdf, expect_math=any(item["equations"] for item in ledger))
    (OUT / f"{stem}-audit.json").write_text(json.dumps(check, ensure_ascii=False, indent=2), encoding="utf-8")
    if check["status"] != "PASS":
        raise ValueError(f"Audit failed. See {OUT / (stem+'-audit.json')}")
    if args.all:
        from proofcheck import audit as case_audit
        from formula_proof import verify
        from edit_probe import prepare, check as check_edit
        from package import paper_and_reference, bundle
        if not case_audit(OUT)[1]:
            raise ValueError("Historical per-page checks failed; inspect Proofcheck.csv.")
        verify()
        edited = WORK / "07-edited.pptx"
        prepare(singles / "07_Diffusion-DPO.pptx", edited)
        edit_pdf = render(edited, edited.with_suffix(".pdf"), args.soffice)
        evidence = check_edit(edit_pdf, pdf)
        (OUT / "Edit_Probe.json").write_text(json.dumps(evidence, indent=2), encoding="utf-8")
        paper_and_reference()
        bundle()
    else:
        compare(pdf, OUT / f"{stem}-comparison.pdf", ordered_refs)
    print(OUT)


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("directory", type=Path, help="New or empty work directory")
    selection = p.add_mutually_exclusive_group(required=True)
    selection.add_argument("--figure", type=int, choices=range(1, 9))
    selection.add_argument("--all", action="store_true")
    p.add_argument("--fonts", help="Path containing the eight Liberation Serif/Sans font files")
    p.add_argument("--soffice", help="LibreOffice executable")
    replay(p.parse_args())
