"""Place each supplied reference beside the corresponding rendered vector PDF page."""
import argparse
from pathlib import Path

import pymupdf


def compare(pdf_path, output, references):
    pdf_path, output = Path(pdf_path).resolve(), Path(output).resolve()
    if output == pdf_path or output in (Path(f).resolve() for f in references):
        raise ValueError("Choose a separate proof PDF.")
    if output.suffix.lower() != ".pdf":
        raise ValueError("The comparison output must end in .pdf")
    output.parent.mkdir(parents=True, exist_ok=True)
    with pymupdf.open(pdf_path) as deck, pymupdf.open() as result:
        if not references or len(references) != len(deck):
            raise ValueError("Supply exactly one ordered reference image per PDF page.")
        for idx, reference in enumerate(references):
            size = deck[idx].rect
            w, h = size.width, size.height
            gap, top = 12, 25
            page = result.new_page(width=2*w+3*gap, height=h+gap+top)
            for x, label in ((gap, f"{idx+1}   Supplied reference"), (w+2*gap, "Editable slide export")):
                page.insert_text((x, 16), label, fontsize=9)
            page.insert_image(pymupdf.Rect(gap, top, gap+w, top+h), filename=str(reference), keep_proportion=True)
            page.show_pdf_page(pymupdf.Rect(w+2*gap, top, 2*w+2*gap, top+h), deck, idx)
        result.save(output, garbage=4, deflate=True)
    return output


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("pdf")
    p.add_argument("proof_pdf")
    p.add_argument("references", nargs="+", help="Source PNG/JPEG paths in presentation order.")
    args = p.parse_args()
    print(compare(args.pdf, args.proof_pdf, args.references))
