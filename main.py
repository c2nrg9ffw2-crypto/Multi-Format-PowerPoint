#!/usr/bin/env python3
"""Multi-Format PowerPoint Generator — xlsx/pdf → .pptx"""
import argparse
import sys
from pathlib import Path

from src.parsers import parse_xlsx, parse_pdf
from src.summarizer import build_slide_data
from src.builder import build_pptx
from src.qa import run_qa


def _infer_label(inputs: list[str]) -> str:
    stems = [Path(p).stem for p in inputs]
    return " & ".join(stems) if stems else "Report"


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Convert .xlsx / .pdf files into a branded .pptx presentation."
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        metavar="FILE",
        help=".xlsx and/or .pdf source files (1 or 2)",
    )
    parser.add_argument(
        "-o", "--output",
        default="output.pptx",
        metavar="OUTPUT",
        help="Output .pptx path (default: output.pptx)",
    )
    parser.add_argument(
        "--no-qa",
        action="store_true",
        help="Skip the QA validation step",
    )
    args = parser.parse_args(argv)

    if len(args.inputs) > 2:
        parser.error("Provide at most 2 input files.")

    xlsx_data = None
    pdf_data = None

    for filepath in args.inputs:
        p = Path(filepath)
        if not p.exists():
            print(f"Error: file not found: {filepath}", file=sys.stderr)
            sys.exit(1)
        suffix = p.suffix.lower()
        if suffix == ".xlsx":
            if xlsx_data is not None:
                parser.error("Only one .xlsx file is supported.")
            print(f"Parsing xlsx: {filepath}")
            xlsx_data = parse_xlsx(filepath)
            print(f"  > {len(xlsx_data)} sheet(s): {list(xlsx_data.keys())}")
        elif suffix == ".pdf":
            if pdf_data is not None:
                parser.error("Only one .pdf file is supported.")
            print(f"Parsing pdf: {filepath}")
            pdf_data = parse_pdf(filepath)
            print(f"  > {len(pdf_data.get('pages', []))} page(s)")
        else:
            print(f"Error: unsupported file type '{suffix}'. Use .xlsx or .pdf.", file=sys.stderr)
            sys.exit(1)

    if xlsx_data is None and pdf_data is None:
        parser.error("No valid input files were parsed.")

    label = _infer_label(args.inputs)
    print(f"Building slide data…")
    slide_data = build_slide_data(xlsx_data, pdf_data, source_label=label)
    print(f"  > {len(slide_data)} slides")

    print(f"Rendering {args.output}…")
    build_pptx(slide_data, args.output)
    print(f"  > saved: {args.output}")

    if not args.no_qa:
        print("Running QA…")
        run_qa(args.output)


if __name__ == "__main__":
    main()
