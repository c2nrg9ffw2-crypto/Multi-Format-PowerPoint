#!/usr/bin/env python3
"""Multi-Format PowerPoint Generator — xlsx/pdf → .pptx"""
import argparse
import sys
from pathlib import Path

from src.parsers import parse_xlsx, parse_pdf, ParseError
from src.summarizer import build_slide_data
from src.builder import build_pptx
from src.qa import run_qa


def _infer_label(inputs: list[str]) -> str:
    stems = [Path(p).stem for p in inputs]
    if len(stems) == 1:
        return stems[0]
    if len(stems) <= 3:
        return " & ".join(stems)
    return f"{stems[0]} & {len(stems) - 1} others"


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Convert .xlsx / .pdf files into a branded .pptx presentation."
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        metavar="FILE",
        help=".xlsx and/or .pdf source files (one or more of each)",
    )
    parser.add_argument(
        "-o", "--output",
        default="output.pptx",
        metavar="OUTPUT",
        help="Output .pptx path (default: output.pptx)",
    )
    parser.add_argument(
        "--template",
        metavar="TEMPLATE.pptx",
        help="Existing .pptx to use as a template; omit to generate from scratch via PptxGenJS",
    )
    parser.add_argument(
        "--no-qa",
        action="store_true",
        help="Skip the QA validation step",
    )
    args = parser.parse_args(argv)

    # Validate template if provided
    if args.template:
        t = Path(args.template)
        if not t.exists():
            print(f"Error: template not found: {args.template}", file=sys.stderr)
            sys.exit(1)
        if t.suffix.lower() != ".pptx":
            print(f"Error: template must be a .pptx file.", file=sys.stderr)
            sys.exit(1)

    # Parse all inputs, collecting results by type
    xlsx_results = []  # list of (Path, dict)
    pdf_results = []   # list of (Path, dict)

    for filepath in args.inputs:
        p = Path(filepath)
        if not p.exists():
            print(f"Error: file not found: {filepath}", file=sys.stderr)
            sys.exit(1)
        suffix = p.suffix.lower()
        try:
            if suffix == ".xlsx":
                print(f"Parsing xlsx: {filepath}")
                data = parse_xlsx(filepath)
                print(f"  > {len(data)} sheet(s): {list(data.keys())}")
                xlsx_results.append((p, data))
            elif suffix == ".pdf":
                print(f"Parsing pdf: {filepath}")
                data = parse_pdf(filepath)
                print(f"  > {len(data.get('pages', []))} page(s)")
                pdf_results.append((p, data))
            else:
                print(f"Error: unsupported file type '{suffix}'. Use .xlsx or .pdf.", file=sys.stderr)
                sys.exit(1)
        except ParseError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

    if not xlsx_results and not pdf_results:
        print("Error: no valid input files were parsed.", file=sys.stderr)
        sys.exit(1)

    # Merge xlsx sheets — prefix sheet names with filename when multiple xlsx files
    xlsx_data = None
    if xlsx_results:
        xlsx_data = {}
        multi = len(xlsx_results) > 1
        for path, data in xlsx_results:
            prefix = f"{path.stem} – " if multi else ""
            for sheet, rows in data.items():
                xlsx_data[prefix + sheet] = rows

    # Merge pdf pages — concatenate in input order
    pdf_data = None
    if pdf_results:
        pdf_data = {"pages": [pg for _, d in pdf_results for pg in d.get("pages", [])]}

    label = _infer_label(args.inputs)
    print("Building slide data…")
    slide_data = build_slide_data(xlsx_data, pdf_data, source_label=label)
    print(f"  > {len(slide_data)} slides")

    print(f"Rendering {args.output}…")
    if args.template:
        from src.builder import build_pptx_from_template
        build_pptx_from_template(slide_data, args.template, args.output)
    else:
        build_pptx(slide_data, args.output)
    print(f"  > saved: {args.output}")

    if not args.no_qa:
        print("Running QA…")
        run_qa(args.output)


if __name__ == "__main__":
    main()
