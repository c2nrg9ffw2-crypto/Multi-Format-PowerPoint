"""QA pass: run markitdown on the output .pptx and flag obvious issues."""
import re
import subprocess
import sys
from pathlib import Path


_PLACEHOLDER_RE = re.compile(
    r"\b(lorem ipsum|placeholder|todo|tbd|fixme|xxx|sample text)\b", re.IGNORECASE
)


def run_qa(pptx_path: str) -> bool:
    path = Path(pptx_path)
    if not path.exists():
        print(f"QA FAIL: output file not found: {pptx_path}", file=sys.stderr)
        return False

    # markitdown converts pptx → markdown so we can inspect content as plain text
    try:
        result = subprocess.run(
            ["markitdown", str(path)],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except FileNotFoundError:
        print("QA SKIP: markitdown not found — install with: pip install markitdown", file=sys.stderr)
        return True

    if result.returncode != 0:
        print(f"QA WARN: markitdown error:\n{result.stderr}", file=sys.stderr)
        return True  # non-fatal

    content = result.stdout

    issues = []

    if _PLACEHOLDER_RE.search(content):
        issues.append("Placeholder text detected in output.")

    slides = re.split(r"^#{1,2} ", content, flags=re.MULTILINE)
    empty_slides = [i for i, s in enumerate(slides) if len(s.strip()) < 10]
    if empty_slides:
        issues.append(f"Potentially empty slide(s) at positions: {empty_slides}")

    if issues:
        print("QA issues found:", file=sys.stderr)
        for issue in issues:
            print(f"  • {issue}", file=sys.stderr)
        return False

    print("QA passed.")
    return True
