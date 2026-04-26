#!/usr/bin/env python3
"""
convert_drive_to_md.py
----------------------
Convert local files to Markdown using markitdown (DOCX, PDF, PPTX, images, etc.).
Output files go into a folder you name.

There is no Google download in this script. For Google Docs/Drive:
  • Open the file in your browser → File → Download (e.g. .docx), then run this tool.
  • For automation at scale, use the Google Drive API with OAuth2 (not included here).

Usage:
    python convert_drive_to_md.py --output MyProject report.docx notes.pdf

    python convert_drive_to_md.py --output MyProject --file paths.txt

paths.txt: one local file path per line (# lines are ignored).
"""

import argparse
import sys
from pathlib import Path

from markitdown import MarkItDown


def collect_paths(args) -> list[str]:
    paths = list(args.files)
    if args.paths_file:
        p = Path(args.paths_file)
        if not p.exists():
            sys.exit(f"Error: paths file not found: {args.paths_file}")
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    paths.append(line)
    if not paths:
        sys.exit(
            "Error: no input files. Pass file paths as arguments or use --file paths.txt."
        )
    return paths


def _looks_like_ooxml_zip(path: Path) -> bool:
    """OOXML (.docx/.xlsx/.pptx) files are ZIP archives and start with PK."""
    try:
        with open(path, "rb") as f:
            return f.read(2) == b"PK"
    except OSError:
        return False


def _guess_bad_ooxml(path: Path) -> str:
    try:
        with open(path, "rb") as f:
            sniff = f.read(240)
    except OSError:
        return ""
    if sniff.strip().startswith(b"<"):
        preview = sniff.decode("utf-8", errors="replace").replace("\n", " ")[:120]
        return f' Looks like HTML/text: {preview!r}...'
    return ""


def convert_to_markdown(local_path: str, md: MarkItDown) -> str | None:
    try:
        result = md.convert(local_path)
        return result.markdown
    except Exception as exc:
        print(f"  [!] Conversion failed: {exc}", file=sys.stderr)
        return None


def convert_one(
    local_path_str: str,
    *,
    md: MarkItDown,
    output_dir: Path,
    prefix: str,
) -> bool:
    lp = Path(local_path_str).expanduser().resolve()
    print(f"  Input: {lp}")

    if not lp.is_file():
        print(f"  [!] Not a file: {lp}", file=sys.stderr)
        return False

    if lp.suffix.lower() in (".docx", ".xlsx", ".pptx") and not _looks_like_ooxml_zip(lp):
        extra = _guess_bad_ooxml(lp)
        print(
            "  [!] Not a valid Office Open XML file (expected ZIP starting with PK)."
            + extra
            + " Re-download from Google as real .docx, or rename if mislabeled.",
            file=sys.stderr,
        )
        return False

    markdown = convert_to_markdown(str(lp), md)
    if markdown is None:
        return False

    stem = lp.stem
    out_file = output_dir / (stem + ".md")
    counter = 1
    while out_file.exists():
        out_file = output_dir / (f"{stem}_{counter}.md")
        counter += 1

    content = (prefix + "\n\n" + markdown) if prefix else markdown
    out_file.write_text(content, encoding="utf-8")
    print(f"  Saved: {out_file}")
    return True


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert local files to Markdown with markitdown.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "files",
        nargs="*",
        metavar="FILE",
        help="Local files to convert",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        metavar="FOLDER",
        help="Output folder (created if missing)",
    )
    parser.add_argument(
        "-f",
        "--file",
        dest="paths_file",
        metavar="PATHS_FILE",
        help="Text file with one local path per line (# comments allowed)",
    )
    parser.add_argument(
        "--prefix",
        metavar="TEXT",
        default="",
        help="Optional text to prepend to every generated Markdown file",
    )
    args = parser.parse_args()

    paths = collect_paths(args)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    md = MarkItDown()
    success, failed = 0, 0

    for idx, raw in enumerate(paths, 1):
        print(f"\n[{idx}/{len(paths)}]")
        if convert_one(
            raw,
            md=md,
            output_dir=output_dir,
            prefix=args.prefix,
        ):
            success += 1
        else:
            failed += 1

    print(f"\nDone. {success} converted, {failed} failed.")
    print(f"Output folder: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
