#!/usr/bin/env python3
"""
convert_drive_to_md.py
----------------------
Download one or more Google Drive files and convert each to Markdown using
markitdown. Output files are placed in a named folder ready to drag into any
Claude Code project as context.

Usage:
    python convert_drive_to_md.py --output MyProject \
        "https://drive.google.com/file/d/ABC123/view" \
        "https://drive.google.com/file/d/DEF456/view"

    # Or supply URLs one-per-line in a text file:
    python convert_drive_to_md.py --output MyProject --file urls.txt
"""

import argparse
import os
import sys
import tempfile
from pathlib import Path

import gdown
from markitdown import MarkItDown


def collect_urls(args) -> list[str]:
    urls = list(args.urls)
    if args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            sys.exit(f"Error: URL file not found: {args.file}")
        with open(file_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    urls.append(line)
    if not urls:
        sys.exit("Error: no URLs provided. Pass URLs as arguments or use --file.")
    return urls


def download_drive_file(url: str, dest_dir: str) -> str | None:
    """Download a Google Drive file into dest_dir and return the local path."""
    try:
        path = gdown.download(url, output=dest_dir + "/", quiet=False, fuzzy=True)
        return path
    except Exception as exc:
        print(f"  [!] Download failed: {exc}", file=sys.stderr)
        return None


def convert_to_markdown(local_path: str, md: MarkItDown) -> str | None:
    """Run markitdown on a local file and return the markdown string."""
    try:
        result = md.convert(local_path)
        return result.markdown
    except Exception as exc:
        print(f"  [!] Conversion failed: {exc}", file=sys.stderr)
        return None


def main():
    parser = argparse.ArgumentParser(
        description="Convert Google Drive files to Markdown for Claude Code projects.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "urls",
        nargs="*",
        metavar="URL",
        help="Google Drive sharing URLs",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        metavar="FOLDER",
        help="Name of the output folder (will be created if missing)",
    )
    parser.add_argument(
        "-f",
        "--file",
        metavar="URLS_FILE",
        help="Path to a text file with one Google Drive URL per line",
    )
    parser.add_argument(
        "--prefix",
        metavar="TEXT",
        default="",
        help="Optional text to prepend to every generated Markdown file",
    )
    args = parser.parse_args()

    urls = collect_urls(args)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    md = MarkItDown()
    success, failed = 0, 0

    for idx, url in enumerate(urls, 1):
        print(f"\n[{idx}/{len(urls)}] {url}")

        with tempfile.TemporaryDirectory() as tmp:
            local_path = download_drive_file(url, tmp)
            if local_path is None:
                print("  Skipped (download failed).")
                failed += 1
                continue

            print(f"  Downloaded: {local_path}")
            markdown = convert_to_markdown(local_path, md)
            if markdown is None:
                print("  Skipped (conversion failed).")
                failed += 1
                continue

            stem = Path(local_path).stem
            out_file = output_dir / (stem + ".md")

            # Avoid silent overwrites — append a counter if name collides
            counter = 1
            while out_file.exists():
                out_file = output_dir / (f"{stem}_{counter}.md")
                counter += 1

            content = (args.prefix + "\n\n" + markdown) if args.prefix else markdown
            out_file.write_text(content, encoding="utf-8")
            print(f"  Saved: {out_file}")
            success += 1

    print(f"\nDone. {success} converted, {failed} failed.")
    print(f"Output folder: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
