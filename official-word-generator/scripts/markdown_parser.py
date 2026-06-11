from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from text_io import read_markdown_safely  # noqa: E402


IMAGE_RE = re.compile(r"^!\[(?P<alt>.*?)\]\((?P<path>.*?)\)\s*$")
HEADING_RE = re.compile(r"^(?P<marks>#{1,6})\s+(?P<text>.+?)\s*$")
TOC_RE = re.compile(r"^(?:\[TOC\]|\[\[TOC\]\])$", re.IGNORECASE)


@dataclass
class Block:
    type: str
    text: str = ""
    level: int | None = None
    rows: list[list[str]] | None = None
    path: str | None = None
    caption: str | None = None


def _is_table_line(line: str) -> bool:
    return line.strip().startswith("|") and line.strip().endswith("|")


def _is_separator_row(row: list[str]) -> bool:
    if not row:
        return False
    return all(re.fullmatch(r":?-{3,}:?", cell.strip()) for cell in row)


def _parse_table_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    return [cell.strip() for cell in stripped.split("|")]


def parse_markdown_text(text: str) -> list[Block]:
    blocks: list[Block] = []
    paragraph: list[str] = []
    lines = text.splitlines()
    i = 0

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            blocks.append(Block(type="paragraph", text=" ".join(part.strip() for part in paragraph).strip()))
            paragraph = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            flush_paragraph()
            i += 1
            continue

        if TOC_RE.match(stripped):
            flush_paragraph()
            blocks.append(Block(type="toc"))
            i += 1
            continue

        heading = HEADING_RE.match(stripped)
        if heading:
            flush_paragraph()
            blocks.append(Block(type="heading", level=len(heading.group("marks")), text=heading.group("text")))
            i += 1
            continue

        image = IMAGE_RE.match(stripped)
        if image:
            flush_paragraph()
            blocks.append(Block(type="image", path=image.group("path").strip(), caption=image.group("alt").strip()))
            i += 1
            continue

        if stripped.startswith(">"):
            flush_paragraph()
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                quote_lines.append(lines[i].strip().lstrip(">").strip())
                i += 1
            blocks.append(Block(type="quote", text=" ".join(quote_lines).strip()))
            continue

        if _is_table_line(line):
            flush_paragraph()
            rows: list[list[str]] = []
            while i < len(lines) and _is_table_line(lines[i]):
                row = _parse_table_row(lines[i])
                if not _is_separator_row(row):
                    rows.append(row)
                i += 1
            if rows:
                blocks.append(Block(type="table", rows=rows))
            continue

        paragraph.append(line)
        i += 1

    flush_paragraph()
    return blocks


def parse_markdown(path: Path) -> list[Block]:
    return parse_markdown_text(read_markdown_safely(path))


def main() -> None:
    parser = argparse.ArgumentParser(description="Parse Markdown into structured blocks for Word generation.")
    parser.add_argument("--content", required=True, type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--normalized-output", type=Path, help="Write a UTF-8 normalized Markdown copy.")
    args = parser.parse_args()

    text = read_markdown_safely(args.content)
    if args.normalized_output:
        args.normalized_output.parent.mkdir(parents=True, exist_ok=True)
        args.normalized_output.write_text(text, encoding="utf-8")
    blocks = parse_markdown_text(text)
    payload = [asdict(block) for block in blocks]
    data = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(data, encoding="utf-8")
    else:
        print(data)


if __name__ == "__main__":
    main()
