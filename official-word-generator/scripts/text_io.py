from __future__ import annotations

from pathlib import Path


TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "gbk", "big5", "utf-16", "utf-16-le", "utf-16-be")


def read_text_safely(path: Path) -> str:
    raw = path.read_bytes()
    last_error: Exception | None = None
    for encoding in TEXT_ENCODINGS:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
    if last_error:
        raise UnicodeDecodeError(
            last_error.encoding,
            last_error.object,
            last_error.start,
            last_error.end,
            f"Cannot decode {path} with supported encodings: {', '.join(TEXT_ENCODINGS)}",
        )
    return raw.decode("utf-8", errors="replace")


def normalize_markdown_text(text: str, *, remove_horizontal_rules: bool = True) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\ufeff", "")
    lines = []
    for line in text.split("\n"):
        if remove_horizontal_rules and line.strip() in {"---", "***", "___"}:
            continue
        lines.append(line.rstrip())
    return "\n".join(lines).strip() + "\n"


def read_markdown_safely(path: Path) -> str:
    return normalize_markdown_text(read_text_safely(path))
