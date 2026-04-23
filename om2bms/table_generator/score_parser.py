from __future__ import annotations

import json

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ScoreEntry:
    title: str
    level: str
    eval: int
    artist: str
    url: str
    url_diff: str
    name_diff: str
    comment: str
    note: int
    total: int
    judge: int
    md5: str
    sha256: str


@dataclass(frozen=True)
class ParsedScore:
    entries: list[ScoreEntry]


def smart_decode(data: bytes) -> str:
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig")
    if data.startswith(b"\xff\xfe"):
        return data.decode("utf-16-le")
    if data.startswith(b"\xfe\xff"):
        return data.decode("utf-16-be")

    for encoding in ("utf-8", "cp932", "shift_jis"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def parse_score_path(score_path: str | Path) -> ParsedScore:
    path = Path(score_path)
    return parse_score_bytes(path.read_bytes())


def parse_score_text(score_text: str) -> ParsedScore:
    return _parse_score_object(_load_score_json(score_text))


def parse_score_bytes(data: bytes) -> ParsedScore:
    return _parse_score_object(_load_score_json(smart_decode(data)))


def parse_score_object(score_object: list[dict[str, Any]] | dict[str, Any] | ParsedScore) -> ParsedScore:
    if isinstance(score_object, ParsedScore):
        return score_object
    return _parse_score_object(score_object)


def _load_score_json(text: str) -> list[dict[str, Any]] | dict[str, Any]:
    stripped = text.strip()
    if not stripped:
        return []

    try:
        return json.loads(stripped)
    except json.JSONDecodeError:
        return json.loads(f"[{stripped}]")


def _parse_score_object(score_object: list[dict[str, Any]] | dict[str, Any]) -> ParsedScore:
    if isinstance(score_object, list):
        return ParsedScore(
            entries=[_parse_score_entry(item) for item in score_object if isinstance(item, dict)]
        )

    if isinstance(score_object, dict):
        if "entries" in score_object and isinstance(score_object["entries"], list):
            return ParsedScore(
                entries=[_parse_score_entry(item) for item in score_object["entries"] if isinstance(item, dict)]
            )
        return ParsedScore(entries=[_parse_score_entry(score_object)])

    raise TypeError("Unsupported score object. Expected ParsedScore, list[dict], or dict.")


def _parse_score_entry(item: dict[str, Any]) -> ScoreEntry:
    return ScoreEntry(
        title=str(item.get("title", "")),
        level=str(item.get("level", "")),
        eval=_safe_int(item.get("eval"), default=0),
        artist=str(item.get("artist", "")),
        url=str(item.get("url", "")),
        url_diff=str(item.get("url_diff", "")),
        name_diff=str(item.get("name_diff", "")),
        comment=str(item.get("comment", "")),
        note=_safe_int(item.get("note"), default=0),
        total=_safe_int(item.get("total"), default=0),
        judge=_safe_int(item.get("judge"), default=0),
        md5=str(item.get("md5", "")),
        sha256=str(item.get("sha256", "")),
    )


def _safe_int(value: Any, default: int) -> int:
    if value is None:
        return default
    try:
        return int(value)
    except (ValueError, TypeError):
        return default
