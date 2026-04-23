from __future__ import annotations

import argparse
import json
import sys

from dataclasses import asdict
from pathlib import Path

from om2bms.analysis.bms_parser import calculate_md5, calculate_sha256, parse_chart_path
from om2bms.table_generator.score_parser import ParsedScore, ScoreEntry, parse_score_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Parse .bms and match score.json by MD5.")
    parser.add_argument("bms_path", type=str, help="Path to .bms file")
    parser.add_argument("score_json_path", type=str, help="Path to score.json")
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty print result as JSON",
    )
    args = parser.parse_args()

    bms_path = Path(args.bms_path)
    score_json_path = Path(args.score_json_path)

    if not bms_path.exists():
        print(f"BMS file not found: {bms_path}", file=sys.stderr)
        sys.exit(1)

    if not score_json_path.exists():
        print(f"Score JSON file not found: {score_json_path}", file=sys.stderr)
        sys.exit(1)

    bms_data = bms_path.read_bytes()
    bms_md5 = calculate_md5(bms_data)
    bms_sha256 = calculate_sha256(bms_data)

    parsed_chart = parse_chart_path(bms_path)
    parsed_scores = parse_score_path(score_json_path)

    score_index = build_score_index_by_md5(parsed_scores)
    matched_score = score_index.get(bms_md5.lower())

    if matched_score is None:
        print("MD5 match failed.", file=sys.stderr)
        print(f"BMS path   : {bms_path}", file=sys.stderr)
        print(f"BMS md5    : {bms_md5}", file=sys.stderr)
        print(f"BMS sha256 : {bms_sha256}", file=sys.stderr)
        print(f"JSON path  : {score_json_path}", file=sys.stderr)
        sys.exit(2)

    if args.pretty:
        result = {
            "matched": True,
            "match_method": "md5",
            "bms_path": str(bms_path),
            "bms_md5": bms_md5,
            "bms_sha256": bms_sha256,
            "score_entry": asdict(matched_score),
            "song_info": asdict(parsed_chart.song_info),
            "timeline_rows": len(parsed_chart.timeline_master),
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    print("MD5 Match Success")
    print("=================")
    print(f"BMS path       : {bms_path}")
    print(f"BMS md5        : {bms_md5}")
    print(f"BMS sha256     : {bms_sha256}")
    print()

    print("BMS Parse Result")
    print("================")
    print(f"Title          : {parsed_chart.song_info.title}")
    print(f"Subtitle       : {parsed_chart.song_info.subtitle}")
    print(f"Artist         : {parsed_chart.song_info.artist}")
    print(f"Subartist      : {parsed_chart.song_info.subartist}")
    print(f"Total          : {parsed_chart.song_info.total}")
    print(f"Total notes    : {parsed_chart.song_info.total_notes}")
    print(f"Song last ms   : {parsed_chart.song_info.song_last_ms}")
    print(f"Timeline rows  : {len(parsed_chart.timeline_master)}")
    print()

    print("Score JSON Entry")
    print("================")
    print_score_entry(matched_score)


def build_score_index_by_md5(scores: ParsedScore) -> dict[str, ScoreEntry]:
    index: dict[str, ScoreEntry] = {}

    for entry in scores.entries:
        entry_md5 = entry.md5.strip().lower()
        if not entry_md5:
            continue
        index[entry_md5] = entry

    return index


def print_score_entry(entry: ScoreEntry) -> None:
    print(f"title     : {entry.title}")
    print(f"level     : {entry.level}")
    print(f"eval      : {entry.eval}")
    print(f"artist    : {entry.artist}")
    print(f"url       : {entry.url}")
    print(f"url_diff  : {entry.url_diff}")
    print(f"name_diff : {entry.name_diff}")
    print(f"comment   : {entry.comment}")
    print(f"note      : {entry.note}")
    print(f"total     : {entry.total}")
    print(f"judge     : {entry.judge}")
    print(f"md5       : {entry.md5}")
    print(f"sha256    : {entry.sha256}")


if __name__ == "__main__":
    main()
