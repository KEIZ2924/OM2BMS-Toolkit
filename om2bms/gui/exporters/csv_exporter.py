from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable


def export_dict_rows_to_csv(
    export_path: str | Path,
    rows: Iterable[dict],
    fieldnames: list[str],
) -> None:
    """
    按 fieldnames 导出 list[dict] 为 CSV。
    """
    with open(export_path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()

        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})


def export_mapped_dict_rows_to_csv(
    export_path: str | Path,
    rows: Iterable[dict],
    columns: list[tuple[str, str]],
) -> None:
    """
    columns: [(内部key, CSV表头), ...]
    用于转换结果导出。
    """
    fieldnames = [label for _, label in columns]

    with open(export_path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()

        for row in rows:
            writer.writerow({
                label: row.get(key, "")
                for key, label in columns
            })
