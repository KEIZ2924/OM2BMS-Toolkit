from __future__ import annotations

from pathlib import Path


def rename_zip_to_osz(path: str) -> int:
    """
    把 .zip 文件改名为 .osz。

    path 可以是：
    1. 单个 .zip 文件
    2. 一个文件夹，处理其直属 .zip 文件

    返回成功改名数量。
    """
    source = Path(path)

    if source.is_file():
        if source.suffix.lower() != ".zip":
            return 0

        target = source.with_suffix(".osz")
        if target.exists():
            return 0

        source.rename(target)
        return 1

    if not source.is_dir():
        return 0

    count = 0

    for zip_file in sorted(source.iterdir()):
        if not zip_file.is_file():
            continue

        if zip_file.suffix.lower() != ".zip":
            continue

        target = zip_file.with_suffix(".osz")
        if target.exists():
            continue

        zip_file.rename(target)
        count += 1

    return count
