from __future__ import annotations

from pathlib import Path
from queue import Queue

# 按你的实际位置调整
from om2bms.table_generator.bms2json import (
    append_missing_entry_if_needed,
    match_or_build_missing_entry,
)


class TableGenWorker:
    """
    score.json 难易度表生成后台 worker。
    """

    def __init__(self, queue: Queue):
        self.queue = queue

    def run(
        self,
        input_path: str,
        score_json_path: str,
        auto_append: bool,
        use_custom_level: bool,
        custom_level: str,
    ) -> None:
        try:
            p = Path(input_path)

            if not p.exists():
                self.queue.put(("error", f"路径不存在：{p}"))
                return

            if p.is_file():
                files = [p]
            else:
                files = self._collect_chart_files(p)

            if not files:
                self.queue.put(("log", "未找到任何可处理的谱面文件"))
                self.queue.put(("done", "处理结束"))
                return

            total = len(files)
            self.queue.put(("log", f"共找到 {total} 个谱面文件"))

            for index, chart_path in enumerate(files, start=1):
                self.queue.put(("log", f"[{index}/{total}] 处理 {chart_path.name} ..."))
                self._process_one(
                    chart_path=chart_path,
                    score_json_path=score_json_path,
                    auto_append=auto_append,
                    use_custom_level=use_custom_level,
                    custom_level=custom_level,
                )

            self.queue.put(("done", "处理完成"))

        except Exception as exc:
            self.queue.put(("error", str(exc)))

    def _collect_chart_files(self, folder: Path) -> list[Path]:
        exts = {".bms", ".bme", ".bml", ".pms"}
        files: list[Path] = []

        for p in folder.rglob("*"):
            if p.is_file() and p.suffix.lower() in exts:
                files.append(p)

        files.sort()
        return files

    def _process_one(
        self,
        chart_path: Path,
        score_json_path: str,
        auto_append: bool,
        use_custom_level: bool,
        custom_level: str,
    ) -> None:
        try:
            if auto_append:
                result = append_missing_entry_if_needed(
                    chart_path,
                    score_json_path,
                    use_custom_level=use_custom_level,
                    custom_level=custom_level,
                )
            else:
                result = match_or_build_missing_entry(
                    chart_path,
                    score_json_path,
                    use_custom_level=use_custom_level,
                    custom_level=custom_level,
                )

            song_info = result.get("song_info") or {}
            score_entry = result.get("score_entry") or {}
            generated_entry = result.get("generated_score_entry") or {}

            title = (
                score_entry.get("title")
                or generated_entry.get("title")
                or song_info.get("title")
                or chart_path.stem
            )

            artist = (
                score_entry.get("artist")
                or generated_entry.get("artist")
                or song_info.get("artist")
                or ""
            )

            level = (
                score_entry.get("level")
                or generated_entry.get("level")
                or song_info.get("level")
                or custom_level
                or ""
            )

            matched = bool(result.get("matched"))
            sha256_verified = result.get("sha256_verified")
            appended = result.get("appended", False)

            if matched:
                if sha256_verified is True:
                    status = "已匹配"
                    self.queue.put(("log", f"匹配成功：{chart_path.name}"))
                elif sha256_verified is False:
                    status = "SHA256不一致"
                    self.queue.put(("log", f"SHA256 不一致：{chart_path.name}"))
                else:
                    status = "已匹配"
                    self.queue.put(("log", f"匹配成功，JSON 未提供 SHA256，已跳过校验：{chart_path.name}"))
            else:
                status = "未收录"
                self.queue.put(("log", f"未收录：{chart_path.name}"))

            if not matched and generated_entry:
                self.queue.put(("log_json", generated_entry))

            if appended:
                self.queue.put(("log", f"已追加到 score.json：{chart_path.name}"))

            row = {
                "Chart": chart_path.name,
                "Title": title,
                "Artist": artist,
                "Level": level,
                "Status": status,
                "Appended": "是" if appended else "否",
            }

            self.queue.put(("result", row))

        except Exception as exc:
            self.queue.put(("log", f"失败: {chart_path.name}: {exc}"))
