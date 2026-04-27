from __future__ import annotations

from pathlib import Path
from queue import Queue

# 按你的实际位置调整
from om2bms.analysis.service import DifficultyAnalyzerService


class AnalyzerWorker:
    """
    BMS 难度分析后台 worker。
    """

    def __init__(self, queue: Queue):
        self.queue = queue
        self.service = DifficultyAnalyzerService()

    def run(self, input_path: str) -> None:
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
                self.queue.put(("log", "未找到任何可分析的谱面文件"))
                self.queue.put(("done", "分析结束"))
                return

            total = len(files)
            self.queue.put(("log", f"共找到 {total} 个谱面文件"))

            for index, chart_path in enumerate(files, start=1):
                self.queue.put(("log", f"[{index}/{total}] 分析 {chart_path.name} ..."))
                self._analyze_one(chart_path)

            self.queue.put(("done", "分析完成"))

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

    def _analyze_one(self, chart_path: Path) -> None:
        try:
            result = self.service.analyze_path(chart_path)

            row = {
                "Chart": chart_path.name,
                "Difficulty": getattr(result, "difficulty_display", ""),
                "Level": getattr(result, "difficulty_label", ""),
                "Raw": getattr(result, "raw_score", ""),
                "Source": getattr(result, "analysis_source", ""),
            }

            self.queue.put(("result", row))
            self.queue.put(("log", f"分析完成：{chart_path.name}"))

        except Exception as exc:
            self.queue.put(("log", f"失败：{chart_path.name}: {exc}"))
