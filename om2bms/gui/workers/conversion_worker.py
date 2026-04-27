from __future__ import annotations

from pathlib import Path
from queue import Queue

from om2bms.utils.zip_utils import rename_zip_to_osz

from om2bms.services.conversion_service import ConversionService
from om2bms.pipeline.types import ConversionOptions


class ConversionWorker:
    """
    转谱后台 worker。
    负责把后台执行结果放入 queue。
    不直接操作 Tkinter 控件。
    """

    def __init__(self, queue: Queue):
        self.queue = queue
        self.conversion_service = ConversionService()

    def run(
        self,
        mode: str,
        input_path: str,
        output_path: str,
        options: ConversionOptions,
    ) -> None:
        try:
            if mode == "zip":
                count = rename_zip_to_osz(input_path)
                self.queue.put(("done", f"已把 {count} 个文件从 .zip 改名为 .osz。"))
                return

            if mode == "single":
                self.queue.put(("log", "开始处理单个压缩包"))
                result = self.conversion_service.convert_osz(
                    input_path,
                    output_path,
                    options,
                )
                self.queue.put(("result", result))
                message = "转换完成。" if result.conversion_success else "转换失败。"
                self.queue.put(("done", message))
                return

            archives = sorted(Path(input_path).glob("*.osz"))
            if not archives:
                raise FileNotFoundError("所选文件夹中没有找到 .osz 文件。")

            total = len(archives)
            success_count = 0

            for index, archive in enumerate(archives, start=1):
                self.queue.put(("status", f"批量任务 {index}/{total}"))
                self.queue.put(("log", f"开始处理 {archive.name}"))

                result = self.conversion_service.convert_osz(
                    str(archive),
                    output_path,
                    options,
                )

                self.queue.put(("result", result))

                if result.conversion_success:
                    success_count += 1

            self.queue.put(("done", f"批量处理完成：成功 {success_count}/{total} 个压缩包。"))

        except Exception as exc:
            self.queue.put(("error", str(exc)))
