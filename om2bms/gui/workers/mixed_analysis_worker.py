import json
import shutil
import subprocess
import tempfile
import threading
import zipfile
from pathlib import Path
from om2bms.pipeline.service import ConversionPipelineService
from om2bms.pipeline.types import ConversionOptions, DifficultyAnalysisMode

class MixedAnalysisWorker:
    def __init__(
        self,
        project_root: Path,
        runner_file: Path,
        input_file: Path,
        output_dir: Path,
        log_callback=None,
        finish_callback=None,
        enable_bms_analysis: bool = True,
        output_bms: bool = False,
        bms_output_dir: Path | None = None,
    ):
        self.project_root = Path(project_root)
        self.runner_file = Path(runner_file)
        self.input_file = Path(input_file)
        self.output_dir = Path(output_dir)

        self.log_callback = log_callback
        self.finish_callback = finish_callback

        self.process = None
        self.temp_dir = None
        self.running = False
        self.thread = None

        # BMS分析参数
        self.enable_bms_analysis = bool(enable_bms_analysis)
        self.output_bms = bool(output_bms)
        self.bms_output_dir = Path(bms_output_dir) if bms_output_dir else None

        # 内部固定参数
        self.speed_rate = 1.0
        self.cvt_flag = ""
        self.with_graph = False
        self.summary_only = True


        

    # ------------------------------------------------------------------
    # public
    # ------------------------------------------------------------------

    def start(self):
        if self.running:
            return False

        self.running = True

        self.thread = threading.Thread(
            target=self._run_safe,
            daemon=True,
        )
        self.thread.start()

        return True

    def stop(self):
        self.running = False

        if self.process is not None:
            try:
                self.process.kill()
                self.log("[GUI] 已停止")
            except Exception as exc:
                self.log(f"[GUI] 停止失败: {exc}")

    # ------------------------------------------------------------------
    # validate
    # ------------------------------------------------------------------

    def validate(self):
        if not self.runner_file.exists() or not self.runner_file.is_file():
            raise FileNotFoundError(f"固定 Node runner 不存在:\n{self.runner_file}")

        if not self.input_file.exists() or not self.input_file.is_file():
            raise FileNotFoundError("输入文件不存在")

        if self.input_file.suffix.lower() not in [".osu", ".osz"]:
            raise ValueError("输入文件必须是 .osu 或 .osz")

        self.output_dir.mkdir(parents=True, exist_ok=True)

        if not self.output_dir.exists() or not self.output_dir.is_dir():
            raise ValueError("JSON 保存路径必须是目录")

    # ------------------------------------------------------------------
    # main flow
    # ------------------------------------------------------------------
    def get_route_mode(self, data):
        if not isinstance(data, dict):
            return None

        route = data.get("route")

        if isinstance(route, dict):
            return route.get("mode")

        return data.get("route.mode")





    def should_run_bms_after_mixed(self, data):
        """
        mixed 分析完成后决定：
        1. 是否需要转换 BMS
        2. 是否需要分析 BMS

        规则：
        - 是否分析 BMS：
            enable_bms_analysis=True 且 route.mode == "RC"

        - 是否转换 BMS：
            只要需要分析 BMS，就必须转换。
            或者 output_bms=True，则无论 route.mode 是什么，都始终转换并输出。
        """
        route_mode = self.get_route_mode(data)

        should_analyze_bms = (
            self.enable_bms_analysis
            and route_mode == "RC"
        )

        should_convert_bms = (
            should_analyze_bms
            or self.output_bms
        )

        if should_convert_bms and should_analyze_bms and self.output_bms:
            return (
                True,
                True,
                'output_bms=True，始终转换输出；且 route.mode == "RC"，执行 BMS 分析',
            )

        if should_convert_bms and should_analyze_bms:
            return (
                True,
                True,
                'route.mode == "RC"，执行临时转换并进行 BMS 分析',
            )

        if should_convert_bms and not should_analyze_bms:
            return (
                True,
                False,
                f'output_bms=True，始终转换输出；但不分析，route.mode={route_mode!r}',
            )

        if not self.enable_bms_analysis and not self.output_bms:
            return (
                False,
                False,
                "BMS 分析未开启，且 BMS 输出未开启",
            )

        return (
            False,
            False,
            f'不满足 BMS 分析条件，route.mode={route_mode!r}',
        )


    def run_bms_convert_and_analysis(
        self,
        osu_file: Path,
        analyze_bms: bool,
    ) -> bool:
        bms_temp_dir = None

        try:
            if self.output_bms:
                output_dir = self.bms_output_dir or self.output_dir
                self.log("[BMS] 输出 BMS: 是")
                self.log(f"[BMS] 输出目录: {output_dir}")
            else:
                bms_temp_dir = Path(tempfile.mkdtemp(prefix="mixed_bms_"))
                output_dir = bms_temp_dir
                self.log("[BMS] 输出 BMS: 否，使用临时目录")
                self.log(f"[BMS] 临时目录: {output_dir}")

            self.log(f"[BMS] BMS 分析: {'是' if analyze_bms else '否'}")

            options = self.build_bms_conversion_options(
                osu_file=osu_file,
                analyze_bms=analyze_bms,
            )

            # self.log(f"[BMS DEBUG] enable_difficulty_analysis={options.enable_difficulty_analysis}")
            # self.log(f"[BMS DEBUG] difficulty_analysis_mode={options.difficulty_analysis_mode}")
            # self.log(f"[BMS DEBUG] resolved_analysis_mode={options.resolved_analysis_mode()}")

            service = ConversionPipelineService()

            result = service.convert_osu_file(
                osu_path=osu_file,
                output_dir=output_dir,
                options=options,
            )

            self.log("")
            self.log("[BMS] convert 服务完成")
            self.log(f"[BMS] 转换成功: {'是' if result.conversion_success else '否'}")

            if result.output_directory:
                self.log(f"[BMS] 输出目录: {result.output_directory}")

            if result.conversion_error:
                self.log(f"[BMS] 转换错误: {result.conversion_error}")

            if result.charts:
                self.log("[BMS] 转换谱面:")
                for chart in result.charts:
                    self.log(
                        f"[BMS] - {chart.chart_id} | "
                        f"{chart.source_chart_name} | "
                        f"{chart.conversion_status}"
                    )

                    if chart.output_path:
                        self.log(f"[BMS]   输出: {chart.output_path}")

                    if chart.conversion_error:
                        self.log(f"[BMS]   错误: {chart.conversion_error}")

            if result.analysis_results:
                self.log("[BMS] 分析结果:")
                for analysis in result.analysis_results:
                    self.log(
                        f"[BMS] - {analysis.chart_id}: "
                        f"{analysis.status}，enabled={analysis.enabled}"
                    )

                    if analysis.estimated_difficulty is not None:
                        self.log(f"[BMS]   estimated_difficulty: {analysis.estimated_difficulty}")

                    if analysis.difficulty_display:
                        self.log(f"[BMS]   difficulty_display: {analysis.difficulty_display}")

                    if analysis.difficulty_table:
                        self.log(f"[BMS]   difficulty_table: {analysis.difficulty_table}")

                    if analysis.error:
                        self.log(f"[BMS]   分析错误: {analysis.error}")

            if result.analysis_error:
                self.log(f"[BMS] 分析总错误: {result.analysis_error}")

            return bool(result.conversion_success)

        except Exception as exc:
            self.log(f"[BMS ERROR] {exc}")
            return False

        finally:
            if bms_temp_dir is not None and not self.output_bms:
                try:
                    shutil.rmtree(bms_temp_dir, ignore_errors=True)
                    self.log(f"[BMS] 已清理临时目录: {bms_temp_dir}")
                except Exception as exc:
                    self.log(f"[BMS] 清理临时目录失败: {exc}")


    def build_bms_conversion_options(
        self,
        osu_file: Path,
        analyze_bms: bool,
    ) -> ConversionOptions:
        return ConversionOptions(
            hitsound=True,
            bg=True,
            offset=0,
            tn_value=0.2,
            judge=3,
            output_folder_name=osu_file.stem,
            enable_difficulty_analysis=bool(analyze_bms),
            difficulty_analysis_mode=(
                DifficultyAnalysisMode.ALL
                if analyze_bms
                else DifficultyAnalysisMode.OFF
            ),
            difficulty_target_id=None,
            include_output_content=False,
        )



    def print_bms_conversion_result(self, result, analyze_bms: bool):
        self.log("")
        self.log("[BMS] convert 服务完成")
        self.log(f"[BMS] 转换成功: {'是' if result.conversion_success else '否'}")
        self.log(f"[BMS] 输出目录: {result.output_directory}")

        if getattr(result, "conversion_error", None):
            self.log(f"[BMS] 转换错误: {result.conversion_error}")

        if getattr(result, "analysis_error", None):
            self.log(f"[BMS] 分析警告: {result.analysis_error}")

        charts = getattr(result, "charts", []) or []
        analysis_results = getattr(result, "analysis_results", []) or []

        if charts:
            self.log("[BMS] 转换谱面:")
            for chart in charts:
                chart_id = getattr(chart, "chart_id", "")
                name = getattr(chart, "source_chart_name", "")
                status = getattr(chart, "conversion_status", "")
                output_path = getattr(chart, "output_path", None)
                error = getattr(chart, "conversion_error", None)

                self.log(f"[BMS] - {chart_id} | {name} | {status}")

                if output_path:
                    self.log(f"[BMS]   输出: {output_path}")

                if error:
                    self.log(f"[BMS]   错误: {error}")

        if not analyze_bms:
            self.log("[BMS] 本次仅转换，不执行 BMS 难度分析")
            return

        if analysis_results:
            self.log("[BMS] 分析结果:")

            for item in analysis_results:
                chart_id = getattr(item, "chart_id", "")
                status = getattr(item, "status", "")
                enabled = getattr(item, "enabled", False)

                if status == "success":
                    display = getattr(item, "difficulty_display", None)
                    estimated = getattr(item, "estimated_difficulty", None)
                    raw_score = getattr(item, "raw_score", None)
                    table = getattr(item, "difficulty_table", None)
                    label = getattr(item, "difficulty_label", None)
                    source = getattr(item, "analysis_source", None)
                    provider = getattr(item, "runtime_provider", None)

                    value = display or estimated or "未知"

                    self.log(f"[BMS] - {chart_id}: success")
                    self.log(f"[BMS]   难度: {value}")

                    if raw_score is not None:
                        self.log(f"[BMS]   rawScore: {raw_score}")

                    if table or label:
                        self.log(f"[BMS]   表: {table or ''} {label or ''}".rstrip())

                    if source:
                        self.log(f"[BMS]   source: {source}")

                    if provider:
                        self.log(f"[BMS]   runtime: {provider}")

                elif status == "failed":
                    error = getattr(item, "error", None)
                    self.log(f"[BMS] - {chart_id}: failed")

                    if error:
                        self.log(f"[BMS]   错误: {error}")

                elif status == "skipped":
                    self.log(
                        f"[BMS] - {chart_id}: skipped"
                        f"{'，enabled=False' if not enabled else ''}"
                    )

                else:
                    self.log(f"[BMS] - {chart_id}: {status}")
        else:
            self.log("[BMS] 没有分析结果")



    def _run_safe(self):
        try:
            self.validate()
            self.run()
        except Exception as exc:
            self.log(f"[GUI ERROR] {exc}")
        finally:
            self.running = False
            self.cleanup_temp_dir()

            if self.finish_callback:
                self.finish_callback()

    def run(self):
        osu_files = self.collect_osu_files()

        self.log("")
        self.log("=" * 70)
        self.log("[GUI] 混合难度分析开始")
        self.log(f"[GUI] 找到 {len(osu_files)} 个 .osu 文件")
        self.log(f"[GUI] JSON 保存目录: {self.output_dir}")
        # self.log(f"[BMS DEBUG] enable_bms_analysis={self.enable_bms_analysis}")
        # self.log(f"[BMS DEBUG] output_bms={self.output_bms}")
        # self.log(f"[BMS DEBUG] bms_output_dir={self.bms_output_dir}")
        self.log("=" * 70)

        completed = 0
        bms_completed = 0

        for idx, osu_file in enumerate(osu_files, 1):
            if not self.running:
                self.log("[GUI] 任务已中断")
                break

            self.log("")
            self.log(f"[{idx}/{len(osu_files)}] {osu_file.name}")
            self.log("-" * 70)

            ok, data = self.run_node_process(osu_file)

            # self.log(f"[BMS DEBUG] mixed ok={ok}")
            # self.log(f"[BMS DEBUG] mixed data type={type(data).__name__}")

            if ok:
                completed += 1

            route_mode = self.get_route_mode(data)
            # self.log(f"[BMS DEBUG] route.mode={route_mode!r}")

            should_convert_bms, should_analyze_bms, reason = self.should_run_bms_after_mixed(data)

            # self.log(f"[BMS DEBUG] should_convert_bms={should_convert_bms}")
            # self.log(f"[BMS DEBUG] should_analyze_bms={should_analyze_bms}")
            # self.log(f"[BMS DEBUG] reason={reason}")

            if should_convert_bms:
                self.log("")
                self.log(f"[BMS] 触发 BMS 转换: {reason}")

                bms_ok = self.run_bms_convert_and_analysis(
                    osu_file=osu_file,
                    analyze_bms=should_analyze_bms,
                )

                if bms_ok:
                    bms_completed += 1
            else:
                self.log("")
                self.log(f"[BMS] 跳过 BMS 转换/分析: {reason}")

        self.log("")
        self.log("=" * 70)

        if self.running:
            self.log(f"[GUI] 全部完成，共 mixed 分析 {completed}/{len(osu_files)} 个谱面")
            self.log(f"[BMS] 共 BMS 转换 {bms_completed}/{len(osu_files)} 个谱面")
        else:
            self.log(f"[GUI] 已停止，已完成 mixed 分析 {completed}/{len(osu_files)} 个谱面")
            self.log(f"[BMS] 已完成 BMS 转换 {bms_completed}/{len(osu_files)} 个谱面")

        self.log("=" * 70)


    # ------------------------------------------------------------------
    # osu / osz
    # ------------------------------------------------------------------

    def collect_osu_files(self):
        if self.input_file.suffix.lower() == ".osz":
            osu_files = self.extract_osz(self.input_file)

            if not osu_files:
                raise ValueError(".osz 中没有找到 .osu 文件")

            return osu_files

        return [self.input_file]

    def extract_osz(self, osz_path: Path):
        self.cleanup_temp_dir()

        self.temp_dir = tempfile.mkdtemp(prefix="osz_")

        with zipfile.ZipFile(osz_path, "r") as z:
            z.extractall(self.temp_dir)

        osu_files = list(Path(self.temp_dir).rglob("*.osu"))
        return sorted(osu_files)

    def cleanup_temp_dir(self):
        if self.temp_dir:
            try:
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            except Exception:
                pass
            finally:
                self.temp_dir = None

    # ------------------------------------------------------------------
    # node
    # ------------------------------------------------------------------

    def run_node_process(self, osu_file: Path):
        cmd = [
            "node",
            str(self.runner_file),
            str(self.project_root),
            str(osu_file),
            str(self.speed_rate),
            self.cvt_flag,
            "true" if self.with_graph else "false",
        ]

        stdout_chunks = []

        try:
            # self.log(f"[GUI] 执行: {' '.join(cmd)}")

            self.process = subprocess.Popen(
                cmd,
                cwd=str(self.project_root),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )

            stdout_thread = threading.Thread(
                target=self.read_stream,
                args=(self.process.stdout, "[stdout]", stdout_chunks),
                daemon=True,
            )

            stderr_thread = threading.Thread(
                target=self.read_stream,
                args=(self.process.stderr, "[log]", None),
                daemon=True,
            )

            stdout_thread.start()
            stderr_thread.start()

            exit_code = self.process.wait()

            stdout_thread.join(timeout=1)
            stderr_thread.join(timeout=1)

            full_stdout = "".join(stdout_chunks)

            data = self.save_json_and_print_summary(
                stdout_text=full_stdout,
                osu_file=osu_file,
                exit_code=exit_code,
            )

            return exit_code == 0, data


        except FileNotFoundError:
            self.log("[GUI ERROR] 找不到 node 命令，请确认 Node.js 已安装并加入 PATH")
            return False

        except Exception as exc:
            self.log(f"[GUI ERROR] {exc}")
            return False

        finally:
            self.process = None

    def read_stream(self, stream, prefix, collect):
        try:
            for line in stream:
                if collect is not None:
                    collect.append(line)
                # summary_only 模式下忽略 stderr / 普通 log
                else:
                    if not self.summary_only:
                        self.log(f"{prefix} {line.rstrip()}")
        except Exception:
            pass


    # ------------------------------------------------------------------
    # json
    # ------------------------------------------------------------------

    def extract_json_from_stdout(self, stdout_text):
        text = stdout_text.strip()

        if not text:
            return None

        try:
            return json.loads(text)
        except Exception:
            pass

        lines = text.splitlines()

        for line in reversed(lines):
            line = line.strip()

            if line.startswith("{") and line.endswith("}"):
                try:
                    return json.loads(line)
                except Exception:
                    pass

        first = text.find("{")
        last = text.rfind("}")

        if first >= 0 and last > first:
            candidate = text[first:last + 1]

            try:
                return json.loads(candidate)
            except Exception:
                return None

        return None

    def make_unique_json_path(self, osu_file: Path):
        self.output_dir.mkdir(parents=True, exist_ok=True)

        base_name = osu_file.stem
        json_path = self.output_dir / f"{base_name}.json"

        if not json_path.exists():
            return json_path

        idx = 1

        while True:
            candidate = self.output_dir / f"{base_name}_{idx}.json"

            if not candidate.exists():
                return candidate

            idx += 1

    def save_json_and_print_summary(self, stdout_text, osu_file: Path, exit_code):
        data = self.extract_json_from_stdout(stdout_text)

        if not data:
            self.log("[GUI] 无法解析 JSON，未保存")

            if stdout_text.strip():
                self.log("[GUI] stdout 内容:")
                self.log(stdout_text.strip())

            return None

        summary_text = data.get("summaryText")

        json_data = dict(data)
        json_data.pop("summaryText", None)

        json_data["_gui"] = {
            "sourceOsu": str(osu_file),
            "exitCode": exit_code,
            "runner": str(self.runner_file),
        }

        json_path = self.make_unique_json_path(osu_file)

        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)

            if data.get("ok"):
                self.log(f"[GUI] JSON 已保存: {json_path}")
            else:
                self.log(f"[GUI] 分析失败 JSON 已保存: {json_path}")
                self.log(f"[GUI] 错误: {data.get('error')}")

        except Exception as exc:
            self.log(f"[GUI] 保存 JSON 失败: {exc}")

            # 即使保存失败，也返回 data，让后续 BMS 判断仍然可以继续
            return data

        if summary_text:
            self.log("")
            self.log(summary_text)

        route_mode = self.get_route_mode(data)
        self.log(f"[GUI] route.mode: {route_mode!r}")

        return data
    # ------------------------------------------------------------------
    # callback
    # ------------------------------------------------------------------
    def log(self, text=""):
        if self.log_callback:
            self.log_callback(text)
