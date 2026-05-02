from __future__ import annotations

from copy import deepcopy
from typing import Any
from om2bms.result_processor.pattern_processor import build_pattern_fields
import json
from pathlib import Path
from typing import Any
import math 

INTERVALS_PATH = Path("config/intervals.json")

LEVEL_PREFIX_SYMBOLS = {
    "RC": "★",
    "LN": "◆",
}

def load_intervals() -> dict[str, list[list[Any]]]:
    with INTERVALS_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


INTERVALS_CONFIG = load_intervals()


def get_by_path(data: Any, path: str, default: Any = None) -> Any:
    """
    支持 dict 和 list 的 dot path 取值。

    Example:
        bms.charts.0.bms_summary.song_info.title
    """

    current: Any = data

    for part in path.split("."):
        if isinstance(current, dict):
            if part not in current:
                return default
            current = current[part]
            continue

        if isinstance(current, list):
            try:
                index = int(part)
            except ValueError:
                return default

            if index < 0 or index >= len(current):
                return default

            current = current[index]
            continue

        return default

    return current



def first_not_none(*values: Any, default: Any = None) -> Any:
    for value in values:
        if value is not None:
            return value

    return default


def to_float(value: Any) -> float | None:
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None


def round_number(value: Any, digits: int = 2) -> float | None:
    number = to_float(value)

    if number is None:
        return None

    return round(number, digits)


def to_int(value: Any) -> int | None:
    number = to_float(value)

    if number is None:
        return None

    return int(round(number))

def format_percent(value: Any, digits: int = 2) -> str | None:
    number = to_float(value)

    if number is None:
        return None

    percent = round(number * 100, digits)

    if percent.is_integer():
        return f"{int(percent)}%"

    return f"{percent}%"


def format_ms_to_min_sec(value: Any) -> str | None:
    """
    把毫秒转换成 分钟:秒。

    Example:
        189964.528 -> "3:10"
    """
    number = to_float(value)

    if number is None:
        return None

    total_seconds = int(round(number / 1000))
    minutes = total_seconds // 60
    seconds = total_seconds % 60

    return f"{minutes}:{seconds:02d}"


def merge_title_subtitle(title: Any, subtitle: Any) -> str | None:
    """
    合并 title 和 subtitle。

    Example:
        title = "Anata Ga Mawaru"
        subtitle = "[Extreme]"
        -> "Anata Ga Mawaru [Extreme]"
    """
    if title is None and subtitle is None:
        return None

    title_text = str(title).strip() if title is not None else ""
    subtitle_text = str(subtitle).strip() if subtitle is not None else ""

    if title_text and subtitle_text:
        return f"{title_text} {subtitle_text}"

    if title_text:
        return title_text

    if subtitle_text:
        return subtitle_text

    return None



def normalize_route_type(route_mode: Any) -> str | None:
    """
    把 route.mode 转成最终输出用的 type。
    """
    if not isinstance(route_mode, str):
        return None

    mode = route_mode.upper()

    mapping = {
        "RC": "RC",
        "LN": "LN",
        "HB": "HB",
        "MIX": "MIX",
    }

    return mapping.get(mode, mode)


def build_dan_estimate(data: dict[str, Any]) -> str | None:
    """
    dan_estimate 只取 compact.estDiff。
    """

    value = get_by_path(data, "compact.estDiff")

    if value is None:
        return None

    return str(value)


def build_derived_fields(data: dict[str, Any]) -> dict[str, Any]:
    """
    从 raw_data 中构建最终输出前需要的派生字段。
    """

    title = get_by_path(data, "bms.charts.0.bms_summary.song_info.title")
    subtitle = get_by_path(data, "bms.charts.0.bms_summary.song_info.subtitle")
    artist = get_by_path(data, "bms.charts.0.bms_summary.song_info.artist")

    keys = get_by_path(data, "compact.columnCount")
    route_mode = get_by_path(data, "route.mode")
    sunny_sr = get_by_path(data, "compact.star")
    ln_ratio = get_by_path(data, "compact.lnRatio")

    total = get_by_path(data, "bms.charts.0.bms_summary.song_info.total")
    song_last_ms = get_by_path(data, "bms.charts.0.bms_summary.song_info.song_last_ms")
    total_notes = get_by_path(data, "bms.charts.0.bms_summary.song_info.total_notes")
    judge = get_by_path(data, "bms.charts.0.bms_summary.song_info.judge")

    bms_difficulty_table = get_by_path(data, "bms.charts.0.analysis.difficulty_table")
    bms_difficulty_label = get_by_path(data, "bms.charts.0.analysis.difficulty_label")
    bms_difficulty_display = get_by_path(data, "bms.charts.0.analysis.difficulty_display")
    bms_raw_score = get_by_path(data, "bms.charts.0.analysis.raw_score")

    osu_url = get_by_path(data, "bms.charts.0.osu_url")
    md5 = get_by_path(data, "bms.charts.0.bms_hashes.md5")
    sha256 = get_by_path(data, "bms.charts.0.bms_hashes.sha256")

    dan_estimate = build_dan_estimate(data)

    level = analyze_level(
        route_mode=route_mode,
        bms_difficulty_display=bms_difficulty_display,
        sunny_sr=sunny_sr,
        bms_difficulty_label = bms_difficulty_label,
        bms_raw_score=bms_raw_score
    )

    derived = {
        "title": merge_title_subtitle(title, subtitle),
        "subtitle": subtitle,
        "artist": artist,
        "keys": keys,
        "type": normalize_route_type(route_mode),

        "star": round_number(sunny_sr, 2),
        "sunny_sr": round_number(sunny_sr, 2),
        "dan_estimate": dan_estimate,
        "level": level,

        "bms_difficulty_table": bms_difficulty_table,
        "bms_difficulty_label": bms_difficulty_label,
        "bms_difficulty_display": bms_difficulty_display,
        "bms_raw_score": bms_raw_score,

        "osu_url": osu_url,
        "md5": md5,
        "sha256": sha256,

        "total_notes": total_notes,
        "song_length": format_ms_to_min_sec(song_last_ms),
        "song_last_ms": song_last_ms,
        "ln_ratio": format_percent(ln_ratio, 2),
        "total": to_int(total),
        "judge": judge,
    }

    pattern_fields = build_pattern_fields(data)
    derived.update(pattern_fields)

    return derived





def remove_none_values(data: dict[str, Any]) -> dict[str, Any]:
    """
    删除值为 None 的字段。
    """
    return {
        key: value
        for key, value in data.items()
        if value is not None
    }

def parse_percent_value(value: Any) -> float | None:
    """
    解析百分比。
    """
    if value is None:
        return None
    try:
        if isinstance(value, str):
            text = value.strip()
            if not text:
                return None
            if text.endswith("%"):
                return float(text[:-1].strip()) / 100.0
            return float(text)
        return float(value)
    except Exception:
        return None
    

def format_level_sr(value: Any) -> str:
    try:
        return f"{float(value):.2f}"
    except Exception:
        return ""


def raw_score_to_dan_score(raw_score: Any, *, clamp: bool = True) -> float | None:
    """
    将 bms_raw_score 映射到新的 dan_score。

    四次拟合：
        y = a*x^4 + b*x^3 + c*x^2 + d*x + e
    """
    if raw_score in (None, ""):
        return None

    try:
        x = float(raw_score)
    except (TypeError, ValueError):
        return None

    if clamp:
        x = max(0.0, min(1.0, x))

    dan_score = (
        10.41028042 * x ** 4
        - 23.6866 * x ** 3
        + 12.61178 * x ** 2
        + 16.08867 * x
        - 0.22817
    )
    dan_score = max(0.0, dan_score)

    return dan_score


def round_dan_score(score: Any) -> float | None:
    if score in (None, ""):
        return None

    try:
        value = float(score)
    except (TypeError, ValueError):
        return None

    # 只有 10-15 使用 .0 / .5 舍入
    if 10 <= value <= 15:
        base = math.floor(value)
        frac = value - base

        if frac < 0.25:
            return float(base)

        if frac < 0.75:
            return base + 0.5

        return float(base + 1)

    # 其他范围四舍五入到整数
    return float(math.floor(value + 0.5))


def sunny_sr_to_dan_score(
    *,
    sunny_sr: Any,
    route_mode: Any,
    intervals_config: dict[str, list[list[Any]]] = INTERVALS_CONFIG,
    clamp: bool = True,
    ndigits: int | None = 4,
) -> float | None:
    """
    根据 sunny_sr 和 route_mode，从 config/intervals 中读取 interval，
    线性插值计算数值型 dan_score。

    route_mode == "RC" 使用 RC_DAN_INTERVALS_7K
    其他情况使用 LN_DAN_INTERVALS_7K

    interval 格式：
        [sr_start, sr_end, label, dan_start, dan_end, dan_mid]
    """

    if sunny_sr in (None, ""):
        return None

    try:
        x = float(sunny_sr)
    except (TypeError, ValueError):
        return None

    mode = str(route_mode).upper() if route_mode is not None else ""

    if mode == "RC":
        intervals = intervals_config.get("RC_DAN_INTERVALS_7K", [])
    else:
        intervals = intervals_config.get("LN_DAN_INTERVALS_7K", [])

    if not intervals:
        return None

    first = intervals[0]
    last = intervals[-1]

    min_sr = float(first[0])
    max_sr = float(last[1])

    min_dan = float(first[3])
    max_dan = float(last[4])

    # 超出最小范围
    if x < min_sr:
        if not clamp:
            return None
        result = min_dan
        return round(result, ndigits) if ndigits is not None else result

    # 超出最大范围
    if x > max_sr:
        if not clamp:
            return None
        result = max_dan
        return round(result, ndigits) if ndigits is not None else result

    # 查找所在 interval 并线性插值
    for i, interval in enumerate(intervals):
        sr_start, sr_end, label, dan_start, dan_end, dan_mid = interval

        sr_start = float(sr_start)
        sr_end = float(sr_end)
        dan_start = float(dan_start)
        dan_end = float(dan_end)

        is_last = i == len(intervals) - 1

        if sr_start <= x < sr_end or is_last and sr_start <= x <= sr_end:
            if sr_end == sr_start:
                result = dan_start
            else:
                ratio = (x - sr_start) / (sr_end - sr_start)
                result = dan_start + ratio * (dan_end - dan_start)

            if ndigits is not None:
                result = round(result, ndigits)

            return result

    return None

def format_dan_score(value: float | None) -> str | None:
    """
    格式化 dan_score。

    示例：
        8.0  -> "8"
        12.0 -> "12"
        12.5 -> "12.5"
        16.0 -> "16"
    """

    if value is None:
        return None

    value = float(value)

    if value.is_integer():
        return str(int(value))

    return f"{value:.1f}"




def analyze_level(
    *,
    route_mode: Any,
    sunny_sr: Any,
    bms_difficulty_label: Any = None,
    bms_difficulty_display: Any = None,
    bms_raw_score: Any = None,
) -> str | None:
    mode = str(route_mode).upper() if route_mode is not None else ""

    # RC 使用 ★，其他默认使用 ◆
    prefix = LEVEL_PREFIX_SYMBOLS["RC"] if mode == "RC" else LEVEL_PREFIX_SYMBOLS["LN"]

    dan_score = None
    dan_score_bms =None
    dan_score_sunny =None

    # RC 优先使用 bms_raw_score
    if mode == "RC" and bms_raw_score not in (None, ""):
        dan_score_bms = raw_score_to_dan_score(bms_raw_score)
        dan_score_sunny = sunny_sr_to_dan_score(
            sunny_sr=sunny_sr,
            route_mode=route_mode,
            intervals_config=INTERVALS_CONFIG,
            clamp=True,
            ndigits=4,
        )
        dan_score =  dan_score_bms*0.6 + dan_score_sunny*0.4  

    # raw_score 不可用时，使用 sunny_sr interval 插值
    if dan_score is None:
        dan_score = sunny_sr_to_dan_score(
            sunny_sr=sunny_sr,
            route_mode=route_mode,
            intervals_config=INTERVALS_CONFIG,
            clamp=True,
            ndigits=4,
        )

    dan_score = round_dan_score(dan_score)
    dan_score_text = format_dan_score(dan_score)

    if dan_score_text is None:
        return None

    return f"{prefix}{dan_score_text}"



def prepare_final_result_source(
    raw_data: dict[str, Any],
    *,
    remove_none: bool = False,
) -> dict[str, Any]:
    """
    最终 JSON 映射前的预处理入口。

    输入:
        raw_data

    输出:
        带 derived 字段的新 dict

    注意:
        不直接修改传入的 raw_data。
    """
    processed = deepcopy(raw_data)

    derived = build_derived_fields(processed)

    if remove_none:
        derived = remove_none_values(derived)

    processed["derived"] = derived

    return processed
