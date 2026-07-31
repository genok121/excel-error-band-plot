"""
从 Excel 读取两组数据，绘制分时段折线、误差带和时段平均线。

推荐的 Excel 长表结构（一行对应一个观测点）：

    period  time  group1_mean  group1_error  group2_mean  group2_error
    2512      1       0.0          0.2           0.0          0.3
    2512      3       0.1          0.3           0.2          0.2
    ...

如果已有误差带上下限，也可以使用 group1_lower/group1_upper 等列，
并在 GROUPS 中把 band_mode 改为 "bounds"。
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FuncFormatter, MultipleLocator


# =============================================================================
# 可调参数区
# =============================================================================

# Excel 输入与图片输出
INPUT_FILE = Path("data.xlsx")
SHEET_NAME = 0                    # 工作表名称，如 "Sheet1"；0 表示第一个工作表
OUTPUT_FILE = Path("error_band_plot.png")
OUTPUT_DPI = 300

# Excel 中用于划分时段和表示时点的列名
PERIOD_COLUMN = "period"          # 例如 2512、2601、2602
TIME_COLUMN = "time"              # 例如 1、3、5……23

# 时段顺序。设为 None 时，按 Excel 中首次出现的顺序排列
PERIOD_ORDER = None               # 例如 ["2512", "2601", "2602"]

# 每个时段之间额外留出的横向空隙，单位约等于相邻数据点间距
PERIOD_GAP = 0.8

# 两组数据配置；不想显示某组时，把 enabled 改为 False
GROUPS = [
    {
        "enabled": True,
        "label": "Group 1",
        "mean_column": "group1_mean",

        # "error": mean ± error；"bounds": 直接读取 lower/upper
        "band_mode": "error",
        "error_column": "group1_error",
        "lower_column": "group1_lower",
        "upper_column": "group1_upper",
        "error_scale": 1.0,       # 例如 1.96 可把标准误近似转为 95% CI

        # "series": 绘制逐点均值曲线
        # "period_mean": 每个时段绘制一条水平平均线
        # "both": 同时绘制逐点曲线和水平时段平均线
        "line_mode": "series",

        "line_color": "#F36F21",
        "band_color": "#F7A57A",
        "band_alpha": 0.40,
        "line_width": 2.2,
        "line_style": "-",
        "marker": None,           # 例如 "o"、"s"；None 表示无标记
        "marker_size": 4.0,

        # 水平时段平均线的样式（line_mode 为 period_mean/both 时使用）
        "period_mean_color": "#F36F21",
        "period_mean_width": 2.0,
        "period_mean_style": "--",
    },
    {
        "enabled": True,
        "label": "Group 2",
        "mean_column": "group2_mean",
        "band_mode": "error",
        "error_column": "group2_error",
        "lower_column": "group2_lower",
        "upper_column": "group2_upper",
        "error_scale": 1.0,
        "line_mode": "series",
        "line_color": "#2878B5",
        "band_color": "#75AADB",
        "band_alpha": 0.25,
        "line_width": 2.2,
        "line_style": "-",
        "marker": None,
        "marker_size": 4.0,
        "period_mean_color": "#2878B5",
        "period_mean_width": 2.0,
        "period_mean_style": "--",
    },
]

# 画布与坐标轴
FIGURE_SIZE = (10.0, 4.6)
Y_LIMITS = (-1, 40)               # 设为 None 时自动确定，例如 None
Y_MAJOR_STEP = 10                 # 设为 None 时使用 Matplotlib 自动刻度
SHOW_PLUS_SIGN = True             # 把正值显示成 +10、+20……
X_TICK_ROTATION = 0

# 字体设置。中文可尝试 "Microsoft YaHei" 或 "SimHei"
FONT_FAMILY = "Arial"
FONT_SIZE = 10

# 图框、网格、图例
AXIS_LINE_WIDTH = 1.4
AXIS_COLOR = "black"
SHOW_GRID = False
GRID_COLOR = "#D9D9D9"
GRID_ALPHA = 0.6
SHOW_LEGEND = True
LEGEND_LOCATION = "upper right"
LEGEND_FRAME = False

# 时段标签和时段之间的竖直分隔线
SHOW_PERIOD_LABELS = True
PERIOD_LABEL_Y = -0.20            # 使用坐标轴高度比例，负值表示在横轴下方
SEPARATOR_BOTTOM = -0.30
SEPARATOR_TOP = 0.0
SEPARATOR_WIDTH = 0.8

# 图片边距
LEFT_MARGIN = 0.09
RIGHT_MARGIN = 0.98
TOP_MARGIN = 0.95
BOTTOM_MARGIN = 0.25


# =============================================================================
# 绘图逻辑
# =============================================================================

def _required_columns() -> set[str]:
    columns = {PERIOD_COLUMN, TIME_COLUMN}
    for group in GROUPS:
        if not group["enabled"]:
            continue
        columns.add(group["mean_column"])
        if group["band_mode"] == "error":
            columns.add(group["error_column"])
        elif group["band_mode"] == "bounds":
            columns.update([group["lower_column"], group["upper_column"]])
        else:
            raise ValueError(
                f'{group["label"]} 的 band_mode 必须是 "error" 或 "bounds"。'
            )
    return columns


def _load_data() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到 Excel 文件：{INPUT_FILE.resolve()}\n"
            "请修改代码顶部的 INPUT_FILE。"
        )

    data = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    missing = sorted(_required_columns() - set(data.columns))
    if missing:
        raise ValueError(
            "Excel 缺少以下列："
            + ", ".join(missing)
            + "\n请检查列名，或修改代码顶部的配置。"
        )

    data = data.copy()
    data[PERIOD_COLUMN] = data[PERIOD_COLUMN].astype(str)
    return data


def _periods_in_order(data: pd.DataFrame) -> list[str]:
    if PERIOD_ORDER is not None:
        return [str(value) for value in PERIOD_ORDER]
    return data[PERIOD_COLUMN].drop_duplicates().tolist()


def _build_x_positions(
    data: pd.DataFrame, periods: list[str]
) -> tuple[dict[tuple[str, object], float], list[float], list[object], list[dict]]:
    position_map: dict[tuple[str, object], float] = {}
    tick_positions: list[float] = []
    tick_labels: list[object] = []
    period_layout: list[dict] = []
    cursor = 0.0

    for period in periods:
        period_data = data.loc[data[PERIOD_COLUMN] == period]
        times = period_data[TIME_COLUMN].drop_duplicates().tolist()
        try:
            times = sorted(times)
        except TypeError:
            pass

        if not times:
            continue

        positions = np.arange(len(times), dtype=float) + cursor
        for time_value, position in zip(times, positions):
            position_map[(period, time_value)] = float(position)

        tick_positions.extend(positions.tolist())
        tick_labels.extend(times)
        period_layout.append(
            {
                "period": period,
                "start": float(positions[0]),
                "end": float(positions[-1]),
                "center": float((positions[0] + positions[-1]) / 2),
            }
        )
        cursor = float(positions[-1] + 1 + PERIOD_GAP)

    return position_map, tick_positions, tick_labels, period_layout


def _band_values(
    period_data: pd.DataFrame, group: dict
) -> tuple[np.ndarray, np.ndarray]:
    mean = pd.to_numeric(
        period_data[group["mean_column"]], errors="coerce"
    ).to_numpy(dtype=float)

    if group["band_mode"] == "error":
        error = pd.to_numeric(
            period_data[group["error_column"]], errors="coerce"
        ).to_numpy(dtype=float)
        error = error * float(group["error_scale"])
        lower = mean - error
        upper = mean + error
    else:
        lower = pd.to_numeric(
            period_data[group["lower_column"]], errors="coerce"
        ).to_numpy(dtype=float)
        upper = pd.to_numeric(
            period_data[group["upper_column"]], errors="coerce"
        ).to_numpy(dtype=float)

    return lower, upper


def _format_y_tick(value: float, _position: int) -> str:
    if SHOW_PLUS_SIGN:
        return f"{value:+g}"
    return f"{value:g}"


def make_plot(data: pd.DataFrame) -> None:
    periods = _periods_in_order(data)
    position_map, tick_positions, tick_labels, period_layout = (
        _build_x_positions(data, periods)
    )

    plt.rcParams.update(
        {
            "font.family": FONT_FAMILY,
            "font.size": FONT_SIZE,
            "axes.unicode_minus": False,
        }
    )
    fig, ax = plt.subplots(figsize=FIGURE_SIZE)

    # 每个时段单独绘制，避免误差带跨越时段间隙。
    for group in GROUPS:
        if not group["enabled"]:
            continue

        legend_used = False
        for period in periods:
            period_data = data.loc[data[PERIOD_COLUMN] == period].copy()
            if period_data.empty:
                continue

            period_data["_x"] = [
                position_map[(period, value)]
                for value in period_data[TIME_COLUMN]
            ]
            period_data = period_data.sort_values("_x")

            x = period_data["_x"].to_numpy(dtype=float)
            mean = pd.to_numeric(
                period_data[group["mean_column"]], errors="coerce"
            ).to_numpy(dtype=float)
            lower, upper = _band_values(period_data, group)

            valid_band = (
                np.isfinite(x) & np.isfinite(lower) & np.isfinite(upper)
            )
            ax.fill_between(
                x,
                lower,
                upper,
                where=valid_band,
                interpolate=True,
                color=group["band_color"],
                alpha=group["band_alpha"],
                linewidth=0,
                zorder=1,
            )

            line_mode = group["line_mode"]
            if line_mode in {"series", "both"}:
                ax.plot(
                    x,
                    mean,
                    color=group["line_color"],
                    linewidth=group["line_width"],
                    linestyle=group["line_style"],
                    marker=group["marker"],
                    markersize=group["marker_size"],
                    label=group["label"] if not legend_used else None,
                    zorder=3,
                )
                legend_used = True

            if line_mode in {"period_mean", "both"}:
                valid_mean = mean[np.isfinite(mean)]
                if valid_mean.size:
                    ax.hlines(
                        y=float(np.mean(valid_mean)),
                        xmin=float(np.min(x)),
                        xmax=float(np.max(x)),
                        color=group["period_mean_color"],
                        linewidth=group["period_mean_width"],
                        linestyle=group["period_mean_style"],
                        label=group["label"] if not legend_used else None,
                        zorder=4,
                    )
                    legend_used = True

            if line_mode not in {"series", "period_mean", "both"}:
                raise ValueError(
                    f'{group["label"]} 的 line_mode 必须是 '
                    '"series"、"period_mean" 或 "both"。'
                )

    # 横轴的时点标签
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, rotation=X_TICK_ROTATION)
    ax.tick_params(axis="x", length=0, pad=7)
    ax.tick_params(axis="y", length=0, pad=7)

    # 时段名称与分隔线
    if SHOW_PERIOD_LABELS:
        transform = ax.get_xaxis_transform()
        for layout in period_layout:
            ax.text(
                layout["center"],
                PERIOD_LABEL_Y,
                layout["period"],
                ha="center",
                va="center",
                transform=transform,
                clip_on=False,
            )

        if period_layout:
            boundaries = [period_layout[0]["start"] - 0.5]
            for left, right in zip(period_layout[:-1], period_layout[1:]):
                boundaries.append((left["end"] + right["start"]) / 2)
            boundaries.append(period_layout[-1]["end"] + 0.5)

            for boundary in boundaries:
                ax.plot(
                    [boundary, boundary],
                    [SEPARATOR_BOTTOM, SEPARATOR_TOP],
                    color=AXIS_COLOR,
                    linewidth=SEPARATOR_WIDTH,
                    transform=transform,
                    clip_on=False,
                    zorder=5,
                )

    if tick_positions:
        ax.set_xlim(min(tick_positions) - 0.5, max(tick_positions) + 0.5)
    if Y_LIMITS is not None:
        ax.set_ylim(*Y_LIMITS)
    if Y_MAJOR_STEP is not None:
        ax.yaxis.set_major_locator(MultipleLocator(Y_MAJOR_STEP))
    ax.yaxis.set_major_formatter(FuncFormatter(_format_y_tick))

    for spine in ax.spines.values():
        spine.set_visible(True)
        spine.set_color(AXIS_COLOR)
        spine.set_linewidth(AXIS_LINE_WIDTH)

    if SHOW_GRID:
        ax.grid(
            axis="y",
            color=GRID_COLOR,
            alpha=GRID_ALPHA,
            linewidth=0.8,
            zorder=0,
        )

    if SHOW_LEGEND:
        ax.legend(loc=LEGEND_LOCATION, frameon=LEGEND_FRAME)

    fig.subplots_adjust(
        left=LEFT_MARGIN,
        right=RIGHT_MARGIN,
        top=TOP_MARGIN,
        bottom=BOTTOM_MARGIN,
    )
    fig.savefig(
        OUTPUT_FILE,
        dpi=OUTPUT_DPI,
        bbox_inches="tight",
        facecolor="white",
    )
    plt.show()
    print(f"图片已保存到：{OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    make_plot(_load_data())
