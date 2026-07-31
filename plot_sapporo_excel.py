"""
读取“計算用（札幌）.xlsx”的宽表数据，绘制 NO/NO2 均值线和误差带。

Excel 结构：
    第 1 行：时段编号（例如 2512、2601、2602）
    第 2 行：时刻（1～24）
    A 列：NO、NO_bottom、NO_top、NO2、NO2_bottom、NO2_top 等行名

运行示例：
    python plot_sapporo_excel.py "/path/to/計算用（札幌）.xlsx"
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FuncFormatter, MultipleLocator


# =============================================================================
# 可调参数
# =============================================================================

DEFAULT_INPUT_FILE = Path("計算用（札幌）.xlsx")
DEFAULT_OUTPUT_FILE = Path("sapporo_error_band_plot.png")
SHEET_NAME = "Sheet2"

# Excel 中各类信息所在位置（从 0 开始计数）
PERIOD_ROW = 0
TIME_ROW = 1
ROW_LABEL_COLUMN = 0
FIRST_DATA_COLUMN = 1

# None 表示按照 Excel 中的顺序绘制所有时段
PERIODS_TO_PLOT = None             # 例如 ["2512", "2601", "2602"]
PERIOD_GAP = 0.8

# 横轴标签。2 表示显示 1、3、5……23；1 表示显示每个小时
TIME_TICK_STEP = 2
TIME_TICK_START = 1
X_TICK_ROTATION = 0

# 两组数据。可独立关闭、换颜色或增加水平时段平均线
GROUPS = [
    {
        "enabled": True,
        "label": "NO",
        "mean_row": "NO",
        "lower_row": "NO_bottom",
        "upper_row": "NO_top",
        "line_color": "#F36F21",
        "band_color": "#F7A57A",
        "line_width": 2.2,
        "line_style": "-",
        "band_alpha": 0.40,
        "band_edge_color": "none",
        "band_edge_width": 0.0,
        "marker": None,
        "marker_size": 4.0,
        "show_period_average": False,
        "average_color": "#B94A00",
        "average_style": "--",
        "average_width": 1.5,
    },
    {
        "enabled": True,
        "label": "NO2",
        "mean_row": "NO2",
        "lower_row": "NO2_bottom",
        "upper_row": "NO2_top",
        "line_color": "#2878B5",
        "band_color": "#75AADB",
        "line_width": 2.2,
        "line_style": "-",
        "band_alpha": 0.25,
        "band_edge_color": "none",
        "band_edge_width": 0.0,
        "marker": None,
        "marker_size": 4.0,
        "show_period_average": False,
        "average_color": "#174A70",
        "average_style": "--",
        "average_width": 1.5,
    },
]

# 将接近 0 的浮点计算残差显示为 0；设为 None 可关闭
ZERO_THRESHOLD = 1e-8

# 画布与坐标轴
FIGURE_SIZE = (11.0, 4.8)
OUTPUT_DPI = 300
Y_LIMITS = (0, 40)                 # 设为 None 时自动确定
Y_MAJOR_STEP = 10                  # 设为 None 时自动确定
SHOW_PLUS_SIGN = True
X_AXIS_LABEL = ""
Y_AXIS_LABEL = ""
TITLE = ""

# 字体、图框与图例
FONT_FAMILY = "Arial"
FONT_SIZE = 10
AXIS_COLOR = "black"
AXIS_LINE_WIDTH = 1.4
SHOW_GRID = False
GRID_COLOR = "#D9D9D9"
GRID_ALPHA = 0.6
SHOW_LEGEND = True
LEGEND_LOCATION = "upper right"
LEGEND_FRAME = False
LEGEND_COLUMNS = 1

# 时段标签和分隔线
SHOW_PERIOD_LABELS = True
PERIOD_LABEL_Y = -0.20
SEPARATOR_BOTTOM = -0.30
SEPARATOR_TOP = 0.0
SEPARATOR_WIDTH = 0.8

# 图片边距
LEFT_MARGIN = 0.08
RIGHT_MARGIN = 0.98
TOP_MARGIN = 0.94
BOTTOM_MARGIN = 0.25


# =============================================================================
# 数据读取
# =============================================================================

def _normalize_period(value: object) -> str:
    """将 Excel 中的 2512.0 转成更适合显示的 2512。"""
    if pd.isna(value):
        return ""
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        return str(int(value))
    return str(value).strip()


def _clean_small_values(values: np.ndarray) -> np.ndarray:
    values = values.astype(float, copy=True)
    if ZERO_THRESHOLD is not None:
        values[np.abs(values) < ZERO_THRESHOLD] = 0.0
    return values


def read_wide_excel(input_file: Path) -> pd.DataFrame:
    """把 Excel 宽表转换成长表，不修改原始文件。"""
    if not input_file.exists():
        raise FileNotFoundError(f"找不到 Excel 文件：{input_file.resolve()}")

    raw = pd.read_excel(input_file, sheet_name=SHEET_NAME, header=None)
    if raw.shape[0] <= TIME_ROW or raw.shape[1] <= FIRST_DATA_COLUMN:
        raise ValueError(f"{SHEET_NAME} 的数据范围过小，无法读取。")

    period_values = pd.to_numeric(
        raw.iloc[PERIOD_ROW, FIRST_DATA_COLUMN:],
        errors="coerce",
    ).ffill()
    time_values = pd.to_numeric(
        raw.iloc[TIME_ROW, FIRST_DATA_COLUMN:], errors="coerce"
    )
    valid_columns = period_values.notna() & time_values.notna()

    labels = raw.iloc[:, ROW_LABEL_COLUMN].astype(str).str.strip()
    required_rows = {
        row_name
        for group in GROUPS
        if group["enabled"]
        for row_name in (
            group["mean_row"],
            group["lower_row"],
            group["upper_row"],
        )
    }
    missing_rows = sorted(required_rows - set(labels))
    if missing_rows:
        raise ValueError(
            "Excel 缺少以下数据行："
            + ", ".join(missing_rows)
            + f"。请检查 {SHEET_NAME} 的 A 列。"
        )

    long_data = pd.DataFrame(
        {
            "period": [
                _normalize_period(value)
                for value in period_values[valid_columns]
            ],
            "time": time_values[valid_columns].to_numpy(dtype=float),
        }
    )

    for group in GROUPS:
        if not group["enabled"]:
            continue
        for key, suffix in (
            ("mean_row", "mean"),
            ("lower_row", "lower"),
            ("upper_row", "upper"),
        ):
            row_index = labels[labels == group[key]].index[0]
            values = pd.to_numeric(
                raw.iloc[row_index, FIRST_DATA_COLUMN:],
                errors="coerce",
            )[valid_columns].to_numpy(dtype=float)
            long_data[f'{group["label"]}_{suffix}'] = _clean_small_values(
                values
            )

    if PERIODS_TO_PLOT is not None:
        selected = {_normalize_period(value) for value in PERIODS_TO_PLOT}
        long_data = long_data[long_data["period"].isin(selected)].copy()

    if long_data.empty:
        raise ValueError("筛选后没有可绘制的数据。")
    return long_data


# =============================================================================
# 绘图
# =============================================================================

def _period_order(data: pd.DataFrame) -> list[str]:
    if PERIODS_TO_PLOT is not None:
        requested = [_normalize_period(value) for value in PERIODS_TO_PLOT]
        present = set(data["period"])
        return [value for value in requested if value in present]
    return data["period"].drop_duplicates().tolist()


def _build_layout(
    data: pd.DataFrame, periods: list[str]
) -> tuple[pd.DataFrame, list[dict]]:
    pieces = []
    layouts = []
    cursor = 0.0

    for period in periods:
        part = data[data["period"] == period].sort_values("time").copy()
        if part.empty:
            continue

        positions = np.arange(len(part), dtype=float) + cursor
        part["x"] = positions
        pieces.append(part)
        layouts.append(
            {
                "period": period,
                "start": float(positions[0]),
                "end": float(positions[-1]),
                "center": float((positions[0] + positions[-1]) / 2),
            }
        )
        cursor = float(positions[-1] + 1 + PERIOD_GAP)

    return pd.concat(pieces, ignore_index=True), layouts


def _format_y_tick(value: float, _position: int) -> str:
    if SHOW_PLUS_SIGN:
        return f"{value:+g}"
    return f"{value:g}"


def make_plot(
    data: pd.DataFrame,
    output_file: Path,
    show: bool = True,
) -> None:
    periods = _period_order(data)
    plot_data, layouts = _build_layout(data, periods)

    plt.rcParams.update(
        {
            "font.family": FONT_FAMILY,
            "font.size": FONT_SIZE,
            "axes.unicode_minus": False,
        }
    )
    fig, ax = plt.subplots(figsize=FIGURE_SIZE)

    for group in GROUPS:
        if not group["enabled"]:
            continue

        legend_used = False
        for period in periods:
            part = plot_data[plot_data["period"] == period]
            if part.empty:
                continue

            x = part["x"].to_numpy(dtype=float)
            mean = part[f'{group["label"]}_mean'].to_numpy(dtype=float)
            lower = part[f'{group["label"]}_lower'].to_numpy(dtype=float)
            upper = part[f'{group["label"]}_upper'].to_numpy(dtype=float)

            band_lower = np.minimum(lower, upper)
            band_upper = np.maximum(lower, upper)
            valid = (
                np.isfinite(x)
                & np.isfinite(mean)
                & np.isfinite(band_lower)
                & np.isfinite(band_upper)
            )

            ax.fill_between(
                x,
                band_lower,
                band_upper,
                where=valid,
                interpolate=True,
                color=group["band_color"],
                alpha=group["band_alpha"],
                edgecolor=group["band_edge_color"],
                linewidth=group["band_edge_width"],
                zorder=1,
            )
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

            if group["show_period_average"]:
                valid_mean = mean[np.isfinite(mean)]
                if valid_mean.size:
                    ax.hlines(
                        float(np.mean(valid_mean)),
                        xmin=float(np.min(x)),
                        xmax=float(np.max(x)),
                        color=group["average_color"],
                        linestyle=group["average_style"],
                        linewidth=group["average_width"],
                        zorder=4,
                    )

    tick_rows = plot_data[
        (
            (plot_data["time"] - TIME_TICK_START)
            % TIME_TICK_STEP
        ).abs()
        < 1e-9
    ]
    ax.set_xticks(tick_rows["x"].to_numpy())
    ax.set_xticklabels(
        [
            f"{value:g}"
            for value in tick_rows["time"].to_numpy(dtype=float)
        ],
        rotation=X_TICK_ROTATION,
    )
    ax.tick_params(axis="x", length=0, pad=7)
    ax.tick_params(axis="y", length=0, pad=7)

    if SHOW_PERIOD_LABELS and layouts:
        transform = ax.get_xaxis_transform()
        for layout in layouts:
            ax.text(
                layout["center"],
                PERIOD_LABEL_Y,
                layout["period"],
                ha="center",
                va="center",
                transform=transform,
                clip_on=False,
            )

        boundaries = [layouts[0]["start"] - 0.5]
        boundaries.extend(
            (left["end"] + right["start"]) / 2
            for left, right in zip(layouts[:-1], layouts[1:])
        )
        boundaries.append(layouts[-1]["end"] + 0.5)
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

    ax.set_xlim(
        float(plot_data["x"].min()) - 0.5,
        float(plot_data["x"].max()) + 0.5,
    )
    if Y_LIMITS is not None:
        ax.set_ylim(*Y_LIMITS)
    if Y_MAJOR_STEP is not None:
        ax.yaxis.set_major_locator(MultipleLocator(Y_MAJOR_STEP))
    ax.yaxis.set_major_formatter(FuncFormatter(_format_y_tick))

    ax.set_xlabel(X_AXIS_LABEL)
    ax.set_ylabel(Y_AXIS_LABEL)
    ax.set_title(TITLE)

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
        ax.legend(
            loc=LEGEND_LOCATION,
            frameon=LEGEND_FRAME,
            ncol=LEGEND_COLUMNS,
        )

    fig.subplots_adjust(
        left=LEFT_MARGIN,
        right=RIGHT_MARGIN,
        top=TOP_MARGIN,
        bottom=BOTTOM_MARGIN,
    )
    output_file.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(
        output_file,
        dpi=OUTPUT_DPI,
        bbox_inches="tight",
        facecolor="white",
    )
    if show:
        plt.show()
    plt.close(fig)
    print(f"图片已保存到：{output_file.resolve()}")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="绘制札幌 NO/NO2 均值线和误差带。"
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=DEFAULT_INPUT_FILE,
        help="输入 Excel 文件路径",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_FILE,
        help="输出图片路径",
    )
    parser.add_argument(
        "--no-show",
        action="store_true",
        help="生成图片后不弹出绘图窗口",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    make_plot(
        read_wide_excel(args.input),
        args.output,
        show=not args.no_show,
    )
