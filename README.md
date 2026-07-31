# Excel Error Band Plot

从 Excel 读取两组时间序列，绘制分时段折线、半透明误差带和时段平均线。

主要特点：

- 两组误差带独立绘制，不会像堆积面积图一样相互累加
- 支持“均值 ± 误差”或直接读取误差带上下限
- 支持逐点均值曲线、分时段水平平均线，或同时显示两者
- 支持多时段横轴、时段标签和竖直分隔线
- 颜色、透明度、线型、坐标范围和图片分辨率均可配置

## 札幌宽表数据

`plot_sapporo_excel.py` 已适配“計算用（札幌）.xlsx”的结构：

- 第 1 行：2512、2601、2602 等时段
- 第 2 行：1～24 时
- A 列：`NO`、`NO_bottom`、`NO_top`、`NO2`、`NO2_bottom`、`NO2_top`

运行：

```bash
python plot_sapporo_excel.py "/path/to/計算用（札幌）.xlsx"
```

指定输出位置且不弹出窗口：

```bash
python plot_sapporo_excel.py "/path/to/計算用（札幌）.xlsx" \
  -o sapporo_error_band_plot.png --no-show
```

脚本顶部可以调节两组颜色、误差带透明度、线宽、纵轴范围、时段间距、小时标签密度和图例。将 `show_period_average` 改为 `True`，可为对应组增加每个时段的水平平均线。

## 通用长表数据

`excel_error_band_plot.py` 适用于以下长表格式：

| period | time | group1_mean | group1_error | group2_mean | group2_error |
|---|---:|---:|---:|---:|---:|
| 2512 | 1 | 0.0 | 0.2 | 0.0 | 0.3 |
| 2512 | 3 | 0.1 | 0.3 | 0.2 | 0.2 |
| 2601 | 1 | 0.0 | 0.1 | 0.0 | 0.2 |

如果 Excel 已经包含上下限列，可以在通用脚本配置区将 `band_mode` 改为 `"bounds"`。

## 安装

```bash
python -m pip install -r requirements.txt
```

## 常用配置

- `band_alpha`：误差带透明度
- `line_color`、`band_color`：折线和误差带颜色
- `line_width`、`line_style`：线宽和线型
- `Y_LIMITS`、`Y_MAJOR_STEP`：纵轴范围和主刻度间隔
- `PERIOD_GAP`：不同时段之间的间距
- `FIGURE_SIZE`、`OUTPUT_DPI`：图片尺寸和输出分辨率

Excel 数据文件与生成的 PNG 默认由 `.gitignore` 排除，避免误传研究数据。
