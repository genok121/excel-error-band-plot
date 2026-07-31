# Excel Error Band Plot

从 Excel 读取两组时间序列，绘制分时段折线、半透明误差带和时段平均线。

主要特点：

- 两组误差带独立绘制，不会像堆积面积图一样相互累加
- 支持“均值 ± 误差”或直接读取误差带上下限
- 支持逐点均值曲线、分时段水平平均线，或同时显示两者
- 支持多时段横轴、时段标签和竖直分隔线
- 颜色、透明度、线型、坐标范围和图片分辨率均可配置

## Excel 数据格式

推荐使用长表格式：

| period | time | group1_mean | group1_error | group2_mean | group2_error |
|---|---:|---:|---:|---:|---:|
| 2512 | 1 | 0.0 | 0.2 | 0.0 | 0.3 |
| 2512 | 3 | 0.1 | 0.3 | 0.2 | 0.2 |
| 2601 | 1 | 0.0 | 0.1 | 0.0 | 0.2 |

如果 Excel 已经包含上下限列，可以在脚本配置区将 `band_mode` 改为
`"bounds"`，并设置相应的 `lower_column` 和 `upper_column`。

## 安装

```bash
python -m pip install -r requirements.txt
```

## 使用

1. 将 Excel 文件放在脚本所在目录。
2. 在 `excel_error_band_plot.py` 顶部配置区修改 `INPUT_FILE`、工作表名称和列名。
3. 运行：

```bash
python excel_error_band_plot.py
```

图片默认保存为 `error_band_plot.png`。

## 常用配置

- `band_alpha`：误差带透明度
- `line_color`、`band_color`：折线和误差带颜色
- `line_width`、`line_style`：线宽和线型
- `line_mode="series"`：逐点均值曲线
- `line_mode="period_mean"`：每个时段绘制水平平均线
- `line_mode="both"`：同时显示两类均值线
- `Y_LIMITS`、`Y_MAJOR_STEP`：纵轴范围和主刻度间隔
- `PERIOD_GAP`：不同时段之间的间距
- `FIGURE_SIZE`、`OUTPUT_DPI`：图片尺寸和输出分辨率
