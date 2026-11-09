# 定位与包完整性参考

## 锚点契约

matplotlib 导出底图时同时保存归一化锚点：

```python
anc.add_px("target", x / 100 * W * DPI, y / 47.5 * H * DPI, W * DPI, H * DPI)
anc.save("figure.anchors.json")
```

统一约定 JSON 使用 `py_from_bottom`。映射到 PPT 时只翻转一次：

```python
def to_in(anc, key, left, top, width, height):
    p = anc["points"][key]
    return (
        left + p["px"] / (anc["fig_w_in"] * anc["dpi"]) * width,
        top + (1 - p["py_from_bottom"] / (anc["fig_h_in"] * anc["dpi"])) * height,
    )
```

`animMotion pathEditMode="relative"` 的坐标按幻灯片宽高归一化：`(目标英寸-起点英寸)/slide_width_in` 与 `/slide_height_in`。

## 叠加与文本

- 底图和 PPT 叠加层不要重复绘制同一描边。
- 颜色掩膜定位避开同色文本；抗锯齿文字会污染包围盒。
- 新建文本框启用 `wrap="square"` 并预留足够高度，避免右侧溢出。

## 包与关系检查

`validate` 已覆盖 ZIP CRC、XML 可解析、关系 Target 存在、Content_Types 与重复 notesSlide 检查；自己写脚本时只需额外核对非预期丢失（嵌入字体、VBA、媒体、备注）。部件数量变化本身不是错误。python-pptx 会重新序列化包，修改前后比较部件清单，只把无法由本次编辑解释的丢失视为问题。定位演示页使用 `prs.slides[i].part.partname`，不要按 `slideN.xml` 文件名猜顺序。重复探测使用新输出路径并关闭文件句柄。
