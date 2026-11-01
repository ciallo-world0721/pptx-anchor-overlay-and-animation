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
- 底图里的公式/变量标签用 mathtext（`$H_{\mathrm{min}}$`），不要写成 `H_min`——下划线会原样印出；出图后放大检查注释区，不要只看整图。
- 新建文本框启用 `wrap="square"` 并预留足够高度，避免右侧溢出。

## 包与关系检查

```python
import zipfile
from lxml import etree

with zipfile.ZipFile(path) as zf:
    assert zf.testzip() is None
    for name in zf.namelist():
        if name.endswith((".xml", ".rels")):
            etree.fromstring(zf.read(name))
```

继续检查关系 Target 是否存在、Content_Types 是否覆盖新增部件、是否出现重复 notesSlide 关系。部件数量变化本身不是错误；重点检查非预期丢失，如嵌入字体、VBA、媒体或备注。

python-pptx 会重新序列化包。修改前后比较部件清单，只把无法由本次编辑解释的丢失视为问题。定位演示页使用 `prs.slides[i].part.partname`，不要按 slideN.xml 文件名猜顺序。

重复探测使用新输出路径并关闭文件句柄，以避免占用和旧文件混淆；不要把文件名本身当成损坏原因。
