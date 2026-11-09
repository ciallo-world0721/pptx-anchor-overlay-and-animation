# 验证、读回与放映清单

## COM 读回

```bash
$PY scripts/pptx_com_audit.py animated.pptx --report com-audit.json
```

输出每页：形状数、`MainSequence` 条数、交互序列条数与内容、每条效果的 shape / effect_type / trigger / duration / delay、页面切换 entry 与时长。

规则：

- `MainSequence.Count` 不含交互序列效果；`interactiveSeq` 要单独数，否则会误判「按钮页没有动画」。
- 用 `DispatchEx` 新建实例，`ReadOnly=True` + `WithWindow=False` 打开，结束时只 `Quit()` 自己启动的实例。
- 读回数字与设计一致，只说明 PowerPoint 认这棵树；节奏、遮挡、卡顿仍要看放映。

典型故障指纹：XML 里有 `<p:timing>` 但 COM 报 `Count = 0` → 主序列缺少 `<p:prevCondLst>` / `<p:nextCondLst>`，timing 树被静默丢弃。

## 放映复核

1. 动画窗格里的顺序、开始方式、效果选项与实际播放一致。
2. 并行与串行节拍符合设计，没有被自动翻页提前截断。
3. 终态位置、尺寸、旋转正确，无遮挡与裁切。
4. 触发器只作用于目标对象，旧序列没有残留。
5. Morph 相邻页按预期配对，没有误配其他对象。

旧版 PowerPoint 不支持 p14 回弹或 p159 Morph 时，应分别降级为无回弹与 fade，而不是整页失效。
