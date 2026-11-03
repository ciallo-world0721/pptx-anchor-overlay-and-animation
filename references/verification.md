# 验证、读回与放映清单

## 六层验证

按顺序执行，任何一层通过都不能替下一层背书；每层能证明/不能证明什么见 `SKILL.md` §5。

1. **输入** — `inspect deck.pptx`：页面顺序、形状名/ID、已有时间线与切换。
2. **包** — `validate out.pptx`：ZIP/OPC、XML、关系、时间节点、形状引用、MCE 前缀。
3. **语义** — `validate out.pptx --spec plan.json`：重新读盘，逐行核对效果、目标、触发、偏移、时长、重复、反向、平滑、回弹、终点状态与自定义路径。
4. **静态渲染** — PowerPoint `Export` 或 LibreOffice 导出 PDF/PNG：只证明页面能渲染。
5. **COM 读回** — `pptx_com_audit.py`：PowerPoint 自己解析这棵时间树。
6. **放映 + 保存往返** — 用户点一遍；另存副本后重跑 `inspect` / `validate`，确认动画未被重写或丢失。

## COM 读回

```bash
$PY scripts/pptx_com_audit.py animated.pptx --report com-audit.json
```

输出每页：形状数、`MainSequence` 条数、交互序列条数与内容、每条效果的 shape / effect_type / trigger / duration / delay、页面切换 entry 与时长。

规则：

- `Effect` 对象**没有** `TriggerType`，必须读 `fx.Timing.TriggerType`。
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

## 流畅度调节

- **同步运动先清零默认平滑**：多对象要严格同时加速/停止时统一 `accelerate` / `decelerate`，不依赖各预设默认值。
- **回弹只给单个焦点对象**：装配、对齐、多对象同步用平滑末端；回弹越界会造成视觉错位。
- **错峰用 delay，不复制短动画**：一个路径加 0.1–0.3 s 错峰，比切成十几片更稳、更好改。
- **串行等最长分支**：后续行动画用 `after-previous`，不要手算固定等待。
- **循环必须算墙钟**：`auto_reverse` 翻倍、`repeat_count` 乘算；自动翻页时间不早于最长分支。
- **Morph 保持对象类型稳定**：换类型就改用对象动画或单独过渡。
