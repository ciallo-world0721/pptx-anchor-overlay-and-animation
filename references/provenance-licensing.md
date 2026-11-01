# 预设来源与复用许可

## 复用

`scripts/vendor/` 中的 `pptx_animations.py`、`pptx_transitions.py`、`pptx_animation_presets.json` 来自 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)，MIT，Copyright (c) 2025–2026 Hugo He，许可全文见 `vendor/LICENSE-ppt-master.txt`。

本技能在其上修正了两处行为：

- 显式零平滑（`accel=0` / `decel=0`）在回读时不再被当作「未设置」并重新套用模板默认缓冲。
- 非零回弹会自动删除路径模板继承的 `decel`，避免平滑末端与回弹同时存在。

预设表的 `source` 记录为 Microsoft PowerPoint 16 的 `MsoAnimEffect` 对象效果：203 个，已排除 4 个纯媒体效果（83/84/85/150）。分项为 53 入场 / 33 强调 / 64 路径 / 53 退出。

## 只吸收事实、未复制代码的来源

| 仓库 | 许可 | 取用 |
|---|---|---|
| BramAlkema/pptx-animation-skill | AGPL-3.0 / 商业双许可 | 只取验证方法：`fill.opacity` 可能解析但放映静默失效；预设 ID 不等于完整播放行为 |
| Yusong-Enceladus/pptx-animation-skill | MIT | 只确认 mainSeq 与点击组的最小结构 |
| caseyh-peach/AnimatedPowerPointCreator | 未标注 | 反例：只有 `presetID` / `presetClass` 没有行为节点 = 不播放 |
| sbroenne/mcp-server-powerpoint | MIT | COM `MainSequence.AddEffect` 的用法事实 |
| iOfficeAI/OfficeCLI | — | Morph 相邻页匹配与静态渲染的局限 |
| wellkilo/SlideCraft | MIT | 边界：网页动画不能当原生能力 |
| anthropics/skills | Proprietary / source-available | 未复制；确认其无完整对象动画写入方案 |
| grapeot/pptx.skill | — | 渲染闭环；明确把复杂动画列为不支持 |

官方结构参考：[Working with animation (Open XML SDK)](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation)。

## 三条硬事实

1. `presetID` + `presetClass` 存在 ≠ 会播放；行内必须有 `set` / `animEffect` / `animMotion` / `animRot` 等行为节点。
2. 部分填充透明度写法能被解析、但放映时静默无效；透明度应走原生可见性 + `animEffect` 组合。
3. COM 数量、XML 校验、静态渲染都不能替代真实放映验收。
