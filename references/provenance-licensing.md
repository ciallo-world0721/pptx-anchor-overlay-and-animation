# 预设来源与复用许可

`scripts/vendor/` 中的 `pptx_animations.py`、`pptx_transitions.py`、`pptx_animation_presets.json` 来自 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)，MIT，Copyright (c) 2025–2026 Hugo He，许可全文见 `vendor/LICENSE-ppt-master.txt`。

预设表覆盖 PowerPoint 16 `MsoAnimEffect` 对象效果 203 个（53 入场 / 33 强调 / 64 路径 / 53 退出），已排除 4 个纯媒体效果（83/84/85/150）。

## 三条硬事实

1. `presetID` + `presetClass` 存在 ≠ 会播放；行内必须有 `set` / `animEffect` / `animMotion` / `animRot` 等行为节点。
2. 部分填充透明度写法能被解析、但放映时静默无效；透明度应走原生可见性 + `animEffect` 组合。
3. COM 数量、XML 校验、静态渲染都不能替代真实放映验收。

官方结构参考：[Working with animation (Open XML SDK)](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation)。
