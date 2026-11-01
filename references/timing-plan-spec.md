# JSON 参数契约

## 顶层与页面

```text
{
  version: 1,                 // 必填，当前仅支持 1
  kiosk: false,               // 可选，设置 kiosk 循环放映
  slides: [SlideSpec, ...]    // 必填，非空，页码不得重复
}

SlideSpec {
  slide: 正整数,               // 按 presentation.xml 的真实页面顺序，从 1 开始
  replace_existing: false,    // 可选，仅影响本页已有对象时间线
  animations: [Animation, ...],
  transition: Transition       // 可选
}
```

未知字段一律报错。未列出的页面和未修改的 package part 保持原样。`replace_existing` 不会清除页面切换，页面切换只由本页的 `transition` 明确控制。

## 目标解析

- `target`：唯一形状名称或正整数形状 ID，必填。
- `trigger_shape`：唯一形状名称或正整数形状 ID，可选。
- 名称匹配区分大小写；重复名称、缺失目标和相同的目标/触发形状都会报错。
- 工具不按 `slide1.xml` 文件名猜页码，而是解析 `presentation.xml → relationships → slide part`。
- 可动画形状包括普通形状、图片、组合、连接符和图形框架。根 `spTree` 不是合法目标。

先用 CLI 检查：

```bash
python3 scripts/native_ppt_anim.py inspect input.pptx
```

## Animation 字段

```text
{
  target: "shape-name" | 2,
  effect: "entrance_fade" | ... | "custom_motion",
  duration: 0.5,               // 正秒数；省略时使用预设作者时长
  delay: 0,                   // 非负秒数，高级层换算为 delay_ms
  trigger: "after-previous", // on-click | with-previous | after-previous
  trigger_shape: "button",   // 需要 trigger=on-click
  effect_options: {},         // 由 describe 查询，逐效果校验

  repeat_count: 2,            // 正数，可与小数次数；与 repeat_duration 互斥
  repeat_duration: 3.0,       // 正秒数，总重复时长
  auto_reverse: false,
  rewind: false,              // true 对应播放后移除终点状态
  accelerate: 0,              // 0–1
  decelerate: 0,              // 0–1，accelerate + decelerate <= 1
  bounce_end: 0,              // 0–1，仅插值行为；与非零 decelerate 互斥
  restart: "when-not-active",// always | when-not-active | never
  after_effect: "none",      // none | dim | hide | hide-on-next-click

  path: "M 0 0 L 0.2 0 E",  // 仅 custom_motion，归一化坐标
  path_px: "M 0 0 L 200 0 E"// 仅 custom_motion，96 DPI 像素坐标
}
```

`duration` 对不可缩放的原生预设不会改写内部行为时长。`emphasis_change_font` 必须提供 `effect_options.font_name`；其他必选参数也通过 `describe` 查明。

当前高级接口不接受 `sound`，避免写入缺少关系或媒体部分的半完成动画。

## 触发与调度

普通行动画进入 `mainSeq`，形状触发动画进入独立 `interactiveSeq`。形状触发器不参与普通主序列的先后调度。

对同一普通点击组：

- `on-click` 开启新组，并把节拍起点和此前最晚结束时间归零。
- `with-previous` 的 `start = beat_start + delay`。
- `after-previous` 的 `beat_start = timeline_end`，然后再加自己的 `delay`。
- 每行结束后更新 `timeline_end = max(timeline_end, start + playback_duration)`。

因此：

```text
A: after-previous, duration 0.65, delay 0     → start 0
B: with-previous,  duration 0.65, delay 0.12  → start 0.12
C: with-previous,  duration 0.65, delay 0.36  → start 0.36
D: after-previous, duration 0.30              → start 1.01
```

D 从 1.01 秒开始，因为最长分支 C 在 `0.36 + 0.65 = 1.01` 秒结束。

实际播放时长：

```text
one_play = duration × (2 if auto_reverse else 1)
repeat_count 存在：playback = one_play × repeat_count
repeat_duration 存在：playback = repeat_duration
否则：playback = one_play
```

`after-previous` 使用上述实际时长，不只看单程 `duration`。

## 平滑、回弹与终点状态

- OOXML 的 `accel` 和 `decel` 使用 0–100000，高级接口使用 0–1。
- 显式零值会删除模板默认平滑，用于匀速和严格同步。
- 非零回弹使用 `mc:AlternateContent`：`Choice Requires="p14"` 保留回弹，`Fallback` 删除回弹但保留其余动画。
- 非零 `bounce_end` 自动删除 `decel`，包括路径预设继承的默认平滑末端。
- 默认 `rewind: false` 保持终点；`true` 对应 `fill="remove"`，播放后可能回到起点。

## 自定义运动路径

`custom_motion` 使用合法原生 `path_right` 行作为载体，再替换其唯一 `p:animMotion@path`。不是预置换叫 `path_custom` 的不存在预设。

支持：

- `M x y`：一个初始移动点，必须恰好一对坐标。
- `L x y`：直线，可连续追加多对坐标。
- `C x1 y1 x2 y2 x y`：三次贝塞尔曲线，可连续追加多组六坐标。
- `E`：结束，必须位于末尾。
- 数字使用普通十进制，不接受科学计数法、Q、弧线、H/V 或隐式相对命令。

`path` 的 x、y 分别除以幻灯片宽和高。`path_px` 先按 96 DPI 将像素乘以 9525 转为 EMU，再除以 `p:sldSz` 的宽和高。常用 1280×720 设计坐标因此可直接用于 16:9 页面。

默认 `pathEditMode="relative"`；`effect_options: {"relative": false}` 使用 `fixed`。所有 MCE 分支中的真实路径都必须与计划完全一致。验证器另用标准路径的内存副本核对其他原生行为，防止自定义路径绕过行为树校验。

## Transition 字段

```text
{
  effect: "fade",             // 原生切换名，或 morph；none 表示移除视觉切换
  duration: 0.5,              // 正秒数
  effect_options: {},         // 用 describe-transition 查询
  advance_after: 2.0,          // 非负秒数，可选
  advance_on_click: true       // 可选
}
```

- 有 `effect`：明确替换本页视觉切换。
- 无 `effect` 但有翻页字段：保留视觉切换，仅修改推进方式。
- 只有 `advance_after`：自动推进并关闭单击推进。
- 同时有 `advance_after` 和 `advance_on_click: true`：两者均可推进。
- 只有 `advance_on_click: true`：改为仅单击。
- `advance_on_click: false` 必须搭配 `advance_after`。

任何 `advance_after` 都会设置 `ppt/presProps.xml` 的 `p:showPr@useTimings=1`，并验证相关关系和内容类型。

## Morph

在目标页设置：

```json
{"slide": 2, "transition": {"effect": "morph", "duration": 0.8}}
```

源页和目标页必须相邻，并共享至少一个顶层 `!!key` 名称。名称在单页唯一，两侧对象 OOXML 类型必须相同。目标页必须为 `p159:morph option="byObject"`，旧渲染器可回退为 fade。

示例：

```text
slide 1: 形状名称 !!hero
slide 2: 形状名称 !!hero，transition=morph
```

## 错误与原子性

以下情况不会发布输出：

- JSON 未知字段、未知效果或非法参数组合。
- 页码越界、重复页面计划、重复/缺失形状名称。
- 未明确允许却遇到已有时间线。
- 时间节点 ID 重复、引用不存在的形状、MCE 前缀未绑定。
- 自定义路径不一致、原生行为树变化、时长/偏移/触发回读不一致。
- ZIP CRC、XML、内部关系目标、Morph 配对或自动翻页元数据错误。

输出先写同目录临时文件，验证失败则删除；成功后原子替换目标。输入与输出可以是同一路径，但仍建议保留原文件副本。
