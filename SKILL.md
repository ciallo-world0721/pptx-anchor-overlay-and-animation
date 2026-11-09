---
name: pptx-anchor-overlay-and-animation
description: "Use for native PPTX animation, Morph, COM, and overlays. Design the state contract first, then write p:timing from a JSON plan and prove it by PowerPoint read-back."
version: 3.2.0
author: Iris
license: MIT
metadata:
  hermes:
    tags: [pptx, animation, powerpoint, morph, native-animation, timing, overlay, positioning]
    related_skills: [powerpoint, nature-figure]
---

# PPTX 原生动画：状态设计、时间线与精确定位

对象动画写 `p:timing`，Morph 是 `p:transition`。两套机制，分别生成、分别验证。

## When to Use

- 给现成 `.pptx` 加原生入场/强调/退出、运动路径、点击或形状触发序列
- 多对象同时旋转、移动、缩放；相邻页之间做 Morph 状态转场
- 跨页转场（视差、画布外预置）是可选项：先判断内容是否需要空间连续，不需要就正常翻页。
- 自动翻页、循环放映、重复、自动反向、回弹
- 要把形状精确叠在自己生成的底图上
- 审查已有动画：分得清「能播」和「XML 里看着有」

不适用，直接说明而不是硬做：动画音效（接口未封装）、任意遮罩形变与逐帧物理（非原生强项）、网页/CSS 动画（另一条路线，不能冒充原生）。用户没要求动画时不要顺手加。

## 0. 四条不变量

1. **交付原生对象**：图片、形状、文本结束后仍可编辑。不得用嵌入视频、GIF、HTML 或整页截图冒充**需要点击推进的**动画；非交互的循环动效（氛围、底纹、装饰）按 PPT 制作技能的判据走烘焙媒体。
2. **先设计状态，再选工具**：先写清起始画面、结束画面、对象对应关系和讲解停留点，不套预设模板。
3. **终态必须显式成立**：最终位置、尺寸、旋转角直接写在结束页对象上。不要把「动画结束后应该归位」当终态保证。
4. **结构通过 ≠ 视觉通过**：XML 存在、`validate` 通过、COM 读回数量正确，只证明结构被识别。连续运动、卡顿、遮挡和额外停顿必须由真实放映判断。

## 1. 先选引擎

| 需求 | 首选 | 原因 |
|---|---|---|
| 两页之间多对象同时移动/缩放/旋转 | **Morph 双页状态** | 起止几何明确，PowerPoint 负责插值 |
| 页内点击序列、强调、短路径、批量生成、精确错峰 | **计划引擎** `scripts/native_ppt_anim.py` | 203 个 PowerPoint 原生预设，时间可计算，写入即回读校验 |
| 用户正开着 PowerPoint 手工微调，或要 COM 独有能力 | **COM MainSequence** | 直接操作原生时间线，但不可批量、不可复现 |
| 同一对象需绕指定枢轴刚体旋转 | 先解决枢轴（见 2.3），再谈效果 | 枢轴错了，参数怎么调都是错的 |
| 任意遮罩形变、粒子、逐帧 | 说明非原生强项，问是否接受媒体 | 不得擅自改成视频 |

默认路线：**静态页 → 计划引擎排时间线 → `validate` → PowerPoint COM 读回 → 用户放映确认**。COM 不再是主力写入路径，但仍是唯一独立的第三方 oracle。

## 2. 设计层（先写这五条，再碰工具）

### 2.1 状态契约

每个状态逐项写清：

- 讲解时保留哪些真实对象与文本；
- 每个对象的唯一身份、位置、尺寸、旋转角、层级；
- 下一状态由哪些对象继续演变，哪些进入或退出；
- 观众应在哪些状态停留。

### 2.2 对象身份与终态

- Morph 写在**目的页**；跨相邻页同一逻辑对象使用相同强制名 `!!HeroImage`、`!!MainPanel`，单页内唯一。
- 未匹配对象通常淡入/淡出。若这不是设计意图，**修正身份映射，不要调时长碰运气**。
- 要全屏的图，目的页就写 `(0, 0, slide_width, slide_height)`；要转 45° 就写 `Rotation = 45`。
- 同一对象在所有相关页保持相同 z-order；Morph 不是修复层级错误的工具。

### 2.3 旋转枢轴

PowerPoint 默认绕形状**包围盒中心**旋转，不是可见图案的中心。需要指定枢轴时：把透明 PNG 裁成以枢轴为中心的等尺寸画布；或与透明辅助形状组成对称组合；或整体成组后旋转。不要把偏心图案放在整页透明 PNG 上再假设它绕图案中心转。

### 2.4 连续转场不留准备页

连续 A→B→C 只需要三页：B 是 A→B 的终点，同时是 B→C 的起点；C 需要但在 B 中应较小的对象直接预置在 B。页面数由**讲解停留状态**决定，不由每个转场机械地占两页。多出来的「准备页」在放映时就是一次多余停顿。

### 2.5 节拍与留白

技术正确不等于好看。默认尺度（按讲解节奏调）：

- 一页只动**一个主角**；同屏并列变化超过三个，观众就看不出你想说什么。
- 同组入场只选一种效果；混用多种入场会让人注意动画本身。
- 错峰 0.1–0.3 s，不要把一个动画切成十几片硬拼。
- 强调每页至多一两次；每行文字都加动画等于没有重点。
- 页面切换 0.3–0.8 s（大幅跨页位移按距离另估）；Morph 只用于对象逻辑连续的两页。
- 讲稿不依赖动画的页，保持静态。

### 2.6 不夸大 Morph

图片从小变到全屏只是**图片缩放/位移动画**，不能称为「任意遮罩逐渐变形」。不同图片、任意 freeform 遮罩或复杂 crop 之间，Morph 可能退化为淡入淡出。承诺必须限定在实际采用的对象变换上。

### 2.7 跨页转场（可选）

需要时可考虑视差分层（各层位移不同）与画布外预置（下一屏对象先放在画外，由 Morph 带入）。

XML 与几何细节见 `references/design-states.md`。

## 3. 执行层：inspect → plan → apply → validate

### 3.1 环境

引擎只依赖 `lxml`；COM 读回需要 `pywin32`。**在任务文件夹建 venv，不要建在技能目录里**：

```bash
uv venv .venv --python 3.11
uv pip install --python .venv/Scripts/python.exe -r scripts/requirements.txt   # Windows
# macOS/Linux: uv pip install --python .venv/bin/python -r scripts/requirements.txt
PY=.venv/Scripts/python.exe        # 之后所有命令都用 $PY
```

这台机器的 `python3` 是 WindowsApps 占位符，`python` 是 Hermes 运行时 venv（没有 lxml / python-pptx）。**不要直接 `python3 scripts/...`**——会报模块缺失或直接打不开解释器。

### 3.2 命令

```bash
# 页面顺序、形状名/ID、已有时间线与切换
$PY scripts/native_ppt_anim.py inspect deck.pptx

# 预设查询：203 个原生对象效果（53 入场 / 33 强调 / 64 路径 / 53 退出）
$PY scripts/native_ppt_anim.py list-effects --category emphasis
$PY scripts/native_ppt_anim.py describe emphasis_spin
$PY scripts/native_ppt_anim.py describe custom_motion

# 页面切换与 Morph
$PY scripts/native_ppt_anim.py list-transitions
$PY scripts/native_ppt_anim.py describe-transition morph

# 按计划写入（原子替换，失败不落盘）
$PY scripts/native_ppt_anim.py apply deck.pptx plan.json -o animated.pptx --report report.json

# 重新读盘校验：包、XML、关系、语义回读、路径、Morph 配对
$PY scripts/native_ppt_anim.py validate animated.pptx --spec plan.json

# 按点击组查看每条动画的开始/结束秒（--slide 只看一页）
$PY scripts/native_ppt_anim.py timeline animated.pptx --slide 4

# 独立 oracle：PowerPoint 自己读回（需要 Office + pywin32）
$PY scripts/pptx_com_audit.py animated.pptx --report com-audit.json
```

### 3.3 最小计划

```json
{
  "version": 1,
  "slides": [
    {
      "slide": 1,
      "animations": [
        {
          "target": "hero",
          "effect": "custom_motion",
          "path_px": "M -240 80 C -160 -120 -60 -100 0 0 E",
          "trigger": "on-click",
          "duration": 1.0,
          "accelerate": 0.1,
          "bounce_end": 0.18
        },
        {
          "target": "hero",
          "effect": "emphasis_spin",
          "trigger": "with-previous",
          "duration": 1.0,
          "accelerate": 0,
          "decelerate": 0
        }
      ],
      "transition": {"effect": "fade", "duration": 0.4}
    }
  ]
}
```

`target` / `trigger_shape` 接受唯一形状名或正整数形状 ID。时间单位是秒。未知字段、未知效果、非法组合一律报错，不会静默替换。字段全集见 `references/timing-plan-spec.md`。

**坐标画布**：`path_px` 按 96 DPI 映射到 `p:sldSz`；16:9 等于 **1280×720**，不是 1920×1080。设计稿是 1920×1080 时，顶层写 `"canvas_px": [1920, 1080]`，否则位移会静默放大 1.5 倍且校验全过。`canvas_px` 的宽高比必须与页面一致，否则报错。

**常用语法糖**（展开成普通条目后走同一条路径）：

- `from_offset_px: [dx, dy]`：入场行内补 `ppt_x`/`ppt_y` 偏移，`tm=0` 在偏移处，`tm=100000` 归位。
- `targets` + `stagger`：按名称、正整数 ID 或 `fnmatch` 批量展开，第 i 条延迟 `原 delay + i×stagger`。
- `repeat: "until-next-click" | "until-end-of-slide"`：无限循环，与 `repeat_count`/`repeat_duration` 互斥。
- `iterate: {by: "letter"|"word", delay}` / `build: "paragraph"`：逐字逐词或逐段，两者互斥。

### 3.4 计划写法

- **先 inspect 再写计划**：名称重复、名称拼错、页码按文件名猜，都会在写入阶段直接失败。
- 一个页面计划里 `replace_existing: true` 会清掉该页**全部**对象时间线（不影响切换）；只给空 `animations` 也合法，等于清空该页时间线。已有时间线默认受保护；PowerPoint 另存后时间线仍在，所以别指望「打开改两下再套计划」——要么一次规划到位，要么整页重排。
- 未知效果报错会列出全表；按 `list-effects` / `describe` 查准名字再写，不要凭印象。
- **显式 0 平滑**：写 `accelerate: 0, decelerate: 0` 才会删掉模板默认缓冲；省略就继承预设，多对象同步时会先后错开。
- **回弹会吃掉 decel**：非零 `bounce_end` 自动删除继承的减速末端，要求精确定位的装配不要用。
- **触发形状不能是目标本身**：`trigger_shape` 与 `target` 相同会直接报错。

### 3.5 静图换循环视频

目的页需要持续动效（光束、星闪、火苗）时，可把静态合成图做成循环视频盖在图层上。注意首帧与静态画面一致，否则替换会闪；用入场动画在切换期间隐藏；计划引擎没有媒体字段，时间线需手写；防鼠标进度条的遮罩用纯色加 100% 透明，「无填充」不挡鼠标。

## 4. 验证：六层，不许跳级

| 层 | 做什么 | 能证明 | 不能证明 |
|---|---|---|---|
| 1 输入 | `inspect`：页面顺序、形状名/ID、已有时间线 | 目标存在且唯一 | 动画存在 |
| 2 包 | `validate`：ZIP/OPC、XML、关系、时间节点 | 文件没坏 | 会按预期播放 |
| 3 语义 | `validate --spec`：逐行效果、目标、偏移、时长、重复、平滑、路径 | 写进去的和计划一致 | 运动连续、无遮挡 |
| 4 静态渲染 | PowerPoint 导出 PDF/PNG，逐页看 | 页面和对象能渲染 | 任何动画行为 |
| 5 COM 读回 | `pptx_com_audit.py`：MainSequence / 交互序列 / 触发 / 时长 / 切换 | PowerPoint 自己认这棵时间树 | 视觉节奏对不对 |
| 6 放映 | 用户在 PowerPoint 里点一遍；再另存副本重跑 `validate` | 真实播放与保存往返 | —— |

第 5 层是唯一不共用计划引擎代码的检查：它问的是 PowerPoint 的解释器，而不是我们自己的读取器。`Effect` 对象没有 `TriggerType` 属性，必须读 `fx.Timing.TriggerType`。

交付话术固定：没放映就说「已生成并做结构、语义与读回检查，未做本机放映验证」；不许把 `validate` 通过写成「动画效果已确认」。

放映复核要点见 `references/verification.md`。

## 5. 边界

- 不做动画音效；不做 COM 写入封装（COM 只作手工路线与验收 oracle）。
- 自定义路径支持 `M`/`L`/`C`/`E`，不支持 `Q`、弧线、隐式相对命令。需要时明确说明并单独实现，不伪造成功。
- 用户正开着 PowerPoint 时不做 COM 自动化；只管理自己启动的实例，绝不 `Quit()` 用户窗口。

## 6. 参考与脚本

- `references/`：`design-states`（Morph 的 OOXML 与几何）、`timing-plan-spec`（字段全集、调度公式、路径语法）、`verification`（COM 读回与放映复核）、`com-timing`（COM 进程安全与手写 timing）、`layout-and-integrity`（锚点叠加与包检查）、`provenance-licensing`（预设来源）
- `scripts/native_ppt_anim.py` + `scripts/vendor/` — 计划引擎与 203 预设（MIT，见 vendor/LICENSE）
- `scripts/tests/test_native_ppt_anim.py` — 回归用例
- `scripts/pptx_com_audit.py` — PowerPoint 读回
