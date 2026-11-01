---
name: pptx-anchor-overlay-and-animation
description: "Use for native PPTX animation, Morph, COM, and overlays. Design the state contract first, then write p:timing from a JSON plan and prove it by PowerPoint read-back."
version: 3.0.0
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
- 自动翻页、循环放映、重复、自动反向、回弹
- 要把形状精确叠在自己生成的底图上
- 审查已有动画：分得清「能播」和「XML 里看着有」

不适用，直接说明而不是硬做：动画音效（接口未封装）、任意遮罩形变与逐帧物理（非原生强项）、网页/CSS 动画（另一条路线，不能冒充原生）。用户没要求动画时不要顺手加。

## 0. 四条不变量

1. **交付原生对象**：图片、形状、文本结束后仍可编辑。不得用嵌入视频、GIF、HTML 或整页截图冒充原生动画；用户明确要媒体路线时才用。
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
- 页面切换 0.3–0.8 s；Morph 只用于对象逻辑连续的两页。
- 讲稿不依赖动画的页，保持静态。

### 2.6 不夸大 Morph

图片从小变到全屏只是**图片缩放/位移动画**，不能称为「任意遮罩逐渐变形」。不同图片、任意 freeform 遮罩或复杂 crop 之间，Morph 可能退化为淡入淡出。承诺必须限定在实际采用的对象变换上。

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

Windows 上常见 `python3` 是 WindowsApps 占位符、`python` 也可能指向一个不含 `lxml` 的运行时环境。**不要直接 `python3 scripts/...`**——会报模块缺失或直接打不开解释器。

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

### 3.4 计划写法

- **先 inspect 再写计划**：名称重复、名称拼错、页码按文件名猜，都会在写入阶段直接失败。
- 一个页面计划里 `replace_existing: true` 会清掉该页**全部**对象时间线（不影响切换）。已有时间线默认受保护；PowerPoint 另存后时间线仍在，所以别指望「打开改两下再套计划」——要么一次规划到位，要么整页重排。
- 未知效果报错会列出全表；按 `list-effects` / `describe` 查准名字再写，不要凭印象。

## 4. 调度与参数语义

**触发与节拍**（同一普通点击组内）：

```text
on-click        : 开新组，节拍起点与最晚结束时间归零
with-previous   : start = beat_start + delay
after-previous  : beat_start = timeline_end，再加自己的 delay
每行结束        : timeline_end = max(timeline_end, start + playback)
```

所以 `A(0.65s) → B(+0.12, 0.65) → C(+0.36, 0.65) → D(after-previous, 0.3)` 里，D 从 1.01 s 开始，等的是最长分支 C，不是上一条 B。

**实际播放时长**：

```text
one_play = duration × (2 if auto_reverse else 1)
repeat_count 存在 → playback = one_play × repeat_count
repeat_duration 存在 → playback = repeat_duration
```

`after-previous` 与自动翻页都按这个实际时长算。轮毂 0.7 s ×2（自动反向）×2（重复）= 2.8 s，自动翻页必须晚于 2.8 s，否则被截断。

**平滑**：OOXML 的 `accel`/`decel` 是 0–100000，接口用 0–1；**显式 0 会删掉模板默认平滑**，匀速同步必须写 `accelerate: 0, decelerate: 0`。`accelerate + decelerate ≤ 1`。

**回弹**：非零 `bounce_end` 会自动删除继承的 `decel`，并生成 p14 Choice + Fallback 双分支（旧版降级为无回弹）。回弹会越过终点，不要用于要求精确定位的同步装配。

**触发器**：`trigger_shape` + `trigger: "on-click"`，触发形状与目标形状必须不同。形状触发效果进独立 `interactiveSeq`，不参与主序列调度。

**Morph**：目的页 `{"effect": "morph"}`；源页与目的页相邻、共享至少一个顶层 `!!key`、两侧 OOXML 类型相同。换对象类型（形状→图片）会破坏强制匹配。

**自动翻页**：`transition.advance_after`；工具同时写放映级 `useTimings=1`。只写 `advTm` 不会自动翻页。

**COM 路线**（在 PowerPoint 里手工时）：`AddEffect(Shape, EffectId, Level, Trigger, Index)` 第 3 个参数是 Level，图片通常传 0；TriggerType `1=OnPageClick, 2=WithPrevious, 3=AfterPrevious, 4=OnShapeClick`；退出用 `effect.Exit = -1`；入场效果会让对象在播放前隐藏。进程安全与 Behavior 写法见 `references/com-timing.md`。

## 5. 验证：六层，不许跳级

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

## 6. 边界

- 不做动画音效；不做 COM 写入封装（COM 只作手工路线与验收 oracle）。
- 自定义路径支持 `M`/`L`/`C`/`E`，不支持 `Q`、弧线、隐式相对命令。需要时明确说明并单独实现，不伪造成功。
- 一页只有一个直接 `p:timing`；已有媒体节点时复用其结构，不追加第二棵。
- 用户正开着 PowerPoint 时不做 COM 自动化；只管理自己启动的实例，绝不 `Quit()` 用户窗口。

## 7. 参考与脚本

- `references/design-states.md` — 状态契约、Morph 身份、终态几何、OOXML 片段
- `references/timing-plan-spec.md` — 计划字段全集、调度公式、路径语法、错误与原子性
- `references/verification.md` — 六层验证、放映复核、流畅度调节、COM 审计
- `references/com-timing.md` — COM 进程安全、AddEffect、Behavior、手写 timing 的必要条件
- `references/layout-and-integrity.md` — 锚点契约、叠加、包与关系检查
- `references/provenance-licensing.md` — 预设来源与复用许可
- `scripts/native_ppt_anim.py` + `scripts/vendor/` — 计划引擎与 203 预设（MIT，见 vendor/LICENSE）
- `scripts/pptx_com_audit.py` — PowerPoint 读回
