# pptx-anchor-overlay-and-animation

**1000万以内最好的PPT技能。**

用真正的 OOXML `p:timing` / `p:transition` 写原生对象动画和 Morph 转场——不是嵌入视频、GIF 或整页截图。写完后所有对象在 PowerPoint 里仍然可编辑。

这是给 AI 助手用的技能（Hermes Agent / Claude 风格的技能目录）：`SKILL.md` 是操作手册，`references/` 是规格，`scripts/` 是引擎。

## 能做什么

- **先设计状态，再写动画。** 起始画面、结束画面、对象身份、讲解停留点，先写清楚再碰 XML。
- **两套引擎，按需求选。** 跨页多对象移动/缩放/旋转用 Morph 双页状态；页内点击序列、强调、运动路径、批量错峰用 JSON 计划引擎写 `p:timing`。
- **203 个 PowerPoint 原生对象效果**：53 入场 / 33 强调 / 64 路径 / 53 退出，带真实行为节点，可查询。
- **六层验证，不跳级。** 包完整性 → 语义回读 → 静态渲染 → PowerPoint COM 读回（独立 oracle）→ 真实放映。
- **严格报错，原子写入。** 未知效果、未知字段、重名目标、非法参数直接报错；写盘先暂存，失败不落盘。

## 目录

```
SKILL.md                        # 何时使用、四条不变量、设计层、验证
references/
  design-states.md              # 状态契约、Morph 身份、终态几何、OOXML
  timing-plan-spec.md           # 计划字段全集、调度公式、路径语法
  verification.md               # 六层验证、放映复核、COM 审计
  com-timing.md                 # COM 进程安全、AddEffect、Behavior
  layout-and-integrity.md       # 锚点、叠加、包与关系检查
  provenance-licensing.md       # 预设来源与许可
scripts/
  native_ppt_anim.py            # 计划引擎：inspect / list-effects / apply / validate
  pptx_com_audit.py             # PowerPoint 读回
  requirements.txt              # lxml（Windows 另需 pywin32）
  vendor/                       # ppt-master 预设表（MIT，见 vendor/LICENSE）
```

## 安装

```bash
uv venv .venv --python 3.11
uv pip install --python .venv/Scripts/python.exe -r scripts/requirements.txt   # Windows
# macOS/Linux: uv pip install --python .venv/bin/python -r scripts/requirements.txt
PY=.venv/Scripts/python.exe      # 下面所有命令都用 $PY
```

只强依赖 `lxml`；`pywin32` 只在 Windows 上做 COM 读回时需要。

## 快速开始

```bash
# 1. 读 deck：页面顺序、形状名/ID、已有时间线
$PY scripts/native_ppt_anim.py inspect deck.pptx

# 2. 查预设，别凭印象写效果名
$PY scripts/native_ppt_anim.py list-effects --category emphasis
$PY scripts/native_ppt_anim.py describe emphasis_spin
$PY scripts/native_ppt_anim.py list-transitions

# 3. 写计划（原子替换，失败不动原文件）
$PY scripts/native_ppt_anim.py apply deck.pptx plan.json -o animated.pptx --report report.json

# 4. 按计划校验，再让 PowerPoint 自己读一遍
$PY scripts/native_ppt_anim.py validate animated.pptx --spec plan.json
$PY scripts/pptx_com_audit.py animated.pptx --report com-audit.json
```

最小计划：

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
        }
      ],
      "transition": { "effect": "fade", "duration": 0.4 }
    }
  ]
}
```

`target` / `trigger_shape` 接受唯一形状名或正整数 ID；时间单位是秒。调度、重复、自动反向、回弹的完整语义见 `references/timing-plan-spec.md`。

## 边界

不做动画音效，不做任意遮罩形变或逐帧物理，不封装 COM 写入（COM 只作手工路线和验收 oracle）。自定义路径只支持 `M`/`L`/`C`/`E`，不支持 `Q`、弧线和隐式相对命令。超范围会直接说明，不伪造成功。

结构通过不等于画面好看：`validate` 全绿只说明文件结构正确、语义与计划一致，不代表动画已经过视觉确认。

## 许可

MIT，见 [LICENSE](LICENSE)。

`scripts/vendor/` 来自 [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master)（MIT，Copyright (c) 2025–2026 Hugo He），许可见 `scripts/vendor/LICENSE-ppt-master.txt`。预设来源与参考实现的许可记录在 `references/provenance-licensing.md`。
