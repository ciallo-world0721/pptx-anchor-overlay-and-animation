# COM 与原生 timing 参考

## Office 进程安全

PowerPoint COM 是单实例。用户正在使用 PowerPoint 时不运行自动化；不得附加后调用 `Quit()` 关闭用户窗口。自动化前检查 `POWERPNT.EXE`，只管理脚本自己启动的实例，演示文稿优先 `ReadOnly=True, WithWindow=False`。

## MainSequence.AddEffect

```python
seq = slide.TimeLine.MainSequence
effect = seq.AddEffect(shape, 10, 0, 1)  # Fade, Level=0, OnPageClick
effect.Timing.Duration = 0.45

with_previous = seq.AddEffect(other_shape, 10, 0, 2)
with_previous.Timing.TriggerDelayTime = 0.20

exit_effect = seq.AddEffect(old_shape, 10, 0, 2)
exit_effect.Exit = -1
```

TriggerType：

| 值 | 含义 |
|---|---|
| 0 | None |
| 1 | OnPageClick |
| 2 | WithPrevious |
| 3 | AfterPrevious |
| 4 | OnShapeClick |

`MainSequence.Count` 不包含交互触发序列中的全部效果，不能拿它推断完整点击模型。

## 自定义 Behavior

```python
custom = seq.AddEffect(shape, 0, 0, 2)
scale = custom.Behaviors.Add(3).ScaleEffect   # msoAnimTypeScale
scale.FromX = 20
scale.FromY = 20
scale.ToX = 100
scale.ToY = 100

motion_fx = seq.AddEffect(shape, 0, 0, 2)
motion = motion_fx.Behaviors.Add(1).MotionEffect
motion.ByX = 0.10  # 相对幻灯片宽度
motion.ByY = 0.00
```

Behavior 常用类型：Motion=1、Scale=3、Rotation=4、Set=8。涉及多个对象的页面级协同转场时，优先重新评估 Morph，而不是继续叠 Behavior。

## 手写 timing 时的必要条件

主序列必须保留：

```xml
<p:seq concurrent="1" nextAc="seek">
  <p:cTn id="2" dur="indefinite" nodeType="mainSeq">…</p:cTn>
  <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
  <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
</p:seq>
```

- 每个点击组第一步的首效果：`clickEffect`。
- 同组后续步骤的首效果：`afterEffect`。
- 同一步并列效果：`withEffect`。
- 延时在效果自己的 `cTn/stCondLst/cond@delay`。
- 形状点击触发使用 `interactiveSeq`，其条件包含 `evt="onClick"` 与目标 `spTgt`。
- 一页只能有一个直接 `p:timing`；`replace_existing` 会保留 `p:audio` / `p:video` 子树并重挂到新 tmRoot。

## COM 读回的正确解释

可读回 EffectType、TriggerType、Duration、Behavior 参数，作为结构审计。不得据此声称运动连续、无卡顿或最终画面正确。
