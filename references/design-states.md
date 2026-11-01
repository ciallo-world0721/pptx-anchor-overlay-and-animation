# 状态设计与 Morph 实现参考

## 状态契约模板

每个状态（页）先填这张表，再决定用哪种引擎：

| 项 | 内容 |
|---|---|
| 讲解停留点 | 观众在这一页需要看见什么 |
| 保留对象 | 真实形状/图片/文本，哪些必须可编辑 |
| 对象身份 | 每页内的唯一名；跨页逻辑对象用 `!!key` |
| 几何终态 | left/top/width/height/rotation，直接写在目的页 |
| 层级 | 关键对象的 z-order 在源页与目的页一致 |
| 进出 | 哪些对象继续演变、哪些进入、哪些退出 |

填不出来，说明这一页还没设计好，不要先写动画。

## Morph 身份

- Morph 写在**目的页**；`!!Name` 在单页内唯一，跨相邻页稳定。
- 自动匹配只作辅助；需要可靠对应时显式强制命名。
- 相邻两页至少核对：`matched`（同名，预期插值）、`source-only`（通常退出/淡出）、`destination-only`（通常进入/淡入）、`duplicates`（单页重复强制名，必须修正）。
- 对象换图片内容、换形状类型或换裁剪方式时，不要凭同名承诺几何形变。

## 等价 OOXML

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                  xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                  spd="slow" p14:dur="1200">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="slow">
      <p:fade/>
    </p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

插入顺序遵循 slide schema：`cSld → clrMapOvr → transition → timing → extLst`。写入前删除已有的直接 `transition`，避免重复。计划引擎会同时写 Choice 与 Fallback。

## 终态几何

- 全屏终态：`left=0, top=0, width=slide_width, height=slide_height`。
- 旋转终态：直接写 `shape.rotation`，不要靠路径终点凑。
- 透明图片的旋转中心是**图片包围盒中心**，不是可见像素中心。
- 同一对象在源页与目的页保持预期 z-order；Morph 不是修复层级错误的工具。
- 连续 A→B→C 不添加额外准备页；下一转场需要的初始对象预置在 B。

## Morph 的能力边界

- 小图变全屏 = 图片缩放/位移，不是遮罩形变。
- 不同图片之间、任意 freeform 遮罩之间，Morph 可能退化为淡入淡出。
- 对象类型改变（形状→图片）会破坏强制匹配，应改为对象动画或单独过渡。
- 旧渲染器不支持 p159 时回退为 fade；整页不应因缺少 Morph 而失效。

## 结构审计的限度

`inspect` / `validate` 能证明：目的页含 Morph、强制名可匹配、终态几何明确、两侧类型一致。
不能证明：PowerPoint 对每一帧采用预期插值、是否短暂卡顿、观众是否看得懂。视觉结论留给真实放映。
