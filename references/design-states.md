# 状态设计与 Morph 实现参考

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

## Morph 的能力边界

- 小图变全屏 = 图片缩放/位移，不是遮罩形变。
- 不同图片之间、任意 freeform 遮罩之间，Morph 可能退化为淡入淡出。
- 对象类型改变（形状→图片）会破坏强制匹配，应改为对象动画或单独过渡。
- 旧渲染器不支持 p159 时回退为 fade；整页不应因缺少 Morph 而失效。

## 结构审计的限度

`inspect` / `validate` 能证明：目的页含 Morph、强制名可匹配、终态几何明确、两侧类型一致。
不能证明：PowerPoint 对每一帧采用预期插值、是否短暂卡顿、观众是否看得懂。视觉结论留给真实放映。
