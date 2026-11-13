# PPT 动画原理（DrawingML，照抄能跑）

动画不是「播放效果」，是一棵写进 `ppt/slides/slideN.xml` 的 `<p:timing>` 树。PowerPoint 和 WPS 都读它，不需要插件、不需要手点面板——AI 可以直接写。

## 目录结构

```xml
<p:sld>
  <p:cSld>…形状…</p:cSld>
  <p:clrMapOvr>…</p:clrMapOvr>
  <p:transition>          ← 页面切换（位置必须在 clrMapOvr 之后）
    <p:fade/>
  </p:transition>
  <p:timing>…动画…</p:timing>
</p:sld>
```

`p:timing` 必须放在 `p:transition` **之后**，放错会让 PowerPoint 忽略整块动画。

## timing 骨架

```
p:timing
└ p:tnLst → p:par → p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"
  └ p:childTnLst → p:seq concurrent="1" nextAc="seek"
    └ p:cTn id="2" dur="indefinite" nodeType="mainSeq"
      └ p:childTnLst
        ├ p:par                          ← 第 1 次点击
        │ └ p:cTn id="3" fill="hold"
        │   └ p:stCondLst → p:cond delay="indefinite"
        │   └ p:childTnLst → p:par
        │     └ p:cTn id="4" fill="hold"
        │       └ p:stCondLst → p:cond delay="0"
        │       └ p:childTnLst
        │         ├ p:par  (cTn id="5" … 第一个效果，node="clickEffect")
        │         └ p:par  (cTn id="6" … 同组第二个效果，node="withEffect")
        └ p:par                          ← 第 2 次点击
```

要点：

- **id 全页唯一**，从 3 开始（1=tmRoot，2=mainSeq）。重复 id → PowerPoint 提示「文件已损坏」。
- 每个「点击」= 一个 `p:par`，其 `p:cond delay="indefinite"`。
- 同一点击里第一个效果 `nodeType="clickEffect"`，其余 `"withEffect"`（用 delay 错开）。
- 每个效果自己的 `p:cTn` 里要有 `<p:stCondLst><p:cond delay="N"/></p:stCondLst>` 和 `<p:endCondLst><p:cond evt="onEnd" delay="0"/></p:endCondLst>`，`p:childTnLst` 放真正的特效节点。
- 每个效果还要在 `p:cBhvr/p:tgtEl/p:spTgt spid="N"` 指到形状 id。

## 四种效果

### 淡入 fade

```xml
<p:set>
  <p:cBhvr><p:cTn id="5" dur="1" fill="hold"/>
    <p:tgtEl><p:spTgt spid="12"/></p:tgtEl>
    <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>
  <p:to><p:strVal val="visible"/></p:to>
</p:set>
<p:animEffect transition="in" filter="fade">
  <p:cBhvr><p:cTn id="6" dur="450"/>
    <p:tgtEl><p:spTgt spid="12"/></p:tgtEl></p:cBhvr>
</p:animEffect>
```

形状不用预先设隐藏——只要有进场动画，PowerPoint 会自动藏住它。

### 擦除 wipe

同上，只把 `filter="fade"` 换成 `filter="wipe(left)"` / `wipe(right)` / `wipe(up)` / `wipe(down)`。箭头、下划线、进度条用它，`dur` 600–700ms 比较像「画出来的」。

### 脉冲 pulse（改尺寸，带 additive）

```xml
<p:anim calcmode="lin" valueType="num">
  <p:cBhvr additive="base"><p:cTn id="7" dur="500" autoRev="1"/>
    <p:tgtEl><p:spTgt spid="20"/></p:tgtEl>
    <p:attrNameLst><p:attrName>ppt_w</p:attrName><p:attrName>ppt_h</p:attrName></p:attrNameLst>
  </p:cBhvr>
  <p:tavLst>
    <p:tav tm="0"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav>
    <p:tav tm="50000"><p:val><p:strVal val="#ppt_w*1.07"/></p:val></p:tav>
    <p:tav tm="100000"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav>
  </p:tavLst>
</p:anim>
```

`additive="base"` + `#ppt_w*1.07` 是相对写法，PowerPoint 认；写死绝对尺寸会被当成「移到别处」。`autoRev="1"` 让它自己弹回来。

**脉冲默认 `nodeType="withEffect"`**——它不该占一次点击，通常跟着某个进场效果一起闪。

### 消失 hide

`<p:animEffect transition="out" filter="fade">`，`dur` 200–250ms。用来做「同一位置轮流换内容」：新的一条 `wipe` 进来前，先把上一条 `hide` 掉。

## 常用组合：箭头轨道

5 个细箭头（高度 ~0.05 英寸）叠在同一条水平线上，宽窄不同，指向不同格子。

第 1 次点击：`wipe(left)` ribbon1 + 两个格子 `pulse`（delay 400ms 错开）
第 k 次点击：`hide` ribbon(k-1)（200ms）+ `wipe(left)` ribbon(k)（420ms，delay 200）

效果是「电流顺着流程跑过去」，讲「A→B→C」时比任何箭头静态图都好用。

## 常用组合：答案藏到点击之后

问题卡（大 `?`）放第 1 次点击：`pulse` 问号卡 → 同组 `fade` 答案组（delay 400）。
答案组用「组形状」（`grpSp`）或几个卡片一组，`wipe(down)` 逐行落下比 `fade` 更有「揭开」感（每行 delay 递增 140ms）。

## 生成顺序（代码里）

1. 收集每页的 `slides` 定义 → 每个内层数组 = 一次点击
2. 每组里给每个效果分配唯一 id（全局递增，别每页从 3 重新数）
3. 目标名（形状 name）→ 形状 id（`spid`），找不到就报错，别静默跳过
4. `p:timing` 整棵替换（旧动画一并清掉）
5. 存盘后 `Presentation(out)` 重开验证

## 切换类型

`<p:fade spd="med"/>` 最稳。其他可用的：

```xml
<p:transition><p:push dir="u"/></p:transition>
<p:transition><p:wipe dir="l"/></p:transition>
<p:transition><p:cover dir="u"/></p:transition>
```

花哨的（`p:gallery`、`p:vortex`、`p:ripple`）在投影仪/低配机上会掉帧，专业场合一律 `fade` 或干脆不做切换。