---
name: dsh-ppt-craft
description: 用 python-pptx 做「能拿出去讲」的 PPT：给卡片加悬浮阴影与细描边、按文字删页重排页码、写真正的 Office 转场与组件动画（淡入/擦除/脉冲闪动/隐藏）、画箭头轨道动画、交付前渲染自查。用户要做 PPT / 改 PPT / 加动画 / 问「这份稿子为什么看着廉价」时用它；配套 DSH 插件 dsh-ppt-craft 提供 ppt_polish / ppt_animate / ppt_render / ppt_audit 四个工具。
whenToUse: 用户要做或改一份 .pptx（尤其是要讲的专业内容），或者抱怨「页面平」「卡片和背景糊在一起」「没有动画」「放映时不动」「PPT 说文件损坏」。纯网页 PPT（HTML 横向翻页）不在本技能范围。
---

# 做 PPT

结论先行：一份能讲的 PPT = **一张纸的底色 + 有层次的面 + 会交代顺序的动画**。三样缺一样就会显得廉价，而缺动画最常见。

## 0. 先定三件事

1. **谁讲给谁**：时长决定页数（约 1 页 / 1.5 分钟），听众决定密度。
2. **一条主线**：每页只回答一个问题；连续两页在答同一个问题就该合并。
3. **配色就一套**：底色、纸、次级面、强调，四个色足够；多一个都会散。

## 1. 结构

按「钩子 → 框架 → 逐条展开 → 修正 → 迁移 → 收」走：

- 封面（大图满版，暗一点，标题压在图上）
- 目录 / 地图（让人知道要走多远）
- 一个具体场景开场（心跳、一个案例、一个反直觉问题）——比「今天我们讲…」强十倍
- 主体框架（一页图讲清全貌，后面每页展开其中一条）
- 每条：现象 → 机制 → 证据（实验/数据）→ 一句话结论
- 修正 / 争议页：把「看起来对但现在不这么认为」讲清楚，这是最显专业的地方
- 迁移 / 应用页：回到开场那个场景再走一遍
- 收尾：三句话总结，不加「谢谢聆听」以外的新内容

## 2. 设计：让卡片「浮起来」

廉价感几乎都来自**面和背景一个颜色、还没有边界**。规则：

- 纸底 `F7F2E8`，卡片 `FFFCF5`，次级面 `EFE7D8`（比底色深两档的暖沙，**不要用灰**），强调 `B0472C`，正文 `2B2A26`
- 每张卡：填充 + 细描边（比填充深一档）+ **朝下的柔和阴影**（blur 9pt / dist 3pt / alpha 17%）
- 圆角统一；同层级同形状；不要一页出现三种圆角
- 灰卡糊在背景里是典型事故：换 `EAE0CC` + 描边 `D8CBB2`
- 一页最多 3 个视觉块；文字块宽度别超过 2/3 版心
- 数字、年份、关键结论用强调色或超大字号，制造视线落点

细节和数值见 [设计默认值](references/design-system.md)；想换气质（瑞士国际 / 电子杂志）见 [另外三套配色](references/palettes.md)——换色不换做法。

## 3. 动画：三件事就够

只做这三种，别贪多：

| 目的 | 效果 | 用在 |
| --- | --- | --- |
| 强调「正在讲这格」 | `pulse` 脉冲 | 流程图里当前节点、表格当前行 |
| 把逻辑关系画出来 | `wipe` 擦除 | 箭头从左到右、下划线、进度条、轨道 |
| 把答案藏到点击之后 | `fade` 淡入 / `hide` | 问题卡的答案、分步展开的列表 |

外加全篇 `<p:fade/>` 切换。一页最多两次点击，超过两次说明内容没想清楚。

**箭头轨道动画**（讲「A 传到 B 再传到 C」最好用）：画几条同宽的细箭头，叠在同一条水平线上，初始只显示第 1 条；每点一次，先 `hide` 上一条、再 `wipe(left)` 下一条，像电流跑过去。分支处让两条同时 `wipe`（同一组里 delay 对齐）。

**「同时发生」**：两个形状放同一组动画，第二个 `delay` 约 400ms，看起来像一起闪——不要真的同时，会看不见。

XML 怎么写见 [动画原理](references/pptx-anim-xml.md)。

## 4. 动手顺序（照抄）

```bash
# 1 体检现有稿
python3 bin/craft.py <<<'{"cmd":"audit","args":{"path":"/sdcard/Download/稿.pptx"}}'

# 2 样式：阴影 + 描边 + 换灰卡 + 重排页码
python3 bin/craft.py <<'J'
{"cmd":"polish","args":{"path":"/sdcard/Download/稿.pptx","out":"/sdcard/Download/稿_s.pptx",
 "recolor":{"F6F3EE":"EAE0CC"},"renumber":true}}
J

# 3 动画（auto 自动瀑布，或 slides 精确控制）
python3 bin/craft.py <<'J'
{"cmd":"animate","args":{"path":"/sdcard/Download/稿_s.pptx","out":"/sdcard/Download/稿_a.pptx",
 "spec":{"transition":"fade","auto":{"pages":[3,4,5,6],"stagger":150}}}}
J

# 4 渲染抽页自查（阴影/描边能看见）
python3 bin/craft.py <<<'{"cmd":"render","args":{"path":"/sdcard/Download/稿_a.pptx","outdir":"/tmp/r","pages":[4,12]}}'
```

有 DSH 插件时对应 `ppt_audit` / `ppt_polish` / `ppt_animate` / `ppt_render` 四个工具，参数一致。

**改稿铁律：按文字删页，绝不按序号。** 删一页后面全变号，按序号删必然删错。

## 5. 交付前必须过一遍

- [ ] `audit` 报 0 个「缺阴影」、0 个「动画 id 重复」（id 重复 PowerPoint 会报文件损坏）
- [ ] 每页都有 `p:transition`
- [ ] 用 `Presentation(out)` 重开一次不报错（`animate` 内置）
- [ ] 抽 3–4 页渲染出来看：有没有文字溢出、白卡压白底、色块叠错
- [ ] 视频卡宽高比 = 素材真实比例，且给了 poster
- [ ] 文件放 `/sdcard/Download/DSH-Go/` 再 present，文件名写清「内容_页数_vN」

## 坑

常见的七个坑在 [踩过的坑](references/pitfalls.md)，最容易翻车的两个先说：

1. `hasattr(sh, 'spPr')` **永远是 False** → 样式循环用 `getattr(sh._element, 'spPr', None)`，写错会整轮静默跳过（看着像「什么都没变」）。
2. LibreOffice / 多数预览器**不渲染 `outerShdw`**。XML 里有就是有，交付时说明「请在 PowerPoint/WPS 里看」，自查用 `render`。