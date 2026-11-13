# dsh-ppt-craft · PPT 工艺刀

给 AI 用的 PPT 制作套件。它把「做一份能拿出去讲的 PPT」拆成四件可以反复调用的事：

| 工具 | 解决什么 |
| --- | --- |
| `ppt_polish` | 一批把卡片做成「纸片放在桌上」：细描边 + 朝下柔和阴影；顺手换掉糊在背景里的灰卡、按文字删页、重写页码 |
| `ppt_animate` | 写真正的 Office 动画：全篇淡入切换 + 点击逐步出现（淡入 / 擦除 / 脉冲闪动 / 隐藏） |
| `ppt_render` | 把页面渲染成 PNG 自查（会画圆角、描边、阴影），不用装 Office |
| `ppt_audit` | 交付前体检：缺阴影的卡片、没切换的页、动画 id 重复、视频比例 |

设计上去掉「反复猜」的部分：动画是直接写进 `ppt/presentation.xml` 与每页的 `p:timing` 里的标准 DrawingML，PowerPoint 和 WPS 都认，不需要插件、不需要手点面板。

## 装

```bash
git clone https://github.com/27Synthesis/dsh-ppt-craft.git
cp -r dsh-ppt-craft "$DSH_HOME/dsh-go/plugins/"
dsh plugin --profile web add "$DSH_HOME/dsh-go/plugins/dsh-ppt-craft"
# 重启引擎后四个工具出现
```

依赖：`python3` + `python-pptx` + `Pillow`（`pip install python-pptx pillow`）。没有 python 的话工具会直接告诉你。**装完要重启引擎**，四个工具才会出现。

### 只装 Skill（不带插件）

技能是纯文档 + 一个可直接跑的脚本，不依赖插件：

```bash
cp -r dsh-ppt-craft/skill ~/.dsh/skills/dsh-ppt-craft
python3 ~/.dsh/skills/dsh-ppt-craft/scripts/craft.py <<< '{"cmd":"audit","args":{"path":"稿.pptx"}}'
```

它教的是「怎么做」：先定结构、怎么让卡片浮起来、动画只做三件事、动手顺序、交付前八项检查；细节在 `skill/references/`（设计默认值 / 动画 XML 原理 / 十个坑）。

## 用

### 1. 先体检，再动手

```
ppt_audit { path: "/sdcard/Download/某稿.pptx" }
```

返回每页卡片数、缺阴影数、有无切换、动画步数、动画 id 是否重复。**id 重复的页会让 PowerPoint 报「文件损坏」**，这一项要盯住。

### 2. 加悬浮感

```
ppt_polish {
  path: "/sdcard/Download/稿.pptx",
  out: "/sdcard/Download/稿_修饰.pptx",
  renumber: true
}
```

默认参数就是这套牌桌感：

- 卡片 `FFFCF5` → 描边 `EAE1D1` + 阴影 `blur 9pt / dist 3pt / alpha 17% / 2B2A26`
- 米色块 `EFE7D8` → 描边 `DED3BE`
- 强调块 `B0472C` → 只加重阴影（`13pt / 3.2pt / 22%`），不加描边
- 细条、小标签、整页背景自动跳过（`min_h 0.55 / min_w 0.9`）

**灰卡和背景糊在一起**是常见毛病，用 `recolor` 换掉：

```
ppt_polish { path: "...", recolor: { "F6F3EE": "EAE0CC" }, border: { "EAE0CC": "D8CBB2" } }
```

### 3. 加动画

最省事的是自动瀑布（每页卡片按上到下、左到右依次浮出）：

```
ppt_animate {
  path: "/sdcard/Download/稿.pptx",
  out: "/sdcard/Download/稿_动画.pptx",
  spec: { transition: "fade", auto: { pages: [3,4,5,6], stagger: 150 } }
}
```

要精确控制就写 `slides`。**每个内层数组 = 一次点击**：

```json
{
  "transition": "fade",
  "slides": {
    "4": [
      [ {"eff": "wipe", "target": ["ribbon1"], "dir": "left", "dur": 650},
        {"eff": "pulse", "target": ["Rounded Rectangle 7"]},
        {"eff": "pulse", "target": ["Rounded Rectangle 23"], "delay": 420} ],
      [ {"eff": "hide", "target": ["ribbon1"], "dur": 200},
        {"eff": "wipe", "target": ["ribbon2"], "dur": 420, "delay": 200},
        {"eff": "wipe", "target": ["ribbon3"], "dur": 420, "delay": 640} ]
    ]
  }
}
```

`eff` 有四种，够用：

- `fade` 淡入（默认 450ms）
- `wipe` 擦除，`dir` 取 `left/right/up/down` —— 箭头、进度条、下划线都用它
- `pulse` 脉冲（放大 `scale` 再回弹，默认 1.07）—— 「闪一闪」表示同时发生、表示正在讲的那格
- `hide` 消失 —— 同一位置轮流换内容（下一条轨道出现时先把上一条收掉）

「点了才出答案」用 `reveal`：

```json
{ "reveal": [ { "page": 12, "trigger": ["Rounded Rectangle 15"], "hide": ["Rounded Rectangle 52"] } ] }
```

每个被藏起来的形状配一个触发器：先脉冲问号卡，再淡入答案。

### 4. 渲染自查

```
ppt_render { path: "/sdcard/Download/稿_动画.pptx", outdir: "/tmp/r1", pages: [4, 12] }
```

渲染器会读 `a:outerShdw` 把阴影画出来——这是唯一能在没有 Office 的机器上确认「阴影真的生效」的办法。注意它**不渲染动画**（动画只在放映时可见），所以动画只能靠 `ppt_audit` 的步数和真机放映确认。

## 动画为什么这么写（原理，省得重造）

每页的动画是 `p:sld` 下的一个 `p:timing` 树：

```
p:timing
└ p:tnLst → p:par → p:cTn id=1 nodeType="tmRoot"
  └ p:seq → p:cTn id=2 nodeType="mainSeq"
    └ 每个「点击」一个 p:par（cTn 的 cond delay="indefinite"）
      └ 效果组 p:par
        └ 每个效果一个 p:par（第一个 nodeType="clickEffect"，其余 "withEffect"）
```

要点：

- **id 全页唯一**，从 3 开始排（1、2 被 tmRoot/mainSeq 占了）。重复 id 会让 PPT 认为文件损坏。
- 进场效果 = `p:set` 把 `style.visibility` 写成 `visible` + `p:animEffect transition="in" filter="fade|wipe(left)"`。形状不需要预先隐藏，PowerPoint 会因为「有进场动画」自动藏住它。
- 脉冲 = `p:anim` 改 `ppt_w/ppt_h`，`additive="base"`，三个 `p:tav`：原值 → 放大 → 原值（`#ppt_w*1.07` 这种相对写法 PowerPoint 认）。
- 切换 = `p:transition` 放在 `p:clrMapOvr` 之后、`p:timing` 之前；全篇统一 `<p:fade spd="med"/>` 最稳，花哨切换在投影仪上会掉帧。

## 坑（都踩过）

1. **别按序号删页**。删一页后面全变号，删错是必然的。按页内文字匹配。
2. **`hasattr(sh, 'spPr')` 是 False**。`Shape` 实例没有这个属性，要用 `getattr(sh._element, 'spPr', None)`；写错会让整轮样式静默跳过（看起来「什么都没发生」）。
3. **阴影写进 `spPr/a:effectLst/a:outerShdw`**，同时 `sh.shadow.inherit = False`，否则母版的主题阴影会盖掉或叠加。
4. **LibreOffice / 各种预览器不显示 `outerShdw`**。XML 里有了就是有了，交付前告诉对方「在 PowerPoint/WPS 里看」；想自己确认就用 `ppt_render`。
5. **视频卡比例**：`add_movie` 给的宽高要和素材真实比例一致（1280×592 → 2.16:1），否则放映黑边；给 `poster` 一张同比例的静帧。
6. **白卡压白底**（`FFFCF5` 卡放 `FFFCF5` 卡上）等于不存在——中间层要么换色，要么靠描边和阴影分层。
7. 存完**重新打开校验**（`Presentation(out)`），能开才算数；`ppt_animate` 已经内置这一步。

## 设计上的几条默认值（可直接抄）

- 底色纸感 `F7F2E8`、卡片 `FFFCF5`、次要面 `EFE7D8`、强调 `B0472C`、正文墨色 `2B2A26`
- 卡片圆角统一，描边比填充深一档，阴影永远朝下（`dir=5400000`）
- 「和背景分不开」的次级面：比背景深两档的暖沙色（`EAE0CC`），不要用灰色
- 动画只做三件事：强调当前在讲的那格（pulse）、把逻辑关系画出来（wipe 箭头）、把答案藏到点击之后（fade）
- 一页最多两次点击；超过两次就是内容没想清楚

## License

MIT