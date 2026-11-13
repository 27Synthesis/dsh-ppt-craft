# 踩过的坑

按「翻了多久才找到」排序。前四条能不踩就别踩。

## 1. 按序号删页 = 必然删错

`prs.slides` 的索引在删除后会全部前移。要删「第 15 页」时，它可能已经是别的页。

```python
# 错
del prs.slides[14]

# 对：按页内文字匹配
def drop_slide_containing(prs, key):
    for i, s in enumerate(prs.slides):
        if any(key in (sh.text_frame.text or '') for sh in s.shapes if sh.has_text_frame):
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[i])
            return i
```

删两页时尤其危险：删掉第一页后，第二页的序号就变了。一次只删一页，删完重新扫。

## 2. `hasattr(shape, 'spPr')` 永远是 False

`pptx.shapes.Shape` 实例没有 `spPr` 属性（它在 `shape._element` 上）。写 `hasattr` 判断会让整个样式循环**静默跳过**——不报错，只是什么都没变，很容易被误判成「阴影 API 不生效」。

```python
spPr = getattr(sh._element, 'spPr', None)
if spPr is None:
    continue
```

## 3. 阴影不出现，多半是本机预览器不画

`spPr/a:effectLst/a:outerShdw` 写进去了，PowerPoint / WPS 会画；**LibreOffice 和各种网页预览器基本不画 `outerShdw`**。别因此改代码——写自写 PIL 渲染器，按 `outerShdw` 的 blur/dist/alpha 画一遍就能确认。另注意 `sh.shadow.inherit = False`，否则母版主题阴影会叠加。

## 4. 动画 id 重复 = PowerPoint 报「文件损坏」

所有效果 `p:cTn` 的 id 必须**全页唯一**（tmRoot=1、mainSeq=2 之后从 3 起）。每页重新从 3 编号最容易撞车——用全局计数器。交付前 `audit` 一遍。

## 5. 形状名不是想象的名字

`Anim` 脚本里按名字找形状，报 `KeyError: 'TextBox 24 not on slide'`。中文版 PowerPoint 自动生成的形状名毫无规律（`Rounded Rectangle 19`、`Connector 24`、`TextBox 25`…）。动手前先 dump 一遍：

```python
for sh in slide.shapes:
    print(repr(sh.name), sh.shape_type)
```

## 6. 视频卡宽高比错 = 放映黑边

`add_movie` 给的 `width/height` 要和素材真实像素比一致（`1280×592` → 2.16:1）。另外给 `poster_frame_image` 一张同比例的静帧，否则点播放前是黑块。

## 7. 批量改色改到了不该改的

按填充色批量加描边时，细条、标签、背景图都会被扫到。用尺寸白名单（`min_h 0.55` / `min_w 0.9`）+ 名字前缀排除（`ribbon`、`tab`、`bg`）。

## 8. 存完不重开

`prs.save()` 成功不代表文件合法——XML 顺序错、id 重复都可能让 PowerPoint 拒绝打开。存盘后 `Presentation(out)` 重开一次，这是最便宜的验证。

## 9. 目录页码 / 页码字段

手写页码要跟实际页序对得上，删页/插页后必须重排。写在右下角 10pt，别用 `#` 字段（python-pptx 不解析，放映时才更新，容易错）。

## 10. 一页塞太多

「内容多 = 专业」是错觉。一页最多 3 个块、两次点击。塞不下就拆页——拆页是免费的，讲不清是要命的。