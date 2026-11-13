# 样例

`spec-example.json` 是 `ppt_animate` / `craft.py animate` 用的动画描述，演示三种最常用的写法：

- `auto.pages`：这些页自动做「按位置从上到下、从左到右」的卡片瀑布（间隔 150ms）
- `slides."4"`：两次点击。第一次箭头轨道 1 跑出来 + 两个格子同时闪；第二次收起轨道 1、轨道 2/3 接力跑
- `slides."12"`：第一次点问号卡（放大脉冲），第二次四行答案依次落下（每行错开 140ms）

`target` 里写的是**形状名**，每份稿子都不一样，动手前先 dump：

```bash
python3 -c "from pptx import Presentation;p=Presentation('稿.pptx');
[print(i+1, [s.name for s in sl.shapes]) for i,sl in enumerate(p.slides)]"
```
