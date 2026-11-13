#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-craft 后端：卡片悬浮阴影 / 页面切换 / 组件动画 / 渲染自检。

协议：stdin 收 {"cmd":..., "args":{...}}，stdout 回 {"ok":true,"data":...} 或 {"ok":false,"error":...}
命令：render / audit / polish / animate / effects / info
"""
import io
import json
import math
import os
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Inches, Pt

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

EMU_IN = 914400.0

# ---------------------------------------------------------------- 基础工具


def emu_in(v):
    return 0.0 if v is None else v / EMU_IN


def _clean(c):
    return str(c).upper().replace('#', '')


def shape_fill_hex(sh):
    try:
        if sh.fill.type == 1:            # solid
            return _clean(sh.fill.fore_color.rgb)
    except Exception:
        pass
    return None


def shape_line_hex(sh):
    try:
        if sh.line.fill.type == 1:
            return _clean(sh.line.color.rgb)
    except Exception:
        pass
    return None


def is_picture(sh):
    return str(sh.shape_type).startswith('PICTURE') or 'MEDIA' in str(sh.shape_type)


def text_of(sh):
    if sh.has_text_frame:
        return sh.text_frame.text
    if getattr(sh, 'has_table', False) and sh.has_table:
        return '\n'.join(c.text for r in sh.table.rows for c in r.cells)
    return ''


# ---------------------------------------------------------------- 设计层：阴影 / 描边


def ensure_shadow(sh, blur=9.0, dist=3.0, alpha=17, color='2B2A26', direction=90):
    """给形状写一个朝下的外阴影（幂等：先清空 effectLst）。

    踩坑：不要用 hasattr(sh, 'spPr') 判断 —— Shape 实例没有这个属性，
    必须 getattr(sh._element, 'spPr', None) is None。
    """
    spPr = getattr(sh._element, 'spPr', None)
    if spPr is None:
        return False
    try:
        sh.shadow.inherit = False
    except Exception:
        pass
    eff = spPr.find(qn('a:effectLst'))
    if eff is None:
        eff = etree.SubElement(spPr, qn('a:effectLst'))
    for c in list(eff):
        eff.remove(c)
    shd = etree.SubElement(eff, qn('a:outerShdw'))
    for k, v in (('blurRad', int(blur * 12700)), ('dist', int(dist * 12700)),
                 ('dir', int(direction * 60000)), ('algn', 'ctr'), ('rotWithShape', 0)):
        shd.set(k, str(v))
    clr = etree.SubElement(shd, qn('a:srgbClr'))
    clr.set('val', _clean(color))
    etree.SubElement(clr, qn('a:alpha')).set('val', str(int(alpha * 1000)))
    return True


def ensure_border(sh, color, w=0.75):
    spPr = getattr(sh._element, 'spPr', None)
    if spPr is None:
        return False
    try:
        sh.line.color.rgb = RGBColor.from_string(_clean(color))
        sh.line.width = Pt(w)
        return True
    except Exception:
        return False


def set_fill_alpha(sh, alpha):
    """给纯色填充加透明度（alpha 传 0-1），用于「轨道 / 划线」这类辅助图形。"""
    spPr = getattr(sh._element, 'spPr', None)
    if spPr is None:
        return False
    sf = spPr.find(qn('a:solidFill'))
    if sf is None:
        return False
    c = sf.find(qn('a:srgbClr'))
    if c is None:
        return False
    for a in c.findall(qn('a:alpha')):
        c.remove(a)
    etree.SubElement(c, qn('a:alpha')).set('val', str(int(alpha * 100000)))
    return True


# ---------------------------------------------------------------- 设计层：批量修饰

DEFAULTS = {
    'card_fills': ['FFFCF5', 'EFE7D8', 'EDE7DB', 'E7DCC6', 'F6F3EE'],
    'border': {'FFFCF5': 'EAE1D1', 'EFE7D8': 'DED3BE', 'EDE7DB': 'D8CBB2', 'E7DCC6': 'D6C7AC'},
    'shadow': {'blur': 9, 'dist': 3, 'alpha': 17, 'color': '2B2A26'},
    'accent_fills': ['B0472C'],
    'accent_shadow': {'blur': 13, 'dist': 3.2, 'alpha': 22},
    'min_h': 0.55,
    'min_w': 0.9,
    'skip_prefix': ['ribbon', 'tab'],
    'recolor': {},
    'all_fills': True,
    'include_pictures': False,
}


def polish_deck(prs, opts, report):
    """把所有「卡片感」的形状统一成：纸白 / 米色填充 + 细描边 + 朝下阴影。

    判定：有纯色填充、尺寸够大（min_h × min_w）、不是整页背景、名字不在 skip_prefix。
    这样不用按页码逐页改，换模板也能复用。
    """
    o = dict(DEFAULTS)
    o.update(opts or {})
    sw, sh_ = emu_in(prs.slide_width), emu_in(prs.slide_height)
    card_fills = {_clean(x) for x in o['card_fills']}
    accent = {_clean(x) for x in o['accent_fills']}
    recolor = {_clean(k): _clean(v) for k, v in (o['recolor'] or {}).items()}
    border_map = {_clean(k): _clean(v) for k, v in o['border'].items()}
    n_card = n_accent = 0
    for idx, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            name = sh.name or ''
            if any(name.startswith(p) for p in o['skip_prefix']):
                continue
            if getattr(sh._element, 'spPr', None) is None:
                continue
            if is_picture(sh):
                if not o['include_pictures']:
                    continue
                ensure_shadow(sh, **o['shadow'])
                n_card += 1
                continue
            fill = shape_fill_hex(sh)
            if fill is None:
                continue
            w, h = emu_in(sh.width), emu_in(sh.height)
            if w >= sw * 0.95 or h >= sh_ * 0.95:
                continue
            if h < o['min_h'] or w < o['min_w']:
                continue
            if fill in recolor:
                sh.fill.solid()
                sh.fill.fore_color.rgb = RGBColor.from_string(recolor[fill])
                newfill = recolor[fill]
                report.append('p%d %s: %s → %s' % (idx, name, fill, newfill))
                fill = newfill
            if fill in accent:
                ensure_shadow(sh, **o['accent_shadow'])
                n_accent += 1
                continue
            if fill not in card_fills and fill not in border_map and not o.get('all_fills'):
                continue
            if fill in border_map:
                ensure_border(sh, border_map[fill])
            else:
                ensure_border(sh, 'DED3BE')
            ensure_shadow(sh, **o['shadow'])
            n_card += 1
    report.append('卡片/对象加了阴影：%d 个，强调块：%d 个' % (n_card, n_accent))
    return n_card + n_accent


def renumber(prs, report):
    """把页脚「n / N」和大号页码「NN」改成当前页号与新总数。"""
    total = len(prs.slides)
    changed = 0
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                txt = ''.join(r.text for r in para.runs).strip()
                if not txt:
                    continue
                if re.fullmatch(r'\d+\s*/\s*\d+', txt) or re.fullmatch(r'\d{2}', txt):
                    new = ('%d / %d' % (i, total)) if '/' in txt else ('%02d' % i)
                    if new == txt:
                        continue
                    if para.runs:
                        para.runs[0].text = new
                        for extra in para.runs[1:]:
                            extra.text = ''
                        changed += 1
    report.append('重写页码：%d 处（总页数 %d）' % (changed, total))
    return changed


def delete_slides_by_text(prs, keys, report):
    """按页内文字删页。**绝不能按序号删**：删一页后面的序号全变。"""
    lst = prs.slides._sldIdLst
    for key in keys:
        hit = None
        for sid, slide in zip(list(lst), prs.slides):
            txt = '\n'.join(text_of(sh) for sh in slide.shapes if text_of(sh))
            if key in txt:
                hit = sid
                break
        if hit is None:
            report.append('⚠️ 没找到包含「%s」的页' % key)
            continue
        prs.part.drop_rel(hit.get(R_NS))
        lst.remove(hit)
        report.append('删页：%s' % key)


# ---------------------------------------------------------------- 动画：效果块


class Ids:
    def __init__(self, start=3):
        self.n = start

    def __call__(self):
        v = self.n
        self.n += 1
        return v


def _tgt(spid):
    return '<p:tgtEl><p:spTgt spid="%s"/></p:tgtEl>' % spid


def eff_fade(spid, dur=450, delay=0, node='clickEffect', ids=None):
    a, b, c = ids(), ids(), ids()
    return (
        '<p:par><p:cTn id="%d" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="%s">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '%s<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set>'
        '<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="%d" dur="%d"/>%s</p:cBhvr></p:animEffect>'
        '</p:childTnLst></p:cTn></p:par>'
    ) % (a, node, delay, b, _tgt(spid), c, dur, _tgt(spid))


def eff_wipe(spid, direction='left', dur=600, delay=0, node='clickEffect', ids=None):
    a, b, c = ids(), ids(), ids()
    sub = {'left': 8, 'right': 2, 'up': 4, 'down': 1}.get(direction, 8)
    return (
        '<p:par><p:cTn id="%d" presetID="22" presetClass="entr" presetSubtype="%d" fill="hold" grpId="0" nodeType="%s">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '%s<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:to><p:strVal val="visible"/></p:to></p:set>'
        '<p:animEffect transition="in" filter="wipe(%s)"><p:cBhvr><p:cTn id="%d" dur="%d"/>%s</p:cBhvr></p:animEffect>'
        '</p:childTnLst></p:cTn></p:par>'
    ) % (a, sub, node, delay, b, _tgt(spid), direction, c, dur, _tgt(spid))


def eff_hide(spid, dur=250, delay=0, node='clickEffect', ids=None):
    a, b = ids(), ids()
    return (
        '<p:par><p:cTn id="%d" presetID="1" presetClass="exit" presetSubtype="0" fill="hold" grpId="0" nodeType="%s">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>'
        '<p:set><p:cBhvr><p:cTn id="%d" dur="%d" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        '%s<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        '<p:to><p:strVal val="hidden"/></p:to></p:set>'
        '</p:childTnLst></p:cTn></p:par>'
    ) % (a, node, delay, b, dur, _tgt(spid))


def eff_pulse(spid, dur=500, delay=0, node='withEffect', scale=1.07, ids=None):
    """放大再回弹 —— 「闪一闪」用它，不需要退出再进场。"""
    a = ids()
    anims = []
    for attr in ('ppt_w', 'ppt_h'):
        anims.append(
            '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            '<p:cTn id="%d" dur="%d" fill="hold"/>%s'
            '<p:attrNameLst><p:attrName>%s</p:attrName></p:attrNameLst></p:cBhvr>'
            '<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#%s"/></p:val></p:tav>'
            '<p:tav tm="50000"><p:val><p:strVal val="#%s*%s"/></p:val></p:tav>'
            '<p:tav tm="100000"><p:val><p:strVal val="#%s"/></p:val></p:tav></p:tavLst></p:anim>'
            % (ids(), dur, _tgt(spid), attr, attr, attr, scale, attr))
    return (
        '<p:par><p:cTn id="%d" presetID="6" presetClass="emph" presetSubtype="0" fill="hold" grpId="0" nodeType="%s">'
        '<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>%s</p:childTnLst></p:cTn></p:par>'
    ) % (a, node, delay, ''.join(anims))


EFFECTS = {'fade': eff_fade, 'wipe': eff_wipe, 'hide': eff_hide, 'pulse': eff_pulse}


def build_timing(groups, effects=EFFECTS):
    ids = Ids()
    body = []
    for group in groups:
        gid, eid = ids(), ids()
        parts = []
        for i, fx in enumerate(group):
            fn = effects[fx['eff']]
            kw = {'spid': fx['spid'], 'dur': fx.get('dur', 450), 'delay': fx.get('delay', 0),
                  'node': 'clickEffect' if i == 0 else 'withEffect', 'ids': ids}
            if fx['eff'] == 'wipe':
                kw['direction'] = fx.get('dir', 'left')
            if fx['eff'] == 'pulse':
                kw['scale'] = fx.get('scale', 1.07)
            parts.append(fn(**kw))
        body.append(
            '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            '<p:childTnLst><p:par><p:cTn id="%d" fill="hold">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>%s</p:childTnLst>'
            '</p:cTn></p:par></p:childTnLst></p:cTn></p:par>' % (gid, eid, ''.join(parts)))
    xml = ('<p:timing %s><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
           '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
           '<p:childTnLst>%s</p:childTnLst></p:cTn>'
           '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
           '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
           '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>') % (nsdecls('p', 'a'), ''.join(body))
    return parse_xml(xml)


TRANSITIONS = ('fade', 'wipe', 'push', 'cover', 'dissolve')


def set_transition(slide, kind='fade'):
    """插入 p:transition（必须在 p:clrMapOvr 之后、p:childTnLst 之外）。"""
    sld = slide._element
    for el in sld.findall(qn('p:transition')):
        sld.remove(el)
    inner = '<p:fade/>' if kind == 'fade' else '<p:%s/>' % kind
    tr = parse_xml('<p:transition %s spd="med">%s</p:transition>' % (nsdecls('p'), inner))
    ovr = sld.find(qn('p:clrMapOvr'))
    pos = list(sld).index(ovr) + 1 if ovr is not None else len(list(sld))
    sld.insert(pos, tr)
    return pos


def set_timing(slide, groups, effects=EFFECTS):
    sld = slide._element
    for el in sld.findall(qn('p:timing')):
        sld.remove(el)
    ovr = sld.find(qn('p:clrMapOvr'))
    pos = list(sld).index(ovr) + 1 if ovr is not None else 0
    tr = sld.find(qn('p:transition'))
    if tr is not None:
        pos = list(sld).index(tr) + 1
    sld.insert(pos, build_timing(groups, effects))


# ---------------------------------------------------------------- 动画：目标解析与卡片自动分组


def shape_index(slide):
    by_name, by_id = {}, {}
    for sh in slide.shapes:
        by_id[sh.shape_id] = sh
        by_name.setdefault(sh.name, sh)
    return by_name, by_id


def resolve(slide, target):
    """target 可以是形状名、形状 id 或它们的列表。"""
    by_name, by_id = shape_index(slide)
    out = []
    targets = target if isinstance(target, list) else [target]
    for t in targets:
        sh = by_name.get(t) if isinstance(t, str) else None
        if sh is None:
            try:
                sh = by_id.get(int(t))
            except Exception:
                sh = None
        if sh is not None and sh.shape_id not in out:
            out.append(sh.shape_id)
    return out


def _bbox(sh):
    return (sh.left or 0, sh.top or 0, (sh.left or 0) + (sh.width or 0), (sh.top or 0) + (sh.height or 0))


def _inside(a, b):
    ax1, ay1, ax2, ay2 = _bbox(a)
    bx1, by1, bx2, by2 = _bbox(b)
    return ax1 >= bx1 and ay1 >= by1 and ax2 <= bx2 and ay2 <= by2 and (a.width, a.height) != (b.width, b.height)


def _center_in(sh, box):
    cx = (box.left or 0) + (box.width or 0) / 2.0
    cy = (box.top or 0) + (box.height or 0) / 2.0
    x1, y1, x2, y2 = _bbox(sh)
    return x1 <= cx <= x2 and y1 <= cy <= y2


def card_units(slide, fills, min_h=0.8, min_w=1.5):
    """找出「卡片 + 卡里的文字」作为一个整体，按从上到下、从左到右排序。"""
    fills = {_clean(f) for f in fills}
    cards = []
    for sh in slide.shapes:
        if getattr(sh._element, 'spPr', None) is None:
            continue
        f = shape_fill_hex(sh)
        if f not in fills:
            continue
        if emu_in(sh.height) < min_h or emu_in(sh.width) < min_w:
            continue
        cards.append(sh)
    cards = [c for c in cards if not any(_inside(c, o) for o in cards if o is not c)]
    cards.sort(key=lambda s: (s.top or 0, s.left or 0))
    units = []
    for c in cards:
        members = [c] + [sh for sh in slide.shapes
                         if sh is not c and getattr(sh._element, 'spPr', None) is not None
                         and _center_in(sh, c)]
        members.sort(key=lambda s: (s.top or 0, s.left or 0))
        units.append([m.shape_id for m in members])
    return units


# ---------------------------------------------------------------- 动画：总入口


def animate_deck(prs, spec, report):
    """spec:
    {
      "transition": "fade",                 # 全篇切换
      "auto": {"pages": [4,5], "fills": ["FFFCF5"], "stagger": 150, "dur": 420},
      "slides": {
        "4": [                               # 页号（字符串或数字）
          [ {"eff":"wipe","target":["ribbon1"],"dir":"left","dur":650},
            {"eff":"pulse","target":["Rounded Rectangle 7"],"delay":0} ],
          [ {...} ]                          # 每个内层数组 = 一次点击
        ]
      },
      "reveal": [{"page": 12, "trigger": ["Rounded Rectangle 15"], "hide": ["Rounded Rectangle 52"]}]
    }
    """
    spec = spec or {}
    tr = spec.get('transition', 'fade')
    auto = spec.get('auto') or {}
    explicit = {}
    for k, v in (spec.get('slides') or {}).items():
        explicit[int(k)] = v
    total_effects = 0
    animated = 0
    for idx, slide in enumerate(prs.slides, 1):
        groups = []
        if idx in explicit:
            for grp in explicit[idx]:
                built = []
                for fx in grp:
                    eff = fx.get('eff', 'fade')
                    if eff not in EFFECTS:
                        report.append('⚠️ p%d 未知效果 %s' % (idx, eff))
                        continue
                    spids = resolve(slide, fx.get('target'))
                    if not spids:
                        report.append('⚠️ p%d 找不到目标 %s' % (idx, fx.get('target')))
                        continue
                    for spid in spids:
                        item = dict(fx)
                        item['eff'] = eff
                        item['spid'] = spid
                        built.append(item)
                if built:
                    groups.append(built)
        elif auto.get('pages') and idx in set(auto['pages']):
            units = card_units(slide, auto.get('fills') or DEFAULTS['card_fills'],
                               auto.get('min_h', 0.8), auto.get('min_w', 1.5))
            stagger, dur = auto.get('stagger', 150), auto.get('dur', 420)
            g = []
            for i, unit in enumerate(units):
                for j, spid in enumerate(unit):
                    g.append({'eff': 'fade', 'spid': spid, 'dur': dur,
                              'delay': i * stagger + j * 60})
            if g:
                groups = [g]
        # reveal 预设：点击触发器 → 隐藏的答案淡入
        for item in (spec.get('reveal') or []):
            if int(item.get('page', 0)) != idx:
                continue
            trig = resolve(slide, item.get('trigger'))
            hidden = resolve(slide, item.get('hide'))
            for k, hid in enumerate(hidden):
                g = []
                if k < len(trig):
                    g.append({'eff': 'pulse', 'spid': trig[k], 'dur': 380, 'delay': 0})
                g.append({'eff': item.get('eff', 'fade'), 'spid': hid,
                          'dur': item.get('dur', 380), 'delay': item.get('delay', 180)})
                groups.append(g)
        if tr:
            set_transition(slide, tr)
        if groups:
            set_timing(slide, groups)
            animated += 1
            total_effects += sum(len(g) for g in groups)
    report.append('切换：%s；带动画的页：%d；动画步数：%d' % (tr or '无', animated, total_effects))
    return animated, total_effects


# ---------------------------------------------------------------- 渲染（带阴影，用于自检）

_FONT_DIRS = ['/usr/share/fonts/opentype/noto/', '/usr/share/fonts/truetype/noto/']
_fc = {}


def _font_file(bold, serif):
    base = 'NotoSerifCJK' if serif else 'NotoSansCJK'
    name = base + ('-Bold' if bold else '-Regular') + '.ttc'
    for d in _FONT_DIRS:
        if os.path.exists(d + name):
            return d + name
    return None


def font(name, bold, size_pt, scale):
    from PIL import ImageFont
    key = (name, bold, round(size_pt, 1), scale)
    if key in _fc:
        return _fc[key]
    serif = bool(name) and ('Serif' in name or '宋' in name)
    px = max(6, int(round(size_pt * scale / 72.0)))
    ft = None
    path = _font_file(bold, serif)
    if path:
        try:
            ft = ImageFont.truetype(path, px, index=2)
        except Exception:
            try:
                ft = ImageFont.truetype(path, px)
            except Exception:
                ft = None
    if ft is None:
        ft = ImageFont.load_default()
    _fc[key] = ft
    return ft


def _rgb(color, default=None):
    try:
        if color and color.type is not None:
            c = color.rgb
            return (c[0], c[1], c[2])
    except Exception:
        pass
    return default


def fill_rgba(sh):
    """返回 (r,g,b,a)；a 为 0-255。"""
    try:
        if sh.fill.type != 1:
            return None
        c = sh.fill.fore_color.rgb
        alpha = 255
        spPr = getattr(sh._element, 'spPr', None)
        if spPr is not None:
            a = spPr.find(qn('a:solidFill') + '/' + qn('a:srgbClr') + '/' + qn('a:alpha'))
            if a is not None:
                alpha = int(int(a.get('val')) / 100000.0 * 255)
        return (c[0], c[1], c[2], alpha)
    except Exception:
        return None


def shadow_spec(sh):
    spPr = getattr(sh._element, 'spPr', None)
    if spPr is None:
        return None
    s = spPr.find(qn('a:effectLst') + '/' + qn('a:outerShdw'))
    if s is None:
        return None
    blur = int(s.get('blurRad') or 0) / EMU_IN
    dist = int(s.get('dist') or 0) / EMU_IN
    ang = int(s.get('dir') or 0) / 60000.0
    clr = s.find(qn('a:srgbClr'))
    alpha = 60
    col = (0x2B, 0x2A, 0x26)
    if clr is not None:
        h = clr.get('val')
        col = (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        a = clr.find(qn('a:alpha'))
        if a is not None:
            alpha = int(int(a.get('val')) / 100000.0 * 255)
    return blur, dist, ang, col, alpha


def _draw_text(dr, sh, box, scale):
    tf = sh.text_frame
    x0, y0, w, h = box

    def m(v, d):
        return (v if v is not None else d) / EMU_IN * scale
    ml, mr = m(tf.margin_left, 91440), m(tf.margin_right, 91440)
    mt, mb = m(tf.margin_top, 45720), m(tf.margin_bottom, 45720)
    wrapw = max(10, w - ml - mr)
    lines = []
    for para in tf.paragraphs:
        ls = para.line_spacing or 1.0
        align = str(para.alignment or 'LEFT')
        runs = [r for r in para.runs if r.text]
        if not runs:
            lines.append(('', font('Noto Sans CJK SC', False, 12, scale), (0, 0, 0), 0, align, 12 * scale / 72))
            continue
        base = max((r.font.size.pt if r.font.size else 12) for r in runs)
        lh = base * scale / 72.0 * ls
        cur, curw = [], 0.0
        for r in runs:
            sz = r.font.size.pt if r.font.size else 12
            ft = font(r.font.name or 'Noto Sans CJK SC', bool(r.font.bold), sz, scale)
            col = _rgb(r.font.color, (30, 30, 30))
            for ch in r.text:
                cw = dr.textlength(ch, font=ft)
                if curw + cw > wrapw and cur:
                    lines.append((None, cur, align, lh))
                    cur, curw = [], 0.0
                if ch == '\n':
                    lines.append((None, cur, align, lh))
                    cur, curw = [], 0.0
                    continue
                cur.append((ch, ft, col, cw))
                curw += cw
        if cur:
            lines.append((None, cur, align, lh))
    total = sum(l[3] for l in lines if isinstance(l[3], float))
    anchor = str(tf.vertical_anchor)
    cy = y0 + mt
    if 'MIDDLE' in anchor:
        cy = y0 + (h - total) / 2.0
    elif 'BOTTOM' in anchor:
        cy = y0 + h - mb - total
    for ln in lines:
        if ln[0] is not None:
            cy += ln[3]
            continue
        _, cur, align, lh = ln
        lw = sum(c[3] for c in cur)
        if 'CENTER' in align:
            cx = x0 + ml + (wrapw - lw) / 2.0
        elif 'RIGHT' in align:
            cx = x0 + ml + (wrapw - lw)
        else:
            cx = x0 + ml
        for ch, ft, col, cw in cur:
            dr.text((cx, cy + (lh - ft.size * 1.16) / 2.0), ch, font=ft, fill=col)
            cx += cw
        cy += lh


def render_deck(prs, outdir, pages=None, scale=130, bg=(250, 246, 240)):
    from PIL import Image, ImageDraw, ImageFilter
    W = int(emu_in(prs.slide_width) * scale)
    H = int(emu_in(prs.slide_height) * scale)
    os.makedirs(outdir, exist_ok=True)
    files = []
    for i, slide in enumerate(prs.slides, 1):
        if pages and i not in pages:
            continue
        img = Image.new('RGB', (W, H), bg)
        dr = ImageDraw.Draw(img)
        for sh in slide.shapes:
            st = str(sh.shape_type)
            box = (emu_in(sh.left) * scale, emu_in(sh.top) * scale,
                   emu_in(sh.width) * scale, emu_in(sh.height) * scale)
            if 'PICTURE' in st:
                try:
                    blob = sh.image.blob
                    im = Image.open(io.BytesIO(blob)).convert('RGB')
                    im = im.resize((max(1, int(box[2])), max(1, int(box[3]))))
                    img.paste(im, (int(box[0]), int(box[1])))
                    dr = ImageDraw.Draw(img)
                except Exception:
                    dr.rectangle([box[0], box[1], box[0] + box[2], box[1] + box[3]], outline=(200, 0, 0))
                continue
            if 'MEDIA' in st:
                dr.rectangle([box[0], box[1], box[0] + box[2], box[1] + box[3]], fill=(40, 40, 40))
                dr.text((box[0] + 8, box[1] + 8), 'VIDEO', font=font('Noto Sans CJK SC', True, 14, scale),
                        fill=(255, 255, 255))
                continue
            fill = fill_rgba(sh)
            lc = shape_line_hex(sh)
            lw = 1
            try:
                if sh.line.width:
                    lw = max(1, int(round(sh.line.width.pt * scale / 72.0)))
            except Exception:
                lw = 1
            if 'LINE' in st or 'CONNECTOR' in st:
                y = box[1]
                dr.line([box[0], y, box[0] + box[2], y],
                            fill=(_rgb_hex(lc) if lc else (200, 190, 180)), width=max(1, lw))
                continue
            rad = 0
            try:
                if sh.adjustments and len(sh.adjustments):
                    rad = int(min(box[2], box[3]) * float(sh.adjustments[0]))
            except Exception:
                rad = 0
            sp = shadow_spec(sh)
            if sp and fill:
                blur, dist, ang, col, alpha = sp
                dx = dist * scale * math.cos(math.radians(ang))
                dy = dist * scale * math.sin(math.radians(ang))
                layer = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                dl = ImageDraw.Draw(layer)
                x1, y1 = box[0] + dx, box[1] + dy
                x2, y2 = x1 + box[2], y1 + box[3]
                if rad > 0:
                    dl.rounded_rectangle([x1, y1, x2, y2], radius=rad, fill=col + (alpha,))
                else:
                    dl.rectangle([x1, y1, x2, y2], fill=col + (alpha,))
                if blur * scale > 0.5:
                    layer = layer.filter(ImageFilter.GaussianBlur(blur * scale))
                img = Image.alpha_composite(img.convert('RGBA'), layer).convert('RGB')
                dr = ImageDraw.Draw(img)
            if fill or (lc and lw):
                layer = Image.new('RGBA', (W, H), (0, 0, 0, 0))
                dl = ImageDraw.Draw(layer)
                fc = fill if fill else (0, 0, 0, 0)
                try:
                    if rad > 0:
                        dl.rounded_rectangle([box[0], box[1], box[0] + box[2], box[1] + box[3]],
                                             radius=rad, fill=fc, outline=((*_rgb_hex(lc), 255) if lc else None),
                                             width=lw)
                    else:
                        dl.rectangle([box[0], box[1], box[0] + box[2], box[1] + box[3]],
                                     fill=fc, outline=((*_rgb_hex(lc), 255) if lc else None), width=lw)
                except Exception:
                    pass
                img = Image.alpha_composite(img.convert('RGBA'), layer).convert('RGB')
                dr = ImageDraw.Draw(img)
            if sh.has_text_frame and sh.text_frame.text.strip():
                _draw_text(dr, sh, box, scale)
        path = os.path.join(outdir, 's%02d.png' % i)
        img.save(path)
        files.append(path)
    return {'files': files, 'width': W, 'height': H, 'count': len(files)}


def _rgb_hex(h):
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ---------------------------------------------------------------- 体检


def audit_deck(prs):
    pages = []
    issues = []
    sw, sh_ = emu_in(prs.slide_width), emu_in(prs.slide_height)
    for i, slide in enumerate(prs.slides, 1):
        cards = no_shadow = 0
        for sh in slide.shapes:
            if getattr(sh._element, 'spPr', None) is None:
                continue
            if emu_in(sh.height) < 0.55 or emu_in(sh.width) < 0.9:
                continue
            if shape_fill_hex(sh) is None:
                continue
            cards += 1
            if shadow_spec(sh) is None:
                no_shadow += 1
        t = slide._element.find(qn('p:timing'))
        tr = slide._element.find(qn('p:transition'))
        effects = 0
        dup = False
        if t is not None:
            effects = len(t.findall('.//' + qn('p:animEffect'))) + len(t.findall('.//' + qn('p:anim')))
            ids = [e.get('id') for e in t.iter(qn('p:cTn'))]
            dup = len(ids) != len(set(ids))
        pages.append({'page': i, 'cards': cards, 'cards_without_shadow': no_shadow,
                      'transition': tr is not None, 'effects': effects, 'dup_ids': dup})
        if no_shadow and cards:
            issues.append('p%d: %d/%d 个卡片没有阴影' % (i, no_shadow, cards))
        if dup:
            issues.append('p%d: 动画 cTn id 重复（文件可能打不开）' % i)
        if tr is None:
            issues.append('p%d: 没有页面切换' % i)
    media = []
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if 'MEDIA' in str(sh.shape_type) or (sh.has_text_frame and sh.name.lower().endswith(('.mp4', '.mov'))):
                media.append('p%d %s %.2f×%.2f (比例 %.2f)' % (
                    i, sh.name, emu_in(sh.width), emu_in(sh.height),
                    emu_in(sh.width) / emu_in(sh.height) if sh.height else 0))
    return {
        'slides': len(prs.slides),
        'size_in': [round(sw, 3), round(sh_, 3)],
        'media': media,
        'pages': pages,
        'issues': issues,
        'with_animation': sum(1 for p in pages if p['effects']),
        'total_effects': sum(p['effects'] for p in pages),
    }


# ---------------------------------------------------------------- 入口


def do_render(args):
    prs = Presentation(args['path'])
    pages = args.get('pages')
    pages = {int(x) for x in pages} if pages else None
    data = render_deck(prs, args.get('outdir') or '/tmp/ppt-render',
                       pages=pages, scale=int(args.get('scale') or 130))
    return data


def do_audit(args):
    prs = Presentation(args['path'])
    return audit_deck(prs)


def do_polish(args):
    prs = Presentation(args['path'])
    report = []
    opts = {k: v for k, v in args.items() if k in DEFAULTS or k in (
        'all_fills', 'include_pictures')}
    if args.get('delete_text'):
        delete_slides_by_text(prs, args['delete_text'], report)
        prs = prs
    polish_deck(prs, opts, report)
    if args.get('renumber'):
        renumber(prs, report)
    out = args.get('out') or args['path']
    prs.save(out)
    return {'out': out, 'slides': len(prs.slides), 'report': report}


def do_animate(args):
    prs = Presentation(args['path'])
    spec = args.get('spec') or {}
    if isinstance(spec, str):
        spec = json.loads(spec)
    report = []
    animate_deck(prs, spec, report)
    out = args.get('out') or args['path']
    prs.save(out)
    # 保存后重新打开一次，确认文件没被写坏
    try:
        again = Presentation(out)
        n = len(again.slides)
    except Exception as e:
        return {'out': out, 'slides': 0, 'report': report, 'warn': '重开失败：%s' % e}
    return {'out': out, 'slides': n, 'report': report}


def do_info(args):
    prs = Presentation(args['path'])
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        texts = [text_of(sh).strip().split('\n')[0] for sh in slide.shapes if text_of(sh).strip()]
        slides.append({'page': i, 'shapes': len(slide.shapes), 'title': texts[1] if len(texts) > 1 else (texts[0] if texts else '')})
    return {'slides': len(prs.slides),
            'size_in': [round(emu_in(prs.slide_width), 3), round(emu_in(prs.slide_height), 3)],
            'pages': slides}


CMDS = {'render': do_render, 'audit': do_audit, 'polish': do_polish,
        'animate': do_animate, 'info': do_info}


def main():
    try:
        req = json.load(sys.stdin)
    except Exception as e:
        print(json.dumps({'ok': False, 'error': 'bad request: %s' % e}))
        return
    cmd = req.get('cmd')
    args = req.get('args') or {}
    fn = CMDS.get(cmd)
    if fn is None:
        print(json.dumps({'ok': False, 'error': 'unknown cmd: %s' % cmd}))
        return
    try:
        print(json.dumps({'ok': True, 'data': fn(args)}))
    except Exception as e:
        import traceback
        print(json.dumps({'ok': False, 'error': '%s: %s' % (type(e).__name__, e),
                          'trace': traceback.format_exc()[-1200:]}))


if __name__ == '__main__':
    main()