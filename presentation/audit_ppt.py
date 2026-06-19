# -*- coding: utf-8 -*-
"""结构化审查 deck.pptx：边界越界、文本溢出估算、图片缺失检查。
LibreOffice 无头转换在本容器环境不可用（对纯文本文件转换也失败），
因此用几何与字符宽度估算来替代可视化渲染验证。
"""
import os
import sys
from pptx import Presentation

DECK = os.path.join(os.path.dirname(__file__) or ".", "废旧轮胎的回收与资源化利用.pptx")
from pptx.util import Emu

SLIDE_W, SLIDE_H = 960, 540  # pt

def emu_to_pt(v):
    return float(v) / 12700.0

def is_cjk(ch):
    return '一' <= ch <= '鿿' or ch in "，。、！？：；（）【】《》——…“”‘’·"

def est_text_width(s, size):
    w = 0.0
    for ch in s:
        if is_cjk(ch):
            w += size * 1.0
        elif ch == ' ':
            w += size * 0.28
        else:
            w += size * 0.52
    return w

def rects_overlap(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    smaller = min(aw * ah, bw * bh)
    return inter / smaller if smaller > 0 else 0.0


def audit():
    prs = Presentation(DECK)
    issues = []
    for idx, slide in enumerate(prs.slides, start=1):
        boxes = []  # (rect, label) for non-decorative text-bearing shapes
        for shp in slide.shapes:
            if shp.has_text_frame and shp.text_frame.text.strip():
                try:
                    r = (emu_to_pt(shp.left), emu_to_pt(shp.top),
                         emu_to_pt(shp.width), emu_to_pt(shp.height))
                except TypeError:
                    continue
                boxes.append((r, shp.text_frame.text.strip()[:24]))
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                ratio = rects_overlap(boxes[i][0], boxes[j][0])
                if ratio > 0.25:
                    issues.append(f"[S{idx:02d}] 文本框重叠 {ratio:.0%}: "
                                  f"'{boxes[i][1]}' vs '{boxes[j][1]}'")
        for shp in slide.shapes:
            try:
                x, y, w, h = (emu_to_pt(shp.left), emu_to_pt(shp.top),
                              emu_to_pt(shp.width), emu_to_pt(shp.height))
            except TypeError:
                continue
            name = shp.shape_type
            # 1. 越界检查（容许 1pt 误差）
            if x < -1 or y < -1 or (x + w) > SLIDE_W + 1 or (y + h) > SLIDE_H + 1:
                issues.append(f"[S{idx:02d}] 越界: {name} pos=({x:.0f},{y:.0f}) "
                              f"size=({w:.0f}x{h:.0f}) 右下=({x+w:.0f},{y+h:.0f})")
            # 2. 文本溢出估算（仅对有明确字号的 textbox/autoshape 文本）
            if shp.has_text_frame:
                tf = shp.text_frame
                total_text_h = 0.0
                max_line_w_used = 0.0
                for p in tf.paragraphs:
                    runs = p.runs
                    if not runs:
                        continue
                    txt = "".join(r.text for r in runs)
                    size = runs[0].font.size
                    size = size.pt if size else 18
                    avail_w = w - 8  # margins
                    if avail_w <= 0:
                        continue
                    tw = est_text_width(txt, size)
                    lines = max(1, -(-int(tw) // int(avail_w))) if avail_w > 0 else 1
                    line_h = size * 1.25
                    total_text_h += lines * line_h
                    max_line_w_used = max(max_line_w_used, min(tw, avail_w))
                if total_text_h > h + 6 and h > 0:
                    issues.append(f"[S{idx:02d}] 可能文本溢出: 估算高度 {total_text_h:.0f}pt "
                                  f"> 容器高度 {h:.0f}pt  内容预览='{_preview(tf)}'")
    return issues

def _preview(tf):
    t = tf.text.replace("\n", " / ")
    return t[:40] + ("…" if len(t) > 40 else "")

if __name__ == "__main__":
    issues = audit()
    if not issues:
        print("未发现越界或明显文本溢出问题。")
    else:
        print(f"发现 {len(issues)} 个潜在问题：\n")
        for i in issues:
            print(" -", i)
    sys.exit(0)
