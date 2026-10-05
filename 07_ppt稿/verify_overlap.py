#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v6 布局重叠自检：文本框两两重叠 / 文本框越界 / 文本框与图片重叠"""
import sys
from pptx import Presentation
from pptx.oxml.ns import qn

EMU = 914400.0

def rect(sh):
    if sh.left is None or sh.top is None or sh.width is None or sh.height is None:
        return None
    return (sh.left / EMU, sh.top / EMU,
            (sh.left + sh.width) / EMU, (sh.top + sh.height) / EMU)

def has_text(sh):
    return sh.has_text_frame and bool(sh.text_frame.text.strip())

def is_pic(sh):
    return sh.shape_type == 13

def overlap_area(a, b):
    x1 = max(a[0], b[0]); y1 = max(a[1], b[1])
    x2 = min(a[2], b[2]); y2 = min(a[3], b[3])
    w = x2 - x1; h = y2 - y1
    if w <= 0 or h <= 0:
        return 0.0
    return w * h

def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "240 D1上_PPT骨架_v6.pptx"
    prs = Presentation(path)
    issues = []
    for i, s in enumerate(prs.slides):
        shapes = list(s.shapes)
        texts = [sh for sh in shapes if has_text(sh)]
        pics = [sh for sh in shapes if is_pic(sh)]
        # 文本框两两重叠（排除占位符 idx 与自建框的设计重叠：标题占位符很高，只查其文字区）
        for a in range(len(texts)):
            ra = rect(texts[a])
            if ra is None:
                continue
            # 越界检查
            if ra[2] > 13.35 or ra[3] > 7.52:
                issues.append("P%02d 越界: %r → %s" % (i + 1, texts[a].name,
                                                       "r=%.2f b=%.2f" % (ra[2], ra[3])))
            for b in range(a + 1, len(texts)):
                rb = rect(texts[b])
                if rb is None:
                    continue
                ov = overlap_area(ra, rb)
                # 允许极小容差（0.02 平方英寸）
                if ov > 0.02:
                    issues.append("P%02d 文本框重叠: %r × %r (%.2f)" % (
                        i + 1, texts[a].name, texts[b].name, ov))
        # 文本框与图片重叠
        for t in texts:
            rt = rect(t)
            if rt is None:
                continue
            for pic in pics:
                rp = rect(pic)
                if rp is None:
                    continue
                ov = overlap_area(rt, rp)
                if ov > 0.02:
                    issues.append("P%02d 文本×图重叠: %r × %r (%.2f)" % (
                        i + 1, t.name, pic.name, ov))
    if issues:
        print("发现问题 %d 处:" % len(issues))
        for it in issues:
            print("  " + it)
    else:
        print("无重叠/越界问题")

if __name__ == "__main__":
    main()
