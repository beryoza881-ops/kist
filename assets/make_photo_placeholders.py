# -*- coding: utf-8 -*-
"""사진 버전용 자리표시자(placeholder) 타일 생성.

photos/box{n}_{1..3}.jpg 가 아직 없을 때 들어갈 '여기에 사진' 타일을 만든다.
실제 사진을 넣으면 자리표시자는 자동으로 대체된다.
"""
import os, sys
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sdl_layout import SUBJECTS, TILE_W, TILE_H

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'photo_placeholders')
os.makedirs(OUT, exist_ok=True)

W = 1200
H = int(round(W * TILE_H / TILE_W))          # 타일 비율(약 3:2)에 정확히 맞춤
BG, EDGE, GLYPH, TXT = (233, 237, 244), (166, 184, 216), (143, 163, 196), (78, 94, 124)
FONT = '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf'


def dashed_rect(d, box, color, width, dash=26, gap=18, r=26):
    x0, y0, x1, y1 = box
    for x in range(x0 + r, x1 - r, dash + gap):
        xe = min(x + dash, x1 - r)
        d.line([(x, y0), (xe, y0)], fill=color, width=width)
        d.line([(x, y1), (xe, y1)], fill=color, width=width)
    for y in range(y0 + r, y1 - r, dash + gap):
        ye = min(y + dash, y1 - r)
        d.line([(x0, y), (x0, ye)], fill=color, width=width)
        d.line([(x1, y), (x1, ye)], fill=color, width=width)


def tile(label, idx):
    im = Image.new('RGB', (W, H), BG)
    d = ImageDraw.Draw(im)
    dashed_rect(d, (14, 14, W - 15, H - 15), EDGE, 5)

    # 사진 아이콘 (산 + 해)
    cx, cy, w, h = W // 2, int(H * 0.40), 300, 210
    fx, fy = cx - w // 2, cy - h // 2
    d.rounded_rectangle([fx, fy, fx + w, fy + h], radius=18, outline=GLYPH, width=9)
    d.ellipse([fx + 42, fy + 36, fx + 92, fy + 86], fill=GLYPH)
    d.polygon([(fx + 26, fy + h - 26), (fx + 120, fy + 74), (fx + 196, fy + h - 26)], fill=GLYPH)
    d.polygon([(fx + 150, fy + h - 26), (fx + 214, fy + 104), (fx + w - 26, fy + h - 26)], fill=GLYPH)

    f_lab = ImageFont.truetype(FONT, 76)
    f_num = ImageFont.truetype(FONT, 60)
    tw = d.textbbox((0, 0), label, font=f_lab)[2]
    while tw > W - 90 and f_lab.size > 40:
        f_lab = ImageFont.truetype(FONT, f_lab.size - 4)
        tw = d.textbbox((0, 0), label, font=f_lab)[2]
    d.text(((W - tw) // 2, int(H * 0.68)), label, font=f_lab, fill=TXT)
    d.ellipse([30, 30, 104, 104], fill=(29, 53, 103))
    nw = d.textbbox((0, 0), str(idx), font=f_num)
    d.text((67 - (nw[2] - nw[0]) / 2, 67 - (nw[3] + nw[1]) / 2), str(idx), font=f_num, fill=(255, 255, 255))
    return im


if __name__ == '__main__':
    for n, labels in SUBJECTS.items():
        for i, lab in enumerate(labels, 1):
            p = os.path.join(OUT, f'box{n}_{i}.png')
            tile(lab, i).save(p)
            print(p)
