# -*- coding: utf-8 -*-
"""사진 버전용 자리표시자 타일 생성.

각 타일에 '넣어야 할 파일명 + 피사체 + 검색어'를 직접 찍어 넣어서,
PPT를 확대하기만 하면 어느 칸에 무엇을 넣어야 하는지 바로 보이게 한다.
"""
import os, sys
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sdl_layout import SLOTS, TILE_W, TILE_H
from _draw_util import font, tw, wrap, center

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'photo_placeholders')
os.makedirs(OUT, exist_ok=True)

W = 1200
H = int(round(W * TILE_H / TILE_W))
BG, EDGE, GLYPH = (233, 237, 244), (166, 184, 216), (150, 169, 201)
NAVY, TXT, SUB = (29, 53, 103), (52, 66, 92), (105, 122, 152)
DIFF = {'easy': ((214, 240, 226), (26, 108, 70), '사진 찾기 쉬움'),
        'mid':  ((219, 231, 250), (25, 72, 150), '화면·장면 사진'),
        'hard': ((253, 232, 213), (162, 84, 12), '사진 부적합 · 일러스트 권장')}


def dashed(d, box, color, width, dash=26, gap=18, r=26):
    x0, y0, x1, y1 = box
    for x in range(x0 + r, x1 - r, dash + gap):
        xe = min(x + dash, x1 - r)
        d.line([(x, y0), (xe, y0)], fill=color, width=width)
        d.line([(x, y1), (xe, y1)], fill=color, width=width)
    for y in range(y0 + r, y1 - r, dash + gap):
        ye = min(y + dash, y1 - r)
        d.line([(x0, y), (x0, ye)], fill=color, width=width)
        d.line([(x1, y), (x1, ye)], fill=color, width=width)


def tile(fname, label, keyword, difficulty):
    im = Image.new('RGB', (W, H), BG)
    d = ImageDraw.Draw(im)
    dashed(d, (14, 14, W - 15, H - 15), EDGE, 5)

    f_file = font(58, True)
    pw = tw(d, fname, f_file) + 76
    d.rounded_rectangle([(W - pw) // 2, 36, (W + pw) // 2, 132], radius=48, fill=NAVY)
    center(d, W // 2, 56, fname, f_file, (255, 255, 255))

    # 사진 아이콘
    gw, gh, gy = 230, 160, 180
    gx = (W - gw) // 2
    d.rounded_rectangle([gx, gy, gx + gw, gy + gh], radius=16, outline=GLYPH, width=8)
    d.ellipse([gx + 32, gy + 26, gx + 74, gy + 68], fill=GLYPH)
    d.polygon([(gx + 20, gy + gh - 20), (gx + 92, gy + 56), (gx + 150, gy + gh - 20)], fill=GLYPH)
    d.polygon([(gx + 116, gy + gh - 20), (gx + 164, gy + 80), (gx + gw - 20, gy + gh - 20)], fill=GLYPH)

    f_lab = font(70, True)
    while tw(d, label, f_lab) > W - 80 and f_lab.size > 44:
        f_lab = font(f_lab.size - 4, True)
    center(d, W // 2, 380, label, f_lab, TXT)

    f_kw = font(44)
    y = 470
    for ln in wrap(d, '검색어  ' + keyword, f_kw, W - 160):
        center(d, W // 2, y, ln, f_kw, SUB)
        y += 58

    bg, fg, txt = DIFF[difficulty]
    f_d = font(40, True)
    cw = tw(d, txt, f_d) + 56
    y = H - 108
    d.rounded_rectangle([(W - cw) // 2, y, (W + cw) // 2, y + 66], radius=33, fill=bg)
    center(d, W // 2, y + 12, txt, f_d, fg)
    return im


if __name__ == '__main__':
    for n, slots in SLOTS.items():
        for i, (label, kw, _desc, diff) in enumerate(slots, 1):
            fname = f'box{n}_{i}.jpg'
            tile(fname, label, kw, diff).save(os.path.join(OUT, f'box{n}_{i}.png'))
            print(fname, '->', diff)
