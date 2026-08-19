# -*- coding: utf-8 -*-
"""PIL 그리기 공용 헬퍼 (한글 줄바꿈 포함)."""
from PIL import ImageFont

F_B = '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf'
F_R = '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'

def font(size, bold=False):
    return ImageFont.truetype(F_B if bold else F_R, size)

def tw(d, s, f):
    b = d.textbbox((0, 0), s, font=f)
    return b[2] - b[0]

def wrap(d, text, f, maxw):
    """공백 우선 줄바꿈, 한 덩어리가 너무 길면 글자 단위로 자름."""
    lines, cur = [], ''
    for word in text.split(' '):
        trial = word if not cur else cur + ' ' + word
        if tw(d, trial, f) <= maxw:
            cur = trial
            continue
        if cur:
            lines.append(cur)
        cur = ''
        for ch in word:                      # 한글/긴 토큰은 글자 단위
            if tw(d, cur + ch, f) <= maxw:
                cur += ch
            else:
                lines.append(cur)
                cur = ch
    if cur:
        lines.append(cur)
    return lines

def text_block(d, xy, text, f, fill, maxw, leading=1.35):
    x, y = xy
    lh = int(f.size * leading)
    for ln in wrap(d, text, f, maxw):
        d.text((x, y), ln, font=f, fill=fill)
        y += lh
    return y

def center(d, cx, y, s, f, fill):
    d.text((cx - tw(d, s, f) / 2, y), s, font=f, fill=fill)
