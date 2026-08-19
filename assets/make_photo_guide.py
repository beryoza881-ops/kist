# -*- coding: utf-8 -*-
"""사진 넣는 안내표(1장) 생성 — 어느 칸에 무슨 사진을 넣는지 한눈에."""
import os, sys
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sdl_layout import SLOTS, CARD_TITLE
from _draw_util import font, tw, text_block, center

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
W, MARGIN, GAP = 2000, 64, 34
CARD_W = (W - 2 * MARGIN - 2 * GAP) // 3
NAVY, BLUE, TXT, SUB = (29, 53, 103), (27, 84, 181), (52, 66, 92), (105, 122, 152)
CARD_BG, CARD_ED, KW_BG = (247, 249, 252), (214, 222, 234), (232, 238, 248)
DIFF = {'easy': ((214, 240, 226), (26, 108, 70), '사진 찾기 쉬움'),
        'mid':  ((219, 231, 250), (25, 72, 150), '화면·장면 사진이면 OK'),
        'hard': ((253, 232, 213), (162, 84, 12), '사진으로는 부적합 · 일러스트 권장')}

im = Image.new('RGB', (W, 3400), 'white')
d = ImageDraw.Draw(im)

d.text((MARGIN, 54), '자율실험 폐루프 — 사진 넣는 안내표', font=font(74, True), fill=NAVY)
y = text_block(d, (MARGIN, 158),
               'photos/ 폴더에 아래 파일명 그대로 저장한 뒤  python3 assets/apply_photo_layout.py  실행. '
               '사진 비율은 아무거나 상관없습니다(가운데 자동 크롭).',
               font(38), SUB, W - 2 * MARGIN)
y += 34

for n, slots in SLOTS.items():
    d.rounded_rectangle([MARGIN, y, W - MARGIN, y + 84], radius=16, fill=NAVY)
    d.text((MARGIN + 30, y + 18), f'{n}번 박스   {CARD_TITLE[n]}', font=font(46, True), fill=(255, 255, 255))
    lab = '카드 안 이미지 밴드 = 아래 3장 (왼쪽 → 오른쪽 순서)'
    d.text((W - MARGIN - 30 - tw(d, lab, font(32)), y + 28), lab, font=font(32), fill=(183, 199, 226))
    y += 108

    for i, (label, kw, desc, diff) in enumerate(slots):
        x = MARGIN + i * (CARD_W + GAP)
        d.rounded_rectangle([x, y, x + CARD_W, y + 424], radius=18, fill=CARD_BG, outline=CARD_ED, width=3)
        cx = x + CARD_W // 2

        fn = f'box{n}_{i+1}.jpg'
        f_fn = font(40, True)
        pw = tw(d, fn, f_fn) + 52
        d.rounded_rectangle([cx - pw // 2, y + 22, cx + pw // 2, y + 90], radius=34, fill=BLUE)
        center(d, cx, y + 34, fn, f_fn, (255, 255, 255))

        f_l = font(46, True)
        while tw(d, label, f_l) > CARD_W - 48 and f_l.size > 32:
            f_l = font(f_l.size - 2, True)
        center(d, cx, y + 108, label, f_l, TXT)

        d.rounded_rectangle([x + 22, y + 172, x + CARD_W - 22, y + 248], radius=12, fill=KW_BG)
        center(d, cx, y + 179, '검색어', font(24, True), SUB)
        f_k = font(32, True)
        while tw(d, kw, f_k) > CARD_W - 60 and f_k.size > 22:
            f_k = font(f_k.size - 2, True)
        center(d, cx, y + 209, kw, f_k, BLUE)

        text_block(d, (x + 24, y + 270), desc, font(31), TXT, CARD_W - 48, 1.30)

        bg, fg, txt = DIFF[diff]
        f_d = font(27, True)
        cw = tw(d, txt, f_d) + 40
        d.rounded_rectangle([cx - cw // 2, y + 372, cx + cw // 2, y + 424 - 10], radius=21, fill=bg)
        center(d, cx, y + 381, txt, f_d, fg)
    y += 424 + 44

NOTE = ('주황색 3칸(box1_1 · box3_2 · box4_2)은 "AI 추천엔진 / 목표충족 자동판정 / 파인튜닝 파이프라인"처럼 '
        '실물이 없는 개념이라, 어떤 사진을 넣어도 결국 "모니터 앞 사람" 사진으로 보입니다. '
        '이 3칸만 일러스트 버전 그림을 그대로 쓰는 혼합안을 권합니다.')
from _draw_util import wrap
f_n = font(33)
nh = len(wrap(d, NOTE, f_n, W - 2 * MARGIN - 56)) * int(33 * 1.36) + 56
d.rounded_rectangle([MARGIN, y + 6, W - MARGIN, y + 6 + nh], radius=16,
                    fill=(253, 246, 238), outline=(238, 214, 186), width=3)
text_block(d, (MARGIN + 28, y + 34), NOTE, f_n, (140, 74, 12), W - 2 * MARGIN - 56, 1.36)

im.crop((0, 0, W, y + nh + 62)).save(os.path.join(ROOT, '사진_넣는_안내표.png'))
print('saved 사진_넣는_안내표.png')
