# -*- coding: utf-8 -*-
"""자율실험 폐루프 1~4번 박스용 '장면 일러스트' 12장 생성.

사진을 구할 수 없는 환경이라 직접 손으로 코딩한 벡터 장면이다.
원근·명암·재질 그라디언트를 넣어 실사에 가깝게 그렸다. (외부 저작물 미사용)
사진 버전과 동일한 3:2 타일 규격이라 그대로 교체해 쓸 수 있다.
"""
import os, sys, math
import cairosvg
from PIL import Image, ImageFilter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sdl_layout import TILE_W, TILE_H

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'scene_images')
os.makedirs(OUT, exist_ok=True)
W = 1200
H = int(round(W * TILE_H / TILE_W))          # 830


class S:
    """작은 SVG 빌더 (그라디언트 id 자동 부여)."""
    def __init__(self):
        self.defs, self.body, self._n = [], [], 0

    def _sid(self, p):
        self._n += 1
        return f'{p}{self._n}'

    @staticmethod
    def _stops(stops):
        out = ''
        for st in stops:
            o, c = st[0], st[1]
            a = st[2] if len(st) > 2 else None
            op = '' if a is None else f' stop-opacity="{a}"'
            out += f'<stop offset="{o}" stop-color="{c}"{op}/>'
        return out

    def lg(self, stops, x1=0, y1=0, x2=0, y2=1):
        i = self._sid('lg')
        self.defs.append(f'<linearGradient id="{i}" x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}">'
                         f'{self._stops(stops)}</linearGradient>')
        return f'url(#{i})'

    def rg(self, stops, cx=.5, cy=.5, r=.5):
        i = self._sid('rg')
        self.defs.append(f'<radialGradient id="{i}" cx="{cx}" cy="{cy}" r="{r}">'
                         f'{self._stops(stops)}</radialGradient>')
        return f'url(#{i})'

    def add(self, *s):
        self.body.extend(s)

    def shadow(self, cx, cy, rx, ry, op=.45):
        g = self.rg([(0, '#000', op), (.55, '#000', op * .45), (1, '#000', 0)])
        self.add(f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="{g}"/>')

    def vignette(self, op=.40):
        g = self.rg([(.52, '#000', 0), (1, '#000', op)], .5, .47, .80)
        self.add(f'<rect width="{W}" height="{H}" fill="{g}"/>')

    def out(self):
        return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
                f'viewBox="0 0 {W} {H}"><defs>{"".join(self.defs)}'
                f'<clipPath id="cp"><rect width="{W}" height="{H}" rx="22"/></clipPath>'
                f'</defs><g clip-path="url(#cp)">{"".join(self.body)}</g></svg>')


def screen_content_glow(s, x, y, w, h, color='#1B54B5', op=.30):
    g = s.rg([(0, color, op), (1, color, 0)], .5, .5, .6)
    s.add(f'<rect x="{x-w*.35}" y="{y-h*.35}" width="{w*1.7}" height="{h*1.7}" fill="{g}"/>')


def bezel(s, x, y, w, h, r=14):
    """모니터 베젤 + 화면 안쪽"""
    body = s.lg([(0, '#2C3446'), (.5, '#1B2231'), (1, '#10151F')])
    s.add(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" fill="{body}"/>')
    s.add(f'<rect x="{x+2}" y="{y+2}" width="{w-4}" height="{h*.5}" rx="{r}" '
          f'fill="#FFFFFF" opacity="0.05"/>')
    pad = max(10, w * 0.018)
    sx, sy, sw, sh = x + pad, y + pad, w - 2 * pad, h - 2 * pad
    scr = s.lg([(0, '#0C1526'), (1, '#070C16')])
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh}" rx="{r*.5}" fill="{scr}"/>')
    return sx, sy, sw, sh


def code_lines(s, x, y, w, rows, lh=18, seed=1):
    cols = ['#5FA8FF', '#8AD6A0', '#E6C07B', '#C678DD', '#7F8FA6', '#61AFEF']
    out, r = [], seed
    for i in range(rows):
        r = (r * 1103515245 + 12345) % 2147483648
        n = 2 + r % 3
        cx = x
        for k in range(n):
            r = (r * 1103515245 + 12345) % 2147483648
            seg = (w / n) * (0.35 + (r % 60) / 100.0)
            out.append(f'<rect x="{cx:.0f}" y="{y+i*lh:.0f}" width="{seg:.0f}" height="{lh*.42:.0f}" '
                       f'rx="{lh*.2:.0f}" fill="{cols[(i+k) % len(cols)]}" opacity="0.85"/>')
            cx += seg + 12
            if cx > x + w:
                break
    s.add(*out)


def peaks(s, x, y, w, h, color='#5FD0FF', pts=None):
    """크로마토그램/스펙트럼 곡선"""
    if pts is None:
        pts = [(0, .05), (.12, .10), (.2, .62), (.27, .12), (.36, .22), (.44, .95),
               (.52, .18), (.63, .40), (.71, .10), (.82, .70), (.9, .12), (1, .06)]
    d = f'M{x},{y+h}'
    for px, py in pts:
        d += f' L{x+px*w:.1f},{y+h-py*h:.1f}'
    d += f' L{x+w},{y+h} Z'
    fill = s.lg([(0, color, .40), (1, color, .02)])
    s.add(f'<path d="{d}" fill="{fill}"/>')
    d2 = 'M' + ' L'.join(f'{x+px*w:.1f},{y+h-py*h:.1f}' for px, py in pts)
    s.add(f'<path d="{d2}" fill="none" stroke="{color}" stroke-width="4" '
          f'stroke-linejoin="round" stroke-linecap="round"/>')


def grid(s, x, y, w, h, nx=8, ny=5, color='#7FA8D8', op=.16):
    for i in range(1, nx):
        s.add(f'<line x1="{x+w*i/nx:.1f}" y1="{y}" x2="{x+w*i/nx:.1f}" y2="{y+h}" '
              f'stroke="{color}" stroke-width="1.5" opacity="{op}"/>')
    for i in range(1, ny):
        s.add(f'<line x1="{x}" y1="{y+h*i/ny:.1f}" x2="{x+w}" y2="{y+h*i/ny:.1f}" '
              f'stroke="{color}" stroke-width="1.5" opacity="{op}"/>')


# =====================================================================
# 1-1  AI 추천엔진 : 어두운 연구실 책상 위 모니터 (코드 + 신경망 오버레이)
# =====================================================================
def scene_1_1():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#141C2C"),(1,"#0A0F1A")])}"/>')
    s.add(f'<rect y="{H*.72}" width="{W}" height="{H*.28}" '
          f'fill="{s.lg([(0,"#1B2434"),(1,"#0D131D")])}"/>')
    screen_content_glow(s, 240, 130, 720, 430, '#2A6BE0', .40)
    s.shadow(600, 640, 330, 34, .55)

    sx, sy, sw, sh = bezel(s, 210, 110, 780, 470, 16)
    grid(s, sx, sy, sw, sh, 9, 6)
    code_lines(s, sx + 26, sy + 30, sw * .42, 11, 30, 7)
    peaks(s, sx + sw * .50, sy + sh * .46, sw * .44, sh * .40, '#5FD0FF')
    # 신경망 오버레이
    lay = [[.58, .30], [.58, .46], [.58, .62]], [[.72, .24], [.72, .40], [.72, .56], [.72, .72]], [[.86, .34], [.86, .58]]
    pts = [[(sx + a * sw, sy + b * sh) for a, b in L] for L in lay]
    for A, B in zip(pts, pts[1:]):
        for p in A:
            for q in B:
                s.add(f'<line x1="{p[0]:.0f}" y1="{p[1]:.0f}" x2="{q[0]:.0f}" y2="{q[1]:.0f}" '
                      f'stroke="#7FB6FF" stroke-width="2" opacity="0.45"/>')
    for i, L in enumerate(pts):
        for p in L:
            s.add(f'<circle cx="{p[0]:.0f}" cy="{p[1]:.0f}" r="11" '
                  f'fill="{"#FFFFFF" if i==1 else "#5FA8FF"}" opacity="0.95"/>')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh*.5}" rx="8" fill="#FFFFFF" opacity="0.045"/>')
    # 스탠드 + 키보드
    s.add(f'<path d="M560,580 L640,580 L664,646 L536,646 Z" fill="{s.lg([(0,"#232C3C"),(1,"#151B27")])}"/>')
    s.add(f'<rect x="470" y="642" width="260" height="16" rx="8" fill="#2A3446"/>')
    kb = s.lg([(0, '#28303F'), (1, '#171D28')])
    s.add(f'<path d="M330,700 L870,700 L910,758 L290,758 Z" fill="{kb}"/>')
    for r_ in range(3):
        for c_ in range(14):
            s.add(f'<rect x="{330+c_*38+r_*7}" y="{710+r_*15}" width="26" height="9" rx="3" '
                  f'fill="#48566E" opacity="0.75"/>')
    s.vignette(.45)
    return s


# =====================================================================
# 1-2  과거 실험 DB : 데이터센터 서버랙 통로 (1점 투시)
# =====================================================================
def scene_1_2():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#131E33"),(.6,"#0A1220"),(1,"#05080F")])}"/>')
    # 바닥
    s.add(f'<rect y="{H*.80}" width="{W}" height="{H*.20}" '
          f'fill="{s.lg([(0,"#16203300"),(0,"#162033"),(1,"#080C14")])}"/>')
    racks = [(56, 236, 96), (330, 250, 78), (620, 250, 78), (912, 232, 96)]
    top, bot = 92, H * .80
    for i, (x, w_, side) in enumerate(racks):
        d = -1 if i < 2 else 1                      # 바깥쪽으로 측면이 보이게
        # 측면(원근)
        sxx = x - side if d < 0 else x + w_
        s.add(f'<path d="M{sxx},{top+26} L{x if d<0 else x+w_},{top} '
              f'L{x if d<0 else x+w_},{bot:.0f} L{sxx},{bot-26:.0f} Z" '
              f'fill="{s.lg([(0,"#131A28"),(1,"#080C14")],0,0,1,0)}"/>')
        # 정면
        face = s.lg([(0, '#3A4惑'.replace('惑','A'), ), ], ) if False else s.lg(
            [(0, '#39455C'), (.18, '#252F42'), (.82, '#1B2334'), (1, '#0E141F')], 0, 0, 1, 0)
        s.add(f'<rect x="{x}" y="{top}" width="{w_}" height="{bot-top:.0f}" rx="6" fill="{face}"/>')
        s.add(f'<rect x="{x}" y="{top}" width="{w_}" height="26" rx="6" fill="#4A5670"/>')
        s.add(f'<rect x="{x+6}" y="{top+4}" width="{w_-12}" height="8" rx="4" fill="#FFFFFF" opacity="0.16"/>')
        # 서버 유닛 + LED
        units = 15
        uh = (bot - top - 46) / units
        for u in range(units):
            uy = top + 34 + u * uh
            s.add(f'<rect x="{x+8}" y="{uy:.1f}" width="{w_-16}" height="{uh*.74:.1f}" rx="3" '
                  f'fill="{s.lg([(0,"#2B3purple"),(1,"#151C2A")]) if False else s.lg([(0,"#2B3548"),(1,"#151C2A")])}"/>')
            s.add(f'<rect x="{x+8}" y="{uy:.1f}" width="{w_-16}" height="{uh*.24:.1f}" rx="3" '
                  f'fill="#FFFFFF" opacity="0.05"/>')
            for k in range(3):
                col = ['#57E39B', '#57E39B', '#5FA8FF', '#E3C557'][(u + k + i) % 4]
                s.add(f'<rect x="{x+16+k*13}" y="{uy+uh*.24:.1f}" width="6" height="{uh*.3:.1f}" rx="3" '
                      f'fill="{col}" opacity="{0.95 if (u+k)%5 else 0.35}"/>')
            s.add(f'<rect x="{x+w_-34}" y="{uy+uh*.24:.1f}" width="22" height="{uh*.28:.1f}" rx="3" '
                  f'fill="#0C1119" opacity="0.8"/>')
        # 바닥 반사
        ref = s.lg([(0, '#5A7BB0', .22), (1, '#5A7BB0', 0)])
        s.add(f'<rect x="{x}" y="{bot:.0f}" width="{w_}" height="{H*.16:.0f}" fill="{ref}"/>')
    # 통로 조명 + 공기감
    for i, (cx0, op) in enumerate(((278, .30), (566, .34), (858, .30))):
        g = s.rg([(0, '#9FC4F0', op), (1, '#9FC4F0', 0)], .5, .5, .5)
        s.add(f'<rect x="{cx0-150}" y="{top-40}" width="300" height="{H*.9:.0f}" fill="{g}"/>')
        s.add(f'<rect x="{cx0-56}" y="{top-34}" width="112" height="14" rx="7" fill="#CFE2FA" opacity="0.75"/>')
    haze = s.lg([(0, '#6E9BD6', .10), (1, '#6E9BD6', 0)])
    s.add(f'<rect width="{W}" height="{H*.6:.0f}" fill="{haze}"/>')
    s.vignette(.52)
    return s


# =====================================================================
# 1-3  가상실험 모델 : 분자 3D 모형 (구-막대)
# =====================================================================
def scene_1_3():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.rg([(0,"#1D2E4C"),(1,"#080C16")],.42,.34,.85)}"/>')
    grid(s, 0, 0, W, H, 10, 7, '#6E9BD6', .07)
    atoms = [(430, 300, 78, '#3E6FD6', '#16305F'), (660, 210, 58, '#D8DEE9', '#6A7385'),
             (620, 470, 92, '#2FA8A0', '#0E4A46'), (860, 360, 62, '#D8DEE9', '#6A7385'),
             (350, 540, 54, '#E0664F', '#6E2418')]
    bonds = [(0, 1), (0, 2), (2, 3), (0, 4), (2, 1)]
    s.shadow(620, 700, 300, 40, .55)
    for a, b in bonds:                                # 결합(막대)
        x1, y1, _, _, _ = atoms[a]
        x2, y2, _, _, _ = atoms[b]
        ang = math.atan2(y2 - y1, x2 - x1)
        nx, ny = -math.sin(ang) * 13, math.cos(ang) * 13
        g = s.lg([(0, '#E4E9F2'), (.5, '#A9B4C6'), (1, '#5D6779')],
                 0.5 - nx / 40, 0.5 - ny / 40, 0.5 + nx / 40, 0.5 + ny / 40)
        s.add(f'<path d="M{x1+nx:.0f},{y1+ny:.0f} L{x2+nx:.0f},{y2+ny:.0f} '
              f'L{x2-nx:.0f},{y2-ny:.0f} L{x1-nx:.0f},{y1-ny:.0f} Z" fill="{g}"/>')
    for (x, y, r, c1, c2) in sorted(atoms, key=lambda a: a[1]):
        g = s.rg([(0, '#FFFFFF', .95), (.18, c1), (.75, c1), (1, c2)], .34, .30, .82)
        s.add(f'<circle cx="{x}" cy="{y}" r="{r}" fill="{g}"/>')
        hl = s.rg([(0, '#FFFFFF', .85), (1, '#FFFFFF', 0)], .5, .5, .5)
        s.add(f'<ellipse cx="{x-r*.32:.0f}" cy="{y-r*.38:.0f}" rx="{r*.32:.0f}" ry="{r*.24:.0f}" fill="{hl}"/>')
    s.vignette(.48)
    return s


# =====================================================================
# 2-1  로봇팔 : 밝은 실험실 벤치 위 6축 로봇 암 (바이알 파지)
# =====================================================================
def scene_2_1():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#DAE2ED"),(.55,"#B3BFD0"),(1,"#8896AC")])}"/>')
    for cx0_, ry_, op_ in ((190, 150, .30), (430, 110, .22), (980, 170, .28)):
        g = s.rg([(0, '#93A2B8', op_), (1, '#93A2B8', 0)], .5, .5, .5)
        s.add(f'<ellipse cx="{cx0_}" cy="230" rx="230" ry="{ry_}" fill="{g}"/>')
    s.add(f'<rect y="{H*.78}" width="{W}" height="{H*.22}" '
          f'fill="{s.lg([(0,"#C2CBD9"),(1,"#7C889C")])}"/>')
    s.add(f'<rect y="{H*.78}" width="{W}" height="7" fill="#FFFFFF" opacity="0.65"/>')
    s.shadow(560, 672, 260, 26, .40)

    steel = lambda a, b, c: s.lg([(0, a), (.42, b), (1, c)], 0, 0, 1, 0)
    base = steel('#D5DCE7', '#9AA6B8', '#5E6A7C')
    s.add(f'<path d="M420,660 L700,660 L668,586 L452,586 Z" fill="{base}"/>')
    s.add(f'<rect x="470" y="556" width="180" height="42" rx="12" fill="{steel("#C9D2DE","#8E9BAE","#59647A")}"/>')

    def link(x1, y1, x2, y2, w_):
        ang = math.atan2(y2 - y1, x2 - x1)
        nx, ny = -math.sin(ang) * w_ / 2, math.cos(ang) * w_ / 2
        g = s.lg([(0, '#EAEFF6'), (.35, '#B4BFCF'), (.75, '#7C8798'), (1, '#4E586A')],
                 0.5 - nx / (w_ * 1.2), 0.5 - ny / (w_ * 1.2), 0.5 + nx / (w_ * 1.2), 0.5 + ny / (w_ * 1.2))
        s.add(f'<path d="M{x1+nx:.0f},{y1+ny:.0f} L{x2+nx:.0f},{y2+ny:.0f} '
              f'L{x2-nx:.0f},{y2-ny:.0f} L{x1-nx:.0f},{y1-ny:.0f} Z" fill="{g}"/>')

    sh, el, wr = (560, 556), (386, 340), (700, 236)
    link(*sh, *el, 70)
    link(*el, *wr, 58)
    for (jx, jy, jr) in ((sh[0], sh[1], 40), (el[0], el[1], 36), (wr[0], wr[1], 30)):
        jg = s.rg([(0, '#FFFFFF'), (.35, '#C3CCDA'), (1, '#5A6577')], .35, .32, .8)
        s.add(f'<circle cx="{jx}" cy="{jy}" r="{jr}" fill="{jg}"/>')
        s.add(f'<circle cx="{jx}" cy="{jy}" r="{jr*.34:.0f}" fill="#37414F"/>')
    # 그리퍼 + 바이알
    s.add(f'<rect x="694" y="206" width="46" height="62" rx="10" fill="{steel("#D8DFE9","#94A0B2","#5A6577")}"/>')
    for dy in (-56, 34):
        s.add(f'<rect x="736" y="{236+dy}" width="70" height="22" rx="9" '
              f'fill="{steel("#CBD4E0","#8A96A8","#525D6E")}"/>')
    s.add(f'<rect x="762" y="184" width="52" height="104" rx="14" fill="#E9F1F7" opacity="0.92"/>')
    s.add(f'<rect x="762" y="228" width="52" height="60" rx="12" fill="#E2A63C" opacity="0.92"/>')
    s.add(f'<rect x="770" y="192" width="12" height="88" rx="6" fill="#FFFFFF" opacity="0.55"/>')
    s.add(f'<rect x="758" y="172" width="60" height="16" rx="7" fill="#8C97A8"/>')
    s.vignette(.40)
    return s


# =====================================================================
# 2-2  액체 핸들러 : 8채널 피펫 헤드 + 96웰 플레이트 (원근)
# =====================================================================
def scene_2_2():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#DDE4EE"),(.55,"#B0BDCF"),(1,"#8593A9")])}"/>')
    g = s.rg([(0, '#93A2B8', .26), (1, '#93A2B8', 0)], .5, .5, .5)
    s.add(f'<ellipse cx="600" cy="150" rx="560" ry="180" fill="{g}"/>')
    s.add(f'<rect y="{H*.74}" width="{W}" height="{H*.26}" '
          f'fill="{s.lg([(0,"#C4CDDA"),(1,"#7B8799")])}"/>')
    s.shadow(600, 690, 330, 30, .38)
    # 헤드
    head = s.lg([(0, '#E7EDF5'), (.4, '#A9B5C7'), (1, '#5C6779')])
    s.add(f'<rect x="300" y="70" width="600" height="120" rx="18" fill="{head}"/>')
    s.add(f'<rect x="300" y="76" width="600" height="40" rx="16" fill="#FFFFFF" opacity="0.35"/>')
    s.add(f'<rect x="360" y="190" width="480" height="46" rx="12" fill="{s.lg([(0,"#8C99AC"),(1,"#4E586A")])}"/>')
    for i in range(8):                                 # 팁
        x = 396 + i * 58
        s.add(f'<path d="M{x-15},236 L{x+15},236 L{x+8},352 L{x-8},352 Z" '
              f'fill="{s.lg([(0,"#FFFFFF"),(1,"#B9C4D3")],0,0,1,0)}" opacity="0.95"/>')
        s.add(f'<path d="M{x-9},300 L{x+9},300 L{x+6},350 L{x-6},350 Z" fill="#3F8FD8" opacity="0.75"/>')
        s.add(f'<ellipse cx="{x}" cy="{378+((i*37)%22)}" rx="7" ry="9" fill="#3F8FD8" opacity="0.85"/>')
    # 96웰 플레이트 (사다리꼴 원근)
    pl = s.lg([(0, '#FFFFFF'), (1, '#C3CDDC')])
    s.add(f'<path d="M250,470 L950,470 L1010,650 L190,650 Z" fill="{pl}"/>')
    s.add(f'<path d="M250,470 L950,470 L960,494 L240,494 Z" fill="#FFFFFF" opacity="0.7"/>')
    rows, cols = 6, 12
    for r_ in range(rows):
        t = r_ / (rows - 1.0)
        y = 500 + t * 128
        half = 340 + t * 62
        rr = 8 + t * 4.5
        for c_ in range(cols):
            u = (c_ + .5) / cols
            x = 600 - half + u * half * 2
            filled = (c_ + r_) % 3 != 2
            s.add(f'<ellipse cx="{x:.0f}" cy="{y:.0f}" rx="{rr:.1f}" ry="{rr*.78:.1f}" '
                  f'fill="{"#4C9BE0" if filled else "#DCE3EC"}" opacity="{0.9 if filled else 1}"/>')
    s.add(f'<path d="M250,470 L950,470 L1010,650 L190,650 Z" fill="none" stroke="#8894A6" stroke-width="4"/>')
    s.vignette(.40)
    return s


# =====================================================================
# 2-3  측정장비 : 분석기기 (전면 패널 + 스펙트럼 화면 + 샘플 트레이)
# =====================================================================
def scene_2_3():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#D8E0EB"),(.55,"#AEBACC"),(1,"#8290A6")])}"/>')
    s.add(f'<rect y="{H*.80}" width="{W}" height="{H*.20}" '
          f'fill="{s.lg([(0,"#BCC6D5"),(1,"#77839A")])}"/>')
    s.shadow(600, 690, 350, 28, .40)
    body = s.lg([(0, '#F3F6FA'), (.35, '#DDE3EC'), (1, '#9FAABB')])
    s.add(f'<rect x="150" y="150" width="900" height="530" rx="26" fill="{body}"/>')
    s.add(f'<rect x="150" y="150" width="900" height="170" rx="26" fill="#FFFFFF" opacity="0.45"/>')
    s.add(f'<rect x="150" y="628" width="900" height="52" rx="20" fill="#7E8B9E"/>')
    sx, sy, sw, sh = 210, 214, 470, 300
    s.add(f'<rect x="{sx-10}" y="{sy-10}" width="{sw+20}" height="{sh+20}" rx="14" fill="#2B3444"/>')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh}" rx="8" '
          f'fill="{s.lg([(0,"#0E1A2E"),(1,"#08101E")])}"/>')
    grid(s, sx, sy, sw, sh, 8, 5)
    peaks(s, sx + 16, sy + 44, sw - 32, sh - 90, '#5FD0FF')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh*.42}" rx="8" fill="#FFFFFF" opacity="0.05"/>')
    # 우측 컨트롤
    for i in range(4):
        s.add(f'<rect x="740" y="{224+i*62}" width="250" height="42" rx="10" fill="#C6CFDC"/>')
        s.add(f'<rect x="748" y="{232+i*62}" width="{80+i*46}" height="26" rx="8" '
              f'fill="{"#2F7DD6" if i<2 else "#8B97A8"}"/>')
    kg = s.rg([(0, '#FFFFFF'), (.4, '#BCC6D4'), (1, '#5C6778')], .34, .3, .85)
    s.add(f'<circle cx="800" cy="540" r="46" fill="{kg}"/>')
    s.add(f'<rect x="796" y="504" width="8" height="24" rx="4" fill="#39424F"/>')
    s.add(f'<circle cx="905" cy="520" r="15" fill="#4ADE80"/>')
    s.add(f'<circle cx="955" cy="520" r="15" fill="#E3C557" opacity="0.8"/>')
    # 샘플 트레이
    s.add(f'<rect x="210" y="546" width="470" height="82" rx="12" fill="#AEB9C9"/>')
    for i in range(9):
        x = 240 + i * 50
        s.add(f'<rect x="{x}" y="556" width="30" height="62" rx="9" fill="#EAF1F7" opacity="0.95"/>')
        s.add(f'<rect x="{x}" y="586" width="30" height="32" rx="8" '
              f'fill="{["#E2A63C","#4C9BE0","#7BC96F"][i%3]}" opacity="0.9"/>')
    s.vignette(.40)
    return s


# =====================================================================
# 3-1  측정데이터 분석 : 모니터 화면 가득 크로마토그램
# =====================================================================
def scene_3_1():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#111A2A"),(1,"#080D16")])}"/>')
    screen_content_glow(s, 120, 90, 960, 620, '#2A6BE0', .32)
    s.shadow(600, 742, 300, 26, .5)
    sx, sy, sw, sh = bezel(s, 90, 70, 1020, 620, 18)
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="52" fill="#122036"/>')
    for i, w_ in enumerate((120, 90, 74)):
        s.add(f'<rect x="{sx+26+i*150}" y="{sy+18}" width="{w_}" height="16" rx="8" '
              f'fill="#4B6AA8" opacity="0.8"/>')
    s.add(f'<circle cx="{sx+sw-40}" cy="{sy+26}" r="9" fill="#4ADE80"/>')
    plot_x, plot_y = sx + 90, sy + 86
    plot_w, plot_h = sw - 140, sh - 190
    grid(s, plot_x, plot_y, plot_w, plot_h, 10, 6)
    s.add(f'<line x1="{plot_x}" y1="{plot_y}" x2="{plot_x}" y2="{plot_y+plot_h}" stroke="#6E8BB8" stroke-width="3"/>')
    s.add(f'<line x1="{plot_x}" y1="{plot_y+plot_h}" x2="{plot_x+plot_w}" y2="{plot_y+plot_h}" stroke="#6E8BB8" stroke-width="3"/>')
    peaks(s, plot_x, plot_y, plot_w, plot_h, '#5FD0FF')
    peaks(s, plot_x, plot_y, plot_w, plot_h, '#8AD6A0',
          [(0,.03),(.1,.06),(.22,.30),(.3,.08),(.42,.55),(.5,.10),(.6,.20),(.7,.42),(.8,.09),(.92,.26),(1,.04)])
    for px in (.44, .82):
        s.add(f'<line x1="{plot_x+plot_w*px:.0f}" y1="{plot_y}" x2="{plot_x+plot_w*px:.0f}" '
              f'y2="{plot_y+plot_h}" stroke="#E3C557" stroke-width="2.5" stroke-dasharray="10 8" opacity="0.9"/>')
    for i in range(6):
        s.add(f'<rect x="{plot_x+i*(plot_w/6)+14:.0f}" y="{plot_y+plot_h+26}" width="{plot_w/6-40:.0f}" '
              f'height="12" rx="6" fill="#3C5480" opacity="0.9"/>')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh*.42}" rx="8" fill="#FFFFFF" opacity="0.04"/>')
    s.add(f'<path d="M540,690 L660,690 L684,742 L516,742 Z" fill="{s.lg([(0,"#232C3C"),(1,"#141A25")])}"/>')
    s.add(f'<rect x="450" y="738" width="300" height="18" rx="9" fill="#2A3446"/>')
    s.vignette(.42)
    return s


# =====================================================================
# 3-2  목표충족 자동판정 : QC 판정 UI (규격선 + PASS)
# =====================================================================
def scene_3_2():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#0F1A2C"),(1,"#070C15")])}"/>')
    screen_content_glow(s, 90, 70, 1020, 690, '#1FA36B', .26)
    sx, sy, sw, sh = bezel(s, 70, 50, 1060, 730, 20)
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="66" fill="#12213A"/>')
    s.add(f'<rect x="{sx+26}" y="{sy+22}" width="210" height="20" rx="10" fill="#5E86C4"/>')
    s.add(f'<rect x="{sx+sw-190}" y="{sy+20}" width="150" height="26" rx="13" fill="#1F7A52"/>')

    px, py = sx + 70, sy + 132
    pw, ph = sw - 420, sh - 260
    grid(s, px, py, pw, ph, 8, 5)
    s.add(f'<line x1="{px}" y1="{py+ph}" x2="{px+pw}" y2="{py+ph}" stroke="#6E8BB8" stroke-width="3"/>')
    # 규격 상/하한 밴드
    s.add(f'<rect x="{px}" y="{py+ph*.18:.0f}" width="{pw}" height="{ph*.52:.0f}" fill="#4ADE80" opacity="0.10"/>')
    for yy, col in ((py + ph * .18, '#4ADE80'), (py + ph * .70, '#4ADE80')):
        s.add(f'<line x1="{px}" y1="{yy:.0f}" x2="{px+pw}" y2="{yy:.0f}" stroke="{col}" '
              f'stroke-width="3" stroke-dasharray="14 10" opacity="0.85"/>')
    vals = [.52, .44, .60, .38, .55, .47, .58, .41, .50, .56]
    ptsd = []
    for i, v in enumerate(vals):
        x = px + pw * (i + .5) / len(vals)
        y = py + ph * (1 - v)
        ptsd.append((x, y))
    s.add('<path d="M' + ' L'.join(f'{x:.0f},{y:.0f}' for x, y in ptsd) +
          f'" fill="none" stroke="#5FD0FF" stroke-width="5" stroke-linejoin="round"/>')
    for x, y in ptsd:
        s.add(f'<circle cx="{x:.0f}" cy="{y:.0f}" r="10" fill="#0B1220" stroke="#5FD0FF" stroke-width="4"/>')
    # 우측 판정 패널
    rx0 = px + pw + 40
    s.add(f'<rect x="{rx0}" y="{py-24}" width="290" height="{ph+60:.0f}" rx="16" fill="#14243E"/>')
    cy0 = py + 110
    ring = s.rg([(0, '#4ADE80', .35), (1, '#4ADE80', 0)], .5, .5, .5)
    s.add(f'<circle cx="{rx0+145}" cy="{cy0}" r="130" fill="{ring}"/>')
    s.add(f'<circle cx="{rx0+145}" cy="{cy0}" r="76" fill="#1F7A52"/>')
    s.add(f'<circle cx="{rx0+145}" cy="{cy0}" r="76" fill="none" stroke="#4ADE80" stroke-width="7"/>')
    s.add(f'<path d="M{rx0+107},{cy0+4} l28,30 l52,-62" fill="none" stroke="#EAFFF3" '
          f'stroke-width="18" stroke-linecap="round" stroke-linejoin="round"/>')
    for i in range(4):
        yy = cy0 + 150 + i * 52
        s.add(f'<rect x="{rx0+28}" y="{yy}" width="150" height="16" rx="8" fill="#3C5480"/>')
        s.add(f'<rect x="{rx0+196}" y="{yy}" width="66" height="16" rx="8" '
              f'fill="{"#4ADE80" if i != 2 else "#E3C557"}" opacity="0.9"/>')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh*.4}" rx="10" fill="#FFFFFF" opacity="0.04"/>')
    s.vignette(.40)
    return s


# =====================================================================
# 3-3  시각화 대시보드 : 대형 디스플레이 대시보드
# =====================================================================
def scene_3_3():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#101B2E"),(1,"#070C15")])}"/>')
    screen_content_glow(s, 80, 60, 1040, 660, '#2A6BE0', .34)
    s.shadow(600, 762, 320, 24, .5)
    sx, sy, sw, sh = bezel(s, 60, 44, 1080, 690, 20)
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="64" fill="#12203A"/>')
    s.add(f'<rect x="{sx+26}" y="{sy+22}" width="240" height="20" rx="10" fill="#6E93D0"/>')
    for i in range(3):
        s.add(f'<circle cx="{sx+sw-46-i*40}" cy="{sy+32}" r="10" fill="#3C5480"/>')
    # KPI 카드 3장
    cw = (sw - 100) / 3
    for i in range(3):
        x = sx + 34 + i * (cw + 16)
        s.add(f'<rect x="{x:.0f}" y="{sy+90}" width="{cw:.0f}" height="120" rx="14" fill="#16273F"/>')
        s.add(f'<rect x="{x+22:.0f}" y="{sy+112}" width="110" height="14" rx="7" fill="#41608F"/>')
        s.add(f'<rect x="{x+22:.0f}" y="{sy+140}" width="{140-i*22}" height="34" rx="9" '
              f'fill="{["#5FD0FF","#4ADE80","#E3C557"][i]}" opacity="0.92"/>')
    # 도넛
    dcx, dcy, dr = sx + 190, sy + 400, 108
    s.add(f'<circle cx="{dcx}" cy="{dcy}" r="{dr}" fill="none" stroke="#1B2E4A" stroke-width="42"/>')
    s.add(f'<path d="M{dcx},{dcy-dr} a{dr},{dr} 0 1 1 -{dr*0.95:.0f},{dr*1.32:.0f}" fill="none" '
          f'stroke="#2F7DD6" stroke-width="42" stroke-linecap="round"/>')
    s.add(f'<path d="M{dcx-dr*0.95:.0f},{dcy+dr*0.32:.0f} a{dr},{dr} 0 0 1 {dr*0.42:.0f},-{dr*1.28:.0f}" '
          f'fill="none" stroke="#5FD0FF" stroke-width="42" stroke-linecap="round"/>')
    # 막대
    bx, by, bh = sx + 380, sy + 500, 210
    for i, v in enumerate((.42, .68, .55, .86, .62, .95, .74)):
        s.add(f'<rect x="{bx+i*54}" y="{by-bh*v:.0f}" width="34" height="{bh*v:.0f}" rx="8" '
              f'fill="{s.lg([(0,"#5FD0FF"),(1,"#1E5FB0")])}"/>')
    s.add(f'<line x1="{bx-16}" y1="{by}" x2="{bx+380}" y2="{by}" stroke="#3C5480" stroke-width="3"/>')
    # 라인 차트
    lx, ly, lw, lh = sx + 790, sy + 300, 250, 200
    grid(s, lx, ly, lw, lh, 5, 4)
    ptsl = [(0, .22), (.18, .45), (.34, .34), (.52, .68), (.7, .58), (.86, .84), (1, .76)]
    s.add('<path d="M' + ' L'.join(f'{lx+a*lw:.0f},{ly+lh-b*lh:.0f}' for a, b in ptsl) +
          '" fill="none" stroke="#4ADE80" stroke-width="5" stroke-linejoin="round"/>')
    s.add(f'<path d="M{lx},{ly+lh} ' + ' '.join(f'L{lx+a*lw:.0f},{ly+lh-b*lh:.0f}' for a, b in ptsl) +
          f' L{lx+lw},{ly+lh} Z" fill="{s.lg([(0,"#4ADE80",.32),(1,"#4ADE80",.02)])}"/>')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh*.4}" rx="10" fill="#FFFFFF" opacity="0.04"/>')
    s.vignette(.40)
    return s


# =====================================================================
# 4-1  모델 재학습 : GPU 서버 내부 (그래픽카드 + 방열핀 + LED)
# =====================================================================
def scene_4_1():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#131A28"),(1,"#070A11")])}"/>')
    case = s.lg([(0, '#2B3546'), (.5, '#1A2231'), (1, '#0E141E')])
    s.add(f'<rect x="60" y="70" width="1080" height="690" rx="20" fill="{case}"/>')
    s.add(f'<rect x="60" y="70" width="1080" height="160" rx="20" fill="#FFFFFF" opacity="0.05"/>')
    s.add(f'<rect x="96" y="106" width="1008" height="618" rx="12" fill="#0B1119"/>')
    # 메인보드
    s.add(f'<rect x="120" y="130" width="960" height="570" rx="8" fill="#123024" opacity="0.85"/>')
    for i in range(26):
        s.add(f'<line x1="{140+i*36}" y1="140" x2="{140+i*36}" y2="690" stroke="#1D6B45" '
              f'stroke-width="2" opacity="0.28"/>')
    # GPU 카드 4장
    for i in range(4):
        y = 168 + i * 132
        g = s.lg([(0, '#3A4457'), (.35, '#242D3C'), (1, '#141A25')])
        s.add(f'<rect x="150" y="{y}" width="800" height="104" rx="10" fill="{g}"/>')
        s.add(f'<rect x="150" y="{y}" width="800" height="34" rx="10" fill="#FFFFFF" opacity="0.06"/>')
        for k in range(2):                                # 팬
            fx, fy = 300 + k * 300, y + 52
            s.add(f'<circle cx="{fx}" cy="{fy}" r="42" fill="#0D1219"/>')
            s.add(f'<circle cx="{fx}" cy="{fy}" r="42" fill="none" stroke="#39445A" stroke-width="5"/>')
            for b in range(7):
                a0 = b * 2 * math.pi / 7
                s.add(f'<path d="M{fx},{fy} L{fx+38*math.cos(a0):.0f},{fy+38*math.sin(a0):.0f} '
                      f'A38,38 0 0 1 {fx+38*math.cos(a0+0.62):.0f},{fy+38*math.sin(a0+0.62):.0f} Z" '
                      f'fill="#2A3purple"'.replace('#2A3purple"', '#2A3346"') + ' opacity="0.95"/>')
            s.add(f'<circle cx="{fx}" cy="{fy}" r="13" fill="#4C5husk"'.replace('#4C5husk"', '#4C5A72"') + '/>')
        for k in range(14):                               # 방열핀
            s.add(f'<rect x="{620+k*22}" y="{y+16}" width="9" height="72" rx="3" fill="#39445A" opacity="0.9"/>')
        s.add(f'<rect x="168" y="{y+22}" width="14" height="60" rx="6" '
              f'fill="{["#4ADE80","#4ADE80","#5FD0FF","#4ADE80"][i]}"/>')
        s.add(f'<rect x="900" y="{y+38}" width="34" height="28" rx="6" fill="#0F1620"/>')
    glow = s.rg([(0, '#5FD0FF', .16), (1, '#5FD0FF', 0)], .5, .5, .5)
    s.add(f'<rect x="96" y="106" width="1008" height="618" fill="{glow}"/>')
    s.vignette(.46)
    return s


# =====================================================================
# 4-2  파인튜닝 파이프라인 : 코드 에디터 + 학습 진행 터미널
# =====================================================================
def scene_4_2():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#111A2A"),(1,"#070C15")])}"/>')
    screen_content_glow(s, 70, 50, 1060, 700, '#2A6BE0', .28)
    sx, sy, sw, sh = bezel(s, 60, 44, 1080, 700, 20)
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="56" fill="#141C2A"/>')
    for i, c in enumerate(('#E06C75', '#E3C557', '#4ADE80')):
        s.add(f'<circle cx="{sx+34+i*34}" cy="{sy+28}" r="11" fill="{c}"/>')
    for i in range(3):
        s.add(f'<rect x="{sx+180+i*160}" y="{sy+14}" width="140" height="42" rx="8" '
              f'fill="{"#1D2838" if i else "#26344A"}"/>')
        s.add(f'<rect x="{sx+198+i*160}" y="{sy+30}" width="{92-i*16}" height="12" rx="6" fill="#5E7BA8"/>')
    # 코드 영역
    cx0, cy0 = sx + 28, sy + 84
    s.add(f'<rect x="{cx0}" y="{cy0}" width="66" height="{sh-260}" fill="#0B1220"/>')
    for i in range(12):
        s.add(f'<rect x="{cx0+22}" y="{cy0+18+i*28}" width="24" height="11" rx="5" fill="#33415C"/>')
    code_lines(s, cx0 + 96, cy0 + 16, sw - 200, 12, 28, 11)
    s.add(f'<rect x="{cx0+96}" y="{cy0+16+4*28-6}" width="{sw-200}" height="24" fill="#5FA8FF" opacity="0.10"/>')
    # 터미널 (학습 로그 + 진행바)
    ty = sy + sh - 148
    s.add(f'<rect x="{sx+28}" y="{ty}" width="{sw-56}" height="120" rx="10" fill="#08101C"/>')
    s.add(f'<rect x="{sx+28}" y="{ty}" width="{sw-56}" height="34" rx="10" fill="#132033"/>')
    s.add(f'<rect x="{sx+48}" y="{ty+11}" width="120" height="12" rx="6" fill="#4B6AA8"/>')
    for i in range(2):
        s.add(f'<rect x="{sx+48}" y="{ty+50+i*24}" width="{300-i*90}" height="11" rx="5" fill="#4ADE80" opacity="0.8"/>')
        s.add(f'<rect x="{sx+368-i*90}" y="{ty+50+i*24}" width="{160+i*40}" height="11" rx="5" fill="#3C5480"/>')
    pbw = sw - 220
    s.add(f'<rect x="{sx+48}" y="{ty+96}" width="{pbw}" height="14" rx="7" fill="#1B2A42"/>')
    s.add(f'<rect x="{sx+48}" y="{ty+96}" width="{pbw*0.68:.0f}" height="14" rx="7" '
          f'fill="{s.lg([(0,"#5FD0FF"),(1,"#2F7DD6")],0,0,1,0)}"/>')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh*.4}" rx="10" fill="#FFFFFF" opacity="0.04"/>')
    s.vignette(.40)
    return s


# =====================================================================
# 4-3  다음조건·종료 결정 : 대형 화면 앞 연구자 2인 (분기 결정)
# =====================================================================
def scene_4_3():
    s = S()
    s.add(f'<rect width="{W}" height="{H}" fill="{s.lg([(0,"#101728"),(1,"#060A12")])}"/>')
    screen_content_glow(s, 170, 70, 860, 470, '#2A6BE0', .40)
    sx, sy, sw, sh = bezel(s, 160, 60, 880, 470, 16)
    grid(s, sx, sy, sw, sh, 8, 5)
    # 분기 다이어그램
    nx0, ny0 = sx + 180, sy + sh * .52
    s.add(f'<path d="M{nx0-40},{ny0-52} L{nx0+52},{ny0} L{nx0-40},{ny0+52} L{nx0-132},{ny0} Z" fill="#2F7DD6"/>')
    s.add(f'<path d="M{nx0-70},{ny0} l20,22 l44,-48" fill="none" stroke="#FFFFFF" '
          f'stroke-width="12" stroke-linecap="round" stroke-linejoin="round"/>')
    for dy, col in ((-110, '#4ADE80'), (110, '#E3C557')):
        s.add(f'<path d="M{nx0+62},{ny0} C{nx0+150},{ny0} {nx0+170},{ny0+dy} {nx0+270},{ny0+dy}" '
              f'fill="none" stroke="{col}" stroke-width="6" opacity="0.9"/>')
        s.add(f'<path d="M{nx0+300},{ny0+dy} l-30,-16 l0,32 z" fill="{col}"/>')
        s.add(f'<rect x="{nx0+312}" y="{ny0+dy-40}" width="200" height="80" rx="14" fill="#16273F" '
              f'stroke="{col}" stroke-width="4"/>')
        for k in range(2):
            s.add(f'<rect x="{nx0+336}" y="{ny0+dy-18+k*24}" width="{150-k*54}" height="13" rx="6" '
                  f'fill="{col}" opacity="{0.9-k*0.45}"/>')
    s.add(f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh*.4}" rx="8" fill="#FFFFFF" opacity="0.05"/>')
    # 바닥 + 인물 실루엣
    s.add(f'<rect y="{H*.72}" width="{W}" height="{H*.28}" fill="{s.lg([(0,"#141C2C"),(1,"#080C14")])}"/>')
    spill = s.rg([(0, '#4C86D8', .30), (1, '#4C86D8', 0)], .5, .1, .75)
    s.add(f'<rect y="{H*.60}" width="{W}" height="{H*.40}" fill="{spill}"/>')

    def person(cx0_, sc, op):
        """뒷모습 실루엣 — 머리·목·어깨를 겹쳐 하나의 덩어리로 이어지게."""
        base = H + 30
        sh_y = base - 190 * sc                      # 어깨 높이
        s.add(f'<ellipse cx="{cx0_:.0f}" cy="{base:.0f}" rx="{132*sc:.0f}" ry="{210*sc:.0f}" '
              f'fill="#04070E" opacity="{op}"/>')
        s.add(f'<rect x="{cx0_-30*sc:.0f}" y="{sh_y-64*sc:.0f}" width="{60*sc:.0f}" '
              f'height="{90*sc:.0f}" rx="{22*sc:.0f}" fill="#04070E" opacity="{op}"/>')
        s.add(f'<circle cx="{cx0_:.0f}" cy="{sh_y-96*sc:.0f}" r="{54*sc:.0f}" '
              f'fill="#04070E" opacity="{op}"/>')
        # 화면빛이 어깨 윤곽에만 살짝 걸리게
        s.add(f'<path d="M{cx0_-128*sc:.0f},{base-150*sc:.0f} a{132*sc:.0f},{210*sc:.0f} 0 0 1 {54*sc:.0f},-{74*sc:.0f}" '
              f'fill="none" stroke="#6FA3E8" stroke-width="{4.5*sc:.1f}" opacity="0.30"/>')

    person(286, 1.00, .97)
    person(902, 0.88, .97)
    s.vignette(.42)
    return s


SCENES = {'box1_1': scene_1_1, 'box1_2': scene_1_2, 'box1_3': scene_1_3,
          'box2_1': scene_2_1, 'box2_2': scene_2_2, 'box2_3': scene_2_3,
          'box3_1': scene_3_1, 'box3_2': scene_3_2, 'box3_3': scene_3_3,
          'box4_1': scene_4_1, 'box4_2': scene_4_2, 'box4_3': scene_4_3}


def build(name, fn, scale=2):
    svg_path = os.path.join(OUT, name + '.svg')
    png_path = os.path.join(OUT, name + '.png')
    open(svg_path, 'w', encoding='utf-8').write(fn().out())
    cairosvg.svg2png(url=svg_path, write_to=png_path,
                     output_width=W * scale, output_height=H * scale)
    im = Image.open(png_path).convert('RGB')
    im = im.filter(ImageFilter.GaussianBlur(0.4)).resize((W, H), Image.LANCZOS)
    im.save(png_path, quality=95)
    return png_path


if __name__ == '__main__':
    for name, fn in SCENES.items():
        print(build(name, fn), os.path.getsize(os.path.join(OUT, name + '.png')))
