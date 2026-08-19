# -*- coding: utf-8 -*-
"""자율실험 폐루프 1~4번 박스용 예시 이미지(플랫 벡터 일러스트) 생성.
   전량 직접 제작한 원본 SVG -> PNG (외부 저작물/스톡 이미지 미사용)."""
import math, os, cairosvg

W, H = 1270, 300                 # 2.6249in x 0.62in (4.234:1)
NAVY, NAVY2, BLUE = "#1D3567", "#21407C", "#1B54B5"
LBLUE, PALE, BG, BORDER = "#C7D4EC", "#E8EEF8", "#F4F6FA", "#DDE0E6"

CW = W / 3.0
def cx(i): return CW * (i + 0.5)
CY = 152

def frame():
    s = f'<rect x="1.5" y="1.5" width="{W-3}" height="{H-3}" rx="16" fill="{BG}" stroke="{BORDER}" stroke-width="3"/>'
    for i in (1, 2):
        s += f'<line x1="{CW*i:.1f}" y1="46" x2="{CW*i:.1f}" y2="254" stroke="{BORDER}" stroke-width="2.5"/>'
    return s

def svg(body):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
            f'viewBox="0 0 {W} {H}">{frame()}{body}</svg>')

# ---------------- 공통 부품 ----------------
def neural(ox, oy, sc=1.0):
    layers = [[-1, 0, 1], [-1.5, -0.5, 0.5, 1.5], [-1, 0, 1]]
    gx, gy, r = 66 * sc, 50 * sc, 13 * sc
    pts = [[(ox + (li - 1) * gx, oy + p * gy) for p in lay] for li, lay in enumerate(layers)]
    s = ""
    for a, b_ in zip(pts, pts[1:]):
        for p in a:
            for q in b_:
                s += (f'<line x1="{p[0]:.1f}" y1="{p[1]:.1f}" x2="{q[0]:.1f}" y2="{q[1]:.1f}" '
                      f'stroke="{LBLUE}" stroke-width="{2.8*sc:.1f}"/>')
    for li, lay in enumerate(pts):
        c = BLUE if li == 1 else NAVY
        for p in lay:
            s += f'<circle cx="{p[0]:.1f}" cy="{p[1]:.1f}" r="{r:.1f}" fill="{c}"/>'
    return s

def arrow(x1, y1, x2, y2, color=BLUE, sw=8, head=19):
    ang = math.atan2(y2 - y1, x2 - x1)
    bx, by = x2 - head * 0.85 * math.cos(ang), y2 - head * 0.85 * math.sin(ang)
    px, py = -math.sin(ang), math.cos(ang)
    p1 = (bx + px * head * 0.62, by + py * head * 0.62)
    p2 = (bx - px * head * 0.62, by - py * head * 0.62)
    return (f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{bx:.1f}" y2="{by:.1f}" stroke="{color}" '
            f'stroke-width="{sw}" stroke-linecap="round"/>'
            f'<path d="M{x2:.1f},{y2:.1f} L{p1[0]:.1f},{p1[1]:.1f} L{p2[0]:.1f},{p2[1]:.1f} z" fill="{color}"/>')

def arc_arrow(ox, oy, r, a0, a1, color=BLUE, sw=10, head=24):
    """a0 -> a1 (도, 시계방향 양수) 원호 + 끝단 화살촉"""
    ra0, ra1 = math.radians(a0), math.radians(a1)
    x0, y0 = ox + r * math.cos(ra0), oy + r * math.sin(ra0)
    x1, y1 = ox + r * math.cos(ra1), oy + r * math.sin(ra1)
    large = 1 if abs(a1 - a0) > 180 else 0
    sweep = 1 if a1 > a0 else 0
    s = (f'<path d="M{x0:.1f},{y0:.1f} A{r},{r} 0 {large} {sweep} {x1:.1f},{y1:.1f}" fill="none" '
         f'stroke="{color}" stroke-width="{sw}" stroke-linecap="round"/>')
    tang = ra1 + (math.pi / 2 if sweep else -math.pi / 2)
    tipx, tipy = x1 + head * 0.75 * math.cos(tang), y1 + head * 0.75 * math.sin(tang)
    px, py = -math.sin(tang), math.cos(tang)
    s += (f'<path d="M{tipx:.1f},{tipy:.1f} L{x1+px*head*0.6:.1f},{y1+py*head*0.6:.1f} '
          f'L{x1-px*head*0.6:.1f},{y1-py*head*0.6:.1f} z" fill="{color}"/>')
    return s

def gear(ox, oy, r=34, teeth=9, color=NAVY2, hole=BG):
    pts, step = [], math.pi / teeth
    for i in range(teeth):
        a = 2 * step * i
        for k, rr in ((0.0, r), (step * 0.42, r), (step * 0.58, r * 0.76), (step * 1.42, r * 0.76),
                      (step * 1.58, r)):
            pass
        pts += [(ox + r * math.cos(a - step * 0.30), oy + r * math.sin(a - step * 0.30)),
                (ox + r * math.cos(a + step * 0.30), oy + r * math.sin(a + step * 0.30)),
                (ox + r * 0.74 * math.cos(a + step * 0.62), oy + r * 0.74 * math.sin(a + step * 0.62)),
                (ox + r * 0.74 * math.cos(a + step * 1.38), oy + r * 0.74 * math.sin(a + step * 1.38))]
    d = " ".join(f"{p[0]:.1f},{p[1]:.1f}" for p in pts)
    return (f'<polygon points="{d}" fill="{color}" stroke="{color}" stroke-width="4" stroke-linejoin="round"/>'
            f'<circle cx="{ox}" cy="{oy}" r="{r*0.30:.1f}" fill="{hole}"/>')

# =========================================================
# 1. AI 실험 설계 : AI 추천엔진 / 과거 실험 DB / 가상실험 모델
# =========================================================
b = neural(cx(0), CY, 1.05)

ox, oy = cx(1), CY                                    # 과거 실험 DB
w, h, ry = 128, 136, 23
b += (f'<path d="M{ox-w/2},{oy-h/2} v{h} a{w/2},{ry} 0 0 0 {w},0 v-{h} z" fill="#FFFFFF" '
      f'stroke="{NAVY}" stroke-width="8" stroke-linejoin="round"/>')
b += f'<ellipse cx="{ox}" cy="{oy-h/2}" rx="{w/2}" ry="{ry}" fill="{NAVY}"/>'
for dy in (h * 0.36, h * 0.70):
    b += (f'<path d="M{ox-w/2},{oy-h/2+dy} a{w/2},{ry} 0 0 0 {w},0" fill="none" stroke="{BLUE}" '
          f'stroke-width="7" stroke-linecap="round"/>')

ox = cx(2)                                            # 가상실험 모델
b += (f'<rect x="{ox-128}" y="{oy-104}" width="256" height="208" rx="16" fill="#FFFFFF" '
      f'stroke="{NAVY}" stroke-width="7" stroke-dasharray="18 12"/>')
atoms = [(ox - 62, oy - 26), (ox - 4, oy - 66), (ox + 2, oy + 22), (ox + 68, oy - 20)]
for a_, b_ in ((0, 1), (1, 2), (2, 3), (0, 2), (1, 3)):
    b += (f'<line x1="{atoms[a_][0]}" y1="{atoms[a_][1]}" x2="{atoms[b_][0]}" y2="{atoms[b_][1]}" '
          f'stroke="{NAVY2}" stroke-width="8"/>')
for i, (ax, ay) in enumerate(atoms):
    b += f'<circle cx="{ax}" cy="{ay}" r="{23 if i % 2 == 0 else 17}" fill="{BLUE if i % 2 == 0 else NAVY}"/>'
b += (f'<path d="M{ox-96},{oy+76} q34,-40 66,-8 t58,-30" fill="none" stroke="{LBLUE}" '
      f'stroke-width="10" stroke-linecap="round"/>')
open('assets/example_images/box1.svg', 'w').write(svg(b))

# =========================================================
# 2. 로봇·장비 실험 수행 : 로봇팔 / 액체 핸들러 / 측정장비
# =========================================================
b = ""
ox, oy = cx(0), CY                                    # 로봇팔
sh, el, wr = (ox - 18, oy + 26), (ox - 68, oy - 44), (ox + 12, oy - 70)
b += f'<rect x="{sh[0]-66}" y="{oy+82}" width="132" height="26" rx="11" fill="{NAVY}"/>'
b += f'<path d="M{sh[0]-40},{oy+82} L{sh[0]-22},{sh[1]} L{sh[0]+22},{sh[1]} L{sh[0]+40},{oy+82} z" fill="{NAVY}"/>'
b += (f'<path d="M{sh[0]},{sh[1]} L{el[0]},{el[1]} L{wr[0]},{wr[1]}" fill="none" stroke="{NAVY2}" '
      f'stroke-width="19" stroke-linecap="round" stroke-linejoin="round"/>')
b += f'<circle cx="{sh[0]}" cy="{sh[1]}" r="14" fill="{BLUE}"/><circle cx="{el[0]}" cy="{el[1]}" r="14" fill="{BLUE}"/>'
b += f'<rect x="{wr[0]-10}" y="{wr[1]-26}" width="22" height="52" rx="8" fill="{NAVY}"/>'   # 손목
b += f'<rect x="{wr[0]+10}" y="{wr[1]-30}" width="44" height="13" rx="6" fill="{NAVY}"/>'   # 그리퍼 상
b += f'<rect x="{wr[0]+10}" y="{wr[1]+17}" width="44" height="13" rx="6" fill="{NAVY}"/>'   # 그리퍼 하
b += (f'<rect x="{wr[0]+24}" y="{wr[1]-22}" width="34" height="44" rx="9" fill="{LBLUE}" '
      f'stroke="{NAVY}" stroke-width="6"/>')                                                # 바이알

ox = cx(1)                                            # 액체 핸들러
b += f'<rect x="{ox-92}" y="{oy-108}" width="184" height="42" rx="12" fill="{NAVY}"/>'
for k in (-57, -19, 19, 57):
    b += f'<rect x="{ox+k-6}" y="{oy-66}" width="12" height="34" fill="{NAVY2}"/>'
    b += f'<path d="M{ox+k-11},{oy-32} L{ox+k+11},{oy-32} L{ox+k},{oy-12} z" fill="{NAVY2}"/>'
    b += f'<circle cx="{ox+k}" cy="{oy+8}" r="9" fill="{BLUE}"/>'
b += f'<rect x="{ox-104}" y="{oy+30}" width="208" height="78" rx="11" fill="#FFFFFF" stroke="{NAVY}" stroke-width="8"/>'
for r_ in range(2):
    for c_ in range(6):
        b += f'<circle cx="{ox-78+c_*31.2:.1f}" cy="{oy+54+r_*30}" r="9.5" fill="{LBLUE}"/>'

ox = cx(2)                                            # 측정장비
b += f'<rect x="{ox-110}" y="{oy-96}" width="220" height="152" rx="14" fill="{PALE}" stroke="{NAVY}" stroke-width="8"/>'
b += f'<rect x="{ox-84}" y="{oy-70}" width="128" height="100" rx="9" fill="#FFFFFF" stroke="{NAVY2}" stroke-width="6"/>'
b += (f'<path d="M{ox-72},{oy+8} l22,0 l14,-50 l15,50 l13,-26 l16,0" fill="none" stroke="{BLUE}" '
      f'stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>')
b += f'<circle cx="{ox+74}" cy="{oy-54}" r="13" fill="{BLUE}"/>'
b += f'<rect x="{ox+59}" y="{oy-22}" width="30" height="46" rx="7" fill="{LBLUE}" stroke="{NAVY2}" stroke-width="5"/>'
b += f'<rect x="{ox-110}" y="{oy+66}" width="220" height="24" rx="10" fill="{NAVY}"/>'
open('assets/example_images/box2.svg', 'w').write(svg(b))

# =========================================================
# 3. AI·분석SW의 결과분석 : 측정데이터 분석 / 자동판정 / 시각화 대시보드
# =========================================================
b = ""
ox, oy = cx(0), CY                                    # 측정데이터 분석
b += f'<rect x="{ox-114}" y="{oy-100}" width="228" height="200" rx="14" fill="#FFFFFF" stroke="{NAVY}" stroke-width="8"/>'
b += f'<path d="M{ox-82},{oy-70} v140 h158" fill="none" stroke="{LBLUE}" stroke-width="7" stroke-linecap="round"/>'
b += (f'<path d="M{ox-76},{oy+56} l32,-18 l26,-70 l24,78 l28,-98 l34,56" fill="none" stroke="{BLUE}" '
      f'stroke-width="9" stroke-linecap="round" stroke-linejoin="round"/>')
for px, py in ((ox-44, oy+38), (ox-18, oy-32), (ox+34, oy-52)):
    b += f'<circle cx="{px}" cy="{py}" r="10" fill="{NAVY}"/>'

ox = cx(1)                                            # 목표 충족여부 자동판정
base = oy + 78
b += f'<line x1="{ox-104}" y1="{base}" x2="{ox+104}" y2="{base}" stroke="{NAVY}" stroke-width="8" stroke-linecap="round"/>'
for i, hgt in enumerate((52, 88, 124)):
    b += f'<rect x="{ox-98+i*54}" y="{base-hgt}" width="38" height="{hgt}" rx="7" fill="{LBLUE if i < 2 else BLUE}"/>'
b += (f'<line x1="{ox-108}" y1="{base-112}" x2="{ox+108}" y2="{base-112}" stroke="{NAVY2}" '
      f'stroke-width="7" stroke-dasharray="18 13" stroke-linecap="round"/>')
b += f'<circle cx="{ox+70}" cy="{oy-72}" r="44" fill="{NAVY}"/>'
b += (f'<path d="M{ox+48},{oy-72} l16,18 l30,-38" fill="none" stroke="#FFFFFF" stroke-width="12" '
      f'stroke-linecap="round" stroke-linejoin="round"/>')

ox = cx(2)                                            # 결과 시각화 대시보드
b += f'<rect x="{ox-118}" y="{oy-100}" width="236" height="200" rx="14" fill="#FFFFFF" stroke="{NAVY}" stroke-width="8"/>'
b += f'<path d="M{ox-118},{oy-86} a14,14 0 0 1 14,-14 h208 a14,14 0 0 1 14,14 v16 h-236 z" fill="{NAVY}"/>'
b += f'<circle cx="{ox-42}" cy="{oy+18}" r="48" fill="none" stroke="{PALE}" stroke-width="20"/>'
b += (f'<path d="M{ox-42},{oy-30} a48,48 0 1 1 -40,74" fill="none" stroke="{BLUE}" '
      f'stroke-width="20" stroke-linecap="round"/>')
for i, hgt in enumerate((44, 74, 104)):
    b += f'<rect x="{ox+26+i*34}" y="{oy+68-hgt}" width="24" height="{hgt}" rx="7" fill="{NAVY2}"/>'
open('assets/example_images/box3.svg', 'w').write(svg(b))

# =========================================================
# 4. AI 모델 갱신 : 모델 재학습 / 파인튜닝 파이프라인 / 다음 실험조건·종료 결정
# =========================================================
b = ""
ox, oy = cx(0), CY                                    # 모델 재학습
b += neural(ox, oy, 0.80)
b += arc_arrow(ox, oy, 112, 118, -150, BLUE, 11, 30)

ox = cx(1)                                            # 파인튜닝 파이프라인
for i in range(3):
    bx = ox - 118 + i * 92
    b += f'<rect x="{bx}" y="{oy+14}" width="62" height="62" rx="12" fill="{NAVY if i < 2 else BLUE}"/>'
    if i < 2:
        b += arrow(bx + 68, oy + 45, bx + 88, oy + 45, NAVY2, 7, 15)
b += gear(ox - 52, oy - 56, 40, 9, NAVY2)
b += gear(ox + 26, oy - 26, 28, 9, BLUE)

ox, dy = cx(2), -14                                   # 다음 실험조건 또는 종료 결정
b += f'<path d="M{ox-30},{oy-92+dy} L{ox+40},{oy-22+dy} L{ox-30},{oy+48+dy} L{ox-100},{oy-22+dy} z" fill="{NAVY}"/>'
b += f'<path d="M{ox-51},{oy-22+dy} l14,16 l30,-34" fill="none" stroke="#FFFFFF" stroke-width="10" stroke-linecap="round" stroke-linejoin="round"/>'
b += arrow(ox + 44, oy - 28 + dy, ox + 102, oy - 72 + dy, BLUE, 9, 21)      # 다음 실험조건
b += arrow(ox - 30, oy + 54 + dy, ox - 30, oy + 100 + dy, NAVY2, 9, 21)     # 종료
b += f'<rect x="{ox+66}" y="{oy-100+dy}" width="42" height="30" rx="8" fill="{LBLUE}" stroke="{NAVY}" stroke-width="6"/>'
b += f'<rect x="{ox-56}" y="{oy+100+dy}" width="52" height="18" rx="8" fill="{NAVY2}"/>'
open('assets/example_images/box4.svg', 'w').write(svg(b))

for i in (1, 2, 3, 4):
    cairosvg.svg2png(url=f'assets/example_images/box{i}.svg', write_to=f'assets/example_images/box{i}.png',
                     output_width=1905, output_height=450)
    print(f'assets/example_images/box{i}.png', os.path.getsize(f'assets/example_images/box{i}.png'))
