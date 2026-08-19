# -*- coding: utf-8 -*-
"""일러스트 버전 빌드 : 1~4번 박스에 직접 제작한 벡터 일러스트 밴드를 삽입."""
import os, sys
from pptx import Presentation
from pptx.util import Emu
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from sdl_layout import (apply_geometry, emu, BAND_X, BAND_TOP, IMG_DY, IMG_H,
                        IMG_W_EMU, ALT)

SRC = os.path.join(ROOT, 'assets', '원본_v4_new_mod.pptx')
OUT = os.path.join(ROOT, '자율실험실_SDL_도식_v5_예시이미지.pptx')
IMG = os.path.join(HERE, 'example_images')


def main():
    prs = Presentation(SRC)
    slide = prs.slides[0]
    apply_geometry(slide)

    for n in (1, 2, 3, 4):
        pic = slide.shapes.add_picture(
            os.path.join(IMG, f'box{n}.png'),
            Emu(BAND_X[n]), Emu(emu(BAND_TOP[n] + IMG_DY)),
            Emu(IMG_W_EMU), Emu(emu(IMG_H)))
        pic.name = f'예시 이미지 {n}'
        pic._element.find('.//' + qn('p:cNvPr')).set('descr', ALT[n])

    prs.save(OUT)
    print('saved', os.path.basename(OUT))


if __name__ == '__main__':
    main()
