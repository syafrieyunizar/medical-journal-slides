#!/usr/bin/env python3
"""
Reusable helper module for generating anti-AI-slop medical journal PowerPoint (.pptx) decks.
Enforces strict horizontal bands, projection typography floors, and capacity assertions.
"""

import math
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from PIL import Image, ImageOps

SLIDE_W    = 13.333
SLIDE_H    = 7.5
MARGIN_X   = 0.8
CONTENT_W  = 11.733
TITLE_TOP  = 0.50
TITLE_H    = 1.50
BODY_TOP   = 2.20
BODY_H     = 4.60
FOOTER_TOP = 6.95
FOOTER_H   = 0.35
COL_GAP    = 0.45
COL_W_3    = 3.60
COL_W_2    = 5.60

CHAR_W_EM = 0.52
LINE_H    = 1.22
RAGGED    = 0.90

BG_THEME   = RGBColor(0x0F, 0x17, 0x2A)
BG_CONTENT = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_LIGHT = RGBColor(0xF8, 0xFA, 0xFC)
TEXT_DARK  = RGBColor(0x0F, 0x17, 0x2A)
TEXT_MUTED = RGBColor(0x47, 0x55, 0x69)
ACCENT     = RGBColor(0x25, 0x63, 0xEB)
TINT       = RGBColor(0xEF, 0xF6, 0xFF)
ALERT_RED  = RGBColor(0xB9, 0x1C, 0x1C)
GREEN_POS  = RGBColor(0x15, 0x80, 0x3D)

HEAD_FONT  = "Calibri"
BODY_FONT  = "Calibri"


def create_medical_presentation() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def chars_per_line(w_in: float, pt: float) -> int:
    return max(1, int(w_in * 72 / (pt * CHAR_W_EM) * RAGGED))


def lines_needed(text: str, w_in: float, pt: float) -> int:
    return math.ceil(len(text) / chars_per_line(w_in, pt)) if text else 1


def set_slide_canvas(prs: Presentation, slide, is_bookend: bool = False):
    bg = slide.shapes.add_shape(
        pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_THEME if is_bookend else BG_CONTENT
    bg.line.fill.background()
    bg.shadow.inherit = False
    return bg


def add_header(slide, title_text: str, kicker_text: str = "CLINICAL STUDY", is_dark: bool = False):
    n = lines_needed(title_text, CONTENT_W, 32)
    assert n <= 2, f"Action title wraps to {n} lines (max 2) -- shorten: {title_text!r}"
    tb = slide.shapes.add_textbox(Inches(MARGIN_X), Inches(TITLE_TOP), Inches(CONTENT_W), Inches(TITLE_H))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p_k = tf.paragraphs[0]
    p_k.alignment = PP_ALIGN.LEFT
    p_k.space_after = Pt(4)
    rk = p_k.add_run()
    rk.text = kicker_text.upper()
    rk.font.size = Pt(13)
    rk.font.name = BODY_FONT
    rk.font.bold = True
    rk.font.color.rgb = ACCENT if not is_dark else RGBColor(0x93, 0xC5, 0xFD)

    p_t = tf.add_paragraph()
    p_t.alignment = PP_ALIGN.LEFT
    rt = p_t.add_run()
    rt.text = title_text
    rt.font.size = Pt(32)
    rt.font.name = HEAD_FONT
    rt.font.bold = True
    rt.font.color.rgb = TEXT_LIGHT if is_dark else TEXT_DARK


def add_footer(slide, citation_text: str, slide_num: int, is_dark: bool = False):
    tb = slide.shapes.add_textbox(Inches(MARGIN_X), Inches(FOOTER_TOP), Inches(CONTENT_W), Inches(FOOTER_H))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = f"{citation_text}   |   Slide {slide_num}"
    r.font.size = Pt(12)
    r.font.name = BODY_FONT
    r.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8) if is_dark else TEXT_MUTED
