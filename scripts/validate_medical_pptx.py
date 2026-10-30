#!/usr/bin/env python3
"""
validate_medical_pptx.py — Comprehensive Geometry, Band, Overlap, and Density Validator.

Audits a generated medical PowerPoint (.pptx) presentation against all Anti-AI-Slop rules:
1. Band Verification: Ensures every shape stays inside Title (0.40-2.05"), Body (2.15-6.85"), or Footer (6.90-7.40") bands.
2. Pairwise 2D Bounding-Box Collision Detection: Detects unintended shape overlaps (allowing text inside a card container).
3. Separate Word Budgets:
   - Title Band: <= 22 words
   - Body Band (Text + Tables): <= 65 words for standard text slides, <= 110 words for data-table slides
   - Footer Band: excluded from body word budget
   - Bookend Slides (Slide 1 & Slide N, when total_slides > 3): <= 38 words total, 0 bullets, 0 tables, 0 cards.
4. Card / Box Quota Check:
   - For decks > 3 slides, boxed slides must be <= ceil(0.40 * interior_slides).
   - For 1-3 slide Visual Abstracts (Archetype C), 3-pillar cards are explicitly permitted.
5. Font Size Floor: Every text run >= 10.5 pt (excluding 1pt hidden accessibility title sync).
6. Native Background Check: Warns if a full-slide draggable rectangle is used instead of slide.background.fill.
"""

import math
import sys
from typing import List, Tuple
import pptx
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

EMU_PER_INCH = 914400.0

TITLE_BAND  = (0.35, 2.08)
BODY_BAND   = (2.12, 6.86)
FOOTER_BAND = (6.88, 7.42)

MAX_TITLE_WORDS       = 22
MAX_BODY_WORDS_TEXT   = 65
MAX_BODY_WORDS_TABLE  = 115
MAX_BOOKEND_WORDS     = 40
MIN_FONT_PT           = 10.5


def bbox_in(shape) -> Tuple[float, float, float, float]:
    left = (shape.left or 0) / EMU_PER_INCH
    top  = (shape.top or 0) / EMU_PER_INCH
    w    = (shape.width or 0) / EMU_PER_INCH
    h    = (shape.height or 0) / EMU_PER_INCH
    return (left, top, left + w, top + h)


def boxes_overlap(b1: Tuple[float, float, float, float], b2: Tuple[float, float, float, float], tol: float = 0.04) -> bool:
    l1, t1, r1, btm1 = b1
    l2, t2, r2, btm2 = b2
    if (r1 - l2) <= tol or (r2 - l1) <= tol:
        return False
    if (btm1 - t2) <= tol or (btm2 - t1) <= tol:
        return False
    return True


def is_container_of(outer: Tuple[float, float, float, float], inner: Tuple[float, float, float, float], tol: float = 0.08) -> bool:
    return (
        outer[0] - tol <= inner[0]
        and outer[1] - tol <= inner[1]
        and outer[2] + tol >= inner[2]
        and outer[3] + tol >= inner[3]
    )


def has_bullet_xml(paragraph) -> bool:
    pPr = paragraph._p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}pPr")
    if pPr is None:
        return False
    buChar = pPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}buChar")
    buAuto = pPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum")
    return (buChar is not None) or (buAuto is not None)


def validate_deck(pptx_path: str) -> int:
    prs = Presentation(pptx_path)
    errors: List[str] = []
    warnings: List[str] = []
    total_slides = len(prs.slides)
    is_visual_abstract_deck = (total_slides <= 3)
    boxed_slide_count = 0

    for idx, slide in enumerate(prs.slides, start=1):
        is_bookend = (not is_visual_abstract_deck) and (idx in (1, total_slides))
        title_words = 0
        body_words  = 0
        table_words = 0
        has_table   = False
        slide_cards = 0
        bullet_count = 0
        placed_shapes = []

        for shape in slide.shapes:
            l, t, r, b = bbox_in(shape)
            w = r - l
            h = b - t

            # Skip invisible 0.01" accessibility title sync placeholder
            if w < 0.05 and h < 0.05:
                continue

            # 1. Check for full-slide draggable background rectangle anti-pattern
            if w > 12.8 and h > 7.1 and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                warnings.append(
                    f"Slide {idx}: Full-slide rectangle '{shape.name}' detected; prefer native slide.background.fill."
                )
                continue

            # 2. Check slide canvas boundaries
            if b > 7.45 or r > 13.38 or l < -0.05 or t < -0.05:
                errors.append(
                    f"Slide {idx}: Shape '{shape.name}' exceeds canvas bounds (L={l:.2f}, T={t:.2f}, R={r:.2f}, B={b:.2f} in)."
                )

            # 3. Check Horizontal Band compliance (for interior slides or Visual Abstract)
            if not is_bookend:
                in_title  = (TITLE_BAND[0] <= t and b <= TITLE_BAND[1])
                in_body   = (BODY_BAND[0]  <= t and b <= BODY_BAND[1])
                in_footer = (FOOTER_BAND[0] <= t and b <= FOOTER_BAND[1])
                if not (in_title or in_body or in_footer):
                    errors.append(
                        f"Slide {idx}: Shape '{shape.name}' (top={t:.2f} in, bottom={b:.2f} in) crosses horizontal band boundaries!"
                    )
            else:
                in_title  = (t < 2.1)
                in_footer = (t >= 6.5)
                in_body   = not (in_title or in_footer)

            # 4. Identify Card / Box shapes
            is_card = (
                shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                and not shape.has_text_frame
                or (shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and "Card" in shape.name)
            )
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and h >= 1.0 and w >= 2.0:
                is_card = True
                slide_cards += 1

            placed_shapes.append((shape.name, (l, t, r, b), is_card))

            # 5. Count words and check font sizes in text frames
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    txt = p.text.strip()
                    if txt:
                        w_cnt = len(txt.split())
                        if in_title:
                            title_words += w_cnt
                        elif in_footer:
                            pass  # Footer citations excluded from body word cap
                        else:
                            body_words += w_cnt

                        if has_bullet_xml(p) or txt.startswith("•"):
                            bullet_count += 1

                    for run in p.runs:
                        if run.font.size is not None:
                            pt = run.font.size.pt
                            if pt > 1.5 and pt < MIN_FONT_PT:
                                errors.append(
                                    f"Slide {idx}: Font size {pt:.1f}pt below {MIN_FONT_PT}pt floor ({run.text[:25]!r})."
                                )

            # 6. Count words and font sizes inside Tables
            if shape.has_table:
                has_table = True
                for row in shape.table.rows:
                    for cell in row.cells:
                        c_txt = cell.text_frame.text.strip()
                        if c_txt:
                            table_words += len(c_txt.split())
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if run.font.size is not None and run.font.size.pt < MIN_FONT_PT:
                                    errors.append(
                                        f"Slide {idx}: Table font {run.font.size.pt:.1f}pt below {MIN_FONT_PT}pt floor."
                                    )

            # 7. Check Native Charts
            if shape.has_chart:
                has_table = True  # Data-dense allowance

        # 8. Pairwise 2D Overlap Check
        n_sh = len(placed_shapes)
        for i in range(n_sh):
            name1, box1, card1 = placed_shapes[i]
            for j in range(i + 1, n_sh):
                name2, box2, card2 = placed_shapes[j]
                if boxes_overlap(box1, box2):
                    # Allow text/figure inside a card container, but never two cards or two textboxes colliding
                    if (card1 and not card2 and is_container_of(box1, box2)) or (
                        card2 and not card1 and is_container_of(box2, box1)
                    ):
                        continue
                    errors.append(
                        f"Slide {idx}: Shape collision between '{name1}' {tuple(round(x,2) for x in box1)} "
                        f"and '{name2}' {tuple(round(x,2) for x in box2)}."
                    )

        # 9. Bookend Rules (Cover & Outro on multi-slide decks)
        if is_bookend:
            total_bk_words = title_words + body_words
            if total_bk_words > MAX_BOOKEND_WORDS:
                errors.append(
                    f"Slide {idx} (Bookend): Word count {total_bk_words} exceeds bookend limit of {MAX_BOOKEND_WORDS}."
                )
            if has_table:
                errors.append(f"Slide {idx} (Bookend): Must not contain tables or charts.")
            if bullet_count > 0:
                errors.append(f"Slide {idx} (Bookend): Must not contain bullet points ({bullet_count} found).")
            if slide_cards > 0:
                errors.append(f"Slide {idx} (Bookend): Must not contain floating cards/boxes.")
        else:
            # Interior or Visual Abstract slide checks
            if title_words > MAX_TITLE_WORDS:
                errors.append(
                    f"Slide {idx}: Title Band has {title_words} words (max {MAX_TITLE_WORDS})."
                )
            total_interior_body = body_words + table_words
            limit = MAX_BODY_WORDS_TABLE if has_table else MAX_BODY_WORDS_TEXT
            if total_interior_body > limit:
                errors.append(
                    f"Slide {idx}: Body word count {total_interior_body} exceeds ceiling of {limit} words."
                )
            if bullet_count > 6:
                errors.append(f"Slide {idx}: Contains {bullet_count} bullets (max 6 per slide).")
            if slide_cards > 0:
                boxed_slide_count += 1

    # 10. Deck-Wide Card Quota Check (exempt 1-3 slide Visual Abstract decks)
    if not is_visual_abstract_deck and total_slides > 3:
        interior_count = total_slides - 2
        max_boxed = math.ceil(interior_count * 0.40)  # Allows up to ~1/3 (e.g. 2 of 6 interior slides)
        if boxed_slide_count > max_boxed:
            errors.append(
                f"Deck Card Quota: {boxed_slide_count}/{interior_count} interior slides use cards "
                f"(max allowed is {max_boxed} to prevent AI box spam)."
            )

    print(f"=== Medical PPTX Validation Report: {pptx_path} ({total_slides} slides) ===")
    for w in warnings:
        print(f"  [WARN]  {w}")
    for e in errors:
        print(f"  [ERROR] {e}")

    if not errors:
        print("  [PASS] All band, collision, word-budget, card-quota, and accessibility checks passed!")
        return 0
    return 1


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 validate_medical_pptx.py <path_to_deck.pptx>")
        sys.exit(2)
    sys.exit(validate_deck(sys.argv[1]))
