#!/usr/bin/env python3
"""
Automated geometry and anti-AI-slop validator for medical PowerPoint (.pptx) presentations.
Checks word count ceilings, font size floors, bookend cleanliness, and horizontal band limits.
"""

import sys
from pptx import Presentation

MAX_WORDS_PER_SLIDE = 65
MAX_WORDS_BOOKEND   = 38
MIN_FONT_PT         = 10.5


def validate_deck(pptx_path: str) -> int:
    prs = Presentation(pptx_path)
    errors = []
    warnings = []
    total_slides = len(prs.slides)

    for idx, slide in enumerate(prs.slides, start=1):
        is_bookend = idx in (1, total_slides)
        words = 0

        for shape in slide.shapes:
            top_in = shape.top / 914400 if shape.top is not None else 0
            h_in   = shape.height / 914400 if shape.height is not None else 0
            if h_in < 7.4 and (top_in + h_in) > 7.42:
                errors.append(f"Slide {idx}: Shape {shape.name} overflows bottom edge (bottom={top_in + h_in:.2f}\")")

            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    text = p.text.strip()
                    if text:
                        words += len(text.split())
                    for run in p.runs:
                        if run.font.size is not None:
                            pt = run.font.size.pt
                            if pt < MIN_FONT_PT:
                                errors.append(
                                    f"Slide {idx}: Font size {pt:.1f}pt is below {MIN_FONT_PT}pt floor ({run.text[:25]!r})"
                                )

            if is_bookend and shape.has_table:
                errors.append(f"Slide {idx}: Bookend slide (Cover/Outro) must not contain tables.")

        limit = MAX_WORDS_BOOKEND if is_bookend else MAX_WORDS_PER_SLIDE
        if words > limit:
            errors.append(f"Slide {idx}: Word count {words} exceeds ceiling of {limit} words.")
        elif words > 52 and not is_bookend:
            warnings.append(f"Slide {idx}: High word density ({words} words; target 35-50).")

    print(f"=== Validation Report for {pptx_path} ({total_slides} slides) ===")
    for w in warnings:
        print(f"  [WARN]  {w}")
    for e in errors:
        print(f"  [ERROR] {e}")

    if not errors:
        print("  [PASS] Zero layout or anti-AI-slop errors detected.")
        return 0
    return 1


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 validate_medical_pptx.py <path_to_deck.pptx>")
        sys.exit(2)
    sys.exit(validate_deck(sys.argv[1]))
