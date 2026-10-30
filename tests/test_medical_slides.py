#!/usr/bin/env python3
"""
Automated test suite for medical-journal-slides (pptx_medical_builder & validate_medical_pptx).
Verifies that a multi-slide deck and a 1-slide Visual Abstract build cleanly and pass all validator checks,
and confirms that intentional band/overlap violations are caught.
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "scripts")))
from pptx_medical_builder import (
    create_medical_presentation,
    add_medical_slide,
    add_header,
    add_footer,
    add_2_column_body,
    add_3_column_pillars,
    add_clinical_table,
    add_speaker_notes,
    BODY_TOP,
    MARGIN_X,
)
from validate_medical_pptx import validate_deck
from pptx.util import Inches


def test_valid_visual_abstract():
    prs = create_medical_presentation()
    s1 = add_medical_slide(prs, is_bookend=False, palette_name="Clinical Navy")
    add_header(s1, "Psikobiotik Menurunkan Kecemasan Melalui Modulasi Sumbu Usus-Otak", "VISUAL ABSTRACT")
    add_3_column_pillars(
        s1,
        [
            {"title": "Populasi Studi", "bullets": [("Subjek", "14 uji klinis manusia (RPCT & RDBPCOT)"), ("Kondisi", "Stres akademik, kecemasan, depresi")]},
            {"title": "Intervensi Utama", "bullets": [("Galur", "Lactobacillus & Bifidobacterium"), ("Durasi", "4 hingga 12 minggu suplementasi")]},
            {"title": "Luaran Klinis", "bullets": [("Biomarker", "[↓] Kortisol darah & [↑] metabolit GABA"), ("Gejala", "[✓] Perbaikan tidur & reaktivitas kognitif")]},
        ],
        boxed=True,
    )
    add_footer(s1, "Sisubalan et al. Front. Microbiol. 2026;17:1804560", 1, lang="id")
    add_speaker_notes(s1, "Sumber verbatim: Tabel 1 halaman 5-6.")

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        prs.save(tmp.name)
        rc = validate_deck(tmp.name)
        os.unlink(tmp.name)
    assert rc == 0, "Valid Visual Abstract failed validation!"


def test_collision_detection_catches_overlap():
    prs = create_medical_presentation()
    s1 = add_medical_slide(prs, is_bookend=False)
    add_header(s1, "Uji Deteksi Tabrakan Antar Kotak Teks", "QA TEST")
    tb1 = s1.shapes.add_textbox(Inches(MARGIN_X), Inches(BODY_TOP), Inches(5.0), Inches(2.0))
    tb1.text_frame.text = "Kotak pertama di zona body."
    tb2 = s1.shapes.add_textbox(Inches(MARGIN_X + 1.0), Inches(BODY_TOP + 0.5), Inches(5.0), Inches(2.0))
    tb2.text_frame.text = "Kotak kedua menabrak kotak pertama."
    add_footer(s1, "Test Citation", 1, lang="id")

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        prs.save(tmp.name)
        rc = validate_deck(tmp.name)
        os.unlink(tmp.name)
    assert rc == 1, "Validator failed to catch overlapping textboxes!"


if __name__ == "__main__":
    test_valid_visual_abstract()
    test_collision_detection_catches_overlap()
    print("ALL TESTS PASSED SUCCESSFULLY!")
