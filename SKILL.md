---
name: medical-journal-slides
description: Converts medical journal articles and clinical research PDFs into professional, editable PowerPoint (PPTX) presentations across multiple archetypes (Clinical Trials, Journal Club Critical Appraisal, Visual Abstracts, and Systematic Reviews) with strict anti-AI-slop layout guardrails and natural Indonesian language humanization (via adhiprib/humanizer-indonesia). Use when a user asks to create, convert, summarize, or generate presentation slides, PowerPoint, or PPTX from medical journals, clinical papers, NEJM/Lancet/JAMA studies, or medical PDFs.
version: 2.1.0
author: Medical Research & AI Agent Workgroup
license: MIT
---

# Medical Journal Slides Generator

An autonomous skill for extracting, synthesizing, and converting peer-reviewed medical publications into publication-grade, widescreen (16:9), fully editable PowerPoint (`.pptx`) presentations. Built with strict anti-AI-slop design rules, geometric layout bands, programmatic collision verification, and natural Indonesian prose humanization ([`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)).

## Table of Contents
1. [Overview & Agent Persona](#1-overview--agent-persona)
2. [Trigger Conditions](#2-trigger-conditions)
3. [User Interaction Protocol](#3-user-interaction-protocol)
4. [Anti-AI-Slop Design & Layout Guardrails](#4-anti-ai-slop-design--layout-guardrails)
5. [Indonesian Language Humanization Protocol (`humanizer-indonesia`)](#5-indonesian-language-humanization-protocol-humanizer-indonesia)
6. [Study Classification & Archetypes](#6-study-classification--archetypes)
7. [Clinical Data Extraction Guide](#7-clinical-data-extraction-guide)
8. [Slide Architecture & Templates](#8-slide-architecture--templates)
9. [Automated Code Implementation (python-pptx)](#9-automated-code-implementation-python-pptx)
10. [Visual Rendering & Quality Assurance Audit](#10-visual-rendering--quality-assurance-audit)
11. [Execution Environment & Fallback Strategy](#11-execution-environment--fallback-strategy)
12. [Curated Open-Source References & Benchmarks](#12-curated-open-source-references--benchmarks)

---

## 1. Overview & Agent Persona

When this skill or repository is provided to an AI agent or LLM, the agent immediately assumes the role of an **Expert Medical Visualizer & Clinical Presentation Specialist**.

The agent understands:
- Medical publication standards (CONSORT for trials, PRISMA for meta-analyses, STROBE for observational studies).
- Clinical trial endpoints (PFS, OS, HR, 95% CI, p-values, NNT, Grade 3/4 AEs).
- High-impact presentation aesthetics: 16:9 widescreen, action titles, open layout hierarchy, zero text walls, and WCAG-compliant contrast.
- Natural academic Indonesian style governed by [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia) whenever Indonesian output is requested.
- Fully editable `.pptx` generation using native shapes, tables, charts, and text frames.

---

## 2. Trigger Conditions

Activate this skill when:
- The user provides a link to this GitHub repository or loads this skill into their LLM/Agent.
- The user provides or references a medical journal PDF and requests slides, a presentation, or `.pptx` output.
- The user asks to prepare a clinical Journal Club presentation, hospital morning report, or conference plenary deck.
- The user asks for a 1-3 slide "Visual Abstract" or "Quick Take" summarizing a medical study.

---

## 3. User Interaction Protocol

Follow this exact multi-stage interaction lifecycle:

### Stage 1: Welcome & Intake
Immediately acknowledge the request and invite the user to provide the source file:
- Prompt the user to upload or provide the medical journal PDF (or paper text).
- Present an optional, non-intrusive intake questionnaire to calibrate the output.

### Stage 2: Optional Calibration Questionnaire
Provide the following questions clearly, explicitly informing the user that **all questions are optional**:

> "Silakan unggah berkas PDF jurnal medis yang ingin dibuatkan slide. Untuk menyesuaikan hasilnya, kamu bisa menjawab pertanyaan opsional di bawah ini (atau lewati jika ingin menggunakan pengaturan otomatis terbaik):"
>
> 1. **Bahasa Slide**: Bahasa Indonesia (menggunakan standar alur luwes [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)) atau Bahasa Inggris (*Original Medical English*)?
> 2. **Target Jumlah Slide**: (Contoh: Ringkas 1-3 slide / Standar 7-9 slide / Lengkap 10-12 slide)
> 3. **Tujuan / Forum Presentasi**: (Contoh: Hospital Journal Club, Konferensi Ilmiah/Plenary, Laporan Kasus Departemen, atau Telaah Cepat)
> 4. **Preferensi Desain & Palet Warna**: (Contoh: Clinical Navy / Deep Slate / Warm Teal / Minimalist High-Contrast)
> 5. **Fokus Khusus**: Apakah ada data atau grafik spesifik yang ingin ditonjolkan? (Contoh: Kurva survival Kaplan-Meier, tabel toksisitas, atau analisis subgrup Forest plot)

**Rule on Missing Responses**: If the user provides only the PDF without answering the optional questions, proceed immediately using intelligent defaults derived from the paper's study design.

---

## 4. Anti-AI-Slop Design & Layout Guardrails

To eliminate generic "AI-generated" presentation tells (text overflow, nested boxes, monotonous card grids, tiny fonts, and verbose prose), enforce these unconditional layout rules:

### A. Strict Horizontal Layout Bands (Zero Overlap Contract)
`python-pptx` text boxes do not clip; text exceeding its box spills onto lower elements. Every element must reside strictly inside one of three horizontal bands on a `13.333" x 7.5"` widescreen canvas:

| Band | Top (Y) | Height (H) | Bottom Limit | Permitted Contents |
| :--- | :--- | :--- | :--- | :--- |
| **Title Band** | `0.50"` | `1.50"` | `2.00"` | Category kicker + 1-2 line Action Title |
| **Body Band** | `2.20"` | `4.60"` | `6.80"` | Columns, native charts, clinical tables, figure insets |
| **Footer Band** | `6.95"` | `0.35"` | `7.30"` | Verbatim journal citation, DOI, trial acronym, slide number |

*Note: Maintain a mandatory `0.20"` collision buffer between bands. Lateral margins are fixed at `0.80"` (`CONTENT_W = 11.7"`).*

### B. Open Canvas by Default (No Box Spam)
- **Do NOT put every bullet or column inside a colored card/box.** Over-boxing is the #1 visual signature of AI-generated slide slop.
- **Open by Default**: Let text columns, tables, and charts sit cleanly on the light background, separated by generous whitespace (`COL_GAP = 0.45"`).
- **Box Only When Meaningful**: Use soft tinted cards (`TINT = #EFF6FF`, zero border, zero drop shadow) on **at most 1/3 of interior slides**—strictly for parallel PICO comparisons, 3-arm trial schemas, or a single hero statistical callout. Never nest boxes inside boxes.

### C. Structural Bookending & Palette Discipline
- **Bookend Canvas**: Slide 1 (Cover) and Slide $N$ (Terminal Take-Home Outro) use a deep clinical background (`BG_THEME = #0F172A` or `#1B365D`) with light text (`#F8FAFC`) and zero floating cards.
- **Interior Canvas**: Slides $2 \dots N-1$ use a crisp light canvas (`BG_CONTENT = #FFFFFF`) but **must carry the theme palette** via uppercase category kickers (`#2563EB`), bold navy action titles (`#0F172A`), and semantic data accents.
- **Semantic Medical Color Coding**:
  - Intervention Arm: Primary Clinical Blue (`#2563EB`)
  - Control / Placebo Arm: Neutral Slate (`#64748B`)
  - Positive Efficacy / Low Bias: Emerald (`#15803D`)
  - Grade $\ge 3$ Toxicity / Safety Warning: Crimson Red (`#B91C1C`)

### D. Projection Typography & Fragment Copywriting
- **Explicit Sizing on Every Run**: Never leave `run.font.size` unset.
  - Action Title (H1): `32-34 pt` Bold (Max 2 lines, ~84 chars)
  - Hero Stat Callout (HR / NNT / Primary %): `44-54 pt` Bold
  - Column / Section Header: `20-22 pt` Bold (`ACCENT` color, >= 4pt larger than body)
  - Bullet Lead-In Keyword: `17-18 pt` Bold (`TEXT_DARK`)
  - Bullet Body Fragment: `15-16 pt` Regular (`TEXT_MUTED`)
  - Table Cells: `13-14 pt` Regular (never below `11 pt` floor)
  - Footer Citation: `12-14 pt` Regular
- **Fragment Copywriting (Zero Prose Walls)**:
  - Maximum **35-50 words per slide** (hard ceiling: 65 words).
  - Maximum **3 bullets per column** and **6 bullets per slide**.
  - Write bullets as `**Bold Clinical Keyword:** Concise telegraphic fragment (max 8-10 words)`.

### E. Authentic Visuals & Native Figures (Ban Decorative AI Art)
- **NEVER generate decorative AI clip-art or fake medical illustrations** (e.g., glowing brains, futuristic stethoscopes, or hallucinated survival curves).
- Extract and crop **actual figures, Kaplan-Meier curves, CONSORT flowcharts, and forest plots** directly from the user's PDF at 300 DPI (`pypdfium2`), or construct **editable native PowerPoint charts and tables** (`slide.shapes.add_chart()` / `add_table()`) from verbatim numbers.

---

## 5. Indonesian Language Humanization Protocol (`humanizer-indonesia`)

When the user requests slides or speaker notes in **Bahasa Indonesia**, the agent MUST apply the writing rules and workflow from **[`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)** (including `SKILL.md` and `references/id-ai-tells.md`) so the Indonesian text reads naturally like an experienced Indonesian clinician/academic—never like a stiff machine translation or LLM output:

### Core Rules from `adhiprib/humanizer-indonesia`
1. **Pangkas Nominalisasi Bertumpuk (Unpack Stacked Nominalizations)**:
   - Avoid chaining `pe-...-an` / `ke-...-an` nouns (*"Pelaksanaan pemberian intervensi menghasilkan penurunan tingkat kematian"*).
   - Use direct active clinical verbs (*"Pemberian obat X menurunkan angka kematian sebesar 28%"*).
2. **Hapus Kalkir Bahasa Inggris (Eliminate English Calques)**:
   - Avoid literal translations of English syntax such as *"Sebuah studi acak"* (use *"Studi acak"* or *"Uji klinis acak"*), *"di mana"* as a relative pronoun (*"pasien di mana ditemukan..."* -> *"pasien yang mengalami..."*), *"memainkan peran penting dalam"*, or *"dalam hal"*.
3. **Buang Konektor Berlebihan & Pembuka Basa-basi (Strip Excessive Connectors & Throat-Clearing)**:
   - Cut formulaic AI transitions (*"Lebih lanjut,"*, *"Selain itu,"*, *"Penting untuk dicatat bahwa"*, *"Secara keseluruhan,"*, *"Patut digarisbawahi bahwa"*). Start bullets directly with the clinical subject or metric.
4. **Hindari Personifikasi Benda Mati Abstrak (Avoid Inanimate Personification)**:
   - Avoid stiff personifications like *"Studi ini berargumen/menyoroti/menunjukkan..."* on slide titles; state the clinical fact directly (*"Terapi X memperpanjang kesintasan bebas progresi"*).
5. **Jaga Istilah Medis Baku & Singkatan Klinis**:
   - Use established Indonesian medical terminology (*kesintasan keseluruhan* for Overall Survival [OS], *kesintasan bebas progresi* for Progression-Free Survival [PFS], *kejadian tidak diinginkan* [KTD] for Adverse Events [AE], *rasio bahaya* / HR, *interval kepercayaan* / IK 95%) while retaining internationally recognized clinical abbreviations (`HR`, `95% CI`, `p`, `ORR`, `ITT`) where standard in Indonesian hospital journal clubs.
6. **Putaran Audit Bahasa (Draft -> Audit `id-ai-tells` -> Revisi)**:
   - Before writing Indonesian strings into `python-pptx`, run a self-audit pass checking for translation stiffness, uniform sentence rhythm, and redundant copulas (*"adalah merupakan"*), rewriting every slide title and bullet fragment into crisp, natural Indonesian.

---

## 6. Study Classification & Archetypes

When the paper is ingested, classify the study design and assign the archetype:

| Study Design in PDF | Primary Objective | Target Archetype | Default Slide Range |
| :--- | :--- | :--- | :--- |
| **Phase II/III RCT, Drug/Device Trial** | Present efficacy, safety, survival data | **Archetype A: Clinical Trial Plenary** | 7-10 Slides |
| **Hospital Journal Club / Teaching** | Evaluate validity, methodology, local application | **Archetype B: Critical Appraisal Deck** | 8-12 Slides |
| **Quick Summary / Executive Briefing** | High-level synthesis for rapid review | **Archetype C: Visual Abstract** | 1-3 Slides |
| **Systematic Review / Meta-Analysis** | Synthesize multi-study evidence & heterogeneity | **Archetype D: PRISMA & Meta-Analysis** | 6-9 Slides |
| **Cohort / Case-Control / Observational** | Report association, exposure, relative risk | **Archetype A or B (STROBE-adapted)** | 6-8 Slides |

---

## 7. Clinical Data Extraction Guide

Prior to slide generation, extract and record the following verified parameters from the source paper:

1. **Trial Metadata**: Full title, trial acronym (e.g., KEYNOTE-042), journal name, publication year, volume/pages, DOI, and funding source.
2. **PICO Architecture**:
   - **P (Population)**: Target pathology, staging, prior lines of therapy, total randomized ($N$).
   - **I (Intervention)**: Experimental agent/technique, dosing schedule, duration.
   - **C (Comparator)**: Standard of care (SOC), active comparator, or placebo.
   - **O (Outcomes)**: Primary endpoint definition and key secondary/exploratory endpoints.
3. **Statistical Indicators**:
   - Time-to-event: Hazard Ratio (HR), Median Survival (months), 95% Confidence Interval (CI), log-rank p-value.
   - Dichotomous outcomes: Relative Risk (RR), Odds Ratio (OR), Absolute Risk Reduction (ARR), NNT/NNH.
   - Categorical / Continuous: Mean $\pm$ SD, Median [IQR], percentage responses.
4. **Safety & Tolerability**:
   - All-cause adverse events (AEs) vs treatment-related AEs (TRAEs).
   - Grade $\ge 3$ adverse events and treatment discontinuation rates due to toxicity.
5. **Critical Methodology (CONSORT / CASP)**:
   - Randomization technique, allocation concealment, blinding status (double-blind, open-label).
   - Intention-to-Treat (ITT) vs Per-Protocol (PP) populations.

---

## 8. Slide Architecture & Templates

### Archetype A: Clinical Trial Plenary Deck
- **Slide 1: Cover (Bookend Dark)**: Trial title, acronym, lead authors, journal citation, and disclosures rail.
- **Slide 2: Clinical Background & Unmet Need (Open 2-Col)**: Disease burden, standard of care limitations, and biological rationale.
- **Slide 3: Study Design & Eligibility Schema (Boxed 3-Col Pillars)**: Population criteria, randomization arms, and primary/secondary endpoints.
- **Slide 4: Baseline Patient Characteristics (Clean Clinical Table)**: Demographics and disease markers across arms (max 6 cols x 8 rows).
- **Slide 5: Primary Efficacy Endpoint (Split Hero Stat + PDF Curve/Chart)**: Action title stating exact risk reduction, hero HR/p-value callout, and Kaplan-Meier plot.
- **Slide 6: Subgroup & Secondary Analyses (Open Layout + Forest Plot)**: Consistency of benefit across prespecified clinical subgroups.
- **Slide 7: Safety & Tolerability Profile (Structured Toxicity Table)**: All-grade vs Grade $\ge 3$ adverse events and discontinuation rates.
- **Slide 8: Terminal Take-Home Outro (Bookend Dark)**: Single 6-12 word practice-changing clinical conclusion and DOI citation.

### Archetype B: Journal Club Critical Appraisal Deck
- **Slide 1: Cover (Bookend Dark)**: Clinical dilemma question, full article citation, and journal Impact Factor.
- **Slide 2: Clinical Context & Prior Evidence**: Why this study was conducted and gaps in prior literature.
- **Slide 3: Methodology & PICO Schema**: Inclusion/exclusion criteria, intervention protocol, and CONSORT flow summary.
- **Slide 4: Primary Efficacy Results**: Absolute event rates, Relative Risk / Hazard Ratio, 95% CI, and NNT calculation.
- **Slide 5: Secondary Outcomes & Safety**: Key secondary endpoints and major adverse event rates.
- **Slide 6: Internal Validity Appraisal (CASP Checklist)**: Randomization integrity, blinding, attrition, and power calculation review.
- **Slide 7: External Validity & Local Applicability**: Can these results be generalized to our local hospital patient population?
- **Slide 8: Terminal Take-Home Outro (Bookend Dark)**: Bottom-line clinical verdict on institutional practice adoption.

### Archetype C: Visual Abstract (1-3 Slides)
- **Top Band**: Research question framed as an active clinical finding.
- **Body Band (3 Parallel Boxed Pillars)**:
  - *Column 1 (Population)*: Target condition, inclusion criteria, total $N$.
  - *Column 2 (Intervention vs Comparator)*: Dosing arms and follow-up duration.
  - *Column 3 (Primary Outcome)*: Hero stat callout (`44 pt` bold) with HR, 95% CI, and p-value.
- **Footer Band**: One-sentence clinical takeaway + full journal DOI citation.

### Archetype D: Systematic Review & Meta-Analysis
- **Slide 1: Cover (Bookend Dark)**: Review title, PICO objective, and PROSPERO registration ID.
- **Slide 2: Search Strategy & PRISMA Flow**: Databases searched, screening counts, and final included studies ($N$).
- **Slide 3: Risk of Bias Assessment**: Cochrane RoB-2 or Newcastle-Ottawa quality summary across studies.
- **Slide 4: Pooled Primary Effect (Forest Plot)**: Aggregate effect size, 95% CI, and heterogeneity metrics ($I^2$, Cochran $Q$).
- **Slide 5: Subgroup Sensitivity & Publication Bias**: Funnel plot symmetry and sensitivity analyses.
- **Slide 6: Terminal Take-Home Outro (Bookend Dark)**: GRADE certainty rating and clinical guideline implication.

---

## 9. Automated Code Implementation (python-pptx)

Use this verified `python-pptx` boilerplate with built-in capacity assertions and band geometry:

```python
import math
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# ── STRICT HORIZONTAL BANDS (Never invent arbitrary Y coordinates) ──
MARGIN_X   = 0.8
CONTENT_W  = 11.7
TITLE_TOP  = 0.5
TITLE_H    = 1.5
BODY_TOP   = 2.2
BODY_H     = 4.6
FOOTER_TOP = 6.95
FOOTER_H   = 0.35
COL_GAP    = 0.45
COL_W_3    = 3.6
COL_W_2    = 5.6

# ── CAPACITY BUDGET MATH (Prevents text overflow) ──
CHAR_W_EM = 0.52
LINE_H    = 1.22
RAGGED    = 0.90

def chars_per_line(w_in, pt):
    return max(1, int(w_in * 72 / (pt * CHAR_W_EM) * RAGGED))

def lines_needed(text, w_in, pt):
    return math.ceil(len(text) / chars_per_line(w_in, pt)) if text else 1

# ── MEDICAL PALETTE ARCHITECTURE ──
BG_THEME   = RGBColor(0x0F, 0x17, 0x2A)  # Deep Clinical Navy (Slide 1 & Slide N)
BG_CONTENT = RGBColor(0xFF, 0xFF, 0xFF)  # Crisp White (Interior Slides)
TEXT_LIGHT = RGBColor(0xF8, 0xFA, 0xFC)
TEXT_DARK  = RGBColor(0x0F, 0x17, 0x2A)
TEXT_MUTED = RGBColor(0x47, 0x55, 0x69)
ACCENT     = RGBColor(0x25, 0x63, 0xEB)  # Primary Intervention Blue
TINT       = RGBColor(0xEF, 0xF6, 0xFF)  # Soft Card Fill (Use on <= 1/3 slides)
ALERT_RED  = RGBColor(0xB9, 0x1C, 0x1C)  # Grade 3/4 Toxicity Alert
GREEN_POS  = RGBColor(0x15, 0x80, 0x3D)  # Positive Efficacy / Low Bias
HEAD_FONT  = "Calibri"
BODY_FONT  = "Calibri"

def set_slide_canvas(slide, is_bookend=False):
    bg = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_THEME if is_bookend else BG_CONTENT
    bg.line.fill.background()
    bg.shadow.inherit = False
    return bg

def add_header(slide, title_text, kicker_text="CLINICAL TRIAL", is_dark=False):
    n = lines_needed(title_text, CONTENT_W, 32)
    assert n <= 2, f"Title wraps to {n} lines (max 2) -- shorten: {title_text!r}"
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

def add_footer(slide, citation_text, slide_num, is_dark=False):
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

def set_bullet(p, char="\u2022", indent_in=0.25):
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(Inches(indent_in)))
    pPr.set('indent', str(-Inches(indent_in)))
    for xml in ('<a:buFont %s typeface="Arial"/>' % nsdecls('a'),
                '<a:buChar %s char="%s"/>' % (nsdecls('a'), char)):
        pPr.insert_element_before(parse_xml(xml), 'a:buBlip', 'a:tabLst', 'a:defRPr', 'a:extLst')

def add_fragment_bullet(tf, keyword, fragment, box_w_in=COL_W_2, body_pt=16, lead_pt=18):
    full = f"{keyword}: {fragment}"
    n = lines_needed(full, box_w_in, body_pt)
    assert n <= 3, f"Bullet wraps to {n} lines (max 3) -- tighten text: {full!r}"
    p = tf.add_paragraph() if len(tf.paragraphs[0].text) > 0 else tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.space_after = Pt(12)
    set_bullet(p)
    rb = p.add_run()
    rb.text = f"{keyword}: "
    rb.font.size = Pt(lead_pt)
    rb.font.name = BODY_FONT
    rb.font.bold = True
    rb.font.color.rgb = TEXT_DARK
    rf = p.add_run()
    rf.text = fragment
    rf.font.size = Pt(body_pt)
    rf.font.name = BODY_FONT
    rf.font.color.rgb = TEXT_MUTED
```

---

## 10. Visual Rendering & Quality Assurance Audit

Never deliver a presentation without running a visual and geometric inspection:

1. **Automated Geometry Check**:
   - Verify that every shape sits within its designated band (`TITLE_TOP`, `BODY_TOP`, `FOOTER_TOP`).
   - Verify total words per slide $\le 65$ and total bullets per slide $\le 6$.
   - Confirm Slide 1 and Slide $N$ contain zero bullets, tables, or cards.
2. **LibreOffice + PDFium Visual Inspection**:
   - Convert the generated `.pptx` to PDF via headless LibreOffice, render each slide to PNG (`pypdfium2`), and inspect the rendered images for text clipping, table cell wrap issues, or low-contrast labels:
   ```bash
   libreoffice --headless --nologo --nolockcheck --norestore --convert-to pdf presentation.pptx
   ```
3. **Clinical Data & Language Audit**:
   - Cross-check every Hazard Ratio, Confidence Interval, p-value, sample size ($N$), and toxicity percentage against the source PDF.
   - If the deck is in Bahasa Indonesia, verify compliance with [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia) (no stacked nominalizations, no English calques, no AI filler transitions).

---

## 11. Execution Environment & Fallback Strategy

### Native Execution (Full Agent Support)
If the host LLM environment has execution tools (`execute_bash`, Python runtime, and `python-pptx`):
1. Execute the Python script to synthesize the presentation.
2. Run the visual rendering check to verify zero overlap.
3. Deliver the editable `.pptx` file directly to the user for instant download.

### Fallback Routing (When Local PPTX Generation is Unsupported)
If the current LLM environment cannot execute code or cannot generate binary files (e.g., standard text-only chat interfaces without a code execution sandbox):
The LLM must provide the structured slide outline first, and then explicitly advise the user with this standard guidance:

> *"Lingkungan LLM saat ini tidak memiliki akses eksekusi kode untuk menghasilkan berkas PowerPoint (`.pptx`) biner secara langsung. Untuk membuat dan mengunduh berkas `.pptx` yang siap pakai dan dapat diedit secara otomatis dari PDF jurnal Anda, gunakan **Gemini (Task Mode)** yang dilengkapi dengan lingkungan eksekusi sandbox Python dan `python-pptx` bawaan."*

---

## 12. Curated Open-Source References & Benchmarks

- **[adhiprib/humanizer-indonesia](https://github.com/adhiprib/humanizer-indonesia)**: Indonesian prose humanizer skill that strips translation calques, stacked nominalizations, excessive connectors, and stiff AI patterns so Indonesian academic text reads naturally.
- **[Chen-Po/research-slides-skill](https://github.com/Chen-Po/research-slides-skill)**: Agent skill dedicated to generating clean journal-club PowerPoint presentations, featuring claim discipline, verbatim numerical tracing in speaker notes, and layout density audits.
- **[HKUDS/Paper2Slides](https://github.com/HKUDS/Paper2Slides)**: Python pipeline converting academic papers (`.pdf`) into structured presentation slide decks with multi-modal parsing.
- **[M1n-n9/academic-ppt-master](https://github.com/M1n-n9/academic-ppt-master)**: Academic PPT skill for converting academic papers and PDFs into editable PowerPoint presentations for journal clubs, thesis defenses, and lab group meetings.
- **[k-dense-ai/claude-scientific-writer](https://github.com/k-dense-ai/claude-scientific-writer)**: Specification for scientific slide generation with strict whitespace, visual-first data density rules, and WCAG contrast thresholds.
- **[DejanKostyszyn/presentation-template-uniklinik-freiburg](https://github.com/DejanKostyszyn/presentation-template-uniklinik-freiburg)**: Formal presentation templates developed for University Medical Center Freiburg clinical case reports and hospital presentations.
