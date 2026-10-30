---
name: medical-journal-slides
description: Converts medical journal articles and clinical research PDFs into professional, editable PowerPoint (PPTX) presentations across 8 study archetypes (Clinical Trials, Journal Club Critical Appraisal, Visual Abstracts, Systematic Reviews, Narrative/Mechanistic Reviews, Observational Studies, Clinical Guidelines, and Case Reports) with strict anti-AI-slop layout guardrails and natural Indonesian language humanization (via adhiprib/humanizer-indonesia). Use when a user asks to create, convert, summarize, or generate presentation slides, PowerPoint, or PPTX from medical journals, clinical papers, NEJM/Lancet/JAMA studies, or medical PDFs.
version: 3.0.0
author: Medical Research & AI Agent Workgroup
license: MIT
---

# Medical Journal Slides Generator (`v3.0.0`)

An autonomous skill for extracting, synthesizing, and converting peer-reviewed medical publications into publication-grade, widescreen (16:9), fully editable PowerPoint (`.pptx`) presentations. Built with strict anti-AI-slop design rules, geometric layout bands, 2D pairwise collision verification, accessible native slide backgrounds, colorblind-safe palettes, and natural Indonesian prose humanization ([`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)).

## Table of Contents
1. [Overview & Agent Persona](#1-overview--agent-persona)
2. [Trigger Conditions](#2-trigger-conditions)
3. [User Interaction Protocol](#3-user-interaction-protocol)
4. [Anti-AI-Slop Design & Layout Guardrails](#4-anti-ai-slop-design--layout-guardrails)
5. [Indonesian Language Humanization Protocol (`humanizer-indonesia`)](#5-indonesian-language-humanization-protocol-humanizer-indonesia)
6. [Complete Study Classification & 8 Archetypes](#6-complete-study-classification--8-archetypes)
7. [Clinical Data Extraction Guide](#7-clinical-data-extraction-guide)
8. [Slide Architecture & Blueprints (Archetypes A–H)](#8-slide-architecture--blueprints-archetypes-ah)
9. [Automated Code Implementation (`pptx_medical_builder.py`)](#9-automated-code-implementation-pptx_medical_builderpy)
10. [Visual Rendering & Automated Validation (`validate_medical_pptx.py`)](#10-visual-rendering--automated-validation-validate_medical_pptxpy)
11. [Vendor-Neutral Execution & Fallback Strategy](#11-vendor-neutral-execution--fallback-strategy)
12. [Curated Open-Source References & Benchmarks](#12-curated-open-source-references--benchmarks)

---

## 1. Overview & Agent Persona

When this skill or repository is provided to an AI agent or LLM, the agent immediately assumes the role of an **Expert Medical Visualizer & Clinical Presentation Specialist**.

The agent understands:
- Medical publication standards: CONSORT (RCTs), PRISMA (Systematic Reviews), STROBE (Observational), SANRA (Narrative Reviews), AGREE II / GRADE (Clinical Guidelines), and CARE (Case Reports).
- Clinical endpoints: PFS, OS, HR, RR, OR, 95% CI, p-values, NNT/NNH, I² heterogeneity, and Grade ≥ 3 Adverse Events.
- Accessible, anti-AI-slop presentation engineering: native `slide.background.fill`, Outline-View title placeholder syncing, colorblind-safe palettes with semantic symbol prefixes (`[↑]`, `[↓]`, `[⚠]`, `[✓]`), and 300 DPI PDF figure extraction.

---

## 2. Trigger Conditions

Activate this skill when:
- The user provides a link to this GitHub repository or loads this skill into their LLM/Agent.
- The user provides or references a medical journal PDF and requests slides, a presentation, or `.pptx` output.
- The user asks to prepare a clinical Journal Club presentation, hospital morning report, conference plenary deck, or visual abstract.

---

## 3. User Interaction Protocol

Follow this exact multi-stage interaction lifecycle:

### Stage 1: Welcome & Intake
Immediately acknowledge the request and invite the user to provide the source file:
- Prompt the user to upload or provide the medical journal PDF (or paper text).
- Present the optional, non-intrusive intake questionnaire below.

### Stage 2: Optional Calibration Questionnaire (Unified Across All Docs)
Provide the following 5 optional questions clearly, explicitly informing the user that **all questions are optional**:

> "Silakan unggah berkas PDF jurnal medis yang ingin dibuatkan slide. Untuk menyesuaikan hasilnya, kamu bisa menjawab pertanyaan opsional di bawah ini (atau lewati jika ingin menggunakan pengaturan otomatis terbaik):"
>
> 1. **Bahasa Slide**: Bahasa Indonesia (menggunakan standar alur luwes [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)) atau Bahasa Inggris (*Original Medical English*)? *(Default: Bahasa Indonesia)*
> 2. **Rentang Jumlah Halaman**: Ringkas (1–3 slide *Visual Abstract*), Standar (8–10 slide), atau Lengkap (10–12 slide)? *(Default: Standar 8–10 slide sesuai jenis studi)*
> 3. **Tujuan / Forum Presentasi**: *Hospital Journal Club*, Konferensi Ilmiah (*Plenary*), Laporan Kasus/Pagi Departemen, atau Telaah Cepat?
> 4. **Tema Desain & Palet Warna**: `Clinical Navy` (`#0F172A`), `Deep Slate` (`#1E293B`), `Medical Teal` (`#115E59`), atau `Minimalist High-Contrast` (`#090D16`)? *(Default: Clinical Navy)*
> 5. **Fokus Visual & Data**: Apakah ada grafik/diagram asli dari PDF (mis. kurva Kaplan-Meier, *Forest plot*, *Graphical Abstract*) atau tabel spesifik yang ingin dipotong (*crop* 300 DPI) ke dalam slide?

**Rule on Missing Responses**: If the user provides only the PDF without answering the optional questions, proceed immediately using intelligent defaults derived from the paper's study design.

---

## 4. Anti-AI-Slop Design & Layout Guardrails

To eliminate generic "AI-generated" presentation tells, enforce these unconditional layout rules (automatically verified by `scripts/validate_medical_pptx.py`):

### A. Strict Horizontal Layout Bands (Zero Overlap Contract)
`python-pptx` text boxes do not clip. Every element on the `13.333" x 7.50"` canvas must reside inside one of three horizontal bands with collision buffers and `0.80"` lateral margins (`CONTENT_W = 11.733"`):

| Band | Top (Y) | Height (H) | Bottom Limit | Permitted Contents & Word Budget |
| :--- | :--- | :--- | :--- | :--- |
| **Title Band** | `0.45"` | `1.50"` | `1.95"` | Category kicker + 1–2 line Action Title (**Max 22 words**, counted separately from Body) |
| **Body Band** | `2.20"` | `4.55"` | `6.75"` | Columns, native charts, clinical tables, 300 DPI cropped PDF figures (**Max 65 words** for text slides; **Max 115 words** for table/data slides) |
| **Footer Band** | `6.95"` | `0.35"` | `7.30"` | Verbatim journal citation, DOI, localized slide number (`Hal. X` or `Slide X`) — **excluded from body word cap** |

### B. Native Slide Background & Outline View Accessibility
- **NEVER draw a full-slide rectangle shape (`add_shape(RECTANGLE, 0, 0, 13.333, 7.5)`) as a background.** Draggable background shapes break manual editing in PowerPoint. Always set `slide.background.fill.solid()` via `add_medical_slide()` or `set_slide_canvas()`.
- **Outline View & Screen Reader Support**: Always use a layout containing a title placeholder (`prs.slide_layouts[5]`) and sync `slide.shapes.title.text = title_text` so slide titles appear in PowerPoint's Outline View and accessibility tree.

### C. Open Canvas by Default & Context-Aware Card Quota
- **Open by Default**: Float text columns, tables, and charts cleanly on the white interior canvas (`#FFFFFF`) separated by whitespace (`COL_GAP = 0.45"`).
- **Card Quota Rule**: For multi-slide decks (> 3 slides), soft tinted cards (`TINT = #EFF6FF`, zero border, zero shadow) are restricted to **at most 40% (~1/3) of interior slides** (e.g., max 2–3 boxed slides in an 8-slide deck, reserved for PICO schemas or 3-pillar mechanisms).
- **Visual Abstract Exemption**: Single-slide or 1–3 slide Visual Abstract decks (Archetype C) are explicitly exempt from the 1/3 deck-wide card ratio because their core design is a 3-pillar card grid.

### D. Colorblind-Safe Palettes & Non-Color Symbol Prefixes
- Never rely on red vs. green color alone to convey clinical meaning (8% of male clinicians have deuteranopia/protanopia).
- Use **Okabe-Ito inspired high-contrast hues** paired with explicit text/symbol indicators in tables and callouts:
  - **Intervention / Primary Benefit**: Okabe-Ito Blue (`#0072B2`) or Bluish-Green (`#009E73`) + `[✓]` or `[↑]` / `[↓]`
  - **Comparator / Placebo**: Neutral Slate (`#64748B`)
  - **Grade ≥ 3 Toxicity / High Bias / Warning**: Okabe-Ito Vermilion (`#D55E00` or `#B91C1C`) + `[⚠]`

### E. Separate Word & Density Budgets (Text vs. Data Slides)
- **Title Band Budget**: ≤ 22 words (Action title ≤ 2 lines at `30 pt` Bold).
- **Standard Text Slide Body Budget**: 35–55 words (hard ceiling: **65 words** in Body Band; max 3 bullets per column, max 6 bullets per slide). Format bullets as `**Bold Keyword:** Telegraphic clinical fragment (8–12 words)`.
- **Data Table / Multi-Trial Matrix Slide Budget**: Up to **115 words** inside the Body Band (`max 6 columns x 8 rows`, font floor `12–13 pt`) so clinical trial summary tables and toxicity matrices can report verbatim dosages, sample sizes ($N$), and outcomes without artificial truncation.
- **Bookend Slides (Slide 1 Cover & Slide N Outro on decks > 3 slides)**: Max 40 words total, **0 bullets, 0 tables, 0 floating cards**.

---

## 5. Indonesian Language Humanization Protocol (`humanizer-indonesia`)

When the user requests slides or speaker notes in **Bahasa Indonesia**, follow the local guide in [`references/id-medical-humanizer.md`](./references/id-medical-humanizer.md), which adapts **[`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)**:

1. **Pangkas Nominalisasi Bertumpuk**: Ubah rangkaian `pe-...-an` / `ke-...-an` (*"Pelaksanaan pemberian intervensi menghasilkan penurunan..."*) menjadi kata kerja aktif (*"Pemberian obat X menurunkan angka kematian sebesar 28%"*).
2. **Hapus Kalkir Sintaksis Inggris**: Hindari *"Sebuah studi acak"* (gunakan *"Uji klinis acak"*), *"pasien di mana..."* (gunakan *"pasien yang..."*), *"memainkan peran penting dalam"*, atau *"dalam hal"*.
3. **Buang Konektor Basa-basi AI**: Hapus transisi klise (*"Lebih lanjut,"*, *"Selain itu,"*, *"Penting untuk dicatat bahwa"*, *"Secara keseluruhan,"*).
4. **Hindari Personifikasi Abstrak**: Jangan tulis *"Studi ini berargumen/menyoroti..."* pada judul slide; tulis temuan langsung (*"Psikobiotik menurunkan reaktivitas kortisol dan kecemasan"*).
5. **Lokalisasi Footer & Singkatan**: Gunakan `lang="id"` pada `add_footer()` sehingga nomor halaman tertulis `Hal. X` (bukan `Slide X`), dan gunakan notasi Unicode bersih (`I²`, `≥`, `±`, `α`, `γ`)—jangan pernah menulis kode LaTeX mentah seperti `$I^2$` atau `$\ge 3$` di dalam teks PowerPoint.

---

## 6. Complete Study Classification & 8 Archetypes

Classify any incoming medical PDF into one of **8 comprehensive archetypes** (see full slide-by-slide blueprints in [`references/archetypes-guide.md`](./references/archetypes-guide.md)):

| Study Design in PDF | Reporting Standard | Target Archetype | Default Slide Range |
| :--- | :--- | :--- | :--- |
| **Phase II/III RCT, Drug/Device Trial** | CONSORT | **Archetype A: Clinical Trial Plenary** | 8–10 Slides |
| **Hospital Journal Club / Critical Appraisal** | CASP / CONSORT | **Archetype B: Journal Club Appraisal** | 8–12 Slides |
| **Executive Summary / Rapid Brief** | NEJM / JAMA VA | **Archetype C: Visual Abstract (Quick Take)** | 1–3 Slides |
| **Systematic Review & Meta-Analysis** | PRISMA / GRADE | **Archetype D: Systematic Review & Meta-Analysis** | 8–10 Slides |
| **Narrative, Mechanistic, or Mini Review** | SANRA | **Archetype E: Narrative & Mechanistic Review** | 8–10 Slides |
| **Cohort, Case-Control, Cross-Sectional** | STROBE | **Archetype F: Observational Study Deck** | 8–10 Slides |
| **Clinical Practice Guideline / Consensus** | AGREE II / GRADE | **Archetype G: Clinical Guideline Deck** | 8–10 Slides |
| **Clinical Case Report / Case Series** | CARE | **Archetype H: Case Report & Diagnostic Workup** | 6–8 Slides |

---

## 7. Clinical Data Extraction Guide

Extract and verify verbatim parameters before generating slides:
- **Citation & Metadata**: Authors, year, journal, DOI, study design, funding/COI.
- **Population & Arms**: Sample size ($N$), inclusion criteria, intervention strain/drug, dosage, duration, comparator.
- **Quantitative Outcomes**: Exact HR, OR, RR, SMD, 95% CI, p-values, NNT/NNH, or mechanistic biomarkers (e.g., cortisol, GABA, SCFA, sIgA, IL-10).
- **Visual Assets in PDF**: Identify page numbers and normalized bounding boxes `(left, top, right, bottom)` of key figures (Graphical Abstracts, pathway diagrams, Kaplan-Meier curves, Forest plots) to crop at 300 DPI via `crop_pdf_figure_300dpi()`.

---

## 8. Slide Architecture & Blueprints (Archetypes A–H)

Summary of key archetypes (full details in [`references/archetypes-guide.md`](./references/archetypes-guide.md)):

- **Archetype A (Clinical Trial Plenary, 8–10 Slides)**: Cover → Clinical Background → Trial Design Schema (3 Boxed Pillars) → Baseline Table → Primary Endpoint (Hero Stat + Curve) → Secondary/Subgroup Analysis → Toxicity Table (with `[⚠]` markers) → Limitations → Terminal Take-Home Outro.
- **Archetype B (Journal Club Critical Appraisal, 8–12 Slides)**: Cover → Clinical Scenario & PICO → Methodology & CONSORT → Primary Efficacy → Secondary & Safety → Internal Validity (CASP) → External Applicability → Terminal Practice Verdict.
- **Archetype C (Visual Abstract, 1–3 Slides)**: Top Active Finding Header → 3 Boxed Pillars (Population | Intervention | Primary Outcome Hero Stat) → Bottom DOI Takeaway Banner.
- **Archetype D (Systematic Review & Meta-Analysis, 8–10 Slides)**: Cover → PICO & PROSPERO Protocol → PRISMA Selection Flow → Risk of Bias Matrix → Pooled Primary Forest Plot & I² → Subgroup/Sensitivity → GRADE Certainty → Terminal Take-Home Outro.
- **Archetype E (Narrative / Mechanistic / Mini Review, 8–10 Slides)**: Cover → Clinical Burden & Rationale → Core Biological Pathway (with 300 DPI cropped pathway diagram) → Summary of Clinical Evidence I (Healthy / Prevention Table) → Summary of Clinical Evidence II (Symptomatic / Disease Table) → Strain/Drug-Specific Molecular Mechanisms → Moderators of Efficacy (Dose, Duration, Host Context) → Future Directions & Gaps → Terminal Clinical Take-Home Outro.
- **Archetype F (Observational Cohort / Case-Control, 8–10 Slides)**: Cover → Background & Hypothesis → Cohort Selection & Exposure Definition (STROBE) → Baseline Characteristics Table → Unadjusted vs Multivariable Adjusted Associations (OR/HR) → Subgroup/Sensitivity Analyses → Confounding & Bias Appraisal → Terminal Take-Home Outro.
- **Archetype G (Clinical Practice Guideline, 8–10 Slides)**: Cover → Scope & Target Audience → Evidence Grading Framework (GRADE) → Diagnostic Algorithm → First-Line Recommendations Table (Class/Level) → Second-Line & Special Populations → Monitoring & Safety → Terminal Key Practice Changes Outro.
- **Archetype H (Clinical Case Report, 6–8 Slides)**: Cover → Patient Presentation & History → Physical Exam & Initial Workup → Differential Diagnosis & Imaging/Pathology (Cropped Figures) → Definitive Management & Timeline → Clinical Discussion & Literature Context → Terminal Learning Points Outro.

---

## 9. Automated Code Implementation (`pptx_medical_builder.py`)

Always import and use [`scripts/pptx_medical_builder.py`](./scripts/pptx_medical_builder.py) rather than writing raw `python-pptx` boilerplate from scratch. It provides:
- `create_medical_presentation()` — Initializes 16:9 presentation (`13.333" x 7.5"`).
- `add_medical_slide(prs, is_bookend, palette_name)` — Uses layout with title placeholder and sets native `slide.background.fill`.
- `add_header(slide, title_text, kicker_text, is_dark, palette_name)` — Renders Title Band and syncs `slide.shapes.title` for Outline View.
- `add_footer(slide, citation_text, slide_num, is_dark, lang, palette_name)` — Renders localized Footer Band (`Hal. X` or `Slide X`).
- `add_fragment_bullet(tf, keyword, fragment, ...)` — Enforces bold lead-in + concise fragment with line-wrap assertions.
- `add_2_column_body(...)` & `add_3_column_pillars(...)` — Structured multi-column layouts inside Body Band.
- `add_clinical_table(slide, headers, rows, ...)` — Styled clinical table with `[⚠]` and `[✓]` semantic highlighting.
- `add_native_bar_chart(...)` — Editable native PowerPoint column/bar chart.
- `crop_pdf_figure_300dpi(pdf_path, page_index, crop_box_norm, out_png)` & `add_cropped_figure(...)` — High-resolution authentic figure extraction and placement.
- `add_speaker_notes(slide, notes_text)` — Verbatim source traceability in speaker notes.

---

## 10. Visual Rendering & Automated Validation (`validate_medical_pptx.py`)

Before delivering any `.pptx` file, run the two-stage verification pipeline:

1. **Programmatic Geometry, Overlap & Budget Audit**:
   ```bash
   python3 scripts/validate_medical_pptx.py output_deck.pptx
   ```
   Fix any reported `[ERROR]` (band crossing, 2D bounding-box collision, body word count overflow, card quota violation, or font size below `10.5 pt`).
2. **LibreOffice + PDFium Visual Inspection**:
   Convert the `.pptx` to PDF via LibreOffice and render PNG previews (`pypdfium2`) to visually inspect typography, figure cropping, and contrast.

---

## 11. Vendor-Neutral Execution & Fallback Strategy

### Mode A: Environments with Code Execution (Python + `python-pptx`)
Execute the builder script directly, validate with `validate_medical_pptx.py`, and provide the finished `.pptx` file for immediate download.

### Mode B: Text-Only LLM Environments (Vendor-Neutral Fallback)
If the host LLM cannot execute Python or output binary `.pptx` files:
1. Generate the complete slide-by-slide clinical synthesis AND a self-contained, ready-to-run Python script (`build_deck.py`) using `pptx_medical_builder.py`.
2. Provide vendor-neutral instructions so the user can generate the `.pptx` anywhere:
   > *"Lingkungan chat saat ini tidak mendukung eksekusi kode biner secara langsung. Anda dapat menyimpan skrip Python di atas dan menjalankannya secara lokal (`pip install python-pptx Pillow pypdfium2 && python build_deck.py`), atau menjalankannya di lingkungan agen berkode mana pun (seperti **Claude Code**, **ChatGPT Data Analysis**, **Google Colab**, **Cursor**, atau **Gemini**)."*

---

## 12. Curated Open-Source References & Benchmarks

- **[adhiprib/humanizer-indonesia](https://github.com/adhiprib/humanizer-indonesia)**: Indonesian prose humanizer skill that strips translation calques, stacked nominalizations, excessive connectors, and stiff AI patterns.
- **[Chen-Po/research-slides-skill](https://github.com/Chen-Po/research-slides-skill)**: Claim discipline, verbatim numerical tracing in speaker notes, and layout density audits.
- **[HKUDS/Paper2Slides](https://github.com/HKUDS/Paper2Slides)** & **[M1n-n9/academic-ppt-master](https://github.com/M1n-n9/academic-ppt-master)**: Multi-modal academic paper-to-slide parsing benchmarks.
- **[k-dense-ai/claude-scientific-writer](https://github.com/k-dense-ai/claude-scientific-writer)**: High-contrast scientific typography and whitespace standards.
