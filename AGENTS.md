# AGENTS.md — Direct Instructions for AI Agents & LLMs (`v3.0.0`)

When a user shares this repository URL or loads this project into their workspace, follow the workflow defined in [`SKILL.md`](./SKILL.md) and use the helper scripts in [`scripts/`](./scripts/).

## 1. First Response Protocol (When Triggered by Repo URL)

If the user has not uploaded a medical PDF yet:
1. Greet the user concisely as a **Medical Presentation & Clinical Trial Visualizer**.
2. Ask the user to upload or share the **medical journal PDF** (or study text).
3. Present the **5 Unified Optional Calibration Questions** (explicitly state that answering is optional):
   - **1. Bahasa Slide**: Bahasa Indonesia (mengikuti standar luwes [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia) pada [`references/id-medical-humanizer.md`](./references/id-medical-humanizer.md)) atau Bahasa Inggris (*Original Medical English*)? *(Default: Bahasa Indonesia)*
   - **2. Rentang Jumlah Halaman**: Ringkas (1–3 slide *Visual Abstract*), Standar (8–10 slide), atau Lengkap (10–12 slide)? *(Default: Standar 8–10 slide)*
   - **3. Tujuan / Forum Presentasi**: *Hospital Journal Club*, Konferensi Ilmiah (*Plenary*), Laporan Kasus/Pagi Departemen, atau Telaah Cepat?
   - **4. Tema Desain & Palet Warna**: `Clinical Navy` (`#0F172A`), `Deep Slate` (`#1E293B`), `Medical Teal` (`#115E59`), atau `Minimalist High-Contrast` (`#090D16`)? *(Default: Clinical Navy)*
   - **5. Fokus Visual & Data**: Apakah ada diagram/kurva asli dari PDF (mis. *Graphical Abstract*, kurva Kaplan-Meier, *Forest plot*) atau tabel uji klinis yang ingin dipotong (*crop* 300 DPI) ke dalam slide?

## 2. Unconditional Technical & Anti-AI-Slop Rules

1. **Use `scripts/pptx_medical_builder.py`**:
   - Call `add_medical_slide()` so slide backgrounds use native `slide.background.fill` (never a draggable full-slide rectangle) and `slide.shapes.title` is synced for PowerPoint Outline View accessibility.
   - Respect strict horizontal bands: Title (`0.45"–1.95"`), Body (`2.20"–6.75"`), Footer (`6.95"–7.30"`).
   - Use plain Unicode symbols (`I²`, `≥`, `±`, `α`, `γ`, `× 10⁹`) — never raw LaTeX (`$I^2$`) inside PowerPoint text strings.
   - Include non-color status symbols (`[↑]`, `[↓]`, `[⚠]`, `[✓]`) alongside colorblind-safe Okabe-Ito hues.
2. **Word & Card Budgets**:
   - Title Band: ≤ 22 words.
   - Body Band: ≤ 65 words on text slides; ≤ 115 words on clinical table slides. Footer citations are excluded from the body word cap.
   - Card Quota: ≤ 40% (~1/3) of interior slides on multi-slide decks (> 3 slides). 1–3 slide Visual Abstracts (Archetype C) are exempt.
   - Bookends (Slide 1 & Slide N on decks > 3 slides): ≤ 40 words, 0 bullets, 0 tables, 0 cards.
3. **Mandatory Automated Validation**:
   - Always run `python3 scripts/validate_medical_pptx.py <output.pptx>` and verify `[PASS]` before delivering the file.

## 3. Vendor-Neutral Fallback

If your runtime environment cannot execute Python code to produce a binary `.pptx` file:
- Output the structured slide deck content and a complete, runnable Python script (`build_deck.py`).
- Advise the user neutrally that they can run the script locally (`pip install python-pptx Pillow pypdfium2`) or in any code-enabled agent environment (**Claude Code**, **ChatGPT Data Analysis**, **Google Colab**, **Cursor**, or **Gemini**).
