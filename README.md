# Medical Journal Slides (`medical-journal-slides`) — v3.0.0

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Format: PPTX](https://img.shields.io/badge/Output-Editable_.pptx-0072B2.svg)](#)
[![Language: EN / ID](https://img.shields.io/badge/Language-English_%7C_Bahasa_Indonesia-009E73.svg)](https://github.com/adhiprib/humanizer-indonesia)
[![Validation: 2D Collision & Band QA](https://img.shields.io/badge/QA-validate__medical__pptx.py-D55E00.svg)](#)

> **GitHub Repository Description**: Autonomous AI Agent Skill & Python toolkit (`python-pptx`) for converting medical journal PDFs (RCTs, Meta-Analyses, Narrative Reviews, Guidelines, Case Reports) into clean, anti-AI-slop, editable PowerPoint (`.pptx`) decks in English or humanized Bahasa Indonesia.
>
> **Recommended GitHub Topics**: `medical-presentations` `journal-club` `python-pptx` `ai-agent-skills` `clinical-trials` `bahasa-indonesia` `humanizer` `scientific-slides`

---

## Fitur Utama (Key Capabilities)

1. **8 Arketipe Studi Medis Lengkap (`Archetypes A–H`)**:
   - **Archetype A — Clinical Trial Plenary (8–10 Slide)**: Uji klinis acak (RCT) fase II/III berstandar CONSORT/ASCO/ESMO.
   - **Archetype B — Hospital Journal Club & Critical Appraisal (8–12 Slide)**: Bedah metodologi PICO, validitas internal (CASP), dan aplikabilitas klinis lokal.
   - **Archetype C — Visual Abstract / Quick Take (1–3 Slide)**: Ringkasan eksekutif 3 pilar bergaya *NEJM* / *JAMA*.
   - **Archetype D — Systematic Review & Meta-Analysis (8–10 Slide)**: Alur PRISMA, matriks *Risk of Bias*, *Forest Plot*, heterogenitas (`I²`), dan kepastian bukti GRADE.
   - **Archetype E — Narrative, Mechanistic, or Mini Review (8–10 Slide)**: Sintesis jalur biologis/molekuler, gambar mekanisme 300 DPI, dan tabel perbandingan uji klinis (berstandar SANRA).
   - **Archetype F — Observational Cohort / Case-Control (8–10 Slide)**: Studi kohort/registri berstandar STROBE.
   - **Archetype G — Clinical Practice Guideline (8–10 Slide)**: Algoritma diagnosis dan kelas rekomendasi terapi (AGREE II / GRADE).
   - **Archetype H — Clinical Case Report (6–8 Slide)**: Kronologi klinis, pemeriksaan penunjang, dan *clinical pearls* (CARE).

2. **Arsitektur Anti-AI-Slop & Aksesibilitas Native**:
   - **Latar Belakang Native (`slide.background.fill`) & Sinkronisasi Outline View**: Tidak menggunakan persegi panjang penuh yang mudah tergeser saat diedit, serta menyinkronkan `slide.shapes.title` agar judul slide terbaca di *Outline View* dan pembaca layar (*screen reader*).
   - **Kontrak Zona Horizontal & Anggaran Kata Terpisah**: Memisahkan kuota kata antara `Title Band` (≤ 22 kata), `Body Band` (≤ 65 kata untuk slide teks; ≤ 115 kata untuk slide tabel klinis), dan `Footer Band` (dikecualikan dari batas kata isi).
   - **Palet Ramah Buta Warna (Okabe-Ito) + Simbol Semantik**: Menggunakan kontras biru (`#0072B2`), hijau kebiruan (`#009E73`), dan vermilion (`#D55E00`) yang selalu didampingi penanda simbol eksplisit (`[✓]`, `[↑]`, `[↓]`, `[⚠]`).
   - **Pemotong Gambar PDF 300 DPI Bawaan**: Fungsi `crop_pdf_figure_300dpi()` di dalam `scripts/pptx_medical_builder.py` mengekstrak kurva atau diagram mekanisme langsung dari PDF asli.

3. **Validator Geometri & Deteksi Tumpang Tindih Nyata (`validate_medical_pptx.py`)**:
   - Mengecek kepatuhan koordinat vertikal tiap zona (`Title`, `Body`, `Footer`).
   - Melakukan **deteksi tabrakan kotak pembatas 2D berpasangan (*pairwise 2D bounding-box collision detection*)** antar-elemen.
   - Menghitung kata di dalam *text frame* maupun sel tabel secara terpisah dari judul/footer.
   - Memverifikasi aturan *bookend* (Cover & Outro bebas bullet/tabel/kartu) serta kuota kartu (≤ 40% slide interior untuk dek > 3 slide; dikecualikan untuk *Visual Abstract* 1–3 slide).

4. **Humanisasi Bahasa Indonesia Medis (`humanizer-indonesia`)**:
   - Mengintegrasikan panduan mandiri di [`references/id-medical-humanizer.md`](./references/id-medical-humanizer.md) yang diadaptasi dari [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia), lengkap dengan lokalisasi nomor halaman (`Hal. X`) dan bebas karakter LaTeX mentah.

---

## Cara Penggunaan (How to Use)

### 1. Lewat AI Agent / LLM (Claude Code, Gemini, Cursor, ChatGPT)
Bagikan tautan repositori ini atau unggah isi folder ke agen AI Anda bersama berkas PDF jurnal:

```text
Gunakan skill dari repo ini untuk membuat slide PowerPoint (.pptx) dari PDF jurnal terlampir.
```

Agen akan menawarkan **5 Pertanyaan Kalibrasi Opsional** (Bahasa slide, Rentang halaman, Forum presentasi, Palet warna: `Clinical Navy` / `Deep Slate` / `Medical Teal` / `Minimalist High-Contrast`, serta Fokus visual) atau langsung mengeksekusi dengan pengaturan otomatis jika Anda memilih melewati pertanyaan.

### 2. Eksekusi Lokal & Pengujian Otomatis (Unit Test)
Anda juga dapat menjalankan pembangun slide dan rangkaian pengujian secara lokal di mesin apa pun:

```bash
pip install python-pptx Pillow pypdfium2
python3 tests/test_medical_slides.py
```

---

## Struktur Repositori

```text
medical-journal-slides/
├── README.md                          # Dokumentasi utama, metadata repo, & panduan penggunaan
├── SKILL.md                           # Spesifikasi lengkap Agent Skill (v3.0.0)
├── AGENTS.md                          # Instruksi langsung untuk AI Agent (lintas platform)
├── LICENSE                            # Lisensi MIT
├── .gitignore                         # Konfigurasi pengecualian berkas sementara
├── references/
│   ├── archetypes-guide.md            # Cetak biru lengkap 8 Arketipe Studi Medis (A–H)
│   └── id-medical-humanizer.md        # Panduan mandiri humanisasi Bahasa Indonesia medis
├── scripts/
│   ├── pptx_medical_builder.py        # Library lengkap: native background, tabel, chart, crop 300 DPI
│   └── validate_medical_pptx.py       # Validator 2D overlap, band, kuota kartu, & anggaran kata terpisah
├── tests/
│   └── test_medical_slides.py         # Suite pengujian otomatis pembangun & validator PPTX
└── examples/
    └── psychobiotics_gut_brain_axis_id.pptx  # Contoh nyata output PPTX (Sisubalan et al., 2026)
```

---

## Kredit & Lisensi

- Diadaptasi untuk kaidah Bahasa Indonesia dari **[adhiprib/humanizer-indonesia](https://github.com/adhiprib/humanizer-indonesia)** (MIT License).
- Referensi arsitektur slide akademik: **[Chen-Po/research-slides-skill](https://github.com/Chen-Po/research-slides-skill)**, **[HKUDS/Paper2Slides](https://github.com/HKUDS/Paper2Slides)**, dan **[k-dense-ai/claude-scientific-writer](https://github.com/k-dense-ai/claude-scientific-writer)**.
- Dirilis di bawah [MIT License](LICENSE).
