# Medical Journal Slides (`medical-journal-slides`)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Format: PPTX](https://img.shields.io/badge/Output-Editable_.pptx-2563EB.svg)](#)
[![Language: EN / ID](https://img.shields.io/badge/Language-English_%7C_Bahasa_Indonesia-15803D.svg)](https://github.com/adhiprib/humanizer-indonesia)

An open-source AI Agent Skill that transforms peer-reviewed medical journal articles and clinical research PDFs (from *NEJM*, *The Lancet*, *JAMA*, *BMJ*, *JCO*, etc.) into publication-grade, widescreen (16:9), **fully editable PowerPoint (`.pptx`) presentations**.

Built with strict **Anti-AI-Slop layout guardrails**, automated study-design classification, verbatim clinical metric extraction, and natural Indonesian prose humanization via [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia).

---

## Fitur Utama (Key Features)

1. **4 Arketipe Presentasi Medis Otomatis**:
   - **Archetype A — Clinical Trial Plenary (7–10 Slide)**: Gaya konferensi ASCO/ESMO/AHA untuk uji klinis acak (RCT) fase II/III.
   - **Archetype B — Hospital Journal Club & Critical Appraisal (8–12 Slide)**: Bedah metodologi PICO, alur CONSORT, validitas internal (CASP), dan aplikabilitas klinis lokal.
   - **Archetype C — Visual Abstract / Quick Take (1–3 Slide)**: Ringkasan eksekutif bergaya *NEJM Quick Take* / *JAMA Visual Abstract*.
   - **Archetype D — Systematic Review & Meta-Analysis (6–9 Slide)**: Alur seleksi PRISMA, matriks *Risk of Bias*, *Forest Plot*, uji heterogenitas ($I^2$), dan kepastian bukti GRADE.
2. **Guardrail Desain Anti-AI-Slop**:
   - **Kontrak Zona Horizontal (*Strict Horizontal Bands*)**: Mencegah teks bertabrakan atau meluber (*overflow*) dengan pembagian zona ketat (`Title Band`, `Body Band`, `Footer Band`).
   - **Open Canvas by Default**: Bebas dari tumpukan kotak-kotak kartu AI yang monoton (*box spam*). Kartu lembut (`#EFF6FF`) hanya digunakan maksimal pada 1/3 jumlah slide.
   - **Visual Medis Otentik**: Melarang ilustrasi AI dekoratif palsu. Grafik kurva Kaplan-Meier, diagram CONSORT, dan *forest plot* dipotong langsung dari PDF asli pada 300 DPI (`pypdfium2`) atau dibangun sebagai grafik/tabel *native* PowerPoint yang dapat diedit.
3. **Dukungan Bahasa Indonesia Luwes (`humanizer-indonesia`)**:
   - Jika pengguna meminta slide berbahasa Indonesia, agen otomatis menerapkan kaidah [`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia) untuk memangkas nominalisasi bertumpuk, menghapus kalkir bahasa Inggris, dan membuang konektor klise khas AI.

---

## Cara Penggunaan (How It Works)

### 1. Cukup Berikan Link Repo Ini ke AI Agent / LLM Anda
Salin tautan repositori GitHub ini ke dalam percakapan dengan LLM/Agen Anda (seperti **Gemini Task Mode**, **Claude**, atau **Cursor / Codex**):

```text
Gunakan skill dari repo ini: https://github.com/syafrieyunizar/medical-journal-slides
Saya ingin membuat presentasi PowerPoint dari jurnal medis.
```

### 2. Alur Interaksi Otomatis
Begitu membaca repositori ini (`SKILL.md` dan `AGENTS.md`), agen AI akan:
1. Meminta Anda mengunggah berkas **PDF jurnal medis** yang ingin dijadikan slide.
2. Menyajikan **5 pertanyaan kalibrasi opsional** (Bahasa slide, target jumlah halaman, forum presentasi, tema warna, dan fokus grafik). Anda bebas menjawabnya atau langsung melewati untuk memakai pengaturan otomatis.
3. Mengekstrak metrik klinis verbatim (PICO, $N$, HR, 95% CI, *p-value*, NNT, dan toksisitas Grade $\ge 3$).
4. Menghasilkan berkas **`.pptx` yang sepenuhnya dapat diedit** lengkap dengan validasi tata letak.

> **Catatan Kompatibilitas LLM**: Jika Anda menggunakan antarmuka LLM yang tidak mendukung eksekusi kode Python untuk membuat file `.pptx` biner, agen akan menyusun kerangka konten slide secara lengkap lalu mengarahkan Anda ke **Gemini (Task Mode)** agar berkas `.pptx` dapat langsung dibuat dan diunduh.

---

## Struktur Repositori (Repository Structure)

```text
medical-journal-slides/
├── README.md                          # Dokumentasi utama & panduan penggunaan
├── SKILL.md                           # Spesifikasi utama Agent Skill (v2.1.0)
├── AGENTS.md                          # Instruksi langsung untuk AI Agent (Gemini, Claude, Codex)
├── LICENSE                            # Lisensi MIT
├── .gitignore                         # Konfigurasi gitignore standar
├── references/
│   ├── archetypes-guide.md            # Detail struktur slide per arketipe studi medis
│   └── id-medical-humanizer.md        # Panduan humanisasi bahasa Indonesia medis (humanizer-indonesia)
└── scripts/
    ├── pptx_medical_builder.py        # Modul helper python-pptx dengan kontrak zona anti-AI-slop
    └── validate_medical_pptx.py       # Skrip audit geometri, batas kata, dan deteksi tumpang tindih
```

---

## Kredit & Referensi Terbuka

- **[adhiprib/humanizer-indonesia](https://github.com/adhiprib/humanizer-indonesia)** — Kaidah penyuntingan teks Bahasa Indonesia agar luwes dan bebas pola terjemahan AI.
- **[Chen-Po/research-slides-skill](https://github.com/Chen-Po/research-slides-skill)** — Disiplin verifikasi angka dan struktur slide *journal club*.
- **[HKUDS/Paper2Slides](https://github.com/HKUDS/Paper2Slides)** & **[M1n-n9/academic-ppt-master](https://github.com/M1n-n9/academic-ppt-master)** — Referensi ekstraksi dokumen ilmiah ke slide presentasi.

---

## Lisensi

Dirilis di bawah [MIT License](LICENSE).
