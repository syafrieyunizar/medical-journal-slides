# Panduan Humanisasi Bahasa Indonesia Medis (`id-medical-humanizer.md`)

Dokumen mandiri (*self-contained*) di dalam repositori ini mengadaptasi kerangka kerja dan katalog pola AI dari **[`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)** karya Pribadi Mumpuni Adhi (lisensi MIT), yang dirancang khusus untuk menyunting teks slide presentasi medis dan *speaker notes* berbahasa Indonesia agar luwes, tajam, dan tidak terbaca seperti terjemahan mesin.

---

## 1. Katalog Pola Khas Tulisan AI Bahasa Indonesia (`id-ai-tells`)

| Kategori Pola AI | Contoh Kalimat Kaku / Terjemahan AI (DILARANG) | Hasil Revisi Luwes & Klinis (GUNAKAN) |
| :--- | :--- | :--- |
| **Nominalisasi Bertumpuk** (`pe-...-an` / `ke-...-an`) | *"Pelaksanaan pemberian kombinasi probiotik menghasilkan penurunan tingkat kecemasan secara signifikan."* | *"Kombinasi probiotik menurunkan tingkat kecemasan secara bermakna."* |
| **Kalkir Sintaksis Inggris** (*"Sebuah studi"*, *"di mana"*) | *"Sebuah uji klinis acak di mana 86 mahasiswa dilibatkan dalam penelitian ini."* | *"Uji klinis acak pada 86 mahasiswa..."* |
| **Konektor Basa-basi AI** (*"Lebih lanjut"*, *"Selain itu"*) | *"Lebih lanjut, penting untuk dicatat bahwa profil keamanan obat dapat ditoleransi dengan baik."* | *"Profil keamanan terjaga: tidak ditemukan kejadian tidak diinginkan serius."* |
| **Personifikasi Abstrak** (*"Studi ini menyoroti/berargumen"*) | *"Tinjauan ini menyoroti pentingnya modulasi sumbu usus-otak pada depresi."* | *"Modulasi sumbu usus-otak meredakan reaktivitas kognitif pada depresi."* |
| **Kopula Ganda & Berlebihan** (*"adalah merupakan"*, *"yaitu adalah"*) | *"Titik akhir primer penelitian ini adalah merupakan kualitas tidur."* | *"Titik akhir primer: kualitas dan durasi tidur (PSQI)."* |
| **Frasa Pembungkus Kosong** (*"dalam hal"*, *"memainkan peran penting"*) | *"Metabolit asam lemak rantai pendek memainkan peran penting dalam hal menjaga sawar usus."* | *"Asam lemak rantai pendek (SCFA) memperkuat integritas sawar mukosa usus."* |

---

## 2. Padanan Istilah Klinis & Statistik Baku Bahasa Indonesia

Saat menyusun slide berbahasa Indonesia, gunakan padanan istilah klinis baku berikut sembari mempertahankan singkatan medis internasional:

- **Randomized Controlled Trial (RCT)**: Uji klinis acak terkontrol (`RPCT` / `RCT`)
- **Cross-over Trial**: Uji silang acak tersamar ganda (`RDBPCOT`)
- **Narrative / Mechanistic Review**: Tinjauan naratif terstruktur / tinjauan mekanistik
- **Gut-Brain Axis (MGBA)**: Sumbu mikrobiota–usus–otak (*gut–brain axis*)
- **Gut Barrier Integrity**: Integritas sawar mukosa usus
- **Short-Chain Fatty Acids (SCFAs)**: Asam lemak rantai pendek (`SCFA`)
- **Primary / Secondary Endpoint**: Titik akhir primer / sekunder (atau *luaran primer / sekunder*)
- **Overall / Progression-Free Survival**: Kesintasan keseluruhan (`OS`) / Kesintasan bebas progresi (`PFS`)
- **Hazard Ratio / Confidence Interval**: Rasio bahaya (`HR`) / Interval kepercayaan 95% (`IK 95%` atau `95% CI`)
- **Adverse Events (AEs)**: Kejadian tidak diinginkan (`KTD`)
- **Cognitive Reactivity & Rumination**: Reaktivitas kognitif dan ruminasi

---

## 3. Aturan Simbol & Lokalisasi PowerPoint

1. **Nomor Halaman Terlokalisasi**: Gunakan parameter `lang="id"` pada fungsi `add_footer()` agar footer menampilkan `Hal. 1`, `Hal. 2`, dst. (bukan `Slide 1`).
2. **Bebas Sintaksis LaTeX Mentah**: Jangan pernah menulis `$I^2$`, `$\ge 3$`, atau `$1 \times 10^9$` di dalam PowerPoint karena PowerPoint tidak merender kode LaTeX. Gunakan karakter Unicode murni: `I²`, `≥ 3`, `1 × 10⁹ CFU`, `α`, `γ-aminobutyric acid (GABA)`.
3. **Penanda Aksesibilitas Buta Warna**: Selalu sertakan simbol eksplisit di samping warna pada tabel klinis:
   - `[✓]` atau `[↑]` / `[↓]` untuk manfaat klinis / perubahan bermakna
   - `[⚠]` untuk efek terbatas, hasil tidak bermakna (*null finding*), atau toksisitas derajat ≥ 3
