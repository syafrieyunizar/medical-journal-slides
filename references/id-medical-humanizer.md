# Panduan Humanisasi Bahasa Indonesia Medis (`humanizer-indonesia`)

Rujukan ini mengadaptasi kaidah dari **[`adhiprib/humanizer-indonesia`](https://github.com/adhiprib/humanizer-indonesia)** (`SKILL.md` dan `references/id-ai-tells.md`) khusus untuk penyusunan teks slide presentasi medis dan *speaker notes* berbahasa Indonesia agar luwes, tajam, dan tidak terbaca seperti terjemahan mesin atau keluaran mentah LLM.

## 1. Pola Khas Tulisan AI Bahasa Indonesia yang Wajib Dihapus (`id-ai-tells`)

| Kategori Pola AI | Contoh Kalimat Kaku / Terjemahan AI (DILARANG) | Hasil Revisi Luwes & Klinis (GUNAKAN) |
| :--- | :--- | :--- |
| **Nominalisasi Bertumpuk** (`pe-...-an` / `ke-...-an`) | *"Pelaksanaan pemberian kombinasi kemoterapi menghasilkan penurunan tingkat kematian secara signifikan."* | *"Kombinasi kemoterapi menurunkan angka kematian sebesar 32%."* |
| **Kalkir Sintaksis Inggris** (*"Sebuah studi"*, *"di mana"*) | *"Sebuah uji klinis acak di mana 450 pasien dilibatkan dalam penelitian ini."* | *"Uji klinis acak pada 450 pasien..."* |
| **Konektor Basa-basi AI** (*"Lebih lanjut"*, *"Selain itu"*) | *"Lebih lanjut, penting untuk dicatat bahwa profil keamanan obat dapat ditoleransi dengan baik."* | *"Profil toksisitas terkendali: KTD derajat >= 3 tercatat pada 14% pasien."* |
| **Personifikasi Abstrak** (*"Studi ini menyoroti/berargumen"*) | *"Studi ini menyoroti pentingnya pemberian terapi dini pada syok septik."* | *"Pemberian vasopresor dini menurunkan mortalitas 28 hari pada syok septik."* |
| **Kopula Ganda & Berlebihan** (*"adalah merupakan"*, *"yaitu adalah"*) | *"Titik akhir primer penelitian ini adalah merupakan kesintasan bebas progresi."* | *"Titik akhir primer: kesintasan bebas progresi (PFS)."* |
| **Frasa Pembungkus Kosong** (*"dalam hal"*, *"memainkan peran penting"*) | *"Biomarker PD-L1 memainkan peran penting dalam hal memprediksi respons klinis."* | *"Ekspresi PD-L1 >= 50% memprediksi respons terapi yang lebih tinggi."* |

## 2. Padanan Istilah Klinis & Statistik Baku Bahasa Indonesia

- **Randomized Controlled Trial (RCT)**: Uji klinis acak terkontrol
- **Primary / Secondary Endpoint**: Titik akhir primer / sekunder (atau *luaran primer / sekunder*)
- **Overall Survival (OS)**: Kesintasan keseluruhan (*Overall Survival* / OS)
- **Progression-Free Survival (PFS)**: Kesintasan bebas progresi (PFS)
- **Objective Response Rate (ORR)**: Laju respons objektif (ORR)
- **Hazard Ratio (HR)**: Rasio bahaya (`HR`)
- **Confidence Interval (95% CI)**: Interval kepercayaan 95% (`IK 95%` atau `95% CI`)
- **Number Needed to Treat (NNT)**: Jumlah pasien yang perlu diobati (`NNT`)
- **Adverse Events (AEs)**: Kejadian tidak diinginkan (`KTD` / *Adverse Events*)
- **Serious Adverse Events (SAEs)**: Kejadian tidak diinginkan serius (`KTDS`)
- **Standard of Care (SoC)**: Terapi standar
- **Intention-to-Treat (ITT)**: Analisis niat-obati (`ITT`)
- **Loss to Follow-up**: Lepas pantau

## 3. Alur Kerja Tiga Putaran (Draf -> Audit -> Revisi)

1. **Putaran 1 (Draf Ekstraksi)**: Susun judul *action title* dan poin-poin data dari PDF jurnal.
2. **Putaran 2 (Audit `id-ai-tells`)**: Periksa apakah masih ada kata *"Sebuah"*, *"di mana"*, *"Lebih lanjut"*, *"Selain itu"*, atau kalimat yang melebihi 10 kata setelah kata kunci tebal.
3. **Putaran 3 (Pemadatan Telegrafis)**: Pangkas menjadi fragmen telegrafis padat (`**Kata Kunci Tebal:** Fragmen data 6–10 kata`) agar muat dalam batas zona `Body Band` tanpa meluber.
