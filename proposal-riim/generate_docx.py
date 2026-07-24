#!/usr/bin/env python3
"""Generate RIIM proposal DOCX from structured content."""
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from pathlib import Path

doc = Document()
for s in doc.sections:
    s.top_margin = Cm(2.5)
    s.bottom_margin = Cm(2.5)
    s.left_margin = Cm(3)
    s.right_margin = Cm(2.5)

style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(12)

def add_center(text, bold=False, size=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"
    return p

def add_para(text, bold=False, justify=True):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(12)
    run.font.name = "Times New Roman"
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    return p

def add_h(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "Times New Roman"
    return h

# COVER
add_center("PROPOSAL", True, 14)
add_center("RISET DAN INOVASI UNTUK INDONESIA MAJU KOMPETISI", True, 14)
add_center("(RIIM KOMPETISI)", True, 14)
doc.add_paragraph()
add_center("TEMA", True)
add_center("Kedaulatan Pangan")
doc.add_paragraph()
add_center("JUDUL", True)
add_center("Integrasi Genomik dan Metabolomik pada Fenotip Teropong,", True, 12)
add_center("Sifat Produksi, dan Reproduksi sebagai Basis Penggaluran", True, 12)
add_center("Bibit Unggul Domba Temanggung untuk Ketahanan Pangan Nasional", True, 12)
doc.add_paragraph()
add_center("KETUA: [Nama Ketua Periset]")
add_center("ANGGOTA:")
add_center("1. [Nama Anggota 1]")
add_center("2. [Nama Anggota 2]")
add_center("3. [Nama Anggota 3]")
doc.add_paragraph()
add_center("INSTITUSI PENGUSUL", True)
add_center("[Nama Institusi Pengusul]")
doc.add_paragraph()
add_center("TAHUN 2026", True)

doc.add_page_break()

# PENGESAHAN
add_center("HALAMAN PENGESAHAN", True, 14)
add_center("PROPOSAL RISET DAN INOVASI UNTUK INDONESIA MAJU KOMPETISI", True, 12)
doc.add_paragraph()
add_para("1. Tema\t: Kedaulatan Pangan", justify=False)
add_para("2. Judul Proposal\t: Integrasi Genomik dan Metabolomik pada Fenotip Teropong, Sifat Produksi, dan Reproduksi sebagai Basis Penggaluran Bibit Unggul Domba Temanggung untuk Ketahanan Pangan Nasional", justify=False)
add_para("3. Ketua Periset", bold=True, justify=False)
for line in [
    "a. Nama Lengkap\t: [diisi]",
    "b. NIP/NIK\t: [diisi]",
    "c. Jabatan\t: [diisi]",
    "d. Institusi Periset\t: [diisi]",
    "e. Unit Kerja Periset\t: [diisi]",
    "f. Alamat\t: [diisi]",
    "g. No. HP/WA\t: [diisi]",
    "h. Email\t: [diisi]",
]:
    add_para(line, justify=False)
add_para("4. Mitra Riset\t: Dinas Ketahanan Pangan, Pertanian dan Perikanan (DKPPP) Kabupaten Temanggung / Komunitas Domba Seni Lokal (DSL) Temanggung", justify=False)
add_para("Alamat Mitra Riset\t: Kabupaten Temanggung, Jawa Tengah", justify=False)
add_para("Peran Mitra Riset\t: penyedia akses populasi ternak dan lahan uji lapangan; fasilitasi sampling dan pencatatan fenotip bersama peternak; pengguna akhir hasil penggaluran bibit; dukungan hilirisasi rekomendasi kebijakan penetapan galur/rumpun lokal", justify=False)
add_para("5. Anggota Periset", bold=True, justify=False)
add_para("[Tabel anggota diisi: Nama | Institusi | No. HP/WA | Email]", justify=False)
add_para("6. Keluaran (ringkas)", bold=True, justify=False)
add_para("Tahun 1 (Teropong & produksi): 1 KTI under review; draft KI panel marka Teropong–produksi; model/standar identitas–produksi.", justify=False)
add_para("Tahun 2 (Reproduksi): KTI T1 accepted + 1 KTI baru under review (reproduksi); KI terdaftar; model seleksi reproduksi.", justify=False)
add_para("Tahun 3 (Usulan galur): minimal 1 KTI published + 1 accepted + 1 under review; KI tambahan; naskah usulan galur siap diajukan.", justify=False)
add_para("7. Pendanaan", bold=True, justify=False)
add_para("Tahun 1: Rp400.000.000; Tahun 2: Rp450.000.000; Tahun 3: Rp400.000.000; Total: Rp1.250.000.000 (dana pendamping Rp0). Disusun berbasis kesetaraan output: 1 keluaran ≈ Rp150–200 juta; paket 1 KTI + 1 KI + 1 model ≈ Rp300–450 juta/tahun.", justify=False)
add_para("Dengan ini menyatakan bahwa proposal yang diajukan bersifat orisinil dan belum pernah memperoleh pendanaan dari lembaga/sumber dana lain, serta tidak mengandung plagiasi.")
add_para("Mengetahui,\t\t\t\t\tKetua Tim", justify=False)
add_para("Kepala Unit Kerja", justify=False)
add_para("(ttd dan cap basah/TTE)\t\t\t(ttd/TTE)", justify=False)
add_para("Nama: ........................\t\tNama: ........................", justify=False)
add_para("NIP: .........................\t\tNIP: .........................", justify=False)

doc.add_page_break()

# ISI
add_h("1. Judul Riset", 1)
add_para("Integrasi Genomik dan Metabolomik pada Fenotip Teropong, Sifat Produksi, dan Reproduksi sebagai Basis Penggaluran Bibit Unggul Domba Temanggung untuk Ketahanan Pangan Nasional")

add_h("2. Abstrak", 1)
add_para(
    "Domba Temanggung merupakan plasma nutfah lokal Jawa Tengah dengan ciri khas fenotip teropong "
    "(warna bulu hitam melingkar di sekitar mata), yang saat ini diusulkan sebagai identitas galur/rumpun lokal unggul. "
    "Namun, diferensiasi galur Teropong versus non-Teropong masih bertumpu pada deskripsi fenotipik visual dan data genetik "
    "maternal terbatas (mtDNA), sehingga belum tersedia bukti genomik nuklear maupun profil metabolomik yang mendukung "
    "penggaluran berbasis sains. Riset ini mengintegrasikan pendekatan genomik (GWAS dan signatures of selection) serta "
    "metabolomik secara bertahap: Tahun 1 fokus fenotip Teropong dan sifat produksi; Tahun 2 fokus sifat reproduksi; "
    "Tahun 3 integrasi multi-omik untuk menghasilkan standar dan naskah usulan galur. Populasi studi berada di Kabupaten "
    "Temanggung bersama mitra DKPPP dan komunitas peternak. Luaran mencakup publikasi internasional bereputasi, kekayaan "
    "intelektual, model seleksi per tahap, serta dokumen usulan galur yang siap diajukan untuk memperkuat bibit lokal unggul "
    "dan ketahanan pangan nasional."
)
add_para("Kata kunci: Domba Temanggung; fenotip teropong; GWAS; signatures of selection; metabolomik; sifat produksi; reproduksi; penggaluran bibit; ketahanan pangan", bold=False)
add_para("Keywords: Temanggung sheep; periocular black phenotype; GWAS; signatures of selection; metabolomics; production traits; reproduction; breed/line development; food security")

add_h("3. Pendahuluan", 1)
add_h("3.1 Latar Belakang", 2)
add_para(
    "Ketahanan pangan nasional menuntut peningkatan ketersediaan protein hewani yang berkelanjutan, efisien, dan berbasis "
    "sumberdaya genetik lokal. Domba merupakan komoditas strategis bagi peternak rakyat karena siklus produksi relatif cepat, "
    "modal relatif terjangkau, serta adaptif pada sistem pemeliharaan ekstensif–semi intensif. Di Kabupaten Temanggung, "
    "Domba Temanggung (sering disebut Domba Teropong) memiliki ciri khas pola bulu hitam di sekitar mata menyerupai "
    "kacamata/teropong, disertai karakter morfologi lokal lain (ekor tipis berbentuk V/ngotes, tanduk melengkung, telinga "
    "semi menggantung). Populasi diperkirakan sekitar 20.000 ekor dan menjadi identitas peternakan daerah, termasuk melalui "
    "komunitas Domba Seni Lokal (DSL). Pemerintah daerah telah menginisiasi pengajuan penetapan sebagai rumpun lokal asli "
    "untuk memperkuat legalitas dan standar genetik."
)
add_para(
    "Penggaluran berbasis fenotip Teropong memiliki nilai strategis ganda: (1) menjaga kemurnian dan identitas genetik lokal; "
    "serta (2) membangun sistem seleksi bibit yang lebih terukur apabila fenotip identitas tersebut dapat dikaitkan dengan "
    "keunggulan produksi dan reproduksi. Namun, informasi genetik Domba Temanggung masih sangat terbatas. Studi awal berbasis "
    "mtDNA cytochrome b menunjukkan keragaman maternal rendah dan campuran haplogroup A/B, tetapi belum dapat dijadikan bukti "
    "diferensiasi genomik nuklear antar fenotip Teropong dan non-Teropong (Ridlo et al., 2026). Studi pada domba lokal Jawa "
    "umumnya masih fokus pada keragaman mtDNA antarrumpun (Ibrahim et al., 2021; Ibrahim et al., 2023), belum pada asosiasi "
    "genom–fenotip untuk penggaluran."
)
add_para(
    "Di tingkat global, warna bulu/pola pigmen pada domba telah berhasil diurai melalui GWAS dan selective sweep, dengan gen "
    "kandidat utama pada jalur melanogenesis seperti MC1R, MITF, dan KIT (Yang et al., 2013; Xu et al., 2023; Zhang et al., 2025; "
    "Liu et al., 2026). Pendekatan serupa juga efektif untuk sifat pertumbuhan dan produksi. Di sisi lain, metabolomik dan "
    "multi-omik semakin digunakan untuk memahami efisiensi pakan, metabolisme, dan prediksi sifat produksi pada domba "
    "(Hess et al., 2020; Wang et al., 2024). Integrasi genomik–metabolomik pada Domba Temanggung belum pernah dilaporkan, "
    "sehingga menjadi peluang kebaruan yang kuat untuk mendukung penggaluran bibit unggul lokal."
)

add_h("3.2 Rumusan Masalah dan Hipotesis Solusi", 2)
add_para("Rumusan masalah:", bold=True)
for i, q in enumerate([
    "Apakah fenotip Teropong Domba Temanggung memiliki dasar genetik nuklear yang terdeteksi melalui GWAS dan signatures of selection?",
    "Bagaimana struktur asosiasi genomik terhadap sifat produksi (bobot badan dan morfometri) serta parameter reproduksi pada populasi Domba Temanggung?",
    "Apakah terdapat profil metabolit diferensial yang berkorelasi dengan fenotip Teropong, performa produksi, dan reproduksi?",
    "Bagaimana integrasi data genomik–metabolomik dapat menghasilkan panel marka dan model seleksi yang operasional untuk penggaluran bibit unggul?",
], 1):
    add_para(f"{i}. {q}")
add_para(
    "Hipotesis solusi: Fenotip Teropong dikendalikan oleh lokus genetik terkait pigmen yang dapat dipetakan genomik; sifat "
    "produksi dan reproduksi memiliki komponen genetik aditif yang terukur; profil metabolom memberikan biomarker pelengkap; "
    "integrasi keduanya menghasilkan basis saintifik penggaluran Domba Temanggung Teropong sebagai bibit unggul pendukung ketahanan pangan."
)

add_h("3.3 State of the Art dan Kebaruan (Novelty)", 2)
add_para("State of the art:", bold=True)
add_para("Karakterisasi Domba Temanggung masih pada tahap fenotipik lapangan dan mtDNA terbatas (Ridlo et al., 2026). GWAS/signatures of selection efektif mengidentifikasi lokus warna bulu dan sifat produksi pada domba di berbagai negara (Xu et al., 2023; Zhang et al., 2025; Liu et al., 2026). Metabolomik dan multi-omik meningkatkan pemahaman efisiensi dan performa produksi ruminansia kecil (Hess et al., 2020; Wang et al., 2024; Martinez et al., 2023). Di Indonesia, penggaluran/rumpun lokal umumnya berbasis morfometri dan penetapan administratif, belum berbasis panel marka multi-omik (Gunawan et al., 2008; Hidayat et al., 2024).")
add_para("Kebaruan riset:", bold=True)
add_para("(1) Pertama kali mengintegrasikan genomik nuklear (GWAS + signatures of selection) dan metabolomik pada Domba Temanggung; (2) memperluas fokus dari identitas fenotip Teropong menuju paket sifat Teropong + sifat produksi + reproduksi sebagai kriteria penggaluran; (3) menghasilkan panel marka dan model seleksi yang dapat dihilirisasi sebagai instrumen bibit unggul lokal untuk ketahanan pangan.")

add_h("3.4 Tujuan dan Sasaran", 2)
add_para("Tujuan umum: Menghasilkan basis saintifik genomik–metabolomik untuk penggaluran Domba Temanggung fenotip Teropong sebagai bibit unggul pendukung ketahanan pangan.")
add_para("Tujuan khusus bertahap: (1) Tahun 1 menetapkan standar dan basis genomik–metabolomik fenotip Teropong serta sifat produksi; (2) Tahun 2 memetakan basis genomik–metabolomik sifat reproduksi dan menyusun model seleksi reproduksi; (3) Tahun 3 mengintegrasikan paket sifat Teropong + produksi + reproduksi menjadi panel marka, standar galur, dan naskah usulan galur yang siap diajukan.")
add_para("Sasaran: Tahun 1 tersedia database Teropong–produksi dan draf standar identitas–produksi; Tahun 2 tersedia parameter serta model seleksi reproduksi; Tahun 3 tersedia naskah usulan galur Domba Temanggung berbasis bukti multi-omik.")

add_h("4. Kerangka Berpikir dan Nilai Strategis", 1)
add_h("4.1 Kerangka Berpikir", 2)
add_para(
    "Riset disusun berjenjang menuju usulan galur: Tahun 1 mengolah fenotip Teropong dan sifat produksi melalui GWAS, "
    "signatures of selection, dan metabolomik menjadi standar identitas–produksi; Tahun 2 memfokuskan pencatatan serta "
    "analisis genomik–metabolomik sifat reproduksi menjadi model seleksi reproduksi; Tahun 3 mengintegrasikan ketiga paket "
    "sifat menjadi panel marka final, standar galur, dan naskah usulan galur."
)
add_para(
    "Secara konseptual, fenotip Teropong berfungsi sebagai breed/line identity marker, sementara sifat produksi dan reproduksi "
    "menjadi economic merit. Integrasi bertahap memastikan pada akhir Tahun 3 bukti ilmiah sudah memadai untuk usulan galur."
)
add_h("4.2 Nilai Strategis", 2)
add_para(
    "Nilai strategis meliputi: penguatan kedaulatan pangan berbasis plasma nutfah lokal; konservasi sekaligus pemanfaatan ekonomi "
    "galur lokal; dukungan evidence-based policy bagi penetapan galur/rumpun; hilirisasi panel marka dan protokol seleksi; serta "
    "peningkatan daya saing peternak rakyat Temanggung melalui bibit bermutu dan identitas produk lokal."
)

add_h("5. Peta Jalan", 1)
add_para("Durasi usulan: 3 tahun (2026–2028), dengan fokus sifat berjenjang.")
add_para("Tahun 1 — Fenotip Teropong & sifat produksi: standardisasi skor Teropong; bobot & morfometri; genotyping; GWAS + signatures of selection; metabolomik identitas/produksi; KTI1 under review; draft KI; model identitas–produksi.")
add_para("Tahun 2 — Sifat reproduksi: pencatatan litter size, interval beranak, kebuntingan, survival; GWAS & metabolomik reproduksi; model seleksi reproduksi; KTI2 under review; KI terdaftar.")
add_para("Tahun 3 — Integrasi & usulan galur: integrasi multi-omik; validasi panel marka; finalisasi standar galur; naskah usulan galur siap diajukan; KTI3 under review; KI tambahan.")
add_para("Produk akhir: Naskah usulan galur Domba Temanggung berbasis bukti genomik–metabolomik pada fenotip Teropong, sifat produksi, dan reproduksi.")

add_h("6. Metodologi", 1)
add_h("6.1 Desain Riset", 2)
add_para("Desain bertahap: Tahun 1 case–control Teropong vs non-Teropong dan kuantitatif produksi; Tahun 2 kuantitatif reproduksi; Tahun 3 integrasi multi-omik untuk standar dan usulan galur.")
add_h("6.2 Lokasi dan Populasi", 2)
add_para("Lokasi: Kabupaten Temanggung. Mitra: DKPPP dan DSL. Genomik kumulatif 100–150 ekor (T1 genotyping inti ±100–120 ekor; T2 pendalaman betina produktif; T3 validasi panel terbatas). Metabolomik 30–50 ekor total (T1 Teropong/produksi; T2 status reproduksi).")
add_h("6.3 Tahapan Pekerjaan", 2)
add_para("A. Persiapan: klirens etik, MoU mitra, SOP, DMP.")
add_para("B. Tahun 1 (Teropong & produksi): skor Teropong; bobot & morfometri; sampling DNA/serum; genotyping; GWAS dan signatures of selection Teropong–produksi; metabolomik tahap I; draf standar identitas–produksi, KTI1, draft KI.")
add_para("C. Tahun 2 (Reproduksi): pencatatan litter size, jarak beranak, kebuntingan, survival (+ USG bila memungkinkan); GWAS reproduksi; metabolomik tahap II; model seleksi reproduksi; KTI2 dan pendaftaran KI.")
add_para("D. Tahun 3 (Usulan galur): integrasi gen–metabolit–fenotip lintas sifat; panel marka final; validasi terbatas; finalisasi standar galur; penyusunan naskah usulan galur bersama mitra; KTI3 dan KI tambahan.")
add_h("6.4 Analisis Statistik dan Manajemen Risiko", 2)
add_para("Analisis menggunakan R/PLINK/GCTA/GEMMA/VCFtools dan perangkat metabolomik standar. Model linier campuran memasukkan efek tetap (sex, age/parity class, farm) dan efek random genetik aditif. Mitigasi mencakup form pencatatan reproduksi sejak T1, sampling menyebar, standardisasi waktu sampling metabolom, serta FGD bertahap dengan pemda sejak T2 untuk kesiapan naskah usulan galur.")

add_h("7. Jangka Waktu Pelaksanaan Riset", 1)
add_para("Riset diusulkan selama 3 (tiga) tahun periode pendanaan (2026–2028). Tahun 3 diarahkan agar hasil riset siap digunakan untuk usulan galur.")

add_h("8. Keluaran dan Indikator Kinerja", 1)
add_para("Tahun 1 (Teropong & produksi): 1 KTI under review; draft KI panel Teropong–produksi; model/standar identitas–produksi; database ≥100 ekor + biobank.")
add_para("Tahun 2 (Reproduksi): 1 KTI accepted (akumulasi) + 1 KTI baru under review; KI terdaftar; model seleksi reproduksi.")
add_para("Tahun 3 (Usulan galur): minimal 1 KTI published + 1 accepted + 1 under review; KI tambahan terdaftar; naskah usulan galur siap diajukan bersama mitra ke otoritas terkait.")

add_h("9. Jadwal Kegiatan Riset", 1)
add_para("Tahun 1: persiapan & etik; fenotip Teropong & produksi; sampling; genotyping; GWAS/selection signatures Teropong–produksi; metabolomik T1; standar & KTI1/draft KI.")
add_para("Tahun 2: pencatatan sifat reproduksi; GWAS reproduksi; metabolomik T2; model seleksi reproduksi; KTI2 & pendaftaran KI.")
add_para("Tahun 3: integrasi multi-omik; validasi panel marka; penyusunan naskah usulan galur & FGD mitra; publikasi, diseminasi, monev.")

add_h("10. Anggaran per Tahun (Ringkasan RAB)", 1)
add_para("Pendanaan berbasis kesetaraan output: 1 keluaran ≈ Rp150–200 juta; paket 1 KTI + 1 KI + 1 model/implementasi ≈ Rp300–450 juta/tahun. Belanja non-personil mengikuti fokus sifat tiap tahun.")
add_para("Proporsi Juknis RIIM: Personil ≤25%; Non-personil ≥70%; Tidak langsung ≤5%.")
add_para("Tahun 1 Rp400.000.000 (Personil 90 jt; Non-personil 290 jt; Tidak langsung 20 jt) — genotyping/fenotip Teropong–produksi; metabolomik T1; KTI + draft KI + model identitas–produksi.")
add_para("Tahun 2 Rp450.000.000 (Personil 100 jt; Non-personil 328 jt; Tidak langsung 22 jt) — pencatatan/USG reproduksi; GWAS & metabolomik reproduksi; KTI + KI terdaftar + model seleksi reproduksi.")
add_para("Tahun 3 Rp400.000.000 (Personil 90 jt; Non-personil 290 jt; Tidak langsung 20 jt) — validasi panel; standar galur; naskah usulan galur; KTI + KI tambahan + implementasi usulan galur.")
add_para("Total usulan Rp1.250.000.000. RAB detail per item dilampirkan pada lembar RAB resmi sistem Pendanaan Risnov.")

add_h("11. Daftar Pustaka", 1)
refs = [
    "Gunawan, A., Jamal, & Sumantri, C. (2008). Pendugaan nilai heritabilitas bobot lahir dan bobot sapih Domba Garut tipe laga. Media Peternakan.",
    "Hess, M. K., Rowe, S. J., Van Stijn, T. C., Henry, H. M., Hickey, S. M., Brauning, R., McCulloch, A. F., Hickey, A. J., Dodds, K. G., & Jonker, A. (2020). Genomic predictions for enteric methane production are improved by metabolome and microbiome data in sheep (Ovis aries). Journal of Animal Science. https://doi.org/10.1093/jas/skaa331",
    "Hidayat, R., et al. (2024). Phenotype variation, correlation, and regression of morphology characteristics in rams and ewes Garut Indonesian local sheep. International Journal of Agriculture, Environment and Bioresearch. https://doi.org/10.35410/ijaeb.2024.5917",
    "Ibrahim, A., Budisatria, I. G. S., Widayanti, R., & Artama, W. T. (2021). The genetic profiles and maternal origin of local sheep breeds on Java Island (Indonesia) based on complete mitochondrial DNA D-loop sequences. Journal of Genetic Engineering and Biotechnology.",
    "Ibrahim, A., Baliarti, E., Budisatria, I. G. S., Artama, W. T., Widayanti, R., Maharani, D., et al. (2023). Genetic diversity and relationship among Indonesian local sheep breeds on Java Island based on mitochondrial cytochrome b gene sequences. Journal of Genetic Engineering and Biotechnology, 21(1), 34. https://doi.org/10.1186/s43141-023-00491-z",
    "Liu, Y., et al. (2026). Whole-genome resequencing revealed genetic diversity and the haplotype containing MC1R associated with black coat color in Liangshan sheep. BMC Genomics. https://doi.org/10.1186/s12864-026-12966-7",
    "Martinez, B., et al. (2023). Combining host and rumen metagenome profiling for selection in sheep: Prediction of methane, feed efficiency, production, and health traits. Genetics Selection Evolution, 55, 53. https://doi.org/10.1186/s12711-023-00822-1",
    "Purwantini, D., et al. (2024). Correlation between morphometric traits and body weight in fat-tailed sheep. BIO Web of Conferences, 88, 00026. https://doi.org/10.1051/bioconf/20248800026",
    "Ridlo, M. R., Ariyanti, F., As Syauqi, A. G., Sari, A. P. Z. N. L., Haryanto, A., Assihhah, N., Panjono, Ermawati, D., Atmoko, B. A., & Ibrahim, A. (2026). Preliminary study on genetic diversity and phylogenetic relationships of Teropong Temanggung sheep based on mitochondrial cytochrome b sequences. Brazilian Journal of Biology, 86, e301893. https://doi.org/10.1590/1519-6984.301893",
    "Wang, X., et al. (2024). Multi-omics revealed the mechanism of feed efficiency in sheep by the combined action of the host and rumen microbiota. https://pmc.ncbi.nlm.nih.gov/articles/PMC11406083/",
    "Xu, S., et al. (2023). Convergent changes in melanocortin receptor 1 gene are associated with black-headed coat color in sheep. Journal of Animal Science, 101, skad084. https://doi.org/10.1093/jas/skad084",
    "Yang, G. L., Fu, D. L., Lang, X., Wang, Y. T., Cheng, S. R., Fang, S. L., & Luo, Y. Z. (2013). Mutations in MC1R gene determine black coat color phenotype in Chinese sheep. The Scientific World Journal, 2013, 675382. https://doi.org/10.1155/2013/675382",
    "Zhang, L., et al. (2025). Revealing the genetic basis of coat color in Tibetan sheep through selective sweep and transcriptomic analyses. Frontiers in Veterinary Science, 12, 1711294. https://doi.org/10.3389/fvets.2025.1711294",
]
for r in refs:
    p = doc.add_paragraph(r)
    p.paragraph_format.first_line_indent = Cm(-0.75)
    p.paragraph_format.left_indent = Cm(0.75)
    for run in p.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(11)

doc.add_page_break()
add_h("Lampiran A. Kompetensi Tim Periset", 1)
add_para("Ketua: kepakaran genetika ternak/pemuliaan — koordinator riset, desain genomik, interpretasi GWAS & selection signatures.")
add_para("Anggota 1: metabolomik/biologi molekuler — analisis metabolomik & integrasi multi-omik.")
add_para("Anggota 2: produksi & reproduksi ruminansia kecil — fenotip produksi, reproduksi, sampling lapangan.")
add_para("Anggota 3: bioinformatika genomik — pipeline QC, GWAS, selection scan, kurasi data.")
add_para("Catatan: lengkapi nama, pendidikan, URL SCOPUS, dan SK tim sebelum unggah.")

add_h("Lampiran B. Catatan Administrasi Unggah", 1)
add_para("Lengkapi identitas periset, pengesahan Kepala Unit Kerja, kesediaan mitra, RAB rinci (termasuk pajak), DMP, biodata, dan setelah lolos seleksi ajukan Klirens Etik Riset sebelum sampling. Unggah melalui https://pendanaan-risnov.brin.go.id/")

out = Path("/workspace/proposal-riim/PROPOSAL_RIIM_Domba_Temanggung.docx")
doc.save(out)
print(f"Saved {out}")
