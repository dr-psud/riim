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
add_center("Genomik dan Metabolomik Fenotip Teropong, Sifat Produksi,", True, 12)
add_center("serta Proliferasi Domba Temanggung sebagai Basis", True, 12)
add_center("Penggaluran Bibit Unggul untuk Ketahanan Pangan", True, 12)
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
add_para("2. Judul Proposal\t: Genomik dan Metabolomik Fenotip Teropong, Sifat Produksi, serta Proliferasi Domba Temanggung sebagai Basis Penggaluran Bibit Unggul untuk Ketahanan Pangan", justify=False)
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
add_para("Tahun 1: 1 KTI jurnal internasional minimal Q3 under review; draft KI/paten sederhana; dataset fenotip–biobank; draft standar fenotip galur.", justify=False)
add_para("Tahun 2: KTI T1 accepted + 1 KTI baru under review; paten sederhana terdaftar; profil metabolom diferensial.", justify=False)
add_para("Tahun 3: minimal 1 KTI published + 1 accepted + 1 under review; KI tambahan terdaftar; rekomendasi teknis penggaluran bibit unggul.", justify=False)
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
add_para("Genomik dan Metabolomik Fenotip Teropong, Sifat Produksi, serta Proliferasi Domba Temanggung sebagai Basis Penggaluran Bibit Unggul untuk Ketahanan Pangan")

add_h("2. Abstrak", 1)
add_para(
    "Domba Temanggung merupakan plasma nutfah lokal Jawa Tengah dengan ciri khas fenotip teropong "
    "(warna bulu hitam melingkar di sekitar mata), yang saat ini diusulkan sebagai identitas galur/rumpun lokal unggul. "
    "Namun, diferensiasi galur Teropong versus non-Teropong masih bertumpu pada deskripsi fenotipik visual dan data genetik "
    "maternal terbatas (mtDNA), sehingga belum tersedia bukti genomik nuklear maupun profil metabolomik yang mendukung "
    "penggaluran berbasis sains. Riset ini mengintegrasikan pendekatan genomik (GWAS dan signatures of selection) serta "
    "metabolomik untuk mengarakterisasi fenotip Teropong, sifat produksi (bobot badan dan morfometri), serta proliferasi "
    "Domba Temanggung sebagai basis penggaluran bibit unggul. Populasi studi dirancang case–control dan kuantitatif pada "
    "Domba Temanggung di Kabupaten Temanggung dengan mitra DKPPP dan komunitas peternak. Kegiatan meliputi standardisasi "
    "fenotip, genotyping/whole-genome sequencing, asosiasi genom–fenotip, deteksi jejak seleksi, profil metabolom "
    "serum/jaringan terkait, serta integrasi multi-omik untuk menghasilkan panel marka dan model seleksi. Luaran utama "
    "mencakup publikasi jurnal internasional bereputasi, kekayaan intelektual (panel marka/paten sederhana), dan rekomendasi "
    "teknis penggaluran yang dapat memperkuat penyediaan bibit lokal unggul dan mendukung ketahanan pangan berbasis protein hewani."
)
add_para("Kata kunci: Domba Temanggung; fenotip teropong; GWAS; signatures of selection; metabolomik; sifat produksi; proliferasi; penggaluran bibit; ketahanan pangan", bold=False)
add_para("Keywords: Temanggung sheep; periocular black phenotype; GWAS; signatures of selection; metabolomics; production traits; proliferation; breed/line development; food security")

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
    "keunggulan produksi dan proliferasi. Namun, informasi genetik Domba Temanggung masih sangat terbatas. Studi awal berbasis "
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
    "Bagaimana struktur asosiasi genomik terhadap sifat produksi (bobot badan dan morfometri) serta parameter proliferasi pada populasi Domba Temanggung?",
    "Apakah terdapat profil metabolit diferensial yang berkorelasi dengan fenotip Teropong, performa produksi, dan proliferasi?",
    "Bagaimana integrasi data genomik–metabolomik dapat menghasilkan panel marka dan model seleksi yang operasional untuk penggaluran bibit unggul?",
], 1):
    add_para(f"{i}. {q}")
add_para(
    "Hipotesis solusi: Fenotip Teropong dikendalikan oleh lokus genetik terkait pigmen yang dapat dipetakan genomik; sifat "
    "produksi dan proliferasi memiliki komponen genetik aditif yang terukur; profil metabolom memberikan biomarker pelengkap; "
    "integrasi keduanya menghasilkan basis saintifik penggaluran Domba Temanggung Teropong sebagai bibit unggul pendukung ketahanan pangan."
)

add_h("3.3 State of the Art dan Kebaruan (Novelty)", 2)
add_para("State of the art:", bold=True)
add_para("Karakterisasi Domba Temanggung masih pada tahap fenotipik lapangan dan mtDNA terbatas (Ridlo et al., 2026). GWAS/signatures of selection efektif mengidentifikasi lokus warna bulu dan sifat produksi pada domba di berbagai negara (Xu et al., 2023; Zhang et al., 2025; Liu et al., 2026). Metabolomik dan multi-omik meningkatkan pemahaman efisiensi dan performa produksi ruminansia kecil (Hess et al., 2020; Wang et al., 2024; Martinez et al., 2023). Di Indonesia, penggaluran/rumpun lokal umumnya berbasis morfometri dan penetapan administratif, belum berbasis panel marka multi-omik (Gunawan et al., 2008; Hidayat et al., 2024).")
add_para("Kebaruan riset:", bold=True)
add_para("(1) Pertama kali mengintegrasikan genomik nuklear (GWAS + signatures of selection) dan metabolomik pada Domba Temanggung; (2) memperluas fokus dari identitas fenotip Teropong menuju paket sifat Teropong + sifat produksi + proliferasi sebagai kriteria penggaluran; (3) menghasilkan panel marka dan model seleksi yang dapat dihilirisasi sebagai instrumen bibit unggul lokal untuk ketahanan pangan.")

add_h("3.4 Tujuan dan Sasaran", 2)
add_para("Tujuan umum: Menghasilkan basis saintifik genomik–metabolomik untuk penggaluran Domba Temanggung fenotip Teropong sebagai bibit unggul pendukung ketahanan pangan.")
add_para("Tujuan khusus: (1) menetapkan standar fenotipik Teropong, sifat produksi, dan proliferasi; (2) mengidentifikasi lokus/kandidat gen terkait fenotip Teropong melalui GWAS dan signatures of selection; (3) memetakan asosiasi genomik sifat bobot badan, morfometri, dan proliferasi; (4) mengarakterisasi profil metabolom terkait kelompok fenotip/performa; (5) menyusun panel marka dan model seleksi multi-omik; (6) menyusun rekomendasi teknis penggaluran bersama mitra daerah.")
add_para("Sasaran: tersedianya data primer fenotipik–genomik–metabolomik terkurasi; kandidat marka genetik dan biomarker metabolit; draf standar galur Teropong dan model seleksi bibit; serta meningkatnya kesiapan hilirisasi menuju penetapan galur/rumpun dan sistem perbibitan lokal.")

add_h("4. Kerangka Berpikir dan Nilai Strategis", 1)
add_h("4.1 Kerangka Berpikir", 2)
add_para(
    "Riset bertolak dari kebutuhan membedakan dan mengembangkan galur Domba Temanggung Teropong secara ilmiah. Input berupa "
    "populasi Teropong vs non-Teropong beserta data sifat produksi dan proliferasi diproses melalui standardisasi fenotip, "
    "analisis genomik (struktur populasi, GWAS, signatures of selection: FST, π-ratio, iHS/XP-EHH, ROH), analisis metabolomik, "
    "serta integrasi multi-omik (gen → metabolit → fenotip). Output berupa lokus kandidat, panel marka, biomarker, model seleksi, "
    "dan rekomendasi penggaluran. Outcome yang diharapkan adalah penguatan sistem bibit lokal unggul yang meningkatkan "
    "produktivitas peternak dan berkontribusi pada penyediaan protein hewani nasional."
)
add_para(
    "Secara konseptual, fenotip Teropong berfungsi sebagai breed/line identity marker, sementara sifat produksi dan proliferasi "
    "menjadi economic merit. Genomik menjelaskan heritabilitas dan lokus kunci; metabolomik menjelaskan status fisiologis/metabolik "
    "yang relevan dengan performa. Integrasi keduanya mengurangi risiko seleksi hanya berbasis penampilan visual."
)
add_h("4.2 Nilai Strategis", 2)
add_para(
    "Nilai strategis meliputi: penguatan kedaulatan pangan berbasis plasma nutfah lokal; konservasi sekaligus pemanfaatan ekonomi "
    "galur lokal; dukungan evidence-based policy bagi penetapan galur/rumpun; hilirisasi panel marka dan protokol seleksi; serta "
    "peningkatan daya saing peternak rakyat Temanggung melalui bibit bermutu dan identitas produk lokal."
)

add_h("5. Peta Jalan", 1)
add_para("Durasi usulan: 3 tahun (2026–2028).")
add_para("Tahun 1 — Fondasi fenotip & data omik: standar fenotip; koleksi sampel ≥100 ekor; genotyping tahap I; metabolom subset; draft manuskrip 1; draft KI.")
add_para("Tahun 2 — Analisis genomik–metabolomik mendalam: hasil GWAS & selection signatures; profil metabolom diferensial; model asosiasi awal; KTI T1 accepted; KTI baru under review; paten sederhana terdaftar.")
add_para("Tahun 3 — Integrasi, validasi & hilirisasi: panel marka tervalidasi; model seleksi multi-omik; rekomendasi penggaluran; publikasi published; KI tambahan; diseminasi mitra.")
add_para("Produk akhir: Sistem penggaluran bibit Domba Temanggung berbasis genomik–metabolomik (standar fenotip + panel marka + model seleksi + rekomendasi kebijakan/teknis) untuk mendukung ketahanan pangan.")

add_h("6. Metodologi", 1)
add_h("6.1 Desain Riset", 2)
add_para("Desain kombinasi case–control untuk fenotip Teropong vs non-Teropong; asosiasi kuantitatif untuk bobot badan, morfometri, dan parameter proliferasi; serta multi-omik integratif (genomik + metabolomik).")
add_h("6.2 Lokasi dan Populasi", 2)
add_para("Lokasi: Kabupaten Temanggung, Jawa Tengah. Mitra: DKPPP Temanggung dan komunitas DSL/peternak. Target sampel genomik 100–150 ekor, seimbang Teropong : non-Teropong, dengan penambahan bertahap sesuai tahun pendanaan. Subset metabolomik 30–50 ekor terpilih berdasarkan kontras fenotip/performa, distandardisasi sampling.")
add_h("6.3 Tahapan Pekerjaan", 2)
add_para("A. Klirens etik dan perizinan sebelum sampling hewan sesuai Peraturan BRIN.")
add_para("B. Standardisasi fenotip: skor Teropong (kehadiran/intensitas/simetri lingkar hitam periokular + foto standar); sifat produksi (bobot badan dan morfometri: tinggi pundak, panjang badan, lingkar dada, dll.); proliferasi (litter size, jarak beranak, angka kebuntingan, survival anak, dan/atau indikator fisiologis terkait); kovariat umur, jenis kelamin, tipe kelahiran, lokasi/peternak, manajemen.")
add_para("C. Sampling biologis: darah EDTA untuk DNA; serum/plasma untuk metabolomik; subset jaringan/swab kulit periokular opsional; cold-chain dan kurasi DMP.")
add_para("D. Analisis genomik: ekstraksi DNA & QC; genotyping SNP chip densitas menengah–tinggi dan/atau low-coverage WGS + imputation; QC genotip; GWAS model mixed linear (GRM/kinship) untuk sifat biner dan kuantitatif; signatures of selection (FST, π-ratio, iHS, XP-EHH, ROH); anotasi gen kandidat dan enrichment jalur.")
add_para("E. Analisis metabolomik: LC-MS/MS untargeted (± targeted); PCA/OPLS-DA; identifikasi metabolit dan pathway enrichment; korelasi dengan fenotip/genotipe (mGWAS/mQTL pada subset bila memadai).")
add_para("F. Integrasi multi-omik: network gen–metabolit–fenotip; panel marka SNP prioritas; model seleksi/indeks genomic breeding value sederhana; validasi internal dan validasi terbatas tahun akhir.")
add_para("G. Hilirisasi: draf standar galur, FGD mitra, rekomendasi teknis, pengajuan KI, dan publikasi.")
add_h("6.4 Analisis Statistik dan Manajemen Risiko", 2)
add_para("Analisis menggunakan R/PLINK/GCTA/GEMMA/VCFtools dan perangkat metabolomik standar. Model linier campuran memasukkan efek tetap (sex, age class, farm) dan efek random genetik aditif. Mitigasi risiko mencakup standardisasi protokol sampling, koreksi farm/batch, desain sampling menyebar, power simulation, serta standardisasi waktu sampling metabolom.")

add_h("7. Jangka Waktu Pelaksanaan Riset", 1)
add_para("Riset diusulkan selama 3 (tiga) tahun periode pendanaan (2026–2028), dengan evaluasi tahunan sesuai ketentuan RIIM Kompetisi.")

add_h("8. Keluaran dan Indikator Kinerja", 1)
add_para("Tahun 1: 1 publikasi minimal Q3 under review; draft paten sederhana/KI; database fenotip ≥100 ekor + biobank; draft standar fenotip galur.")
add_para("Tahun 2: 1 KTI accepted (akumulasi) + 1 KTI baru under review; paten sederhana terdaftar; laporan profil metabolom diferensial.")
add_para("Tahun 3: minimal 1 KTI published + 1 accepted + 1 under review; KI tambahan terdaftar; dokumen rekomendasi teknis penggaluran bibit unggul untuk mitra/pemda.")

add_h("9. Jadwal Kegiatan Riset", 1)
add_para("Tahun 1: persiapan & etik; survei & sampling fenotip; koleksi sampel; ekstraksi DNA & genotyping; metabolomik tahap I; analisis struktur populasi dan manuskrip 1.")
add_para("Tahun 2: GWAS fenotip Teropong; signatures of selection; GWAS produksi & proliferasi; metabolomik tahap II & integrasi; penyusunan panel marka dan pengajuan KI.")
add_para("Tahun 3: validasi marka & model seleksi; finalisasi standar galur dan FGD mitra; publikasi, diseminasi, monev, dan pelaporan.")

add_h("10. Anggaran per Tahun (Ringkasan RAB)", 1)
add_para("Pendanaan mengikuti prinsip kesetaraan output: 1 keluaran ≈ Rp150–200 juta. Paket tahunan 1 KTI + 1 KI + 1 model/implementasi dianggarkan pada rentang Rp300–450 juta/tahun.")
add_para("Proporsi mengikuti Juknis RIIM: Biaya Langsung Personil ≤25%; Biaya Langsung Non-Personil ≥70%; Biaya Tidak Langsung ≤5%.")
add_para("Tahun 1 Rp400.000.000 (Personil 90 jt; Non-personil 290 jt; Tidak langsung 20 jt) untuk 1 KTI under review + 1 draft KI + 1 draft model.")
add_para("Tahun 2 Rp450.000.000 (Personil 100 jt; Non-personil 328 jt; Tidak langsung 22 jt) untuk 1 KTI baru under review + 1 KI terdaftar + 1 model seleksi.")
add_para("Tahun 3 Rp400.000.000 (Personil 90 jt; Non-personil 290 jt; Tidak langsung 20 jt) untuk 1 KTI baru under review + 1 KI tambahan + 1 implementasi/rekomendasi teknis.")
add_para("Total usulan Rp1.250.000.000. Status accepted/published KTI tahun sebelumnya merupakan akumulasi kewajiban RIIM dan tidak dihitung sebagai keluaran baru untuk penambahan plafon dana. RAB detail per item dilampirkan pada lembar RAB resmi sistem Pendanaan Risnov, termasuk pemisahan harga pokok dan pajak sesuai status PKP/Non-PKP institusi.")

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
add_para("Anggota 2: produksi & reproduksi ruminansia kecil — fenotip produksi, proliferasi, sampling lapangan.")
add_para("Anggota 3: bioinformatika genomik — pipeline QC, GWAS, selection scan, kurasi data.")
add_para("Catatan: lengkapi nama, pendidikan, URL SCOPUS, dan SK tim sebelum unggah.")

add_h("Lampiran B. Catatan Administrasi Unggah", 1)
add_para("Lengkapi identitas periset, pengesahan Kepala Unit Kerja, kesediaan mitra, RAB rinci (termasuk pajak), DMP, biodata, dan setelah lolos seleksi ajukan Klirens Etik Riset sebelum sampling. Unggah melalui https://pendanaan-risnov.brin.go.id/")

out = Path("/workspace/proposal-riim/PROPOSAL_RIIM_Domba_Temanggung.docx")
doc.save(out)
print(f"Saved {out}")
