from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

doc = Document()
section = doc.sections[0]
section.page_height = Inches(11.69)
section.page_width = Inches(8.27)
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1.5)
section.right_margin = Inches(1.5)

def add_heading_custom(doc, text, level=1):
    p = doc.add_heading(level=level)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.color.rgb = RGBColor(0,0,0)
    if level == 1: run.font.size = Pt(14); run.bold = True; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif level == 2: run.font.size = Pt(12); run.bold = True
    else: run.font.size = Pt(11); run.bold = True; run.italic = True
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    return p

def add_center_text(doc, text, bold=False, size_pt=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    if bold: run.bold = True
    run.font.color.rgb = RGBColor(0,0,0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    return p

def add_para_custom(doc, text, bold=False, italic=False, align='justify', indent_first=True, size_pt=11):
    p = doc.add_paragraph()
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'left': p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    else: p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    if bold: run.bold = True
    if italic: run.italic = True
    run.font.color.rgb = RGBColor(0,0,0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    if indent_first: p.paragraph_format.first_line_indent = Inches(0.5)
    else: p.paragraph_format.first_line_indent = Inches(0)
    return p

# ================= COVER =================
add_center_text(doc, "POLA KOMUNIKASI INTERPERSONAL ANAK DENGAN ORANG TUA", bold=True, size_pt=14)
add_center_text(doc, "DALAM PENDAMPINGAN BELAJAR DARI RUMAH (BDR) DI BANTUL", bold=True, size_pt=14)
doc.add_paragraph()
add_center_text(doc, "PROPOSAL PENELITIAN", bold=True, size_pt=14)
doc.add_paragraph()
doc.add_paragraph()
add_center_text(doc, "Ditulis untuk Memenuhi Tugas Ujian Tengah Semester (UTS) Bahasa Indonesia", size_pt=11)
add_center_text(doc, "Program Studi Ilmu Komunikasi", size_pt=11)
doc.add_paragraph()
add_center_text(doc, "Oleh:", size_pt=11)
add_center_text(doc, "SETIA KARTIKA HAPSARI NINGRUM", bold=True, size_pt=12)
add_center_text(doc, "NIM: 202612340001", size_pt=11)
add_center_text(doc, "FAKULTAS ILMU SOSIAL DAN ILMU POLITIK", size_pt=11)
add_center_text(doc, "UNIVERSITAS NEGERI YOGYAKARTA", size_pt=11)
add_center_text(doc, "2026", size_pt=11)
doc.add_page_break()

# ================= KATA PENGANTAR =================
add_heading_custom(doc, "KATA PENGANTAR", 1)
add_para_custom(doc, "Puji syukur penulis panjatkan ke hadirat Tuhan Yang Maha Esa atas segala rahmat dan karunia-Nya sehingga penulis dapat menyelesaikan proposal penelitian yang berjudul \"Pola Komunikasi Interpersonal Anak dengan Orang Tua dalam Pendampingan Belajar dari Rumah (BDR) di Bantul\" dengan baik.")
add_para_custom(doc, "Proposal penelitian ini disusun untuk memenuhi tugas Ujian Tengah Semester (UTS) mata kuliah Bahasa Indonesia pada Program Studi Ilmu Komunikasi, Fakultas Ilmu Sosial dan Ilmu Politik, Universitas Negeri Yogyakarta tahun 2026. Penelitian ini berfokus pada pola komunikasi interpersonal antara anak dan orang tua dalam proses pendampingan belajar dari rumah (BDR), yang menjadi isu penting sejak pandemi COVID-19 dan terus berlanjut hingga saat ini.")
add_para_custom(doc, "Dalam penyusunan proposal ini, penulis menyadari bahwa tanpa bantuan dan dukungan dari berbagai pihak, penelitian ini tidak akan terselesaikan dengan baik. Oleh karena itu, penulis menyampaikan ucapan terima kasih kepada:")
items_thanks = [
    "Bapak/Ibu dosen pengampu mata kuliah Bahasa Indonesia yang telah memberikan bimbingan dan arahan.",
    "Keluarga penulis yang senantiasa memberikan dukungan moral dan materiil.",
    "Teman-teman Program Studi Ilmu Komunikasi yang telah memberikan masukan dan semangat.",
    "Semua pihak yang telah membantu, baik secara langsung maupun tidak langsung."
]
for item in items_thanks:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run("\u2022  " + item)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)
add_para_custom(doc, "Penulis menyadari bahwa proposal ini masih jauh dari sempurna. Oleh karena itu, kritik dan saran yang membangun sangat penulis harapkan demi perbaikan di masa yang akan datang. Akhir kata, penulis berharap semoga proposal ini dapat memberikan manfaat bagi pengembangan ilmu komunikasi, khususnya dalam memahami dinamika komunikasi interpersonal dalam konteks pendidikan di rumah.")
doc.add_paragraph()
add_center_text(doc, "Penulis", bold=False, size_pt=11)
add_center_text(doc, "Setia Kartika Hapsari Ningrum", bold=False, size_pt=11)
doc.add_page_break()

# ================= DAFTAR ISI =================
add_heading_custom(doc, "DAFTAR ISI", 1)
items_isi = [
    ("KATA PENGANTAR", "2"),
    ("DAFTAR ISI", "3"),
    ("DAFTAR TABEL", "4"),
    ("DAFTAR GAMBAR", "4"),
    ("BAB I PENDAHULUAN", "5"),
    ("   A. Latar Belakang Masalah", "5"),
    ("   B. Identifikasi Masalah", "6"),
    ("   C. Batasan Penelitian", "6"),
    ("   D. Rumusan Masalah", "7"),
    ("   E. Tujuan Penelitian", "7"),
    ("   F. Manfaat Penelitian", "8"),
    ("      1. Manfaat Teoritis", "8"),
    ("      2. Manfaat Praktis", "8"),
    ("BAB II KAJIAN PUSTAKA", "9"),
    ("   A. Konsep Komunikasi Interpersonal", "9"),
    ("   B. Pola Komunikasi dalam Keluarga", "10"),
    ("   C. Pendampingan Belajar dari Rumah (BDR)", "11"),
    ("BAB III METODE PENELITIAN", "13"),
    ("   A. Jenis Penelitian", "13"),
    ("   B. Tempat dan Waktu Penelitian", "13"),
    ("   C. Subjek Penelitian", "14"),
    ("   D. Desain Penelitian", "14"),
    ("   E. Teknik Pengumpulan Data", "14"),
    ("   F. Instrumen Penelitian", "15"),
    ("   G. Teknik Analisis Data", "15"),
    ("   H. Indikator Keberhasilan", "16"),
    ("DAFTAR PUSTAKA", "17"),
]
for title, pg in items_isi:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.first_line_indent = Inches(0)
    run = p.add_run(title + "\t" + pg)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)
doc.add_page_break()

# ================= BAB I =================
add_heading_custom(doc, "BAB I", 1)
add_heading_custom(doc, "PENDAHULUAN", 1)

add_heading_custom(doc, "A. Latar Belakang Masalah", 2)
add_para_custom(doc, "Komunikasi interpersonal merupakan proses pertukaran pesan antara dua individu atau lebih yang terjadi secara langsung atau tidak langsung, dengan tujuan saling memahami dan mempengaruhi satu sama lain. Dalam konteks keluarga, komunikasi interpersonal antara anak dan orang tua memiliki peran yang sangat penting dalam membentuk karakter, motivasi belajar, serta kesejahteraan psikologis anak (DeVito, 2016). Sejak pandemi COVID-19, sistem pembelajaran di Indonesia mengalami perubahan signifikan dengan diterapkannya pembelajaran jarak jauh atau Belajar dari Rumah (BDR). Perubahan ini tidak hanya mempengaruhi proses pembelajaran di sekolah, tetapi juga mengubah dinamika interaksi antara anak dan orang tua di rumah.")
add_para_custom(doc, "Di Kabupaten Bantul, Daerah Istimewa Yogyakarta, banyak orang tua yang harus berperan ganda: sebagai pengasuh, pendamping belajar, sekaligus komunikator utama bagi anak-anak mereka selama masa BDR. Fenomena ini menimbulkan berbagai tantangan dalam pola komunikasi interpersonal, seperti perubahan frekuensi interaksi, pergeseran topik percakapan dari aktivitas sehari-hari menjadi fokus pada tugas belajar, serta munculnya konflik akibat tekanan akademik yang dirasakan bersama oleh anak dan orang tua. Oleh karena itu, penelitian ini bertujuan untuk memahami bagaimana pola komunikasi interpersonal anak dengan orang tua berlangsung dalam pendampingan BDR di Bantul.")

add_heading_custom(doc, "B. Identifikasi Masalah", 2)
add_para_custom(doc, "Berdasarkan latar belakang di atas, beberapa masalah yang dapat diidentifikasi antara lain:")
items_id = [
    "Adanya perubahan signifikan dalam frekuensi dan kualitas komunikasi antara anak dan orang tua sejak diterapkannya BDR.",
    "Munculnya ketegangan atau konflik komunikasi akibat peran ganda orang tua sebagai pendamping belajar dan pengasuh.",
    "Kurangnya pemahaman orang tua terhadap strategi komunikasi yang efektif dalam mendukung proses belajar anak dari rumah.",
    "Terbatasnya penelitian yang secara spesifik mengkaji pola komunikasi interpersonal anak-orang tua dalam konteks BDR di wilayah Bantul."
]
for item in items_id:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run("\u2022  " + item)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)

add_heading_custom(doc, "C. Batasan Penelitian", 2)
add_para_custom(doc, "Penelitian ini memiliki batasan sebagai berikut:")
batasan_items = [
    "Fokus penelitian terbatas pada pola komunikasi interpersonal anak dengan orang tua dalam konteks pendampingan BDR.",
    "Lokasi penelitian dibatasi di wilayah Kabupaten Bantul, Daerah Istimewa Yogyakarta.",
    "Subjek penelitian adalah anak usia sekolah dasar (SD) hingga sekolah menengah pertama (SMP) bersama orang tua mereka.",
    "Waktu pengambilan data dibatasi pada semester genap tahun akademik 2025/2026.",
    "Penelitian ini tidak mengkaji aspek pembelajaran daring secara teknis, melainkan hanya pada dinamika komunikasinya."
]
for item in batasan_items:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run("\u2022  " + item)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)

add_heading_custom(doc, "D. Rumusan Masalah", 2)
add_para_custom(doc, "Berdasarkan identifikasi masalah di atas, rumusan masalah dalam penelitian ini adalah:")
rumusan_items = [
    "Bagaimana pola komunikasi interpersonal anak dengan orang tua dalam pendampingan BDR di Bantul?",
    "Faktor apa saja yang mempengaruhi efektivitas komunikasi interpersonal anak-orang tua selama proses BDR?",
    "Bagaimana dampak pola komunikasi tersebut terhadap motivasi belajar dan kesejahteraan psikologis anak?"
]
for item in rumusan_items:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run("\u2022  " + item)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)

add_heading_custom(doc, "E. Tujuan Penelitian", 2)
add_para_custom(doc, "Tujuan penelitian ini antara lain:")
tujuan_items = [
    "Mendeskripsikan pola komunikasi interpersonal anak dengan orang tua dalam pendampingan BDR di Bantul.",
    "Mengidentifikasi faktor-faktor yang mempengaruhi efektivitas komunikasi interpersonal selama proses BDR.",
    "Menganalisis dampak pola komunikasi tersebut terhadap motivasi belajar dan kesejahteraan psikologis anak."
]
for item in tujuan_items:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run("\u2022  " + item)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)

add_heading_custom(doc, "F. Manfaat Penelitian", 2)
add_heading_custom(doc, "1. Manfaat Teoritis", 3)
add_para_custom(doc, "Penelitian ini diharapkan dapat memberikan kontribusi teoretis dalam pengembangan ilmu komunikasi, khususnya pada bidang komunikasi interpersonal dan komunikasi keluarga. Hasil penelitian dapat memperkaya literatur mengenai dinamika komunikasi dalam konteks krisis sosial seperti pandemi, serta membuka ruang diskusi tentang peran orang tua sebagai agen komunikasi dalam proses pendidikan informal.")
add_heading_custom(doc, "2. Manfaat Praktis", 3)
add_para_custom(doc, "Secara praktis, penelitian ini dapat memberikan rekomendasi bagi orang tua, pendidik, dan pembuat kebijakan dalam menyusun program pendampingan BDR yang lebih efektif. Selain itu, hasil penelitian dapat menjadi acuan bagi konselor keluarga dalam merancang intervensi komunikasi yang mendukung kesehatan mental anak selama pembelajaran di rumah.")
doc.add_page_break()

# ================= BAB II =================
add_heading_custom(doc, "BAB II", 1)
add_heading_custom(doc, "KAJIAN PUSTAKA", 1)

add_heading_custom(doc, "A. Konsep Komunikasi Interpersonal", 2)
add_para_custom(doc, "Komunikasi interpersonal menurut DeVito (2016) adalah proses pertukaran pesan antara dua individu atau lebih yang terjadi dalam konteks sosial tertentu, dengan tujuan saling memahami dan mempengaruhi satu sama lain. Dalam komunikasi interpersonal, terdapat beberapa elemen utama: pengirim (sender), pesan (message), saluran (channel), penerima (receiver), umpan balik (feedback), serta konteks fisik, sosial, dan psikologis. Efektivitas komunikasi interpersonal sangat bergantung pada kemampuan individu untuk menyampaikan pesan secara jelas, mendengarkan secara aktif, serta memberikan umpan balik yang konstruktif.")
add_para_custom(doc, "Dalam konteks keluarga, komunikasi interpersonal memiliki karakteristik unik karena melibatkan hubungan emosional yang mendalam, sejarah interaksi yang panjang, serta peran sosial yang saling terkait. Menurut Koerner dan Fitzpatrick (2002), komunikasi dalam keluarga dapat dikategorikan menjadi beberapa pola: pluralistik, protektif, laissez-faire, dan konsensual. Setiap pola memiliki implikasi berbeda terhadap perkembangan anak, termasuk dalam hal motivasi belajar, rasa percaya diri, serta kemampuan mengatasi stres.")

add_heading_custom(doc, "B. Pola Komunikasi dalam Keluarga", 2)
add_para_custom(doc, "Pola komunikasi dalam keluarga tidak hanya mencakup isi pesan, tetapi juga frekuensi, durasi, serta kualitas interaksi. Penelitian oleh Vangelisti (2012) menunjukkan bahwa keluarga dengan pola komunikasi terbuka (open communication) cenderung memiliki anak yang lebih resilien dan memiliki hubungan yang lebih sehat dengan orang tua. Sebaliknya, pola komunikasi tertutup atau otoriter seringkali menimbulkan ketegangan emosional dan menurunkan motivasi anak dalam menjalani aktivitas belajar.")
add_para_custom(doc, "Dalam konteks BDR, pola komunikasi keluarga mengalami perubahan signifikan. Orang tua yang sebelumnya hanya berperan sebagai pengasuh, kini juga harus berfungsi sebagai fasilitator belajar. Perubahan peran ini dapat memicu konflik komunikasi jika tidak diimbangi dengan pemahaman terhadap strategi komunikasi yang adaptif. Oleh karena itu, penting untuk memahami bagaimana keluarga dapat mengembangkan pola komunikasi yang mendukung proses pembelajaran anak di rumah.")

add_heading_custom(doc, "C. Pendampingan Belajar dari Rumah (BDR)", 2)
add_para_custom(doc, "Belajar dari Rumah (BDR) adalah sistem pembelajaran yang dilakukan di luar lingkungan sekolah, dengan menggunakan media teknologi atau pendekatan mandiri. Konsep ini menjadi populer sejak pandemi COVID-19, ketika sekolah-sekolah di seluruh dunia, termasuk Indonesia, terpaksa menutup akses pembelajaran tatap muka. Di Indonesia, Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi (Kemendikbudristek) mengeluarkan kebijakan pembelajaran jarak jauh yang kemudian berkembang menjadi model BDR yang lebih fleksibel.")
add_para_custom(doc, "Pendampingan BDR oleh orang tua bukan hanya sekadar membantu anak menyelesaikan tugas sekolah, melainkan juga mencakup pengawasan emosional, pengaturan jadwal belajar, serta komunikasi yang mendukung motivasi anak. Menurut penelitian oleh Epstein (2018), keterlibatan orang tua dalam proses belajar anak memiliki korelasi positif terhadap prestasi akademik dan kesejahteraan psikologis anak. Namun, keterlibatan yang tidak seimbang - misalnya terlalu mengontrol atau kurang mendukung - dapat menimbulkan efek sebaliknya.")
add_para_custom(doc, "Di Kabupaten Bantul, yang memiliki karakteristik masyarakat semi-urban dengan akses teknologi yang bervariasi, pendampingan BDR menjadi tantangan tersendiri. Beberapa orang tua mungkin belum terbiasa dengan teknologi pembelajaran digital, sementara anak-anak mereka justru lebih cepat beradaptasi. Perbedaan ini dapat memicu kesenjangan komunikasi yang perlu dipahami lebih dalam melalui penelitian empiris.")
doc.add_page_break()

# ================= BAB III =================
add_heading_custom(doc, "BAB III", 1)
add_heading_custom(doc, "METODE PENELITIAN", 1)

add_heading_custom(doc, "A. Jenis Penelitian", 2)
add_para_custom(doc, "Penelitian ini menggunakan pendekatan kualitatif dengan desain studi kasus. Pendekatan kualitatif dipilih karena tujuan penelitian adalah untuk memahami secara mendalam pola komunikasi interpersonal dalam konteks sosial tertentu, bukan untuk menguji hipotesis secara statistik. Studi kasus memungkinkan peneliti untuk menggali data secara intensif dari sekelompok subjek dalam lingkungan alami mereka.")

add_heading_custom(doc, "B. Tempat dan Waktu Penelitian", 2)
add_para_custom(doc, "Penelitian ini dilakukan di wilayah Kabupaten Bantul, Daerah Istimewa Yogyakarta. Pemilihan lokasi didasarkan pada ketersediaan akses, keragaman kondisi sosial-ekonomi masyarakat, serta relevansi dengan isu BDR. Waktu penelitian direncanakan pada semester genap tahun akademik 2025/2026, dengan durasi pengumpulan data selama tiga bulan (Maret - Mei 2026).")

add_heading_custom(doc, "C. Subjek Penelitian", 2)
add_para_custom(doc, "Subjek penelitian terdiri dari 10 pasang anak dan orang tua yang tinggal di wilayah Bantul. Kriteria pemilihan subjek meliputi: (1) anak berusia 9-15 tahun dan sedang menjalani pembelajaran dari rumah; (2) orang tua berperan aktif dalam pendampingan belajar anak; (3) keluarga bersedia menjadi partisipan dengan memberikan persetujuan tertulis. Pemilihan subjek dilakukan secara purposive sampling untuk memastikan representasi kondisi sosial yang beragam.")

add_heading_custom(doc, "D. Desain Penelitian", 2)
add_para_custom(doc, "Desain penelitian ini mengikuti model studi kasus eksploratori yang dikembangkan oleh Yin (2018). Proses penelitian melibatkan tiga tahap utama: (1) persiapan dan pengembangan instrumen; (2) pengumpulan data melalui wawancara mendalam, observasi partisipatif, dan dokumentasi; serta (3) analisis data menggunakan teknik analisis tematik. Setiap tahap dirancang untuk memastikan validitas dan reliabilitas data yang diperoleh.")

add_heading_custom(doc, "E. Teknik Pengumpulan Data", 2)
add_para_custom(doc, "Teknik pengumpulan data yang digunakan dalam penelitian ini meliputi:")
teknik_items = [
    "Wawancara mendalam (in-depth interview): Dilakukan secara semi-terstruktur dengan anak dan orang tua secara terpisah untuk memperoleh perspektif yang berbeda.",
    "Observasi partisipatif: Peneliti mengamati interaksi komunikasi antara anak dan orang tua dalam aktivitas belajar sehari-hari selama minimal tiga kali pertemuan.",
    "Dokumentasi: Pengumpulan dokumen pendukung seperti jadwal belajar, tugas sekolah, serta catatan komunikasi (misalnya pesan teks atau catatan harian)."
]
for item in teknik_items:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run("\u2022  " + item)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)

add_heading_custom(doc, "F. Instrumen Penelitian", 2)
add_para_custom(doc, "Instrumen utama dalam penelitian ini adalah panduan wawancara semi-terstruktur yang telah divalidasi melalui konsultasi dengan dosen pembimbing. Panduan wawancara mencakup pertanyaan tentang: (a) frekuensi dan durasi komunikasi; (b) topik percakapan; (c) strategi komunikasi yang digunakan; (d) hambatan komunikasi yang dialami; serta (e) dampak komunikasi terhadap motivasi belajar. Selain itu, peneliti juga menggunakan catatan lapangan dan rekaman audio (dengan persetujuan partisipan) sebagai instrumen pendukung.")

add_heading_custom(doc, "G. Teknik Analisis Data", 2)
add_para_custom(doc, "Data yang terkumpul akan dianalisis menggunakan teknik analisis tematik (thematic analysis) yang dikembangkan oleh Braun dan Clarke (2006). Proses analisis meliputi enam langkah: (1) familiarisasi dengan data; (2) pengkodean awal; (3) pencarian tema; (4) peninjauan tema; (5) penamaan tema; serta (6) penyusunan laporan. Analisis dilakukan secara manual dengan bantuan perangkat lunak NVivo 12 untuk mempermudah pengelolaan kode dan tema.")

add_heading_custom(doc, "H. Indikator Keberhasilan", 2)
add_para_custom(doc, "Keberhasilan penelitian ini akan diukur berdasarkan beberapa indikator:")
indikator_items = [
    "Teridentifikasinya pola komunikasi interpersonal yang dominan dalam pendampingan BDR di Bantul.",
    "Ditemukannya faktor-faktor kunci yang mempengaruhi efektivitas komunikasi anak-orang tua.",
    "Tersusunnya rekomendasi praktis untuk meningkatkan kualitas komunikasi dalam konteks BDR.",
    "Tercapainya validitas data melalui triangulasi sumber (anak, orang tua, dan observasi)."
]
for item in indikator_items:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(0.5)
    run = p.add_run("\u2022  " + item)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)

doc.add_page_break()

# ================= DAFTAR PUSTAKA =================
add_heading_custom(doc, "DAFTAR PUSTAKA", 1)
refs = [
    "Braun, V., & Clarke, V. (2006). Using thematic analysis in psychology. Qualitative Research in Psychology, 3(2), 77-101.",
    "DeVito, J. A. (2016). Interpersonal Communication (14th ed.). Pearson Education.",
    "Epstein, J. L. (2018). School, family, and community partnerships: Preparing educators and improving schools (3rd ed.). Westview Press.",
    "Koerner, A. F., & Fitzpatrick, M. A. (2002). Toward a theory of family communication. Communication Theory, 12(1), 70-91.",
    "Vangelisti, A. L. (2012). Interpersonal processes in family communication. In M. B. Oliver & M. M. Mark (Eds.), The handbook of communication science (pp. 229-248). Sage Publications.",
    "Yin, R. K. (2018). Case study research and applications: Design and methods (6th ed.). Sage Publications.",
    "Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi Republik Indonesia. (2021). Panduan Pembelajaran Jarak Jauh. Jakarta: Kemendikbudristek."
]
for r in refs:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.first_line_indent = Inches(-0.25)
    p.paragraph_format.left_indent = Inches(0.25)
    run = p.add_run(r)
    run.font.name = 'Times New Roman'; run.font.size = Pt(11); run.font.color.rgb = RGBColor(0,0,0)

doc.save('Downloads/Proposal Penelitian Tia Lengkap.docx')
print("File berhasil disimpan: Downloads/Proposal Penelitian Tia Lengkap.docx")
