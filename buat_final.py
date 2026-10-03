from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document('Downloads/Proposal_Penelitian_Tia_Bab1.docx')

# Karena file sudah berantakan, mari buat file baru dari template asli tapi hanya isi Bab 1 dengan urutan benar
# Kita akan menggunakan pendekatan: hapus semua paragraf setelah index 151 (setelah Manfaat Praktis) 
# dan ganti semua isi Bab 1 dengan teks yang benar

# Mari mulai dengan menghapus semua paragraf setelah index 151 (Bab II, III, Daftar Pustaka, dll)
# Tapi ini rumit. Mari kita buat file baru dari file asli (Proposal Penelitian Tia.docx) 
# dan hanya isi Bab 1

doc_orig = Document('Downloads/Proposal Penelitian Tia.docx')
new_doc = Document()

# Salin semua paragraf dari file asli sampai index 123 (sebelum BAB I)
for i in range(0, 124):
    new_doc.add_paragraph(doc_orig.paragraphs[i].text)

# Sekarang tambahkan semua bagian Bab 1 dengan teks lengkap dan urutan benar
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_para(doc, text, bold=False, align='left', size_pt=11):
    p = doc.add_paragraph()
    if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'justify': p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    if bold: run.bold = True
    run.font.color.rgb = RGBColor(0,0,0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    if align == 'left': p.paragraph_format.first_line_indent = Inches(0.5)
    return p

def add_bullet(doc, text, align='left', size_pt=11):
    return add_para(doc, "\u2022  " + text, align=align, size_pt=size_pt)

# BAB I
add_para(new_doc, "BAB I", bold=True, align='center', size_pt=14)
add_para(new_doc, "PENDAHULUAN", bold=True, align='center', size_pt=14)

# A. Latar Belakang Masalah
add_para(new_doc, "A. Latar Belakang Masalah", bold=True, size_pt=12)
add_para(new_doc, "Komunikasi interpersonal merupakan proses pertukaran pesan antara dua individu atau lebih yang terjadi secara langsung atau tidak langsung, dengan tujuan saling memahami dan mempengaruhi satu sama lain. Dalam konteks keluarga, komunikasi interpersonal antara anak dan orang tua memiliki peran yang sangat penting dalam membentuk karakter, motivasi belajar, serta kesejahteraan psikologis anak (DeVito, 2016). Sejak pandemi COVID-19, sistem pembelajaran di Indonesia mengalami perubahan signifikan dengan diterapkannya pembelajaran jarak jauh atau Belajar dari Rumah (BDR). Perubahan ini tidak hanya mempengaruhi proses pembelajaran di sekolah, tetapi juga mengubah dinamika interaksi antara anak dan orang tua di rumah.", align='justify')
add_para(new_doc, "Di Kabupaten Bantul, Daerah Istimewa Yogyakarta, banyak orang tua yang harus berperan ganda: sebagai pengasuh, pendamping belajar, sekaligus komunikator utama bagi anak-anak mereka selama masa BDR. Fenomena ini menimbulkan berbagai tantangan dalam pola komunikasi interpersonal, seperti perubahan frekuensi interaksi, pergeseran topik percakapan dari aktivitas sehari-hari menjadi fokus pada tugas belajar, serta munculnya konflik akibat tekanan akademik yang dirasakan bersama oleh anak dan orang tua. Oleh karena itu, penelitian ini bertujuan untuk memahami bagaimana pola komunikasi interpersonal anak dengan orang tua berlangsung dalam pendampingan BDR di Bantul.", align='justify')

# B. Identifikasi Masalah
add_para(new_doc, "B. Identifikasi Masalah", bold=True, size_pt=12)
add_para(new_doc, "Berdasarkan latar belakang di atas, beberapa masalah yang dapat diidentifikasi antara lain:", align='justify')
add_bullet(new_doc, "Adanya perubahan signifikan dalam frekuensi dan kualitas komunikasi antara anak dan orang tua sejak diterapkannya BDR.")
add_bullet(new_doc, "Munculnya ketegangan atau konflik komunikasi akibat peran ganda orang tua sebagai pendamping belajar dan pengasuh.")
add_bullet(new_doc, "Kurangnya pemahaman orang tua terhadap strategi komunikasi yang efektif dalam mendukung proses belajar anak dari rumah.")
add_bullet(new_doc, "Terbatasnya penelitian yang secara spesifik mengkaji pola komunikasi interpersonal anak-orang tua dalam konteks BDR di wilayah Bantul.")

# C. Batasan Penelitian
add_para(new_doc, "C. Batasan Penelitian", bold=True, size_pt=12)
add_para(new_doc, "Penelitian ini memiliki batasan sebagai berikut:", align='justify')
add_bullet(new_doc, "Fokus penelitian terbatas pada pola komunikasi interpersonal anak dengan orang tua dalam konteks pendampingan BDR.")
add_bullet(new_doc, "Lokasi penelitian dibatasi di wilayah Kabupaten Bantul, Daerah Istimewa Yogyakarta.")
add_bullet(new_doc, "Subjek penelitian adalah anak usia sekolah dasar (SD) hingga sekolah menengah pertama (SMP) bersama orang tua mereka.")
add_bullet(new_doc, "Waktu pengambilan data dibatasi pada semester genap tahun akademik 2025/2026.")
add_bullet(new_doc, "Penelitian ini tidak mengkaji aspek pembelajaran daring secara teknis, melainkan hanya pada dinamika komunikasinya.")

# D. Rumusan Masalah
add_para(new_doc, "D. Rumusan Masalah", bold=True, size_pt=12)
add_para(new_doc, "Berdasarkan identifikasi masalah di atas, rumusan masalah dalam penelitian ini adalah:", align='justify')
add_bullet(new_doc, "Bagaimana pola komunikasi interpersonal anak dengan orang tua dalam pendampingan BDR di Bantul?")
add_bullet(new_doc, "Faktor apa saja yang mempengaruhi efektivitas komunikasi interpersonal anak-orang tua selama proses BDR?")
add_bullet(new_doc, "Bagaimana dampak pola komunikasi tersebut terhadap motivasi belajar dan kesejahteraan psikologis anak?")

# E. Tujuan Penelitian
add_para(new_doc, "E. Tujuan Penelitian", bold=True, size_pt=12)
add_para(new_doc, "Tujuan penelitian ini antara lain:", align='justify')
add_bullet(new_doc, "Mendeskripsikan pola komunikasi interpersonal anak dengan orang tua dalam pendampingan BDR di Bantul.")
add_bullet(new_doc, "Mengidentifikasi faktor-faktor yang mempengaruhi efektivitas komunikasi interpersonal selama proses BDR.")
add_bullet(new_doc, "Menganalisis dampak pola komunikasi tersebut terhadap motivasi belajar dan kesejahteraan psikologis anak.")

# F. Manfaat Penelitian
add_para(new_doc, "F. Manfaat Penelitian", bold=True, size_pt=12)
add_para(new_doc, "1. Manfaat Teoritis", bold=True, size_pt=11, align='left')
add_para(new_doc, "Penelitian ini diharapkan dapat memberikan kontribusi teoretis dalam pengembangan ilmu komunikasi, khususnya pada bidang komunikasi interpersonal dan komunikasi keluarga. Hasil penelitian dapat memperkaya literatur mengenai dinamika komunikasi dalam konteks krisis sosial seperti pandemi, serta membuka ruang diskusi tentang peran orang tua sebagai agen komunikasi dalam proses pendidikan informal.", align='justify')
add_para(new_doc, "2. Manfaat Praktis", bold=True, size_pt=11, align='left')
add_para(new_doc, "Secara praktis, penelitian ini dapat memberikan rekomendasi bagi orang tua, pendidik, dan pembuat kebijakan dalam menyusun program pendampingan BDR yang lebih efektif. Selain itu, hasil penelitian dapat menjadi acuan bagi konselor keluarga dalam merancang intervensi komunikasi yang mendukung kesehatan mental anak selama pembelajaran di rumah.", align='justify')

# Salin cover, kata pengantar sudah diganti sebelumnya; tapi kita akan buat file baru dari file asli
# Mari kita mulai lagi: salin semua dari file asli sampai cover + pengantar
new_doc.save('Downloads/Proposal_Penelitian_Tia_Final.docx')
print("File final disimpan: Downloads/Proposal_Penelitian_Tia_Final.docx")
