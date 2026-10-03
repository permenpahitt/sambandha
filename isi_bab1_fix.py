from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document('Downloads/Proposal_Penelitian_Tia_Bab1.docx')

def insert_text_after_index(doc, idx, text, align='justify', size_pt=11):
    p = doc.paragraphs[idx]._element
    new_p = doc.add_paragraph()
    p.getparent().insert(p.getparent().index(p) + 1, new_p._element)
    new_p.clear()
    run = new_p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0,0,0)
    if align == 'center': new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'justify': new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else: new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    new_p.paragraph_format.space_after = Pt(6)
    new_p.paragraph_format.line_spacing = 1.5
    new_p.paragraph_format.first_line_indent = Inches(0.5)

def insert_bullet_after_index(doc, idx, text, align='justify', size_pt=11):
    p = doc.paragraphs[idx]._element
    new_p = doc.add_paragraph()
    p.getparent().insert(p.getparent().index(p) + 1, new_p._element)
    new_p.clear()
    run = new_p.add_run("\u2022  " + text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0,0,0)
    if align == 'center': new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'justify': new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else: new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    new_p.paragraph_format.space_after = Pt(6)
    new_p.paragraph_format.line_spacing = 1.5
    new_p.paragraph_format.first_line_indent = Inches(0.5)

# Kita akan menggunakan pendekatan: temukan setiap heading di doc dan tambahkan teks setelahnya
# Tetapi karena index berubah setiap kali kita menambah, kita akan bekerja dari bawah ke atas

# Mari mulai dari bagian bawah Bab I (Manfaat Praktis) lalu naik ke atas

# Cari index Manfaat Praktis
for i, p in enumerate(doc.paragraphs):
    if p.text == "Manfaat Praktis":
        # Tambahkan teks manfaat praktis setelah index ini
        insert_text_after_index(doc, i, "Secara praktis, penelitian ini dapat memberikan rekomendasi bagi orang tua, pendidik, dan pembuat kebijakan dalam menyusun program pendampingan BDR yang lebih efektif. Selain itu, hasil penelitian dapat menjadi acuan bagi konselor keluarga dalam merancang intervensi komunikasi yang mendukung kesehatan mental anak selama pembelajaran di rumah.")
        print("Manfaat praktis ditambahkan pada index", i)
        break

# Manfaat Teoritis
for i, p in enumerate(doc.paragraphs):
    if p.text == "Manfaat Teoritis":
        insert_text_after_index(doc, i, "Penelitian ini diharapkan dapat memberikan kontribusi teoretis dalam pengembangan ilmu komunikasi, khususnya pada bidang komunikasi interpersonal dan komunikasi keluarga. Hasil penelitian dapat memperkaya literatur mengenai dinamika komunikasi dalam konteks krisis sosial seperti pandemi, serta membuka ruang diskusi tentang peran orang tua sebagai agen komunikasi dalam proses pendidikan informal.")
        print("Manfaat teoritis ditambahkan pada index", i)
        break

# Manfaat Penelitian (heading) - kita akan menambahkan teks sebelum sub-heading
# Tidak perlu, sudah ada sub-heading

# Tujuan Penelitian
for i, p in enumerate(doc.paragraphs):
    if p.text == "Tujuan Penelitian":
        # Tambah setelah index ini
        for j in range(3):
            insert_bullet_after_index(doc, i, [
                "Mendeskripsikan pola komunikasi interpersonal anak dengan orang tua dalam pendampingan BDR di Bantul.",
                "Mengidentifikasi faktor-faktor yang mempengaruhi efektivitas komunikasi interpersonal selama proses BDR.",
                "Menganalisis dampak pola komunikasi tersebut terhadap motivasi belajar dan kesejahteraan psikologis anak."
            ][j])
            # Karena setiap insert mengubah index, ini akan berantakan; mari sederhanakan
            # Kita akan insert semua sekaligus tapi ini rumit
            pass
        # Mari coba cara lain: insert satu per satu dengan index yang sama (akan terus bertambah)
        # Sebenarnya kita akan insert setelah index i, tapi setiap insert membuat index i bergeser
        # Jadi kita akan insert setelah index yang sama berulang kali
        insert_text_after_index(doc, i, "Tujuan penelitian ini antara lain:")
        items_tujuan = [
            "Mendeskripsikan pola komunikasi interpersonal anak dengan orang tua dalam pendampingan BDR di Bantul.",
            "Mengidentifikasi faktor-faktor yang mempengaruhi efektivitas komunikasi interpersonal selama proses BDR.",
            "Menganalisis dampak pola komunikasi tersebut terhadap motivasi belajar dan kesejahteraan psikologis anak."
        ]
        current_idx = i
        for item in items_tujuan:
            insert_bullet_after_index(doc, current_idx, item)
            current_idx += 1
        print("Tujuan penelitian ditambahkan")
        break

# Rumusan Masalah
for i, p in enumerate(doc.paragraphs):
    if p.text == "Rumusan Masalah":
        insert_text_after_index(doc, i, "Berdasarkan identifikasi masalah di atas, rumusan masalah dalam penelitian ini adalah:")
        items_rumusan = [
            "Bagaimana pola komunikasi interpersonal anak dengan orang tua dalam pendampingan BDR di Bantul?",
            "Faktor apa saja yang mempengaruhi efektivitas komunikasi interpersonal anak-orang tua selama proses BDR?",
            "Bagaimana dampak pola komunikasi tersebut terhadap motivasi belajar dan kesejahteraan psikologis anak?"
        ]
        current_idx = i
        for item in items_rumusan:
            insert_bullet_after_index(doc, current_idx, item)
            current_idx += 1
        print("Rumusan masalah ditambahkan")
        break

# Batasan Penelitian
for i, p in enumerate(doc.paragraphs):
    if p.text == "Batasan Penelitian":
        insert_text_after_index(doc, i, "Penelitian ini memiliki batasan sebagai berikut:")
        items_batasan = [
            "Fokus penelitian terbatas pada pola komunikasi interpersonal anak dengan orang tua dalam konteks pendampingan BDR.",
            "Lokasi penelitian dibatasi di wilayah Kabupaten Bantul, Daerah Istimewa Yogyakarta.",
            "Subjek penelitian adalah anak usia sekolah dasar (SD) hingga sekolah menengah pertama (SMP) bersama orang tua mereka.",
            "Waktu pengambilan data dibatasi pada semester genap tahun akademik 2025/2026.",
            "Penelitian ini tidak mengkaji aspek pembelajaran daring secara teknis, melainkan hanya pada dinamika komunikasinya."
        ]
        current_idx = i
        for item in items_batasan:
            insert_bullet_after_index(doc, current_idx, item)
            current_idx += 1
        print("Batasan penelitian ditambahkan")
        break

# Identifikasi Masalah
for i, p in enumerate(doc.paragraphs):
    if p.text == "Identifikasi Masalah":
        insert_text_after_index(doc, i, "Berdasarkan latar belakang di atas, beberapa masalah yang dapat diidentifikasi antara lain:")
        items_id = [
            "Adanya perubahan signifikan dalam frekuensi dan kualitas komunikasi antara anak dan orang tua sejak diterapkannya BDR.",
            "Munculnya ketegangan atau konflik komunikasi akibat peran ganda orang tua sebagai pendamping belajar dan pengasuh.",
            "Kurangnya pemahaman orang tua terhadap strategi komunikasi yang efektif dalam mendukung proses belajar anak dari rumah.",
            "Terbatasnya penelitian yang secara spesifik mengkaji pola komunikasi interpersonal anak-orang tua dalam konteks BDR di wilayah Bantul."
        ]
        current_idx = i
        for item in items_id:
            insert_bullet_after_index(doc, current_idx, item)
            current_idx += 1
        print("Identifikasi masalah ditambahkan")
        break

# Latar Belakang Masalah
for i, p in enumerate(doc.paragraphs):
    if p.text == "Latar Belakang Masalah":
        insert_text_after_index(doc, i, "Komunikasi interpersonal merupakan proses pertukaran pesan antara dua individu atau lebih yang terjadi secara langsung atau tidak langsung, dengan tujuan saling memahami dan mempengaruhi satu sama lain. Dalam konteks keluarga, komunikasi interpersonal antara anak dan orang tua memiliki peran yang sangat penting dalam membentuk karakter, motivasi belajar, serta kesejahteraan psikologis anak (DeVito, 2016). Sejak pandemi COVID-19, sistem pembelajaran di Indonesia mengalami perubahan signifikan dengan diterapkannya pembelajaran jarak jauh atau Belajar dari Rumah (BDR). Perubahan ini tidak hanya mempengaruhi proses pembelajaran di sekolah, tetapi juga mengubah dinamika interaksi antara anak dan orang tua di rumah.")
        insert_text_after_index(doc, i+1, "Di Kabupaten Bantul, Daerah Istimewa Yogyakarta, banyak orang tua yang harus berperan ganda: sebagai pengasuh, pendamping belajar, sekaligus komunikator utama bagi anak-anak mereka selama masa BDR. Fenomena ini menimbulkan berbagai tantangan dalam pola komunikasi interpersonal, seperti perubahan frekuensi interaksi, pergeseran topik percakapan dari aktivitas sehari-hari menjadi fokus pada tugas belajar, serta munculnya konflik akibat tekanan akademik yang dirasakan bersama oleh anak dan orang tua. Oleh karena itu, penelitian ini bertujuan untuk memahami bagaimana pola komunikasi interpersonal anak dengan orang tua berlangsung dalam pendampingan BDR di Bantul.")
        print("Latar belakang ditambahkan")
        break

doc.save('Downloads/Proposal_Penelitian_Tia_Bab1.docx')
print("File disimpan: Downloads/Proposal_Penelitian_Tia_Bab1.docx")
