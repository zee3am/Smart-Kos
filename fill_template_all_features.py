import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import shutil

def populate_whitebox_docx():
    # Always load from pristine backup so runs are idempotent
    shutil.copyfile(r'_White Box Testing_original_backup.docx', r'_White Box Testing.docx')
    doc = docx.Document(r'_White Box Testing.docx')

    PRIMARY_RGB = RGBColor(44, 79, 161)     # #2C4FA1
    DARK_RGB = RGBColor(15, 23, 42)
    WHITE_RGB = RGBColor(255, 255, 255)
    MUTED_RGB = RGBColor(100, 116, 139)
    COLOR_PRIMARY = "2C4FA1"
    COLOR_LIGHT_BG = "F8FAFC"
    COLOR_BORDER = "CBD5E1"

    def set_cell_background(cell, hex_color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        tcPr.append(shd)

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)

    def set_table_borders(table, color=COLOR_BORDER, sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def format_table_header(row, col_widths, bg_color=COLOR_PRIMARY):
        for i, cell in enumerate(row.cells):
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=140, bottom=140, left=150, right=150)
            cell.width = Inches(col_widths[i])
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.bold = True
                    run.font.size = Pt(9.5)
                    run.font.color.rgb = WHITE_RGB
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    def format_data_row(row, col_widths, is_even=False, font_size=8.5):
        bg_color = COLOR_LIGHT_BG if is_even else "FFFFFF"
        for i, cell in enumerate(row.cells):
            set_cell_background(cell, bg_color)
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            cell.width = Inches(col_widths[i])
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(font_size)
                    if not run.font.color.rgb:
                        run.font.color.rgb = DARK_RGB
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    # 1. Update Project Subtitle (P14)
    p14 = doc.paragraphs[14]
    p14.text = "Smart-Kos: Sistem Manajemen Kos Modern Berbasis Web (Role Under Test: ADMIN)"
    p14.runs[0].font.name = "Calibri"
    p14.runs[0].font.size = Pt(13)
    p14.runs[0].font.bold = True
    p14.runs[0].font.color.rgb = RGBColor(71, 85, 105)

    # 2. Update Table 0 (Document Control)
    t0 = doc.tables[0]
    t0.rows[1].cells[0].paragraphs[0].text = "1.0"
    t0.rows[1].cells[1].paragraphs[0].text = "Tim Praktikum PKPL (Zidane / zee3am)"
    t0.rows[1].cells[2].paragraphs[0].text = "White-Box Testing & Automated Unit Testing Suite untuk 8 Fitur Admin Smart-Kos Sesuai Dokumen Test Plan Modul 1 (43 Test Cases, 100% Statement Coverage)"
    for cell in t0.rows[1].cells:
        for r in cell.paragraphs[0].runs:
            r.font.name = "Calibri"
            r.font.size = Pt(9.0)

    # 3. Update P17, P19, P20, P21
    doc.paragraphs[17].text = "A. PRE-LAB: Manual White-Box Testing Analysis"
    doc.paragraphs[17].runs[0].font.bold = True
    doc.paragraphs[17].runs[0].font.size = Pt(13)
    doc.paragraphs[17].runs[0].font.color.rgb = PRIMARY_RGB

    p19 = doc.paragraphs[19]
    p19._element.clear_content()
    r19 = p19.add_run(
        "Fitur: F-05 Verifikasi Booking Kamar (Admin Verification & Lease Management)\n"
        "Fungsi / Method: approveBooking(bookingId: string)\n"
        "Lokasi File: app/admin/verifikasi/actions.ts\n"
        "Peran Pengujian (Role Under Test): ADMIN (Pengelola Kos)"
    )
    r19.font.name = "Calibri"
    r19.font.size = Pt(9.5)

    p20 = doc.paragraphs[20]
    p20.text = "Target Coverage: 100% Statement Coverage & 100% Branch Coverage"
    p20.runs[0].font.name = "Calibri"
    p20.runs[0].font.size = Pt(9.5)
    p20.runs[0].font.bold = True

    p21 = doc.paragraphs[21]
    p21.text = "Total Statement: 14 Pernyataan Eksekusi (S1 – S14) | Titik Keputusan: 2 (D1, D2 / 4 Cabang: B1 – B4) | V(G) = 3"
    p21.runs[0].font.name = "Calibri"
    p21.runs[0].font.size = Pt(9.5)

    p22 = doc.paragraphs[22]
    p22.text = (
        "Dekomposisi Pernyataan (S1–S14) & Titik Keputusan:\n"
        "• S1: Query prisma.booking.findUnique() mengambil data pemesanan berdasarkan bookingId.\n"
        "• S2 (D1): Evaluasi kondisi percabangan `if (!booking)` (D1).\n"
        "• S3 (B1 - True): Eksekusi `return { error: 'Pesanan tidak ditemukan' }` saat booking bernilai null.\n"
        "• S4 (B2 - False): Deklarasi `startDate = new Date()` (booking valid).\n"
        "• S5: Query prisma.booking.findFirst() mencari sewa aktif terakhir penghuni pada kamar terkait.\n"
        "• S6 (D2): Evaluasi `const isExtension = lastActiveBooking && new Date(lastActiveBooking.endDate) > startDate`.\n"
        "• S7 (B3 - True): `baseDate = new Date(lastActiveBooking.endDate)` (akumulasi perpanjangan sewa).\n"
        "• S8 (B4 - False): `baseDate = startDate` (mulai dari hari ini untuk sewa baru).\n"
        "• S9: Perhitungan tanggal berakhir: `endDate.setMonth(baseDate.getMonth() + booking.durasi)`.\n"
        "• S10: Mutasi prisma.booking.update() -> status: 'APPROVED', startDate, endDate.\n"
        "• S11: Mutasi prisma.kamar.update() -> status: 'TERISI'.\n"
        "• S12: Mutasi prisma.user.update() -> role: 'PENGHUNI'.\n"
        "• S13: Invalidation cache halaman Next.js: `revalidatePath('/admin/verifikasi')`.\n"
        "• S14: Invalidation cache detail kamar: `revalidatePath('/admin/kamar')`."
    )
    p22.runs[0].font.name = "Calibri"
    p22.runs[0].font.size = Pt(8.5)

    # 4. Table 1 (Statement Coverage)
    t1 = doc.tables[1]
    sc_widths = [0.8, 1.8, 1.4, 1.2, 1.3]

    t1.rows[0].cells[0].paragraphs[0].text = "ID Test"
    t1.rows[0].cells[1].paragraphs[0].text = "Input & Kondisi Dependensi"
    t1.rows[0].cells[2].paragraphs[0].text = "Expected Output"
    t1.rows[0].cells[3].paragraphs[0].text = "Covered Statements"
    t1.rows[0].cells[4].paragraphs[0].text = "Cumulative Coverage"
    format_table_header(t1.rows[0], sc_widths)

    sc_rows = [
        (
            "TC-01",
            "bookingId: \"bk-invalid-999\"\n(booking: null / tidak ditemukan)",
            "return {\n  error:\n  \"Pesanan tidak ditemukan\"\n}",
            "S1, S2, S3",
            "{S1, S2, S3}"
        ),
        (
            "TC-02",
            "bookingId: \"bk-new-001\"\n(durasi: 1 bln, lastActiveBooking: null)",
            "Success (void)\n• Booking APPROVED\n• Kamar TERISI\n• User PENGHUNI\n• baseDate = hari ini",
            "S1, S2, S4, S5, S6, S8, S9, S10, S11, S12, S13, S14",
            "{S1, S2, S3, S4, S5, S6, S8, S9, S10, S11, S12, S13, S14}"
        ),
        (
            "TC-03",
            "bookingId: \"bk-ext-002\"\n(durasi: 3 bln, lastActiveBooking end date > now)",
            "Success (void)\n• Booking APPROVED\n• Kamar TERISI\n• User PENGHUNI\n• baseDate = sewa aktif lama",
            "S1, S2, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14",
            "{S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14}\n(Mencakup S7)"
        )
    ]

    for idx, (tcid, inp, exp, cov, cum) in enumerate(sc_rows):
        row = t1.rows[idx + 1]
        row.cells[0].paragraphs[0].text = tcid
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        row.cells[1].paragraphs[0].text = inp
        row.cells[2].paragraphs[0].text = exp
        row.cells[3].paragraphs[0].text = cov
        row.cells[4].paragraphs[0].text = cum
        format_data_row(row, sc_widths, is_even=(idx % 2 == 1))

    # Remove Row 4 (TC-04 placeholder from original doc)
    tr4 = t1.rows[4]._tr
    t1._tbl.remove(tr4)

    # Row 4 is now Summary Statements Covered
    row_sum1 = t1.rows[4]
    row_sum1.cells[0].paragraphs[0].text = "Statements Covered After All Test Cases:"
    row_sum1.cells[0].paragraphs[0].runs[0].font.bold = True
    row_sum1.cells[4].paragraphs[0].text = "{S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14} (14/14)"
    row_sum1.cells[4].paragraphs[0].runs[0].font.bold = True
    set_cell_background(row_sum1.cells[0], "F1F5F9")
    set_cell_background(row_sum1.cells[4], "E2E8F0")

    # Row 5 is now Coverage Percentage
    row_sum2 = t1.rows[5]
    row_sum2.cells[0].paragraphs[0].text = "Coverage Percentage:"
    row_sum2.cells[0].paragraphs[0].runs[0].font.bold = True
    row_sum2.cells[4].paragraphs[0].text = "100% (14 / 14 × 100%)"
    row_sum2.cells[4].paragraphs[0].runs[0].font.bold = True
    row_sum2.cells[4].paragraphs[0].runs[0].font.color.rgb = RGBColor(22, 101, 52)
    set_cell_background(row_sum2.cells[0], "DCFCE7")
    set_cell_background(row_sum2.cells[4], "BBF7D0")

    # 5. Update P23, P24, P25
    doc.paragraphs[23].text = "Branch Coverage & Basis Path Analysis (Fungsi approveBooking):"
    doc.paragraphs[23].runs[0].font.bold = True
    doc.paragraphs[23].runs[0].font.size = Pt(10)

    doc.paragraphs[24].text = (
        "• Titik Keputusan D1 (if !booking): Cabang B1 (True: Pesanan tidak ditemukan -> return error S3), Cabang B2 (False: Lanjut ke S4).\n"
        "• Titik Keputusan D2 (if isExtension): Cabang B3 (True: Perpanjang sewa aktif -> S7), Cabang B4 (False: Sewa baru dari hari ini -> S8).\n"
        "• Branch Coverage: 4 / 4 Cabang tereksekusi = 100% (Ditutupi oleh TC-01, TC-02, TC-03).\n"
        "• Cyclomatic Complexity: V(G) = Predicate Nodes (2) + 1 = 3 Jalur Independen Basis Path:\n"
        "   - Path 1: 1 -> 2 (D1:True) -> 3 [Selesai error]\n"
        "   - Path 2: 1 -> 2 (D1:False) -> 4 -> 5 -> 6 (D2:False) -> 8 -> 9 -> 10 -> 11 -> 12 -> 13 -> 14 [Selesai sewa baru]\n"
        "   - Path 3: 1 -> 2 (D1:False) -> 4 -> 5 -> 6 (D2:True) -> 7 -> 8 -> 9 -> 10 -> 11 -> 12 -> 13 -> 14 [Selesai perpanjangan]"
    )
    doc.paragraphs[24].runs[0].font.size = Pt(8.5)

    doc.paragraphs[25].text = (
        "Ringkasan 8 Fitur Admin Berdasarkan Dokumen Test Plan Modul 1 (Total 43 Unit Tests, 100% Statements):\n"
        "1. F-01: Login Admin (loginUser) - 6 Test Cases (100% Stmts)\n"
        "2. F-02: Register Admin (registerUser) - 5 Test Cases (100% Stmts)\n"
        "3. F-03: Forgot & Reset Password Admin (submitForgotPassword, submitResetPassword) - 9 Test Cases (100% Stmts)\n"
        "4. F-04: Kelola Data Kamar Kos (hapusKamar) - 1 Test Case (100% Stmts)\n"
        "5. F-05: Verifikasi Booking Kamar (approveBooking, rejectBooking) - 5 Test Cases (100% Stmts)\n"
        "6. F-06: Verifikasi Tagihan & Pengaturan Kos (updateKostProfile) - 3 Test Cases (100% Stmts)\n"
        "7. F-07: Kelola Tiket Keluhan Penghuni (updateTiketStatus) - 2 Test Cases (100% Stmts)\n"
        "8. F-08: Kelola Data Penghuni Kos (tambahPenghuniManual, hapusPenghuni) - 7 Test Cases (100% Stmts)"
    )
    doc.paragraphs[25].runs[0].font.size = Pt(8.5)

    # 6. Section B Header
    doc.paragraphs[28].text = "B. LAB SESSION: Unit Testing Implementation (8 Fitur Admin)"
    doc.paragraphs[28].runs[0].font.bold = True
    doc.paragraphs[28].runs[0].font.size = Pt(13)
    doc.paragraphs[28].runs[0].font.color.rgb = PRIMARY_RGB

    doc.paragraphs[29].text = (
        "Implementasi otomatisasi Unit Testing mencakup seluruh 8 fitur peran ADMIN pada Smart-Kos menggunakan Jest, ts-jest, dan pola Arrange-Act-Assert (AAA). Setiap fitur memiliki unit test suite tersendiri dengan 100% Statement Coverage."
    )
    doc.paragraphs[29].runs[0].font.size = Pt(9.5)

    # Update Table 2 for Feature 1 (F-01 Login Admin)
    t2 = doc.tables[2]
    lab_widths = [1.8, 4.7]
    format_table_header(t2.rows[0], lab_widths)
    t2.rows[0].cells[0].paragraphs[0].text = "Judul"
    t2.rows[0].cells[1].paragraphs[0].text = "Fitur F-01: Login Admin (loginUser)"

    t2.rows[1].cells[0].paragraphs[0].text = "Testing"
    t2.rows[1].cells[1].paragraphs[0].text = (
        "[Screenshot Kode Unit Test di VS Code: tests/f01_login.test.ts]\n\n"
        "File: tests/f01_login.test.ts (Menguji fungsi loginUser di app/(auth)/login/actions.ts)\n"
        "Pola AAA: Menyiapkan mock Prisma findUnique & bcrypt.compare (Arrange), memanggil loginUser dengan formData (Act), dan memeriksa kembalian error/redirect URL (Assert)."
    )

    t2.rows[2].cells[0].paragraphs[0].text = "Description"
    t2.rows[2].cells[1].paragraphs[0].text = (
        "Pengujian fungsionalitas autentikasi login mencakup 6 skenario komprehensif:\n"
        "• TC-AUTH-01 (Negatif): Validasi form input kosong (email/password kosong melempar pesan error).\n"
        "• TC-AUTH-02 (Negatif): Validasi email yang tidak terdaftar pada database.\n"
        "• TC-AUTH-03 (Negatif): Validasi kecocokan password menggunakan stub bcrypt.compare yang mengembalikan false.\n"
        "• TC-AUTH-04 (Positif): Pengalihan otentikasi peran ADMIN menuju dashboard '/admin'.\n"
        "• TC-AUTH-05 (Positif): Pengalihan pengguna dengan peran USER biasa menuju '/dashboard'.\n"
        "• TC-AUTH-06 (Negatif): Penanganan graceful error blok try-catch saat query database mengalami kegagalan."
    )

    t2.rows[3].cells[0].paragraphs[0].text = "Result"
    t2.rows[3].cells[1].paragraphs[0].text = (
        "[Screenshot Hasil Test di Terminal: tests/f01_login.test.ts PASS]\n"
        "Status: PASS | 6 Passed, 0 Failed | Statements Coverage: 100% | Duration: 2ms"
    )

    t2.rows[4].cells[0].paragraphs[0].text = "Commit Link"
    t2.rows[4].cells[1].paragraphs[0].text = "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"

    for r_idx in range(1, 5):
        format_data_row(t2.rows[r_idx], lab_widths, is_even=(r_idx % 2 == 1), font_size=8.5)

    # 7. Helper to create additional feature tables in Section B
    # Let's define metadata for features 2 to 8
    f_meta = [
        (
            "Fitur F-02: Register Admin (registerUser)",
            "[Screenshot Kode Unit Test di VS Code: tests/f02_register.test.ts]\n\n"
            "File: tests/f02_register.test.ts (Menguji registerUser di app/(auth)/register/actions.ts)\n"
            "Pola AAA: Menyiapkan mock formData, stub prisma.user.findUnique, dan mock bcrypt.hash & prisma.user.create.",
            "Pengujian registrasi akun admin mencakup 5 skenario:\n"
            "• TC-REG-01 (Negatif): Form input wajib belum lengkap (nama, email, password kosong).\n"
            "• TC-REG-02 (Negatif): Email sudah terdaftar sebelumnya di database.\n"
            "• TC-REG-03 (Positif): Pendaftaran akun ADMIN berhasil jika kode otorisasi sesuai ('KOSADMIN2026') -> role: 'ADMIN'.\n"
            "• TC-REG-04 (Positif): Pendaftaran berhasil dengan role default 'USER' jika kode otorisasi tidak sesuai.\n"
            "• TC-REG-05 (Negatif): Penanganan error database saat proses user.create gagal.",
            "[Screenshot Hasil Test di Terminal: tests/f02_register.test.ts PASS]\n"
            "Status: PASS | 5 Passed, 0 Failed | Statements Coverage: 100% | Duration: 1ms",
            "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"
        ),
        (
            "Fitur F-03: Lupa & Reset Password (submitForgotPassword & submitResetPassword)",
            "[Screenshot Kode Unit Test di VS Code: tests/f03_forgot_password.test.ts]\n\n"
            "File: tests/f03_forgot_password.test.ts (Menguji app/(auth)/forgot-password/actions.ts & reset-password/actions.ts)\n"
            "Pola AAA: Menyiapkan stub database untuk verifikasi email, token reset, dan mock prisma.user.update.",
            "Pengujian fitur pemulihan akun mencakup 9 skenario lengkap:\n"
            "• TC-FP-01 s/d TC-FP-04: Validasi email kosong, email tidak ditemukan, pembuatan token & reset link berhasil, dan penanganan database exception.\n"
            "• TC-RP-01 s/d TC-RP-05: Validasi token/password kosong, konfirmasi password tidak cocok, token tidak valid, pembaruan password berhasil (redirect login), dan error handling update gagal.",
            "[Screenshot Hasil Test di Terminal: tests/f03_forgot_password.test.ts PASS]\n"
            "Status: PASS | 9 Passed, 0 Failed | Statements Coverage: 100% | Duration: 3ms",
            "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"
        ),
        (
            "Fitur F-04: Kelola Data Kamar Kos (hapusKamar)",
            "[Screenshot Kode Unit Test di VS Code: tests/f04_kamar.test.ts]\n\n"
            "File: tests/f04_kamar.test.ts (Menguji hapusKamar di app/admin/kamar/actions.ts)\n"
            "Pola AAA: Menyiapkan mock prisma.kamar.delete dan mock revalidatePath.",
            "Pengujian manajemen data kamar mencakup penghapusan data:\n"
            "• TC-ROOM-01 (Positif): Memverifikasi pemanggilan fungsi prisma.kamar.delete dengan parameter id kamar yang tepat, diikuti pembersihan cache Next.js revalidatePath('/admin/kamar').",
            "[Screenshot Hasil Test di Terminal: tests/f04_kamar.test.ts PASS]\n"
            "Status: PASS | 1 Passed, 0 Failed | Statements Coverage: 100% | Duration: 4ms",
            "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"
        ),
        (
            "Fitur F-05: Verifikasi Booking Kamar (approveBooking & rejectBooking)",
            "[Screenshot Kode Unit Test di VS Code: tests/approveBooking.test.ts & tests/f05_verifikasi_booking.test.ts]\n\n"
            "File: tests/approveBooking.test.ts (Menguji approveBooking & rejectBooking di app/admin/verifikasi/actions.ts)\n"
            "Pola AAA: Menggunakan stub booking data & sewa aktif, serta mock prisma.booking.update, kamar.update, dan user.update.",
            "Pengujian alur persetujuan dan penolakan pesanan kamar sewa mencakup 5 skenario:\n"
            "• TC-BOOK-01 (Negatif): Booking ID tidak ditemukan di database (mengembalikan error).\n"
            "• TC-BOOK-02 (Positif): Penyetujuan booking baru tanpa riwayat sewa aktif sebelumnya (baseDate dihitung dari hari ini).\n"
            "• TC-BOOK-03 (Positif): Penyetujuan perpanjangan sewa booking aktif (baseDate diakumulasikan dari endDate sewa lama).\n"
            "• TC-BOOK-04 (Negatif): rejectBooking gagal jika booking ID tidak ditemukan.\n"
            "• TC-BOOK-05 (Positif): rejectBooking sukses memperbarui status booking menjadi 'REJECTED' dan merevalidasi cache.",
            "[Screenshot Hasil Test di Terminal: tests/approveBooking.test.ts PASS]\n"
            "Status: PASS | 5 Passed, 0 Failed | Statements Coverage: 100% | Duration: 2ms",
            "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"
        ),
        (
            "Fitur F-06: Verifikasi Tagihan & Pengaturan Pembayaran (updateKostProfile)",
            "[Screenshot Kode Unit Test di VS Code: tests/f06_pengaturan_pembayaran.test.ts]\n\n"
            "File: tests/f06_pengaturan_pembayaran.test.ts (Menguji updateKostProfile di app/admin/pengaturan/actions.ts)\n"
            "Pola AAA: Menyiapkan mock FormData simulasi upload berkas foto kost & QRIS, stub profil lama, dan mock create/update profil.",
            "Pengujian konfigurasi profil kost dan rekening pembayaran mencakup 3 skenario:\n"
            "• TC-BILL-01 (Negatif): Validasi form informasi wajib (nama, alamat, kontak, rekening kosong melempar error).\n"
            "• TC-BILL-02 (Positif): Pembaruan profil kost dan rekening pembayaran yang sudah ada, menangani upload foto dan QRIS.\n"
            "• TC-BILL-03 (Positif): Pembuatan profil kost dan rekening pembayaran baru jika ID entitas belum terdaftar di database.",
            "[Screenshot Hasil Test di Terminal: tests/f06_pengaturan_pembayaran.test.ts PASS]\n"
            "Status: PASS | 3 Passed, 0 Failed | Statements Coverage: 100% | Duration: 4ms",
            "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"
        ),
        (
            "Fitur F-07: Kelola Tiket Keluhan Penghuni (updateTiketStatus)",
            "[Screenshot Kode Unit Test di VS Code: tests/f07_keluhan.test.ts]\n\n"
            "File: tests/f07_keluhan.test.ts (Menguji updateTiketStatus di app/admin/keluhan/actions.ts)\n"
            "Pola AAA: Menggunakan mock prisma.keluhan.update dan mock revalidatePath('/admin/keluhan').",
            "Pengujian siklus penanganan tiket keluhan fasilitas kos mencakup 2 skenario transisi status:\n"
            "• TC-TKT-01 (Positif): Perubahan status tiket dari PENDING menjadi 'PROSES' saat admin mulai menindaklanjuti keluhan.\n"
            "• TC-TKT-02 (Positif): Perubahan status tiket menjadi 'SELESAI' setelah keluhan terselesaikan, memvalidasi pembersihan cache.",
            "[Screenshot Hasil Test di Terminal: tests/f07_keluhan.test.ts PASS]\n"
            "Status: PASS | 2 Passed, 0 Failed | Statements Coverage: 100% | Duration: 1ms",
            "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"
        ),
        (
            "Fitur F-08: Kelola Data Penghuni Kos (tambahPenghuniManual & hapusPenghuni)",
            "[Screenshot Kode Unit Test di VS Code: tests/f08_penghuni.test.ts]\n\n"
            "File: tests/f08_penghuni.test.ts (Menguji tambahPenghuniManual & hapusPenghuni di app/admin/penghuni/actions.ts)\n"
            "Pola AAA: Menyiapkan stub validasi kamar/email, mock bcrypt.hash, mock prisma.user.create, kamar.update, booking.create.",
            "Pengujian manajemen penghuni kos mencakup 7 skenario komprehensif:\n"
            "• TC-TNT-01 (Negatif): Form input tambah penghuni tidak lengkap.\n"
            "• TC-TNT-02 (Negatif): Email calon penghuni sudah terdaftar di sistem.\n"
            "• TC-TNT-03 (Negatif): Kamar yang dipilih tidak kosong (sudah TERISI) atau tidak valid.\n"
            "• TC-TNT-04 (Positif): Penambahan penghuni berhasil, kamar diubah jadi TERISI, user role PENGHUNI, dan booking APPROVED dibuat otomatis.\n"
            "• TC-TNT-05 (Negatif): Penanganan error transaksi saat proses pembuatan data melempar exception.\n"
            "• TC-TNT-06 (Negatif): Penghapusan penghuni gagal jika id penghuni tidak ditemukan di database.\n"
            "• TC-TNT-07 (Positif): Penghapusan penghuni berhasil, kamar terkait otomatis diubah kembali menjadi status KOSONG.",
            "[Screenshot Hasil Test di Terminal: tests/f08_penghuni.test.ts PASS]\n"
            "Status: PASS | 7 Passed, 0 Failed | Statements Coverage: 100% | Duration: 3ms",
            "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa"
        ),
    ]

    # Insert remaining 7 feature tables after Table 2
    # In python-docx, to insert a table right after an existing table, we can insert XML after table._tbl
    current_element = t2._tbl
    for feat_idx, (f_title, f_testing, f_desc, f_result, f_commit) in enumerate(f_meta, start=2):
        # Create an empty table with 5 rows and 2 cols
        new_tbl = doc.add_table(rows=5, cols=2)
        set_table_borders(new_tbl)
        format_table_header(new_tbl.rows[0], lab_widths)
        new_tbl.rows[0].cells[0].paragraphs[0].text = "Judul"
        new_tbl.rows[0].cells[1].paragraphs[0].text = f_title

        new_tbl.rows[1].cells[0].paragraphs[0].text = "Testing"
        new_tbl.rows[1].cells[1].paragraphs[0].text = f_testing

        new_tbl.rows[2].cells[0].paragraphs[0].text = "Description"
        new_tbl.rows[2].cells[1].paragraphs[0].text = f_desc

        new_tbl.rows[3].cells[0].paragraphs[0].text = "Result"
        new_tbl.rows[3].cells[1].paragraphs[0].text = f_result

        new_tbl.rows[4].cells[0].paragraphs[0].text = "Commit Link"
        new_tbl.rows[4].cells[1].paragraphs[0].text = f_commit

        for r_idx in range(1, 5):
            format_data_row(new_tbl.rows[r_idx], lab_widths, is_even=(r_idx % 2 == 1), font_size=8.5)

        # Move new_tbl right after current_element
        new_tbl_elem = new_tbl._tbl
        current_element.addnext(new_tbl_elem)
        
        # Add spacing paragraph
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_after = Pt(6)
        new_tbl_elem.addnext(p_space._element)

        current_element = p_space._element

    # 8. Update Section C Header & Table 3
    doc.paragraphs[32].text = "C. TEST DOUBLE IMPLEMENTATION"
    doc.paragraphs[32].runs[0].font.bold = True
    doc.paragraphs[32].runs[0].font.size = Pt(13)
    doc.paragraphs[32].runs[0].font.color.rgb = PRIMARY_RGB

    doc.paragraphs[33].text = (
        "Pengujian unit mengimplementasikan dua tipe Test Double utama untuk mengisolasi fungsionalitas unit under test dari dependensi eksternal (Prisma ORM database, Bcrypt hashing, Next.js revalidatePath, dan redirect):"
    )
    doc.paragraphs[33].runs[0].font.size = Pt(9.5)

    # Note on Table 3:
    # Table 3 is the 11th table in the doc now because we inserted 7 tables after Table 2.
    # Find table that has "Test Double" in header:
    t3 = None
    for tbl in doc.tables:
        if "Test Double" in tbl.rows[0].cells[0].text:
            t3 = tbl
            break

    td_widths = [1.1, 1.0, 2.0, 2.4]
    format_table_header(t3.rows[0], td_widths)
    t3.rows[0].cells[0].paragraphs[0].text = "Test Double"
    t3.rows[0].cells[1].paragraphs[0].text = "Type"
    t3.rows[0].cells[2].paragraphs[0].text = "Purpose"
    t3.rows[0].cells[3].paragraphs[0].text = "Reason for Selection"

    t3.rows[1].cells[0].paragraphs[0].text = "Test Double 1\n(Prisma Queries & Bcrypt Compare)"
    t3.rows[1].cells[1].paragraphs[0].text = "Stub"
    t3.rows[1].cells[2].paragraphs[0].text = (
        "Menyediakan respon data tiruan tetap (canned responses) untuk query pembacaan database prisma.*.findUnique() dan findFirst(), serta fungsi verifikasi password bcrypt.compare()."
    )
    t3.rows[1].cells[3].paragraphs[0].text = (
        "Mengisolasi unit test dari database PostgreSQL riil. Mengembalikan nilai null, entitas pengguna, status kamar, dan riwayat booking secara konsisten, deterministik, dan instan tanpa dependensi koneksi DB."
    )

    t3.rows[2].cells[0].paragraphs[0].text = "Test Double 2\n(Prisma Mutations & Next.js Revalidate)"
    t3.rows[2].cells[1].paragraphs[0].text = "Mock"
    t3.rows[2].cells[2].paragraphs[0].text = (
        "Memverifikasi interaksi pemanggilan efek samping (side-effects) pada fungsi mutasi database prisma.*.create(), update(), delete(), serta revalidatePath() dan redirect()."
    )
    t3.rows[2].cells[3].paragraphs[0].text = (
        "Memvalidasi perilaku sistem (behavior verification). Memastikan bahwa Server Action memanggil query mutasi database dengan argumen data dan parameter relasi yang tepat menggunakan asersi toHaveBeenCalledWith()."
    )

    for r_idx in range(1, 3):
        format_data_row(t3.rows[r_idx], td_widths, is_even=(r_idx % 2 == 1), font_size=8.5)

    # 9. Update Section D (GitHub Issues)
    p_d_header = None
    p_d_content = None
    for p in doc.paragraphs:
        if "GITHUB ISSUES" in p.text.upper():
            p_d_header = p
        elif "GITHUB ISSUE URL" in p.text.upper():
            p_d_content = p

    if p_d_header:
        p_d_header.text = "D. GITHUB ISSUES"
        p_d_header.runs[0].font.bold = True
        p_d_header.runs[0].font.size = Pt(13)
        p_d_header.runs[0].font.color.rgb = PRIMARY_RGB

    if p_d_content:
        p_d_content.text = "GitHub Issue URL: https://github.com/zee3am/Smart-Kos/issues/4 (Issue: White-Box Testing & Automated Unit Testing Implementation for 8 Admin Features)"
        p_d_content.runs[0].font.name = "Consolas"
        p_d_content.runs[0].font.size = Pt(9.5)
        p_d_content.runs[0].font.color.rgb = RGBColor(199, 37, 78)

    # 10. Append Section E (Built-in Test Coverage Report) at the end of the document
    p_cov_h = doc.add_paragraph()
    p_cov_h.paragraph_format.space_before = Pt(16)
    p_cov_h.paragraph_format.space_after = Pt(6)
    r_cov_h = p_cov_h.add_run("E. LAMPIRAN: BUILT-IN TEST COVERAGE REPORT (JEST & ISTANBUL)")
    r_cov_h.font.name = "Calibri"
    r_cov_h.font.size = Pt(13)
    r_cov_h.font.bold = True
    r_cov_h.font.color.rgb = PRIMARY_RGB

    p_cov_desc = doc.add_paragraph()
    p_cov_desc.paragraph_format.space_after = Pt(8)
    p_cov_desc.add_run(
        "Berikut adalah ringkasan built-in code coverage report yang dihasilkan secara otomatis oleh Jest & Istanbul CLI engine (perintah `npx jest --coverage`). Seluruh Server Action untuk 8 fitur peran Admin mencapai 100% Statement Coverage."
    ).font.size = Pt(9.5)

    # Coverage Table
    t_cov = doc.add_table(rows=10, cols=5)
    set_table_borders(t_cov)
    cov_widths = [2.2, 1.1, 1.1, 1.1, 1.1]

    cov_headers = ["File / Target Action", "% Stmts", "% Branch", "% Funcs", "% Lines"]
    for c_i, h_t in enumerate(cov_headers):
        t_cov.rows[0].cells[c_i].paragraphs[0].text = h_t
    format_table_header(t_cov.rows[0], cov_widths)

    cov_data = [
        ("app/(auth)/login/actions.ts (F-01)", "100%", "87.5%", "100%", "100%"),
        ("app/(auth)/register/actions.ts (F-02)", "100%", "100%", "100%", "100%"),
        ("app/(auth)/forgot-password/actions.ts (F-03)", "100%", "100%", "100%", "100%"),
        ("app/(auth)/reset-password/actions.ts (F-03)", "100%", "87.5%", "100%", "100%"),
        ("app/admin/kamar/actions.ts (F-04)", "100%", "100%", "100%", "100%"),
        ("app/admin/verifikasi/actions.ts (F-05)", "100%", "100%", "100%", "100%"),
        ("app/admin/pengaturan/actions.ts (F-06)", "100%", "92.85%", "100%", "100%"),
        ("app/admin/keluhan/actions.ts (F-07)", "100%", "100%", "100%", "100%"),
        ("app/admin/penghuni/actions.ts (F-08)", "100%", "91.66%", "100%", "100%"),
    ]

    for r_i, r_data in enumerate(cov_data, start=1):
        for c_i, val in enumerate(r_data):
            p_cell = t_cov.rows[r_i].cells[c_i].paragraphs[0]
            p_cell.text = val
            if c_i == 0:
                p_cell.runs[0].font.name = "Consolas"
                p_cell.runs[0].font.size = Pt(8.0)
            elif c_i == 1:
                p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cell.runs[0].font.bold = True
                p_cell.runs[0].font.color.rgb = RGBColor(22, 101, 52)
            else:
                p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_data_row(t_cov.rows[r_i], cov_widths, is_even=(r_i % 2 == 1), font_size=8.5)

    # Save directly into _White Box Testing.docx
    doc.save(r"_White Box Testing.docx")
    print("SUCCESS: Populated _White Box Testing.docx with all 8 features, 11 tables, and full coverage report!")

if __name__ == "__main__":
    populate_whitebox_docx()
