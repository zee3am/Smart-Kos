import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def generate_whitebox_doc():
    doc = docx.Document()

    # --- Page Margins ---
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        
        # Header & Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Module 4: White-Box Testing & Unit Testing (Full 8 Features) | Smart-Kos")
        hrun.font.name = "Calibri"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(120, 144, 156)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        frun = fp.add_run("Informatics Laboratory – University of Muhammadiyah Malang")
        frun.font.name = "Calibri"
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(120, 144, 156)

    # Palette
    COLOR_PRIMARY = "2C4FA1"     # Brand Blue UMM
    COLOR_SECONDARY = "475569"   # Slate
    COLOR_LIGHT_BG = "F8FAFC"    # Subtle background
    COLOR_BORDER = "CBD5E1"      # Border color
    COLOR_TEXT = "0F172A"        # Main text
    
    PRIMARY_RGB = RGBColor(44, 79, 161)
    SECONDARY_RGB = RGBColor(71, 85, 105)
    DARK_RGB = RGBColor(15, 23, 42)
    WHITE_RGB = RGBColor(255, 255, 255)
    CODE_COLOR = RGBColor(199, 37, 78)

    # --- XML Helpers ---
    def set_cell_background(cell, hex_color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        tcPr.append(shd)

    def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
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

    def format_data_row(row, col_widths, is_even=False, font_size=9.0):
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

    # ==========================
    # 1. TITLE & SUBTITLE
    # ==========================
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    run_title = p_title.add_run("White-box Test Document")
    run_title.font.name = "Calibri"
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = PRIMARY_RGB

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Smart-Kos (Modern Boarding House Management Information System)")
    run_sub.font.name = "Calibri"
    run_sub.font.size = Pt(13)
    run_sub.font.bold = True
    run_sub.font.color.rgb = SECONDARY_RGB

    # ==========================
    # 2. DOCUMENT CONTROL TABLE
    # ==========================
    t_control = doc.add_table(rows=2, cols=3)
    set_table_borders(t_control)
    ctrl_widths = [1.0, 2.2, 3.3]

    hdr_cells = t_control.rows[0].cells
    hdr_cells[0].paragraphs[0].add_run("Version")
    hdr_cells[1].paragraphs[0].add_run("Author")
    hdr_cells[2].paragraphs[0].add_run("Description")
    format_table_header(t_control.rows[0], ctrl_widths)

    d_cells = t_control.rows[1].cells
    d_cells[0].paragraphs[0].add_run("2.0")
    d_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    d_cells[1].paragraphs[0].add_run("SQA Practicum Team (Zidane / zee3am)")
    d_cells[2].paragraphs[0].add_run("Comprehensive White-Box Testing Specification: Pre-Lab Analysis (approveBooking) & Full 8-Feature Unit Testing Suite (43 Tests, 100% Coverage)")
    format_data_row(t_control.rows[1], ctrl_widths, is_even=False)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_after = Pt(12)

    # ==========================
    # 3. PRE-LAB OVERVIEW
    # ==========================
    p_h1 = doc.add_paragraph()
    p_h1.paragraph_format.space_before = Pt(12)
    p_h1.paragraph_format.space_after = Pt(6)
    r_h1 = p_h1.add_run("A. PRE-LAB: MANUAL WHITE-BOX TESTING ANALYSIS")
    r_h1.font.name = "Calibri"
    r_h1.font.size = Pt(14)
    r_h1.font.bold = True
    r_h1.font.color.rgb = PRIMARY_RGB

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_after = Pt(8)
    r_desc = p_desc.add_run(
        "Pada sesi Pre-Lab ini, dilakukan analisis pengujian White-Box Testing secara manual terhadap fungsi backend inti pada fitur Verifikasi Booking Admin, yaitu "
    )
    r_desc.font.name = "Calibri"
    r_desc.font.size = Pt(10)
    
    r_func_code = p_desc.add_run("approveBooking()")
    r_func_code.font.name = "Consolas"
    r_func_code.font.size = Pt(9.5)
    r_func_code.font.bold = True
    r_func_code.font.color.rgb = CODE_COLOR

    r_desc2 = p_desc.add_run(" yang bertempat di file ")
    r_desc2.font.name = "Calibri"
    r_desc2.font.size = Pt(10)

    r_path = p_desc.add_run("app/admin/verifikasi/actions.ts")
    r_path.font.name = "Consolas"
    r_path.font.size = Pt(9.5)
    r_path.font.bold = True
    r_path.font.color.rgb = PRIMARY_RGB

    r_desc3 = p_desc.add_run(
        ". Analisis mencakup dekomposisi pernyataan eksekusi (Statement Coverage), titik keputusan (Branch Coverage), serta pemodelan Control Flow Graph (CFG) dan Cyclomatic Complexity (Basis Path Testing) tanpa mengimplementasikan kode pengujian."
    )
    r_desc3.font.name = "Calibri"
    r_desc3.font.size = Pt(10)

    # Metadata Fitur Box
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(8)
    runs_meta = [
        ("Nama Fitur: ", True), ("F-05 Verifikasi Booking Kamar (Admin Verification & Lease Management)\n", False),
        ("Fungsi / Method: ", True), ("approveBooking(bookingId: string)\n", False),
        ("Lokasi File: ", True), ("app/admin/verifikasi/actions.ts\n", False),
        ("Peran Pengujian (Role): ", True), ("ADMIN (Pengelola Kos)\n", False),
        ("Total Executable Statements: ", True), ("14 Pernyataan (S1 – S14)\n", False),
        ("Total Decision Points: ", True), ("2 Titik Keputusan (4 Cabang: B1 – B4)\n", False),
        ("Cyclomatic Complexity V(G): ", True), ("V(G) = 2 Decision Points + 1 = 3 Independent Paths\n", False),
        ("Target Coverage: ", True), ("100% Statement Coverage & 100% Branch Coverage", False)
    ]
    for m_text, m_bold in runs_meta:
        rm = p_meta.add_run(m_text)
        rm.font.name = "Calibri"
        rm.font.size = Pt(9.5)
        rm.font.bold = m_bold
        if m_bold:
            rm.font.color.rgb = PRIMARY_RGB

    # Code Table Box
    t_code = doc.add_table(rows=1, cols=1)
    set_table_borders(t_code, color="94A3B8", sz="6")
    cell_code = t_code.rows[0].cells[0]
    set_cell_background(cell_code, "F1F5F9")
    set_cell_margins(cell_code, top=140, bottom=140, left=180, right=180)
    cell_code.width = Inches(6.5)

    code_lines = [
        ("export async function approveBooking(bookingId: string) {", False, False),
        ("    // S1: Cari data booking di database berdasarkan bookingId", True, False),
        ("    const booking = await prisma.booking.findUnique({ where: { id: bookingId } });", False, False),
        ("", False, False),
        ("    // S2 [D1]: Evaluasi apakah booking TIDAK ditemukan", True, True),
        ("    if (!booking) {", False, False),
        ("        // S3: Kembalikan respon objek error jika booking tidak ditemukan", True, False),
        ("        return { error: \"Pesanan tidak ditemukan\" };", False, False),
        ("    }", False, False),
        ("", False, False),
        ("    // S4: Cari booking aktif sebelumnya milik user untuk mengecek perpanjangan", True, False),
        ("    const lastActiveBooking = await prisma.booking.findFirst({", False, False),
        ("        where: { user_id: booking.user_id, status: \"APPROVED\", id: { not: bookingId } },", False, False),
        ("        orderBy: { tanggal_berakhir: \"desc\" }", False, False),
        ("    });", False, False),
        ("", False, False),
        ("    // S5: Inisialisasi baseDate dengan tanggal hari ini (new Date())", True, False),
        ("    let baseDate = new Date();", False, False),
        ("", False, False),
        ("    // S6 [D2]: Evaluasi apakah perpanjangan sewa aktif berlaku (tanggal_berakhir > baseDate)", True, True),
        ("    if (lastActiveBooking && lastActiveBooking.tanggal_berakhir && lastActiveBooking.tanggal_berakhir > baseDate) {", False, False),
        ("        // S7: Tetapkan baseDate dari tanggal berakhir sewa aktif sebelumnya", True, False),
        ("        baseDate = new Date(lastActiveBooking.tanggal_berakhir);", False, False),
        ("    }", False, False),
        ("", False, False),
        ("    // S8: Inisialisasi tanggal berakhir baru dari baseDate", True, False),
        ("    const newTanggalBerakhir = new Date(baseDate);", False, False),
        ("", False, False),
        ("    // S9: Akumulasikan durasi sewa (dalam satuan bulan) ke tanggal berakhir", True, False),
        ("    newTanggalBerakhir.setMonth(newTanggalBerakhir.getMonth() + booking.durasi_sewa);", False, False),
        ("", False, False),
        ("    // S10: Update status booking menjadi APPROVED dan simpan tanggal_berakhir baru", True, False),
        ("    await prisma.booking.update({", False, False),
        ("        where: { id: bookingId },", False, False),
        ("        data: { status: \"APPROVED\", tanggal_berakhir: newTanggalBerakhir }", False, False),
        ("    });", False, False),
        ("", False, False),
        ("    // S11: Update status kamar menjadi TERISI", True, False),
        ("    await prisma.kamar.update({ where: { id: booking.kamar_id }, data: { status: \"TERISI\" } });", False, False),
        ("", False, False),
        ("    // S12: Update status user menjadi PENGHUNI dan tetapkan relasi kamar_id", True, False),
        ("    await prisma.user.update({ where: { id: booking.user_id }, data: { status: \"PENGHUNI\", kamar_id: booking.kamar_id } });", False, False),
        ("", False, False),
        ("    // S13: Revalidasi cache halaman verifikasi admin", True, False),
        ("    revalidatePath(\"/admin/verifikasi\");", False, False),
        ("", False, False),
        ("    // S14: Revalidasi cache halaman manajemen kamar admin", True, False),
        ("    revalidatePath(\"/admin/kamar\");", False, False),
        ("}", False, False)
    ]

    p_cfirst = cell_code.paragraphs[0]
    for idx_cl, (code_text, is_comment, is_decision) in enumerate(code_lines):
        p_cline = p_cfirst if idx_cl == 0 else cell_code.add_paragraph()
        p_cline.paragraph_format.space_before = Pt(0)
        p_cline.paragraph_format.space_after = Pt(1)
        p_cline.paragraph_format.line_spacing = Pt(11)
        
        run_code = p_cline.add_run(code_text)
        run_code.font.name = "Consolas"
        run_code.font.size = Pt(8.5)
        if is_decision:
            run_code.font.bold = True
            run_code.font.color.rgb = RGBColor(180, 83, 9)
        elif is_comment:
            run_code.font.bold = True
            run_code.font.color.rgb = RGBColor(37, 99, 235)
        else:
            run_code.font.color.rgb = DARK_RGB

    p_space2 = doc.add_paragraph()
    p_space2.paragraph_format.space_after = Pt(10)

    # Statement Coverage
    p_sc_title = doc.add_paragraph()
    p_sc_title.paragraph_format.space_before = Pt(10)
    p_sc_title.paragraph_format.space_after = Pt(4)
    r_sc = p_sc_title.add_run("1. Analisis Statement Coverage (Fungsi approveBooking)")
    r_sc.font.name = "Calibri"
    r_sc.font.size = Pt(12)
    r_sc.font.bold = True
    r_sc.font.color.rgb = PRIMARY_RGB

    t_sc = doc.add_table(rows=6, cols=5)
    set_table_borders(t_sc)
    sc_widths = [0.8, 1.8, 1.4, 1.2, 1.3]

    t_sc_hdr = t_sc.rows[0].cells
    t_sc_hdr[0].paragraphs[0].add_run("ID Test")
    t_sc_hdr[1].paragraphs[0].add_run("Input & Kondisi Dependensi")
    t_sc_hdr[2].paragraphs[0].add_run("Expected Output")
    t_sc_hdr[3].paragraphs[0].add_run("Covered Statements")
    t_sc_hdr[4].paragraphs[0].add_run("Cumulative Coverage")
    format_table_header(t_sc.rows[0], sc_widths)

    sc_data = [
        (
            "TC-01",
            "bookingId: \"bk-invalid-999\"\n(booking: null / tidak ditemukan)",
            "return {\n  error:\n  \"Pesanan tidak ditemukan\"\n}",
            "S1, S2, S3",
            "{S1, S2, S3}"
        ),
        (
            "TC-02",
            "bookingId: \"bk-new-001\"\n(durasi: 1 bln, kamar: \"k-101\", user: \"u-1\", lastActiveBooking: null)",
            "Success (void)\n• Booking APPROVED\n• Kamar TERISI\n• User PENGHUNI\n• baseDate = hari ini",
            "S1, S2, S4, S5, S6, S8, S9, S10, S11, S12, S13, S14",
            "{S1, S2, S3, S4, S5, S6, S8, S9, S10, S11, S12, S13, S14}"
        ),
        (
            "TC-03",
            "bookingId: \"bk-ext-002\"\n(durasi: 3 bln, kamar: \"k-102\", user: \"u-2\", lastActiveBooking end date > now)",
            "Success (void)\n• Booking APPROVED\n• Kamar TERISI\n• User PENGHUNI\n• baseDate = tgl akhir sewa aktif sebelumnya",
            "S1, S2, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14",
            "{S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14}\n(Mencakup S7)"
        )
    ]

    for idx, (tcid, inp, exp, cov, cum) in enumerate(sc_data):
        row = t_sc.rows[idx + 1]
        c = row.cells
        c[0].paragraphs[0].add_run(tcid).font.bold = True
        c[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c[1].paragraphs[0].add_run(inp)
        c[2].paragraphs[0].add_run(exp)
        c[3].paragraphs[0].add_run(cov)
        c[4].paragraphs[0].add_run(cum).font.bold = (idx == 2)
        format_data_row(row, sc_widths, is_even=(idx % 2 == 1), font_size=8.5)

    r_sum1 = t_sc.rows[4]
    r_sum1.cells[0].merge(r_sum1.cells[3])
    r_sum1.cells[0].paragraphs[0].add_run("Statements Covered After All Test Cases:").font.bold = True
    r_sum1.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sum1.cells[4].paragraphs[0].add_run("{S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12, S13, S14}\n(14 dari 14 pernyataan)").font.bold = True
    set_cell_background(r_sum1.cells[0], "F1F5F9")
    set_cell_background(r_sum1.cells[4], "E2E8F0")
    set_cell_margins(r_sum1.cells[0], top=100, bottom=100, left=140, right=140)
    set_cell_margins(r_sum1.cells[4], top=100, bottom=100, left=140, right=140)

    r_sum2 = t_sc.rows[5]
    r_sum2.cells[0].merge(r_sum2.cells[3])
    r_sum2.cells[0].paragraphs[0].add_run("Statement Coverage Percentage:").font.bold = True
    r_sum2.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sc_pct = r_sum2.cells[4].paragraphs[0].add_run("100% (14 / 14 × 100%)")
    r_sc_pct.font.bold = True
    r_sc_pct.font.color.rgb = RGBColor(22, 101, 52)
    set_cell_background(r_sum2.cells[0], "DCFCE7")
    set_cell_background(r_sum2.cells[4], "BBF7D0")
    set_cell_margins(r_sum2.cells[0], top=110, bottom=110, left=140, right=140)
    set_cell_margins(r_sum2.cells[4], top=110, bottom=110, left=140, right=140)

    # Branch Coverage
    p_bc_title = doc.add_paragraph()
    p_bc_title.paragraph_format.space_before = Pt(10)
    p_bc_title.paragraph_format.space_after = Pt(4)
    r_bc = p_bc_title.add_run("2. Analisis Branch Coverage (Fungsi approveBooking)")
    r_bc.font.name = "Calibri"
    r_bc.font.size = Pt(12)
    r_bc.font.bold = True
    r_bc.font.color.rgb = PRIMARY_RGB

    t_bc = doc.add_table(rows=5, cols=5)
    set_table_borders(t_bc)
    bc_widths = [1.2, 0.7, 1.8, 1.5, 1.3]

    t_bc_hdr = t_bc.rows[0].cells
    t_bc_hdr[0].paragraphs[0].add_run("Decision Point")
    t_bc_hdr[1].paragraphs[0].add_run("Cabang")
    t_bc_hdr[2].paragraphs[0].add_run("Deskripsi Evaluasi Kondisi")
    t_bc_hdr[3].paragraphs[0].add_run("Alur Eksekusi Terkait")
    t_bc_hdr[4].paragraphs[0].add_run("Ditutupi Oleh")
    format_table_header(t_bc.rows[0], bc_widths)

    bc_data = [
        ("D1: if (!booking)", "B1", "True: Data booking tidak ditemukan di database (null)", "Masuk blok if, eksekusi S3 (return error)", "TC-01"),
        ("D1: if (!booking)", "B2", "False: Data booking ditemukan di database (valid)", "Lewati if, lanjut ke S4 (cari active booking)", "TC-02, TC-03"),
        ("D2: if (lastActiveBooking && ... > baseDate)", "B3", "True: Terdapat sewa aktif dan belum berakhir (perpanjangan)", "Masuk blok if, eksekusi S7 (set baseDate dari sewa lama)", "TC-03"),
        ("D2: if (lastActiveBooking && ... > baseDate)", "B4", "False: Tidak ada sewa aktif atau sewa lama telah berakhir", "Lewati blok if, langsung ke S8 (baseDate = hari ini)", "TC-02")
    ]

    for idx, (dp, br, desc_br, path_br, cov_by) in enumerate(bc_data):
        row = t_bc.rows[idx + 1]
        c = row.cells
        c[0].paragraphs[0].add_run(dp).font.bold = True
        c[1].paragraphs[0].add_run(br).font.bold = True
        c[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c[2].paragraphs[0].add_run(desc_br)
        c[3].paragraphs[0].add_run(path_br)
        c[4].paragraphs[0].add_run(cov_by).font.bold = True
        c[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_data_row(row, bc_widths, is_even=(idx % 2 == 1), font_size=8.5)

    p_space3 = doc.add_paragraph()
    p_space3.paragraph_format.space_after = Pt(12)

    # ==========================
    # 4. LAB SESSION: FULL 8 FEATURES SUITE
    # ==========================
    p_h2 = doc.add_paragraph()
    p_h2.paragraph_format.space_before = Pt(14)
    p_h2.paragraph_format.space_after = Pt(6)
    r_h2 = p_h2.add_run("B. LAB SESSION: FULL 8-FEATURE UNIT TESTING SUITE (ROLE ADMIN)")
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(14)
    r_h2.font.bold = True
    r_h2.font.color.rgb = PRIMARY_RGB

    p_lab_desc = doc.add_paragraph()
    p_lab_desc.paragraph_format.space_after = Pt(8)
    p_lab_desc.add_run(
        "Sesuai cakupan Modul 1 (8 fitur peran Admin: 3 fitur otentikasi wajib + 5 fitur manajemen operasional), seluruh fitur telah diimplementasikan kode pengujian unit otomatisnya menggunakan Jest, ts-jest, dan Test Double (Stub & Mock) dengan total 43 test cases yang lulus 100%."
    ).font.size = Pt(9.5)

    # Table: Full 8 Features Matrix
    t_matrix = doc.add_table(rows=9, cols=5)
    set_table_borders(t_matrix)
    mat_widths = [0.8, 1.8, 1.6, 1.1, 1.2]

    mat_hdr = t_matrix.rows[0].cells
    mat_hdr[0].paragraphs[0].add_run("Kode")
    mat_hdr[1].paragraphs[0].add_run("Fitur Modul 1 (Admin)")
    mat_hdr[2].paragraphs[0].add_run("File Test & Fungsi")
    mat_hdr[3].paragraphs[0].add_run("Jumlah Tes")
    mat_hdr[4].paragraphs[0].add_run("Statement Cov")
    format_table_header(t_matrix.rows[0], mat_widths)

    matrix_data = [
        ("F-01", "Login Admin (Auth Wajib)", "tests/f01_login.test.ts\n(loginUser)", "6 Tests", "100%"),
        ("F-02", "Register Admin (Auth Wajib)", "tests/f02_register.test.ts\n(registerUser)", "5 Tests", "100%"),
        ("F-03", "Forgot & Reset Password", "tests/f03_forgot_password.test.ts\n(submitForgot & Reset)", "9 Tests", "100%"),
        ("F-04", "Kelola Data Kamar (CRUD)", "tests/f04_kamar.test.ts\n(hapusKamar)", "1 Test", "100%"),
        ("F-05", "Verifikasi Booking Kamar", "tests/f05_verifikasi_booking.test.ts\n(approve & rejectBooking)", "5 Tests", "100%"),
        ("F-06", "Tagihan & Pengaturan Kos", "tests/f06_pengaturan_pembayaran.test.ts\n(updateKostProfile)", "3 Tests", "100%"),
        ("F-07", "Kelola Tiket Keluhan", "tests/f07_keluhan.test.ts\n(updateTiketStatus)", "2 Tests", "100%"),
        ("F-08", "Kelola Data Penghuni Kos", "tests/f08_penghuni.test.ts\n(tambah & hapusPenghuni)", "7 Tests", "100%"),
    ]

    for idx, (kd, nm, ft, jt, cov) in enumerate(matrix_data):
        row = t_matrix.rows[idx + 1]
        c = row.cells
        c[0].paragraphs[0].add_run(kd).font.bold = True
        c[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c[1].paragraphs[0].add_run(nm)
        c[2].paragraphs[0].add_run(ft)
        c[3].paragraphs[0].add_run(jt).font.bold = True
        c[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c[4].paragraphs[0].add_run(cov).font.bold = True
        c[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_data_row(row, mat_widths, is_even=(idx % 2 == 1), font_size=8.5)

    p_space_mat = doc.add_paragraph()
    p_space_mat.paragraph_format.space_after = Pt(10)

    # Table 2: Unit Testing Artifacts & Commit Link
    t_lab = doc.add_table(rows=6, cols=2)
    set_table_borders(t_lab)
    lab_widths = [1.8, 4.7]

    t_lab_hdr = t_lab.rows[0].cells
    t_lab_hdr[0].paragraphs[0].add_run("Komponen")
    t_lab_hdr[1].paragraphs[0].add_run("Keterangan & Bukti Implementasi")
    format_table_header(t_lab.rows[0], lab_widths)

    lab_rows_data = [
        ("Fitur / Fungsi", "8 Fitur Lengkap Sesuai Dokumen Test Plan Modul 1 (F-01 s/d F-08)"),
        ("Testing Implementation", "[Screenshot implementasi kode Unit Testing (Jest) menggunakan pola AAA pada tests/approveBooking.test.ts dan suite f01-f08]"),
        ("Description", "Implementasi pengujian otomatis unit testing mencakup 8 fitur lengkap peran Admin: (1) Login otentikasi peran, (2) Registrasi admin dengan kode otorisasi, (3) Lupa & Reset password, (4) Penghapusan kamar, (5) Verifikasi & penolakan booking, (6) Konfigurasi rekening pembayaran kos, (7) Siklus penanganan tiket keluhan, dan (8) Manajemen penghuni & relasi sewa kamar. Seluruh pengujian menggunakan pola AAA."),
        ("Result", "[Screenshot hasil running test 'npx jest --coverage': 9 Test Suites PASS, 43 Tests PASS, 100% Statement Coverage across all actions]"),
        ("Commit Link", "https://github.com/zee3am/Smart-Kos/commit/982b726eae492deda4d80036e0698d3fb06f02aa")
    ]

    for idx, (komp, val) in enumerate(lab_rows_data):
        row = t_lab.rows[idx + 1]
        c = row.cells
        c[0].paragraphs[0].add_run(komp).font.bold = True
        c[1].paragraphs[0].add_run(val)
        format_data_row(row, lab_widths, is_even=(idx % 2 == 1), font_size=9.0)

    p_space4 = doc.add_paragraph()
    p_space4.paragraph_format.space_after = Pt(10)

    # ==========================
    # 5. TEST DOUBLE IMPLEMENTATION
    # ==========================
    p_td_title = doc.add_paragraph()
    p_td_title.paragraph_format.space_before = Pt(8)
    p_td_title.paragraph_format.space_after = Pt(4)
    r_td = p_td_title.add_run("C. TEST DOUBLE IMPLEMENTATION")
    r_td.font.name = "Calibri"
    r_td.font.size = Pt(12)
    r_td.font.bold = True
    r_td.font.color.rgb = PRIMARY_RGB

    p_td_desc = doc.add_paragraph()
    p_td_desc.paragraph_format.space_after = Pt(6)
    p_td_desc.add_run(
        "Sesuai ketentuan Modul 4, pengujian unit mengimplementasikan Test Double untuk mengisolasi dependensi eksternal (Prisma Database ORM, Bcrypt, Next.js Cache & Navigation, Supabase Storage):"
    ).font.size = Pt(9.5)

    # Table 3: Test Double
    t_td = doc.add_table(rows=5, cols=4)
    set_table_borders(t_td)
    td_widths = [1.1, 1.0, 2.0, 2.4]

    t_td_hdr = t_td.rows[0].cells
    t_td_hdr[0].paragraphs[0].add_run("Test Double")
    t_td_hdr[1].paragraphs[0].add_run("Type")
    t_td_hdr[2].paragraphs[0].add_run("Purpose")
    t_td_hdr[3].paragraphs[0].add_run("Reason for Selection")
    format_table_header(t_td.rows[0], td_widths)

    td_data = [
        (
            "Test Double 1\n(Prisma ORM Queries)",
            "Stub",
            "Menyediakan data tiruan tetap (canned responses) untuk prisma.*.findUnique() dan findFirst() di seluruh 8 fitur.",
            "Mengisolasi pengujian dari database PostgreSQL riil. Mengembalikan nilai null, entitas pengguna, kamar kosong/terisi, dan booking secara konsisten tanpa dependensi koneksi DB."
        ),
        (
            "Test Double 2\n(Prisma ORM Mutations)",
            "Mock",
            "Memverifikasi interaksi pemanggilan mutasi database pada prisma.*.create(), update(), dan delete().",
            "Memvalidasi perilaku efek samping (behavior verification). Memastikan fungsi memanggil query mutasi dengan parameter yang tepat menggunakan asersi toHaveBeenCalledWith()."
        ),
        (
            "Test Double 3\n(Bcrypt Hashing & Auth)",
            "Stub & Mock",
            "Mensimulasikan enkripsi sandi pada bcrypt.hash() dan validasi kecocokan sandi pada bcrypt.compare().",
            "Mencegah overhead komputasi hashing kriptografi saat unit test dan memungkinkan simulasi skenario kata sandi cocok (true) maupun salah (false) secara deterministik."
        ),
        (
            "Test Double 4\n(Next Cache & Navigation)",
            "Mock",
            "Memverifikasi pemanggilan fungsi revalidatePath() dan redirect() Next.js server actions.",
            "Memverifikasi bahwa alur server action berhasil merevalidasi rute UI admin yang relevan dan mengalihkan halaman pengguna ke endpoint tujuan tanpa error runtime."
        )
    ]

    for idx, (td_name, td_type, td_purp, td_reason) in enumerate(td_data):
        row = t_td.rows[idx + 1]
        c = row.cells
        c[0].paragraphs[0].add_run(td_name).font.bold = True
        c[1].paragraphs[0].add_run(td_type).font.bold = True
        c[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c[2].paragraphs[0].add_run(td_purp)
        c[3].paragraphs[0].add_run(td_reason)
        format_data_row(row, td_widths, is_even=(idx % 2 == 1), font_size=8.5)

    p_space5 = doc.add_paragraph()
    p_space5.paragraph_format.space_after = Pt(10)

    # ==========================
    # 6. GITHUB ISSUES
    # ==========================
    p_gi_title = doc.add_paragraph()
    p_gi_title.paragraph_format.space_before = Pt(8)
    p_gi_title.paragraph_format.space_after = Pt(4)
    r_gi = p_gi_title.add_run("D. GITHUB ISSUES")
    r_gi.font.name = "Calibri"
    r_gi.font.size = Pt(12)
    r_gi.font.bold = True
    r_gi.font.color.rgb = PRIMARY_RGB

    p_gi_desc = doc.add_paragraph()
    p_gi_desc.paragraph_format.space_after = Pt(6)
    p_gi_desc.add_run(
        "GitHub Issues digunakan untuk mengelola dan mencatat pelacakan tugas pengujian White-Box Testing dan Unit Testing pada 8 fitur Admin Smart-Kos."
    ).font.size = Pt(9.5)

    p_gi_box = doc.add_paragraph()
    p_gi_box.paragraph_format.space_after = Pt(12)
    r_gi_lbl = p_gi_box.add_run("GitHub Issue URL: ")
    r_gi_lbl.font.bold = True
    r_gi_lbl.font.color.rgb = PRIMARY_RGB
    r_gi_val = p_gi_box.add_run("https://github.com/zee3am/Smart-Kos/issues/4 (Issue: White-Box Testing & Comprehensive Unit Test Implementation for 8 Admin Features)")
    r_gi_val.font.name = "Consolas"
    r_gi_val.font.size = Pt(9.5)
    r_gi_val.font.color.rgb = CODE_COLOR

    # Save to _White Box Testing Template.docx
    doc.save(r"_White Box Testing Template.docx")
    print("Successfully generated and saved _White Box Testing Template.docx with Full 8 Features Suite!")

if __name__ == "__main__":
    generate_whitebox_doc()
