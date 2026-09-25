from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageChops


ROOT = Path(__file__).resolve().parent
SOURCE_PATH = ROOT / "Nội dung bài báo cáo CĐTN.docx"
IMAGE_SOURCE_DIR = ROOT / "tmp" / "source_images"
IMAGE_DIR = ROOT / "tmp" / "prepared_images"
OUTPUT_DIR = ROOT / "output"
OUTPUT_PATH = OUTPUT_DIR / "Bao_cao_chuyen_de_tot_nghiep_50_trang.docx"

FONT = "Times New Roman"
BODY_SIZE = 13
TABLE_WIDTH_DXA = 8788
BLACK = RGBColor(0, 0, 0)
PLACEHOLDER_FILL = "FFF2CC"


def set_run_font(run, size=BODY_SIZE, bold=None, italic=None, color=BLACK):
    run.font.name = FONT
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), FONT)
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=100, bottom=80, end=100):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_geometry(table, widths_dxa):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths_dxa)))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_layout = tbl_pr.find(qn("w:tblLayout"))
    if tbl_layout is None:
        tbl_layout = OxmlElement("w:tblLayout")
        tbl_pr.append(tbl_layout)
    tbl_layout.set(qn("w:type"), "fixed")
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        for cell, width in zip(row.cells, widths_dxa):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)


def add_field(paragraph, instruction):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "Cập nhật trường trong Word"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, placeholder, end])
    set_run_font(run, size=12)


def set_page_number_start(section, start=1):
    sect_pr = section._sectPr
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num)
    pg_num.set(qn("w:start"), str(start))


def configure_section(section, body=True):
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    if body:
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(3.5)
        section.right_margin = Cm(2)
    else:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.5)


def configure_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    normal.font.size = Pt(BODY_SIZE)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.first_line_indent = Cm(1.27)

    for style_name, size, before, after in (
        ("Heading 1", 14, 0, 12),
        ("Heading 2", 13, 12, 6),
        ("Heading 3", 13, 8, 4),
    ):
        style = doc.styles[style_name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = BLACK
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.first_line_indent = Cm(0)
    doc.styles["Heading 1"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.styles["Heading 1"].paragraph_format.page_break_before = True
    doc.styles["Heading 2"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.styles["Heading 3"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    for name in ("Caption Figure", "Caption Table"):
        if name not in doc.styles:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        else:
            style = doc.styles[name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(11)
        style.font.italic = True
        style.font.color.rgb = BLACK
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        style.paragraph_format.first_line_indent = Cm(0)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.space_before = Pt(4)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.keep_together = True

    for name in ("List Bullet", "List Number"):
        style = doc.styles[name]
        style.font.name = FONT
        style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        style.font.size = Pt(BODY_SIZE)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.space_after = Pt(0)

    for name in ("TOC 1", "TOC 2", "TOC 3"):
        if name in doc.styles:
            style = doc.styles[name]
            style.font.name = FONT
            style._element.rPr.rFonts.set(qn("w:ascii"), FONT)
            style._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
            style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
            style.font.size = Pt(11)
            style.paragraph_format.line_spacing = 1.0
            style.paragraph_format.space_before = Pt(0)
            style.paragraph_format.space_after = Pt(0)


def add_placeholder_run(paragraph, text):
    run = paragraph.add_run(text)
    set_run_font(run, bold=True)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), PLACEHOLDER_FILL)
    run._element.get_or_add_rPr().append(shd)
    return run


def add_centered(doc, text="", size=13, bold=False, after=0, before=0, placeholder=False):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.line_spacing = 1.15
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    if placeholder:
        add_placeholder_run(paragraph, text)
    else:
        set_run_font(paragraph.add_run(text), size=size, bold=bold)
    return paragraph


def add_cover(doc, inner=False):
    add_centered(doc, "ĐẠI HỌC QUỐC GIA TP. HỒ CHÍ MINH", 13, True)
    add_centered(doc, "TRƯỜNG ĐẠI HỌC CÔNG NGHỆ THÔNG TIN", 13, True)
    add_centered(doc, "[TÊN KHOA]", 13, True, after=36, placeholder=True)
    add_centered(doc, "[HỌ VÀ TÊN SINH VIÊN]", 14, True, after=52, placeholder=True)
    add_centered(doc, "BÁO CÁO CHUYÊN ĐỀ TỐT NGHIỆP", 16, True, after=18)
    add_centered(doc, "XÂY DỰNG HỆ THỐNG PHÁT HIỆN SAI LỆCH NGOẠI QUAN", 16, True)
    add_centered(doc, "CỦA SẢN PHẨM ĐỒ TRANG TRÍ XUẤT KHẨU", 16, True)
    add_centered(doc, "SỬ DỤNG THỊ GIÁC MÁY TÍNH", 16, True, after=8)
    add_centered(doc, "Computer Vision-Based Visual Anomaly Detection", 13, False, after=36)
    add_centered(doc, "CỬ NHÂN/KỸ SƯ NGÀNH [TÊN NGÀNH]", 13, True, placeholder=True)
    if inner:
        add_centered(doc, "GIẢNG VIÊN HƯỚNG DẪN", 13, True, before=34)
        add_centered(doc, "[HỌC HÀM, HỌ TÊN GIẢNG VIÊN]", 13, True, placeholder=True)
    else:
        add_centered(doc, "MÃ SỐ SINH VIÊN: [MSSV]", 13, True, before=34, placeholder=True)
    add_centered(doc, "TP. HỒ CHÍ MINH, 2026", 13, True, before=54)


def add_front_matter(doc):
    add_cover(doc, inner=False)
    doc.add_page_break()
    add_cover(doc, inner=True)
    doc.add_page_break()

    add_centered(doc, "THÔNG TIN HỘI ĐỒNG CHẤM CHUYÊN ĐỀ TỐT NGHIỆP", 14, True, after=24)
    p = doc.add_paragraph(
        "Hội đồng chấm chuyên đề tốt nghiệp được thành lập theo Quyết định số "
        "[SỐ QUYẾT ĐỊNH], ngày [NGÀY/THÁNG/NĂM] của Hiệu trưởng Trường Đại học Công nghệ Thông tin."
    )
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for run in p.runs:
        set_run_font(run)
    rows = [
        ["STT", "Họ và tên", "Chức danh trong hội đồng"],
        ["1", "[HỌ TÊN]", "Chủ tịch"],
        ["2", "[HỌ TÊN]", "Thư ký"],
        ["3", "[HỌ TÊN]", "Ủy viên"],
    ]
    add_table(doc, "", [700, 4800, 3288], rows, numbered=False)
    doc.add_page_break()

    add_centered(doc, "LỜI CẢM ƠN", 14, True, after=18)
    thanks = (
        "Em xin chân thành cảm ơn quý thầy cô Trường Đại học Công nghệ Thông tin đã trang bị nền tảng kiến thức và tạo điều kiện "
        "để em thực hiện chuyên đề tốt nghiệp. Em đặc biệt biết ơn giảng viên hướng dẫn đã định hướng phương pháp nghiên cứu, góp ý "
        "về thiết kế thí nghiệm và hỗ trợ em kiểm tra các giả định trong quá trình xây dựng hệ thống phát hiện sai lệch ngoại quan. "
        "Em cũng cảm ơn đơn vị cung cấp dữ liệu và những người đã hỗ trợ thu thập, rà soát hình ảnh sản phẩm. Do giới hạn về thời gian, "
        "quy mô dữ liệu và kinh nghiệm thực hiện, báo cáo khó tránh khỏi thiếu sót; em mong nhận được góp ý để tiếp tục hoàn thiện nghiên cứu."
    )
    p = doc.add_paragraph(thanks)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for run in p.runs:
        set_run_font(run)
    sign = doc.add_paragraph()
    sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sign.paragraph_format.first_line_indent = Cm(0)
    set_run_font(sign.add_run("TP. Hồ Chí Minh, tháng 9 năm 2026\nSinh viên thực hiện\n\n\n[HỌ VÀ TÊN]"), bold=False)
    doc.add_page_break()

    add_centered(doc, "MỤC LỤC", 14, True, after=12)
    add_field(doc.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u')
    doc.add_page_break()

    add_centered(doc, "DANH MỤC HÌNH", 14, True, after=12)
    add_field(doc.add_paragraph(), 'TOC \\h \\z \\t "Caption Figure,1"')
    add_centered(doc, "DANH MỤC BẢNG", 14, True, before=12, after=12)
    add_field(doc.add_paragraph(), 'TOC \\h \\z \\t "Caption Table,1"')
    doc.add_page_break()

    add_centered(doc, "DANH MỤC TỪ VIẾT TẮT", 14, True, after=12)
    rows = [
        ["Từ viết tắt", "Nội dung"],
        ["AUROC", "Area Under the Receiver Operating Characteristic Curve"],
        ["CNN", "Convolutional Neural Network"],
        ["FP", "False Positive - dương tính giả"],
        ["GLCM", "Gray-Level Co-occurrence Matrix"],
        ["IoU", "Intersection over Union"],
        ["LBP", "Local Binary Pattern"],
        ["OC-MAR", "Object-Constrained Multi-scale Anomaly Recognition"],
        ["PRO", "Per-Region Overlap"],
    ]
    add_table(doc, "", [2000, 6788], rows, numbered=False)


def add_page_number_footer(section):
    section.footer.is_linked_to_previous = False
    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    add_field(paragraph, "PAGE")


def add_heading(doc, text, level=1):
    paragraph = doc.add_paragraph(style=f"Heading {level}")
    paragraph.paragraph_format.first_line_indent = Cm(0)
    if level == 1:
        text = text.upper()
    set_run_font(paragraph.add_run(text), size=14 if level == 1 else 13, bold=True)
    return paragraph


def add_body_paragraph(doc, text, style=None):
    text = normalize_text(text)
    if not text:
        return None
    paragraph = doc.add_paragraph(style=style)
    if style is None:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.first_line_indent = Cm(1.27)
    for run in paragraph.runs:
        set_run_font(run)
    if not paragraph.runs:
        set_run_font(paragraph.add_run(text))
    else:
        paragraph.text = ""
        set_run_font(paragraph.add_run(text))
    return paragraph


def normalize_text(text):
    text = text.replace("\u00ad", "").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    replacements = {
        "bài báo này": "báo cáo này",
        "bài báo cáo này": "báo cáo này",
        "trong bài báo": "trong báo cáo",
        "mẫu sản phầm": "mẫu sản phẩm",
        "mất cân bằng giai cấp": "mất cân bằng dữ liệu",
        "Hệ mét": "Chỉ số đánh giá",
        "nhiều loại hạt": "nhiều hạt giống ngẫu nhiên",
        "loại bỏ mô nền bằng phương pháp đốt điện": "ảnh hưởng của bước loại bỏ nền",
        "phương pháp che phủ đối tượng": "phương pháp tạo mặt nạ đối tượng",
        "khẩu trang": "mặt nạ",
        "Nghĩa là": "Trung bình",
        "lớp học": "mẫu sản phẩm",
        "các lớp học": "các mẫu sản phẩm",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    text = re.sub(r"([,.;:!?])(\S)", r"\1 \2", text)
    return text


def add_table(doc, caption, widths, rows, numbered=True):
    if caption:
        cap = doc.add_paragraph(style="Caption Table")
        set_run_font(cap.add_run(caption), size=11, italic=True)
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    set_table_geometry(table, widths)
    for row_index, row_values in enumerate(rows):
        row = table.rows[row_index]
        for col_index, value in enumerate(row_values):
            cell = row.cells[col_index]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.text = str(value)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.first_line_indent = Cm(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_index else WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    set_run_font(run, size=10.5, bold=(row_index == 0))
            if row_index == 0:
                set_cell_shading(cell, "D9E2F3")
        if row_index == 0:
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def crop_images():
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    meaningful = [2, 3, 4, 5, 6, 8, 12, 16, 17, 18, 19, 20, 21]
    for number in meaningful:
        source = IMAGE_SOURCE_DIR / f"image{number}.jpeg"
        target = IMAGE_DIR / f"image{number}.png"
        with Image.open(source) as image:
            image = image.convert("RGB")
            white = Image.new("RGB", image.size, "white")
            diff = ImageChops.difference(image, white).convert("L")
            diff = diff.point(lambda px: 255 if px > 18 else 0)
            bbox = diff.getbbox()
            if bbox:
                left, top, right, bottom = bbox
                pad = 18
                bbox = (
                    max(0, left - pad),
                    max(0, top - pad),
                    min(image.width, right + pad),
                    min(image.height, bottom + pad),
                )
                image = image.crop(bbox)
            image.save(target)


def add_figure(doc, image_number, caption, width_cm=14.2):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.keep_with_next = True
    run = paragraph.add_run()
    run.add_picture(str(IMAGE_DIR / f"image{image_number}.png"), width=Cm(width_cm))
    cap = doc.add_paragraph(style="Caption Figure")
    set_run_font(cap.add_run(caption), size=11, italic=True)


def add_figure_pair(doc, first_number, second_number, caption, width_cm=7.0):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.keep_with_next = True
    first = paragraph.add_run()
    first.add_picture(str(IMAGE_DIR / f"image{first_number}.png"), width=Cm(width_cm))
    paragraph.add_run("   ")
    second = paragraph.add_run()
    second.add_picture(str(IMAGE_DIR / f"image{second_number}.png"), width=Cm(width_cm))
    cap = doc.add_paragraph(style="Caption Figure")
    set_run_font(cap.add_run(caption), size=11, italic=True)


def should_skip(index, text):
    text = normalize_text(text)
    if not text or re.fullmatch(r"\d+", text):
        return True
    if re.match(r"^(Hình|Bảng)\s+\d+", text, re.I):
        return True
    if re.match(r"^(Lớp|FA-\d+|Trung bình|Nghĩa là)\b", text):
        return True
    ranges = [
        (70, 72), (97, 98), (122, 123), (150, 152), (180, 183),
        (277, 288), (304, 305), (315, 319), (337, 340), (346, 349),
        (377, 390), (431, 432), (440, 442), (486, 487), (490, 493),
        (583, 586),
    ]
    return any(start <= index <= end for start, end in ranges)


def add_source_range(doc, source_doc, start, end, keep_ratio=0.60):
    buffer = []
    items = []

    def flush():
        nonlocal buffer
        if buffer:
            items.append((None, " ".join(buffer)))
            buffer = []

    for index in range(start, end):
        text = normalize_text(source_doc.paragraphs[index].text)
        if should_skip(index, text):
            continue
        if text.startswith("•"):
            flush()
            text = text.lstrip("• ")
            items.append(("List Bullet", text))
            continue
        if re.match(r"^\(\d+\)\s", text):
            flush()
            items.append(("List Number", re.sub(r"^\(\d+\)\s*", "", text)))
            continue
        if re.match(r"^[1-8](?:\.\d+){0,2}\.?\s", text):
            continue
        if text.startswith("Phương pháp 4"):
            continue
        buffer.append(text)
        if len(" ".join(buffer)) >= 520 and re.search(r"[.!?]$", text):
            flush()
        elif len(" ".join(buffer)) >= 760:
            flush()
    flush()

    if len(items) > 2 and keep_ratio < 1:
        target_chars = int(sum(len(text) for _, text in items) * keep_ratio)
        selected = []
        used = 0
        for item in items:
            if not selected or used + len(item[1]) <= target_chars:
                selected.append(item)
                used += len(item[1])
            else:
                break
        items = selected

    for style, text in items:
        add_body_paragraph(doc, text, style=style)


def add_abstract(doc):
    add_heading(doc, "TÓM TẮT", 1)
    paragraphs = [
        "Chuyên đề nghiên cứu bài toán phát hiện và định vị sai lệch ngoại quan trên tập dữ liệu công nghiệp thực tế gồm 12 mẫu sản phẩm đồ trang trí xuất khẩu. Mỗi mẫu có 100 ảnh bình thường; ảnh lỗi chỉ được dùng để đánh giá, dao động từ 30 đến 53 ảnh và không có mặt nạ lỗi ở mức pixel. Trong điều kiện dữ liệu khan hiếm này, báo cáo lựa chọn PatchCore - phương pháp phát hiện bất thường dựa trên bộ nhớ, không cần huấn luyện lại bộ trích xuất đặc trưng - làm mô hình chính. Quy trình được tổ chức theo khung OC-MAR gồm tiền xử lý, tạo mặt nạ đối tượng, suy luận PatchCore, giới hạn tiền cảnh, đánh giá đa tỷ lệ, ngưỡng thích ứng và tinh chỉnh vùng.",
        "Cấu hình cuối cùng đạt AUROC trung bình cấp ảnh 0,996, ổn định với độ lệch chuẩn của giá trị trung bình khoảng 0,0019 trên năm hạt giống ngẫu nhiên và cao hơn đường cơ sở thống kê cổ điển 0,898. Phân tích cũng chỉ ra một lỗi quan trọng khi tái tạo mặt nạ bằng ngưỡng màu trên ảnh đã loại bỏ nền: ba mẫu tối nhất bị bỏ sót 78-96% pixel đối tượng, với hệ số tương quan giữa tỷ lệ âm tính giả và độ sáng đối tượng là r = -0,938. Việc dùng trực tiếp kênh alpha của mô hình tách nền khắc phục được hạn chế này.",
        "Đối với định vị, ngưỡng thích ứng theo từng mẫu sản phẩm làm giảm 66% diện tích dương tính giả trung bình mà không ảnh hưởng đến AUROC phân loại. Tổng hợp đa tỷ lệ không tạo cải thiện ròng so với độ phân giải 384 x 384, trong khi chi phí suy luận tăng gần ba lần. Trên tập tham chiếu 192 ảnh được chú thích thủ công, IoU trung bình đạt 0,220 và độ chính xác pointing game đạt 0,875, cho thấy hệ thống thường xác định đúng vùng lỗi tổng quát nhưng chưa mô tả chính xác biên lỗi. Kết quả khẳng định PatchCore kết hợp tiền xử lý được kiểm chứng là lựa chọn thực tế cho kiểm tra ngoại quan khi dữ liệu lỗi và nhãn pixel rất hạn chế.",
    ]
    for text in paragraphs:
        add_body_paragraph(doc, text)
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    set_run_font(p.add_run("Từ khóa: "), bold=True)
    set_run_font(p.add_run("phát hiện bất thường; PatchCore; thị giác máy tính; kiểm tra ngoại quan; dữ liệu khan hiếm; OC-MAR."))


def add_english_abstract(doc):
    add_heading(doc, "ABSTRACT", 1)
    paragraphs = [
        "This graduation project investigates visual anomaly detection and localization on a real industrial dataset containing 12 exported decorative product types. Each product type provides 100 normal images, while 30 to 53 defective images are reserved exclusively for evaluation. No defect labels or pixel-level masks are used during model training. Under this data-scarce setting, PatchCore is selected as the primary detector because it builds a memory bank from normal-image features without gradient-based training of the feature extractor.",
        "The proposed workflow follows the nine-stage OC-MAR framework, including orientation normalization, background removal, object-mask extraction, multi-resolution inference, PatchCore scoring, foreground-constrained thresholding, adaptive threshold calibration, connected-component filtering, and final defect-mask generation. The final configuration reaches a mean image-level AUROC of 0.996 and remains stable across five random seeds, clearly outperforming the classical Isolation Forest baseline with a mean AUROC of 0.898.",
        "The study also identifies an important preprocessing failure: reconstructing an object mask by applying a color threshold to a background-removed image misses 78-96% of true object pixels for the three darkest product types. Using the segmentation model's alpha output resolves this problem. Class-adaptive thresholding reduces the average false-positive area by 66% without changing classification AUROC, whereas maximum-based multi-scale aggregation does not improve AUROC over a carefully selected 384 x 384 input. On a manually annotated reference subset of 192 defective images, the system obtains a mean IoU of 0.220 and a pointing-game accuracy of 0.875, indicating reliable coarse defect localization but limited boundary precision.",
    ]
    for text in paragraphs:
        add_body_paragraph(doc, text)
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    set_run_font(p.add_run("Keywords: "), bold=True)
    set_run_font(p.add_run("anomaly detection; PatchCore; computer vision; visual inspection; data scarcity; OC-MAR."))


def add_intro(doc):
    add_heading(doc, "CHƯƠNG 1. MỞ ĐẦU", 1)
    add_heading(doc, "1.1. Lý do chọn đề tài", 2)
    add_body_paragraph(doc, "Kiểm tra ngoại quan là công đoạn quan trọng đối với sản phẩm đồ trang trí xuất khẩu, nơi sai lệch về màu sắc, chi tiết in hoặc bộ phận bị thiếu có thể làm giảm chất lượng lô hàng. Kiểm tra thủ công phụ thuộc nhiều vào kinh nghiệm, dễ bị ảnh hưởng bởi mệt mỏi và khó duy trì tiêu chuẩn đồng nhất khi số lượng mẫu sản phẩm thay đổi. Trong khi đó, các hệ thống học sâu giám sát thường cần số lượng lớn ảnh lỗi đã gán nhãn, điều không phù hợp với dây chuyền quy mô nhỏ vì lỗi hiếm, hình dạng lỗi khó dự đoán và chi phí chú thích pixel cao.")
    add_body_paragraph(doc, "Bối cảnh thực tế của chuyên đề chỉ cung cấp 100 ảnh bình thường cho mỗi trong 12 mẫu sản phẩm và 30-53 ảnh lỗi để đánh giá. Đối tượng không cố định hoàn toàn về vị trí, tỷ lệ và hướng chụp. Vì vậy, đề tài tập trung vào phát hiện bất thường một lớp: mô hình học biểu diễn từ ảnh bình thường, sau đó nhận biết ảnh và vùng có đặc trưng lệch khỏi phân bố chuẩn.")
    add_heading(doc, "1.2. Mục tiêu nghiên cứu", 2)
    for item in [
        "Xây dựng quy trình phát hiện ảnh bình thường và ảnh có sai lệch ngoại quan cho từng mẫu sản phẩm.",
        "Định vị vùng nghi ngờ lỗi ở mức pixel mà không sử dụng mặt nạ lỗi trong giai đoạn huấn luyện.",
        "Đánh giá ảnh hưởng của loại bỏ nền, độ phân giải, giới hạn tiền cảnh, đa tỷ lệ và ngưỡng thích ứng.",
        "So sánh PatchCore với đường cơ sở thống kê cổ điển và ghi nhận kết quả thử nghiệm SuperSimpleNet.",
        "Xác định các giới hạn thực tế và đề xuất hướng hoàn thiện để có thể triển khai trong sản xuất.",
    ]:
        add_body_paragraph(doc, item, style="List Bullet")
    add_heading(doc, "1.3. Đối tượng và phạm vi nghiên cứu", 2)
    add_body_paragraph(doc, "Đối tượng nghiên cứu là ảnh JPEG của 12 mã sản phẩm FA-504, FA-808, FA-838, FA-924, FA-1041, FA-1042, FA-1043, FA-1045, FA-1047, FA-1113, FA-1114 và FA-1120. Nghiên cứu giới hạn ở phát hiện bất thường ngoại quan trên ảnh tĩnh, không thiết kế cơ khí dây chuyền, không đánh giá tốc độ thời gian thực trên thiết bị nhúng và không sử dụng ảnh lỗi để huấn luyện mô hình chính.")
    add_heading(doc, "1.4. Phương pháp nghiên cứu", 2)
    add_body_paragraph(doc, "Nghiên cứu kết hợp khảo sát tài liệu, phân tích dữ liệu thăm dò, xây dựng mô hình thực nghiệm và so sánh định lượng. PatchCore được triển khai độc lập theo từng mẫu sản phẩm; 80 ảnh bình thường dùng để xây dựng ngân hàng bộ nhớ và 20 ảnh bình thường giữ lại cho hiệu chuẩn. Toàn bộ ảnh lỗi chỉ tham gia đánh giá. Các chỉ số chính gồm AUROC cấp ảnh, IoU, độ phủ vùng và pointing game. Phân tích định tính dùng bản đồ nhiệt và đường biên vùng được giữ lại.")
    add_heading(doc, "1.5. Đóng góp của chuyên đề", 2)
    for item in [
        "Hoàn thiện quy trình OC-MAR chín giai đoạn cho dữ liệu công nghiệp khan hiếm, đạt AUROC trung bình 0,996.",
        "Phát hiện có hệ thống lỗi tạo mặt nạ trên vật thể tối và chứng minh mối liên hệ với độ sáng đối tượng.",
        "Đánh giá riêng tác động của đa tỷ lệ và ngưỡng thích ứng, tránh quy kết cải thiện khi chưa có bằng chứng.",
        "Xây dựng tập tham chiếu 192 ảnh chú thích thủ công để bổ sung đánh giá định vị định lượng.",
        "Báo cáo minh bạch kết quả không hội tụ của SuperSimpleNet trong điều kiện 80 ảnh huấn luyện mỗi mẫu.",
    ]:
        add_body_paragraph(doc, item, style="List Bullet")
    add_heading(doc, "1.6. Bố cục báo cáo", 2)
    add_body_paragraph(doc, "Báo cáo gồm tám chương. Sau phần mở đầu, Chương 2 tổng quan các hướng nghiên cứu liên quan; Chương 3 mô tả dữ liệu và phát biểu bài toán; Chương 4 trình bày phương pháp và thiết kế quy trình; Chương 5 nêu thiết lập thực nghiệm; Chương 6 báo cáo kết quả; Chương 7 thảo luận các phát hiện, giới hạn và khả năng triển khai; Chương 8 kết luận và đề xuất hướng phát triển. Tài liệu tham khảo và phụ lục được trình bày ở cuối báo cáo.")


def build_report():
    crop_images()
    source = Document(SOURCE_PATH)
    doc = Document()
    configure_styles(doc)
    configure_section(doc.sections[0], body=False)
    add_front_matter(doc)

    body_section = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section(body_section, body=True)
    body_section.header.is_linked_to_previous = False
    set_page_number_start(body_section, 1)
    add_page_number_footer(body_section)

    add_abstract(doc)
    add_english_abstract(doc)
    add_intro(doc)

    add_heading(doc, "CHƯƠNG 2. TỔNG QUAN NGHIÊN CỨU", 1)
    chapter2 = [
        ("2.1. Phát hiện bất thường dựa trên ngân hàng bộ nhớ và đặc trưng nhúng", 36, 45),
        ("2.2. Phương pháp tái cấu trúc và huấn luyện khuyết tật tổng hợp", 46, 57),
        ("2.3. Đường cơ sở thống kê cổ điển", 58, 63),
        ("2.4. Bộ dữ liệu chuẩn và khác biệt của bối cảnh nghiên cứu", 64, 85),
    ]
    for heading, start, end in chapter2:
        add_heading(doc, heading, 2)
        add_source_range(doc, source, start, end)
    add_figure(doc, 2, "Hình 2.1. Ví dụ ảnh bình thường và ảnh lỗi của một số mẫu sản phẩm.", 13.5)
    rows = [
        ["Nhóm phương pháp", "Dữ liệu huấn luyện", "Ưu điểm", "Hạn chế trong đề tài"],
        ["PatchCore", "Chỉ ảnh bình thường", "Không huấn luyện gradient; phù hợp dữ liệu ít", "Bản đồ định vị còn thô"],
        ["DRAEM / SuperSimpleNet", "Ảnh bình thường và lỗi giả", "Có đầu phân đoạn trực tiếp", "Khó hội tụ khi mỗi mẫu chỉ có 80 ảnh huấn luyện"],
        ["Isolation Forest", "Thống kê thủ công từ ảnh bình thường", "Đơn giản; chi phí thấp", "Khả năng khái quát không đồng đều giữa các mẫu"],
    ]
    add_table(doc, "Bảng 2.1. So sánh các nhóm phương pháp được xem xét.", [1700, 1900, 2600, 2588], rows)

    add_heading(doc, "CHƯƠNG 3. DỮ LIỆU VÀ PHÁT BIỂU BÀI TOÁN", 1)
    chapter3 = [
        ("3.1. Tập dữ liệu", 87, 102),
        ("3.2. Sự thay đổi vị trí và kích thước đối tượng", 103, 109),
        ("3.3. Quy mô khuyết tật", 110, 114),
        ("3.4. Mất cân bằng dữ liệu", 115, 127),
        ("3.5. Phát biểu bài toán", 128, 132),
    ]
    for heading, start, end in chapter3:
        add_heading(doc, heading, 2)
        add_source_range(doc, source, start, end)
        if heading.startswith("3.2"):
            add_figure(doc, 3, "Hình 3.1. Phân bố vị trí và kích thước hộp bao đối tượng trong dữ liệu.", 12.5)
        if heading.startswith("3.4"):
            add_figure(doc, 4, "Hình 3.2. Số lượng ảnh bình thường và ảnh lỗi theo từng mẫu sản phẩm.", 13.5)
    rows = [["Thuộc tính", "Giá trị / mô tả"],
            ["Số mẫu sản phẩm", "12 mã sản phẩm độc lập"],
            ["Ảnh bình thường", "100 ảnh cho mỗi mẫu; 80 ảnh huấn luyện và 20 ảnh hiệu chuẩn"],
            ["Ảnh lỗi", "30-53 ảnh cho mỗi mẫu; chỉ dùng đánh giá"],
            ["Tổng số ảnh", "1.705 ảnh JPEG, dung lượng khoảng 474 MB"],
            ["Độ phân giải phổ biến", "956 x 1.276 pixel; một phần ảnh xoay ngang"],
            ["Nhãn pixel", "Không có trong dữ liệu gốc"],
            ["Nhiệm vụ", "Phân loại cấp ảnh và định vị vùng lỗi"]]
    add_table(doc, "Bảng 3.1. Tóm tắt tập dữ liệu nghiên cứu.", [2400, 6388], rows)

    add_heading(doc, "CHƯƠNG 4. PHƯƠNG PHÁP VÀ THIẾT KẾ HỆ THỐNG", 1)
    chapter4 = [
        ("4.1. Tổng quan quy trình OC-MAR", 134, 141),
        ("4.2. Tiền xử lý ảnh", 142, 155),
        ("4.3. Loại bỏ nền và tạo mặt nạ đối tượng", 156, 166),
        ("4.4. Mô hình PatchCore", 167, 194),
        ("4.5. Đầu vào và tổng hợp đa tỷ lệ", 195, 204),
        ("4.6. Lọc giới hạn đối tượng và tinh chỉnh vùng", 205, 217),
        ("4.7. Phân ngưỡng", 218, 224),
    ]
    for heading, start, end in chapter4:
        add_heading(doc, heading, 2)
        add_source_range(doc, source, start, end)
        if heading.startswith("4.1"):
            add_figure(doc, 5, "Hình 4.1. Quy trình OC-MAR gồm chín giai đoạn.", 14.2)
        if heading.startswith("4.3"):
            add_figure(doc, 8, "Hình 4.2. So sánh mặt nạ alpha và mặt nạ ngưỡng màu trên vật thể tối.", 13.6)
        if heading.startswith("4.4"):
            add_figure(doc, 6, "Hình 4.3. Nguyên lý xây dựng ngân hàng bộ nhớ và suy luận của PatchCore.", 14.2)
    stages = [
        ["1", "Tiếp nhận ảnh", "Chuẩn hóa hướng và kích thước"],
        ["2", "Tách nền", "Lấy ảnh đối tượng và mặt nạ alpha"],
        ["3", "Tạo nhiều tỷ lệ", "256, 384 và 512 pixel"],
        ["4", "Phát hiện bất thường", "PatchCore độc lập theo mẫu sản phẩm"],
        ["5", "Tổng hợp", "Đưa bản đồ về cùng kích thước và lấy cực đại"],
        ["6", "Giới hạn tiền cảnh", "Loại phản hồi ngoài mặt nạ đối tượng"],
        ["7", "Ngưỡng thích ứng", "Hiệu chuẩn phân vị theo từng mẫu"],
        ["8", "Tinh chỉnh vùng", "Giữ thành phần liên thông lớn nhất"],
        ["9", "Xuất kết quả", "Điểm cấp ảnh, bản đồ nhiệt và mặt nạ nhị phân"],
    ]
    add_table(doc, "Bảng 4.1. Chức năng của các giai đoạn OC-MAR.", [700, 2500, 5588], [["Giai đoạn", "Tên", "Vai trò"]] + stages)

    add_heading(doc, "CHƯƠNG 5. THIẾT LẬP THỰC NGHIỆM", 1)
    chapter5 = [
        ("5.1. Quy trình phân chia dữ liệu", 226, 230),
        ("5.2. Chỉ số đánh giá", 231, 237),
        ("5.3. Tỷ lệ coreset", 238, 241),
        ("5.4. Các cấu hình so sánh", 242, 255),
        ("5.5. Đường cơ sở thống kê cổ điển", 256, 267),
        ("5.6. Phương pháp thay thế", 268, 272),
    ]
    for heading, start, end in chapter5:
        add_heading(doc, heading, 2)
        add_source_range(doc, source, start, end)
    rows = [
        ["Cấu hình", "Ảnh đầu vào", "Độ phân giải", "Hậu xử lý"],
        ["C1 - Cơ sở", "Ảnh gốc", "224 x 224", "Ngưỡng phân vị toàn ảnh"],
        ["C2 - Tách nền", "Ảnh đã loại nền", "224 x 224", "Ngưỡng phân vị toàn ảnh"],
        ["C3 - Giới hạn tiền cảnh", "Ảnh đã loại nền", "224 x 224", "Mặt nạ alpha và thành phần lớn nhất"],
        ["C4 - Cấu hình cuối", "Ảnh đã loại nền", "384 x 384", "Giới hạn tiền cảnh và thành phần lớn nhất"],
        ["C5 - Đa tỷ lệ", "Ảnh đã loại nền", "256/384/512", "Tổng hợp cực đại và giới hạn tiền cảnh"],
    ]
    add_table(doc, "Bảng 5.1. Các cấu hình thực nghiệm chính.", [1200, 2200, 1700, 3688], rows)

    add_heading(doc, "CHƯƠNG 6. KẾT QUẢ VÀ ĐÁNH GIÁ", 1)
    add_heading(doc, "6.1. Hiệu suất phát hiện cơ sở", 2)
    add_source_range(doc, source, 274, 294)
    add_heading(doc, "6.2. So sánh với đường cơ sở thống kê", 2)
    add_source_range(doc, source, 295, 306)
    classic = [
        ["Mẫu", "Isolation Forest", "PatchCore"],
        ["FA-504", "0,993", "1,000"], ["FA-808", "0,799", "0,970"],
        ["FA-838", "0,814", "0,996"], ["FA-924", "0,799", "1,000"],
        ["FA-1041", "0,928", "0,996"], ["FA-1042", "0,845", "0,981"],
        ["FA-1043", "0,965", "0,991"], ["FA-1045", "0,833", "0,995"],
        ["FA-1047", "0,958", "0,964"], ["FA-1113", "0,946", "1,000"],
        ["FA-1114", "0,978", "1,000"], ["FA-1120", "0,922", "1,000"],
        ["Trung bình", "0,898", "0,991"],
    ]
    add_table(doc, "Bảng 6.1. AUROC của Isolation Forest và PatchCore cơ sở.", [2800, 2994, 2994], classic)

    add_heading(doc, "6.3. Ảnh hưởng của bước loại bỏ nền", 2)
    add_source_range(doc, source, 307, 320)
    add_heading(doc, "6.4. Sai lệch mặt nạ đối với vật thể tối", 2)
    add_source_range(doc, source, 321, 364)
    mask = [
        ["Mẫu", "Tỷ lệ âm tính giả", "Độ sáng trung bình"],
        ["FA-1120", "0,963", "4,7"], ["FA-1047", "0,825", "28,7"],
        ["FA-1114", "0,780", "31,6"], ["FA-1113", "0,064", "133,2"],
        ["FA-1042", "0,047", "132,8"], ["FA-1043", "0,042", "139,5"],
        ["FA-1045", "0,040", "132,4"], ["FA-838", "0,007", "152,5"],
        ["FA-1041", "0,003", "171,6"], ["FA-924", "0,003", "128,7"],
        ["FA-504", "0,003", "117,6"], ["FA-808", "0,002", "200,9"],
    ]
    add_table(doc, "Bảng 6.2. Sai lệch của mặt nạ ngưỡng màu so với mặt nạ alpha.", [2900, 2944, 2944], mask)
    add_figure(doc, 8, "Hình 6.1. Vật thể tối bị bỏ sót khi tái tạo mặt nạ bằng ngưỡng màu.", 13.6)

    add_heading(doc, "6.5. Cấu hình cuối cùng và độ phân giải đầu vào", 2)
    add_source_range(doc, source, 365, 413)
    final_auc = [
        ["Mẫu", "AUROC", "Mẫu", "AUROC"],
        ["FA-504", "1,000", "FA-1043", "0,996"],
        ["FA-808", "0,983", "FA-1045", "1,000"],
        ["FA-838", "1,000", "FA-1047", "1,000"],
        ["FA-924", "1,000", "FA-1113", "0,995"],
        ["FA-1041", "1,000", "FA-1114", "0,994"],
        ["FA-1042", "1,000", "FA-1120", "0,987"],
        ["Trung bình", "0,996", "", ""],
    ]
    add_table(doc, "Bảng 6.3. AUROC của cấu hình cuối cùng theo từng mẫu sản phẩm.", [2100, 1600, 2100, 2988], final_auc)

    add_heading(doc, "6.6. Độ ổn định theo hạt giống ngẫu nhiên", 2)
    add_source_range(doc, source, 414, 433)
    seeds = [
        ["Mẫu", "AUROC trung bình", "Độ lệch chuẩn"],
        ["FA-1120", "0,982", "0,014"], ["FA-1042", "0,992", "0,005"],
        ["FA-924", "0,992", "0,011"], ["FA-1113", "0,993", "0,005"],
        ["FA-808", "0,996", "0,008"], ["FA-1114", "0,998", "0,003"],
        ["FA-1045", "0,999", "0,002"], ["FA-1043", "0,999", "0,001"],
        ["FA-1047", "1,000", "0,000"], ["FA-1041", "1,000", "0,000"],
        ["FA-504", "1,000", "0,000"], ["FA-838", "1,000", "0,000"],
    ]
    add_table(doc, "Bảng 6.4. AUROC trung bình và độ lệch chuẩn trên năm hạt giống.", [2900, 2944, 2944], seeds)

    add_heading(doc, "6.7. Đánh giá tổng hợp đa tỷ lệ", 2)
    add_source_range(doc, source, 434, 465)
    multiscale = [
        ["Mẫu", "384 x 384", "Đa tỷ lệ"],
        ["FA-504", "1,000", "1,000"], ["FA-808", "1,000", "1,000"],
        ["FA-838", "1,000", "1,000"], ["FA-924", "1,000", "1,000"],
        ["FA-1041", "1,000", "0,999"], ["FA-1042", "0,995", "0,994"],
        ["FA-1043", "0,994", "0,999"], ["FA-1045", "0,997", "0,998"],
        ["FA-1047", "0,999", "0,999"], ["FA-1113", "1,000", "0,993"],
        ["FA-1114", "0,995", "0,995"], ["FA-1120", "0,977", "0,973"],
        ["Trung bình", "0,996", "0,996"],
    ]
    add_table(doc, "Bảng 6.5. So sánh đơn tỷ lệ và tổng hợp đa tỷ lệ.", [2900, 2944, 2944], multiscale)
    add_figure(doc, 12, "Hình 6.2. Bản đồ bất thường theo từng tỷ lệ và kết quả tổng hợp.", 14.2)

    add_heading(doc, "6.8. Ngưỡng thích ứng theo từng mẫu", 2)
    add_source_range(doc, source, 466, 505)
    adaptive = [
        ["Mẫu", "Phân vị chọn", "FP tại p92", "FP sau hiệu chỉnh"],
        ["FA-1114", "99", "1.244,5", "196,4"], ["FA-1113", "99", "1.171,8", "187,3"],
        ["FA-1047", "98", "846,3", "253,1"], ["FA-1043", "98", "634,8", "219,1"],
        ["FA-1042", "98", "618,4", "223,0"], ["FA-1041", "98", "562,9", "175,2"],
        ["FA-1045", "96", "577,9", "293,7"], ["FA-808", "98", "543,7", "195,8"],
        ["FA-1120", "98", "530,7", "163,2"], ["FA-924", "96", "432,7", "260,2"],
        ["FA-838", "96", "410,7", "240,2"], ["FA-504", "94", "374,5", "288,0"],
        ["Trung bình", "-", "662,4", "224,6"],
    ]
    add_table(doc, "Bảng 6.6. Hiệu quả ngưỡng thích ứng theo mẫu sản phẩm.", [2100, 1600, 2500, 2588], adaptive)

    add_heading(doc, "6.9. Đánh giá định tính khả năng định vị", 2)
    add_source_range(doc, source, 506, 541)
    add_figure_pair(
        doc, 16, 17,
        "Hình 6.3. Kết quả định vị trên ảnh lỗi: (a) FA-808; (b) FA-1047.",
        7.0,
    )
    add_figure_pair(
        doc, 18, 19,
        "Hình 6.4. Đối chiếu định vị: (a) ảnh bình thường FA-808; (b) các mẫu sản phẩm còn lại.",
        7.0,
    )

    add_heading(doc, "6.10. Đánh giá định lượng trên tập chú thích thủ công", 2)
    add_source_range(doc, source, 542, 594)
    localization = [
        ["Mẫu", "n", "IoU", "Độ phủ", "Pointing game"],
        ["FA-808", "20", "0,323", "0,564", "0,800"], ["FA-1041", "20", "0,289", "0,361", "0,950"],
        ["FA-1043", "16", "0,258", "0,329", "1,000"], ["FA-1047", "12", "0,222", "0,229", "1,000"],
        ["FA-1042", "16", "0,214", "0,369", "1,000"], ["FA-1113", "12", "0,212", "0,225", "1,000"],
        ["FA-1045", "12", "0,209", "0,237", "1,000"], ["FA-504", "20", "0,206", "0,352", "0,700"],
        ["FA-924", "12", "0,199", "0,242", "0,917"], ["FA-838", "20", "0,194", "0,269", "0,900"],
        ["FA-1120", "12", "0,179", "0,206", "1,000"], ["FA-1114", "20", "0,116", "0,213", "0,500"],
        ["Trung bình", "192", "0,220", "0,313", "0,875"],
    ]
    add_table(doc, "Bảng 6.7. Chỉ số định vị trên 192 ảnh được chú thích thủ công.", [1850, 800, 1700, 1800, 2638], localization)
    add_figure(doc, 20, "Hình 6.5. Ví dụ đối chiếu vùng dự đoán với mặt nạ chú thích thủ công.", 10.5)

    add_heading(doc, "CHƯƠNG 7. THẢO LUẬN", 1)
    add_heading(doc, "7.1. Ý nghĩa của kết quả", 2)
    add_source_range(doc, source, 595, 620)
    add_heading(doc, "7.2. Ảnh hưởng của tiền xử lý và hậu xử lý", 2)
    add_source_range(doc, source, 620, 628)
    add_heading(doc, "7.3. Lựa chọn giữa mô hình không huấn luyện và mô hình có huấn luyện", 2)
    add_source_range(doc, source, 628, 637)
    add_heading(doc, "7.4. Khả năng ứng dụng và giới hạn", 2)
    add_body_paragraph(doc, "Kết quả cho thấy quy trình phù hợp cho bước sàng lọc hoặc hỗ trợ người kiểm tra trong môi trường sản xuất có ít ảnh lỗi. Tuy nhiên, AUROC cao không đồng nghĩa vùng lỗi được phân đoạn chính xác. Khi triển khai thực tế cần tách hai ngưỡng: ngưỡng cảnh báo cấp ảnh và ngưỡng hiển thị vùng lỗi; đồng thời lưu lại ảnh mới để giám sát sự thay đổi của điều kiện chiếu sáng, phông nền và thiết kế sản phẩm.")
    add_body_paragraph(doc, "Các giới hạn chính gồm quy mô dữ liệu nhỏ, không có mặt nạ lỗi do chuyên gia xây dựng, chưa đo độ đồng thuận giữa người chú thích, chưa đánh giá tốc độ suy luận trên dây chuyền và chưa thực hiện kiểm thử theo thời gian. Tập tham chiếu 192 ảnh có chủ ý được chú thích thô nên các giá trị IoU và độ phủ chỉ nên được hiểu như giới hạn dưới. Việc mở rộng kết luận sang sản phẩm mới cần thêm thí nghiệm độc lập.")
    add_heading(doc, "7.5. Khuyến nghị triển khai thử nghiệm", 2)
    add_body_paragraph(doc, "Giai đoạn thử nghiệm nên bắt đầu ở một trạm chụp cố định với nền, khoảng cách camera và điều kiện chiếu sáng được kiểm soát. Mỗi mã sản phẩm cần có ngân hàng bộ nhớ riêng và bộ ảnh hiệu chuẩn bình thường được rà soát trước khi đưa vào sử dụng. Hệ thống nên lưu đồng thời ảnh gốc, điểm bất thường, bản đồ nhiệt, mặt nạ đối tượng và quyết định cuối để có thể truy vết khi phát sinh cảnh báo sai. Khi đổi thiết kế sản phẩm, vật liệu, camera hoặc bố trí ánh sáng, dữ liệu phải được đánh giá lại thay vì tiếp tục dùng ngưỡng cũ.")
    add_body_paragraph(doc, "Trong vận hành, mô hình nên đóng vai trò sàng lọc và cung cấp vùng nghi ngờ cho nhân viên kiểm tra, chưa nên tự động loại sản phẩm chỉ từ một dự đoán. Các cảnh báo cần được phân nhóm thành đúng lỗi, báo động giả và trường hợp chưa chắc chắn; số liệu này được tổng hợp định kỳ để theo dõi tỷ lệ báo động giả theo từng mã hàng. Chỉ nên mở rộng sang nhiều trạm sau khi ngưỡng cảnh báo đạt độ ổn định trong một giai đoạn chạy song song với quy trình thủ công và các trường hợp lỗi quan trọng không bị bỏ sót.")

    add_heading(doc, "CHƯƠNG 8. KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", 1)
    add_heading(doc, "8.1. Kết luận", 2)
    add_source_range(doc, source, 638, 665)
    add_heading(doc, "8.2. Hướng phát triển", 2)
    add_source_range(doc, source, 665, 672)
    for item in [
        "Xây dựng bộ mặt nạ lỗi do nhiều chuyên gia chú thích và đo mức độ đồng thuận giữa người chú thích.",
        "Đánh giá PRO đa ngưỡng, đường cong precision-recall và chi phí lỗi theo yêu cầu kiểm soát chất lượng.",
        "Thử nghiệm mô hình trên dữ liệu thu thập theo thời gian để phát hiện trôi dữ liệu và thay đổi điều kiện chụp.",
        "Tối ưu hóa ngân hàng bộ nhớ, thời gian suy luận và cơ chế cập nhật đối với mã sản phẩm mới.",
        "Thiết kế giao diện xác nhận cảnh báo để phản hồi của nhân viên kiểm tra trở thành dữ liệu hiệu chuẩn có kiểm soát.",
    ]:
        add_body_paragraph(doc, item, style="List Bullet")

    add_heading(doc, "TÀI LIỆU THAM KHẢO", 1)
    references = [
        "[1] P. Bergmann, M. Fauser, D. Sattlegger và C. Steger, 'MVTec AD - A Comprehensive Real-World Dataset for Unsupervised Anomaly Detection,' Proc. IEEE/CVF CVPR, pp. 9584-9592, 2019.",
        "[2] F. T. Liu, K. M. Ting và Z.-H. Zhou, 'Isolation Forest,' Proc. IEEE International Conference on Data Mining, 2008.",
        "[3] B. Rolih, M. Fučka và D. Skočaj, 'SuperSimpleNet: Unifying Unsupervised and Supervised Learning for Fast and Reliable Surface Defect Detection,' arXiv:2408.03143, 2024.",
        "[4] K. Roth, L. Pemula, J. Zepeda, B. Schölkopf, T. Brox và P. Gehler, 'Towards Total Recall in Industrial Anomaly Detection,' Proc. IEEE/CVF CVPR, 2022.",
        "[5] V. Zavrtanik, M. Kristan và D. Skočaj, 'DRAEM - A Discriminatively Trained Reconstruction Embedding for Surface Anomaly Detection,' Proc. IEEE/CVF ICCV, pp. 8330-8339, 2021.",
        "[6] J. Zhang, S. A. Bargal, Z. Lin, J. Brandt, X. Shen và S. Sclaroff, 'Top-Down Neural Attention by Excitation Backprop,' International Journal of Computer Vision, vol. 126, no. 10, pp. 1084-1102, 2018.",
    ]
    for ref in references:
        p = add_body_paragraph(doc, ref)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.left_indent = Cm(0.75)

    add_heading(doc, "PHỤ LỤC A. THỬ NGHIỆM SUPERSIMPLENET", 1)
    add_heading(doc, "A.1. Thiết lập", 2)
    add_body_paragraph(doc, "SuperSimpleNet được thử nghiệm như một phương án có thể huấn luyện nhằm tạo trực tiếp bản đồ phân đoạn lỗi. Hai mẫu có AUROC PatchCore cơ sở thấp nhất, FA-1047 và FA-808, được chọn để đánh giá. Mỗi mô hình dùng 80 ảnh bình thường, WideResNet-50 làm bộ trích xuất đặc trưng đóng băng, lỗi giả tổng hợp bằng nhiễu Gaussian có mặt nạ Perlin và lịch huấn luyện 300 epoch. Ảnh lỗi thực và mặt nạ lỗi không tham gia huấn luyện.")
    add_heading(doc, "A.2. Hiệu chỉnh theo triển khai tham chiếu", 2)
    add_body_paragraph(doc, "Ba vòng hiệu chỉnh được thực hiện bằng cách đối chiếu với mã tham chiếu. Vòng đầu sửa hàm mất mát phân đoạn từ khoảng cách L1 trên đầu ra sigmoid sang hàm biên độ bản lề trên logit. Vòng tiếp theo bổ sung cơ chế dừng gradient giữa hai đầu ra, dùng đặc trưng xương sống thô cho đầu phân loại, thay đầu phân đoạn bằng khối tích chập 1 x 1 hai lớp có chiều ẩn 1.024 và bổ sung lịch giảm tốc độ học tại 80% và 90% tiến trình. Các thay đổi này loại trừ phần lớn sai khác kiến trúc có thể xác định được.")
    add_heading(doc, "A.3. Phân tích kết quả không hội tụ", 2)
    add_body_paragraph(doc, "Sau hiệu chỉnh, tổn thất giảm mạnh trong khoảng 20 epoch đầu rồi dao động quanh 1,05-1,11 cho đến hết 300 epoch; AUROC cuối đạt 0,384 trên FA-1047 và 0,891 trên FA-808. Quỹ đạo tổn thất tương tự nhưng chất lượng hai mẫu khác xa nhau cho thấy mô hình chưa học ổn định. Báo cáo xem đây là kết quả tiêu cực có giá trị: với 80 ảnh bình thường cho mỗi mẫu, phương pháp có thể huấn luyện này chưa chứng minh được lợi thế so với PatchCore. Nghiên cứu tiếp theo cần kiểm tra cơ chế sao chép theo lô trong bước chèn lỗi giả và đánh giá trên quy mô 200-300 ảnh bình thường cho mỗi mẫu.")
    add_figure(doc, 21, "Hình A.1. Hàm mất mát huấn luyện SuperSimpleNet trên FA-1047 và FA-808.", 8.0)

    core_props = doc.core_properties
    core_props.title = "Báo cáo chuyên đề tốt nghiệp - Phát hiện sai lệch ngoại quan"
    core_props.subject = "Thị giác máy tính và phát hiện bất thường công nghiệp"
    core_props.author = "[HỌ VÀ TÊN SINH VIÊN]"
    core_props.keywords = "PatchCore, anomaly detection, computer vision, OC-MAR"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT_PATH)
    print(OUTPUT_PATH)


if __name__ == "__main__":
    build_report()
