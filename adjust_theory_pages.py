from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parent


def remove_selected_breaks(filename, selected, chapter2, chapter3):
    path = ROOT / filename
    doc = Document(path)
    paragraphs = doc.paragraphs
    end = next(
        i for i, p in enumerate(paragraphs)
        if p.text.strip() == chapter3 and not p.style.name.lower().startswith("toc")
    )
    chapter2_candidates = [
        i for i, p in enumerate(paragraphs[:end])
        if p.text.strip() == chapter2 and not p.style.name.lower().startswith("toc")
    ]
    start = chapter2_candidates[-1]
    headings = [p for p in paragraphs[start:end] if p.text.strip().startswith("2.5")]
    if len(headings) != 27:
        raise RuntimeError(f"{filename}: dự kiến 27 đề mục, thực tế {len(headings)}")
    for idx in selected:
        p = headings[idx]
        for br in list(p._p.iter(qn("w:br"))):
            if br.get(qn("w:type")) == "page":
                br.getparent().remove(br)
                break
    doc.save(path)


remove_selected_breaks(
    "IE400.F3.CN2.CNTT_CDTN_NguyenThiHien_25410207.docx",
    [2, 5, 8, 11, 16, 20, 23, 26],
    "CHƯƠNG 2. TỔNG QUAN NGHIÊN CỨU VÀ CƠ SỞ LÝ THUYẾT",
    "CHƯƠNG 3. DỮ LIỆU VÀ PHÁT BIỂU BÀI TOÁN",
)
remove_selected_breaks(
    "hienmain (1)_L.docx",
    [2, 4, 7, 10, 14, 16, 19, 22, 25],
    "2. Các công trình nghiên cứu liên quan và cơ sở lý thuyết",
    "3. Thiết lập bộ dữ liệu và bài toán",
)
print("Đã cân chỉnh ngắt trang cho hai tài liệu.")
