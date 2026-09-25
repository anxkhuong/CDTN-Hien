from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parent


def trim(filename, chapter2, chapter3, selected):
    path = ROOT / filename
    doc = Document(path)
    ps = doc.paragraphs
    start = [i for i, p in enumerate(ps) if p.text.strip() == chapter2 and not p.style.name.lower().startswith("toc")][-1]
    end = [i for i, p in enumerate(ps) if i > start and p.text.strip() == chapter3 and not p.style.name.lower().startswith("toc")][0]
    headings = [p for p in ps[start:end] if p.text.strip().startswith("2.5")]
    if len(headings) != 27:
        raise RuntimeError(f"{filename}: không tìm đủ 27 đề mục")
    for idx in selected:
        hp = headings[idx]
        node = hp._p.getnext()
        body = []
        while node is not None:
            text = "".join(node.itertext()).strip()
            if text.startswith("2.5") or text == chapter3:
                break
            if node.tag.endswith("}p") and text and not text.startswith("Bảng 2."):
                body.append(node)
            node = node.getnext()
        if len(body) >= 3:
            body[2].getparent().remove(body[2])
        else:
            raise RuntimeError(f"{filename}: đề mục {idx} không có đủ ba đoạn")
    doc.save(path)


trim(
    "IE400.F3.CN2.CNTT_CDTN_NguyenThiHien_25410207.docx",
    "CHƯƠNG 2. TỔNG QUAN NGHIÊN CỨU VÀ CƠ SỞ LÝ THUYẾT",
    "CHƯƠNG 3. DỮ LIỆU VÀ PHÁT BIỂU BÀI TOÁN",
    [2, 5, 8, 11, 14, 17, 20, 23, 26],
)
trim(
    "hienmain (1)_L.docx",
    "2. Các công trình nghiên cứu liên quan và cơ sở lý thuyết",
    "3. Thiết lập bộ dữ liệu và bài toán",
    [0, 2, 4, 5, 7, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26],
)
print("Đã rút gọn các đoạn trùng ý để cân đối số trang.")
