from __future__ import annotations

from pathlib import Path

from docx import Document


def squash_adjacent(values: list[str]) -> list[str]:
    result: list[str] = []
    for value in values:
        cleaned = " ".join(value.split())
        if not result or result[-1] != cleaned:
            result.append(cleaned)
    return result


doc = Document("Nội dung bài báo cáo CĐTN.docx")
lines: list[str] = []
for table_index, table in enumerate(doc.tables, 1):
    lines.append(f"TABLE {table_index}")
    for row in table.rows:
        lines.append("\t".join(squash_adjacent([cell.text for cell in row.cells])))
    lines.append("")
Path("tmp/source_extracts/content_tables.txt").write_text(
    "\n".join(lines), encoding="utf-8"
)
