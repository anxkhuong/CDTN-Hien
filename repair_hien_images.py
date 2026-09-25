from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "backups" / "hienmain (1)_L_before_theory.docx"
TARGET = ROOT / "hienmain (1)_L.docx"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


with ZipFile(TARGET, "r") as current, ZipFile(SOURCE, "r") as backup:
    entries = {name: current.read(name) for name in current.namelist()}
    backup_media = {
        name: backup.read(name)
        for name in backup.namelist()
        if name.startswith("word/media/")
    }

rels_name = "word/_rels/document.xml.rels"
root = ET.fromstring(entries[rels_name])
for rel in root.findall(f"{{{REL_NS}}}Relationship"):
    target = rel.get("Target", "")
    if rel.get("Type", "").endswith("/image") and target.startswith("ooxWord://word/media/"):
        rel.set("Target", "media/" + target.rsplit("/", 1)[-1])
        rel.attrib.pop("TargetMode", None)
entries[rels_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

ct_name = "[Content_Types].xml"
ct_root = ET.fromstring(entries[ct_name])
if not any(n.get("Extension", "").lower() in {"jpg", "jpeg"} for n in ct_root.findall(f"{{{CT_NS}}}Default")):
    ET.SubElement(ct_root, f"{{{CT_NS}}}Default", Extension="jpeg", ContentType="image/jpeg")
entries[ct_name] = ET.tostring(ct_root, encoding="utf-8", xml_declaration=True)

entries.update(backup_media)
with NamedTemporaryFile(dir=ROOT, suffix=".docx", delete=False) as tmp:
    tmp_path = Path(tmp.name)
with ZipFile(tmp_path, "w", ZIP_DEFLATED) as out:
    for name, data in entries.items():
        out.writestr(name, data)
tmp_path.replace(TARGET)
print("Đã khôi phục ảnh nhúng của hienmain.")
