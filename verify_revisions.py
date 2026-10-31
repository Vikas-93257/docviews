"""Three independent cross-checks of the revision numbers.

1. revised paragraphs: docviews vs a hand-written XML rule
2. revision containers: len(doc.revisions()) vs elements counted in the XML
3. files where the XML has markup docviews never reports (parts we cannot see)
"""
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, "src")
from docviews import open_docx                                  # noqa: E402

REV = {"ins", "del", "moveFrom", "moveTo"}


def tag(el):
    return el.tag.rsplit("}", 1)[-1]


def has_revision(el):
    for child in el:
        t = tag(child)
        if t in REV:
            return True
        if t == "txbxContent":
            continue
        if has_revision(child):
            return True
    return False


def count_paragraphs(xml: bytes) -> int:
    root = ET.fromstring(xml)
    return sum(1 for p in root.iter() if tag(p) == "p" and has_revision(p))


def count_containers(xml: bytes) -> int:
    root = ET.fromstring(xml)
    return sum(1 for el in root.iter() if tag(el) in REV)


tot = {"paras_dv": 0, "paras_xml": 0, "revs_dv": 0, "containers_xml": 0}
mismatch = []
for path in sorted(Path("corpus").glob("*.docx")):
    try:
        zf = zipfile.ZipFile(path)
    except Exception:
        continue
    parts = [n for n in zf.namelist()
             if n.startswith("word/") and n.endswith(".xml") and "/" not in n[5:]]
    xml_paras = xml_containers = 0
    for name in parts:
        try:
            data = zf.read(name)
            xml_paras += count_paragraphs(data)
            xml_containers += count_containers(data)
        except ET.ParseError:
            continue
    try:
        doc = open_docx(path)
        paras_dv = sum(1 for p in doc.iter_paragraphs("accepted") if p.is_revised)
        revs_dv = len(doc.revisions())
        seen_parts = {p.name for p in doc._package.text_parts()}
    except Exception as exc:                                    # noqa: BLE001
        if xml_containers:
            mismatch.append((path.name, "open failed", str(exc)[:40]))
        continue
    tot["paras_dv"] += paras_dv
    tot["paras_xml"] += xml_paras
    tot["revs_dv"] += revs_dv
    tot["containers_xml"] += xml_containers
    if (paras_dv, revs_dv) != (xml_paras, xml_containers):
        unseen = [n for n in parts if n not in seen_parts
                  and count_containers(zf.read(n))]
        mismatch.append((path.name, f"paras {paras_dv}/{xml_paras} "
                                    f"revs {revs_dv}/{xml_containers}",
                         f"parts not read: {unseen}"))

print("totals: revised paragraphs docviews={paras_dv} xml={paras_xml} | "
      "revisions docviews={revs_dv} containers={containers_xml}".format(**tot))
print(f"\nfiles that disagree: {len(mismatch)}")
for name, a, b in mismatch:
    print(f"  {name[:60]:60s} {a:34s} {b}")
