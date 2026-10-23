import sys
import zipfile
import shutil
from lxml import etree

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
M = "{%s}" % NS["m"]


def sheet_map(z):
    wbx = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    tgt = {r.get("Id"): r.get("Target") for r in rels.findall("pr:Relationship", NS)}
    out = {}
    for s in wbx.find("m:sheets", NS):
        t = tgt[s.get("{%s}id" % NS["r"])].lstrip("/")
        out[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
    return out


def shared(z):
    try:
        sst = etree.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(si.itertext()) for si in sst.findall("m:si", NS)]


src, twin, out = sys.argv[1:4]
zs, zt = zipfile.ZipFile(src), zipfile.ZipFile(twin)
ms, mt = sheet_map(zs), sheet_map(zt)
tss = shared(zt)
newparts = {}
filled = 0
for name, part in ms.items():
    sx = etree.fromstring(zs.read(part))
    tx = etree.fromstring(zt.read(mt[name]))
    tv = {}
    for c in tx.iter(M + "c"):
        v = c.find(M + "v")
        tv[c.get("r")] = (c.get("t"), v.text if v is not None else None)
    changed = False
    for c in sx.iter(M + "c"):
        if c.find(M + "f") is None or c.get("r") not in tv:
            continue
        t, val = tv[c.get("r")]
        if t == "s":
            t, val = "str", tss[int(val)]
        if t == "e":
            continue
        for old in c.findall(M + "v"):
            c.remove(old)
        v = etree.SubElement(c, M + "v")
        v.text = val if val is not None else ""
        if t in ("str", "b"):
            c.set("t", t)
        elif "t" in c.attrib:
            del c.attrib["t"]
        filled += 1
        changed = True
    if changed:
        newparts[part] = etree.tostring(sx, xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
    for item in zs.infolist():
        zo.writestr(item, newparts.get(item.filename, zs.read(item.filename)))
print("cached values written:", filled)
