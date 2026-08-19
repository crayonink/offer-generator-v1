"""Check the things Word refuses to open a .docx over.

Not a schema validator — a check of the handful of structural rules that a
generated document actually trips: a container ending in a table, a
relationship id with nothing behind it, a part with no declared content type.
Run it against a known-good document too, so a clean result means something.
"""
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def check(path):
    faults = []
    z = zipfile.ZipFile(path)

    if z.testzip():
        faults.append("zip is corrupt")

    # 1. every xml part parses
    parts = {}
    for n in z.namelist():
        if not n.endswith((".xml", ".rels")):
            continue
        try:
            parts[n] = ET.fromstring(z.read(n))
        except Exception as e:
            faults.append(f"{n}: will not parse — {e}")

    # 2. a body, header, footer or table cell must not end with a table
    for n, root in parts.items():
        if not (n.endswith("document.xml") or "/header" in n or "/footer" in n):
            continue
        for tag in (f"{W}body", f"{W}hdr", f"{W}ftr"):
            for el in root.iter(tag):
                kids = [k for k in el if k.tag in (f"{W}p", f"{W}tbl")]
                if kids and kids[-1].tag == f"{W}tbl":
                    faults.append(f"{n}: {tag.split('}')[1]} ends with a table")
        for tc in root.iter(f"{W}tc"):
            kids = [k for k in tc if k.tag in (f"{W}p", f"{W}tbl")]
            if kids and kids[-1].tag == f"{W}tbl":
                faults.append(f"{n}: a table cell ends with a table")
            if not kids:
                faults.append(f"{n}: an empty table cell")

    # 3. every relationship id used has a relationship behind it
    for n in list(parts):
        if n.endswith(".rels"):
            continue
        rels_name = n.rsplit("/", 1)[0] + "/_rels/" + n.rsplit("/", 1)[-1] + ".rels"
        have = set()
        if rels_name in parts:
            have = {r.get("Id") for r in parts[rels_name]}
        used = set(re.findall(r'r:(?:id|embed|link)="([^"]+)"',
                              z.read(n).decode("utf-8", "replace")))
        for rid in sorted(used - have):
            faults.append(f"{n}: uses {rid} with no relationship")

    # 4. every part has a content type
    ct = z.read("[Content_Types].xml").decode("utf-8", "replace")
    defaults = set(re.findall(r'Extension="([^"]+)"', ct))
    overrides = set(re.findall(r'PartName="([^"]+)"', ct))
    for n in z.namelist():
        if n.startswith("_rels") or n.endswith("/") or n == "[Content_Types].xml":
            continue
        ext = n.rsplit(".", 1)[-1].lower()
        if ext not in defaults and "/" + n not in overrides:
            faults.append(f"{n}: no content type declared")

    # 5. every media file a relationship points at exists
    names = set(z.namelist())
    for n, root in parts.items():
        if not n.endswith(".rels"):
            continue
        base = n.split("_rels/")[0]
        for rel in root:
            tgt = rel.get("Target", "")
            if rel.get("TargetMode") == "External" or tgt.startswith("http"):
                continue
            # posixpath.normpath, so word/ + ../customXML lands at the root
            import posixpath
            resolved = posixpath.normpath(base + tgt).lstrip("/")
            if resolved not in names:
                faults.append(f"{n}: target missing — {tgt}")

    z.close()
    return faults


if __name__ == "__main__":
    worst = 0
    for path in sys.argv[1:]:
        faults = check(path)
        name = path.rsplit("\\", 1)[-1].rsplit("/", 1)[-1]
        if faults:
            worst = 1
            print(f"FAIL  {name}")
            for f in faults[:12]:
                print(f"        {f}")
        else:
            print(f"ok    {name}")
    sys.exit(worst)
