"""Validate a .pptx written by lib/local-pptx.js (used by tests/slides.test.js and tests/e2e/slides.mjs).

Checks: every XML part parses; every part has a content type; every internal
relationship target exists; cNvPr ids are unique per slide; python-pptx opens the
file and reads every slide, shape, text frame, table and notes page.
Prints a JSON summary; exit code 1 on any problem.
"""
import json, posixpath, sys, zipfile
from lxml import etree
from pptx import Presentation

def main(path):
    problems = []
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    ct = etree.fromstring(z.read("[Content_Types].xml"))
    ns = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    defaults = {d.get("Extension").lower() for d in ct.iter(ns + "Default")}
    overrides = {o.get("PartName").lstrip("/") for o in ct.iter(ns + "Override")}
    for o in overrides:
        if o not in names:
            problems.append(f"override for missing part {o}")
    for n in names:
        if n.endswith("/"):
            continue
        if n != "[Content_Types].xml" and n not in overrides and n.rsplit(".", 1)[-1].lower() not in defaults:
            problems.append(f"no content type: {n}")
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                etree.fromstring(z.read(n))
            except Exception as e:
                problems.append(f"bad xml {n}: {e}")
    for n in names:
        if not n.endswith(".rels"):
            continue
        base = posixpath.dirname(posixpath.dirname(n))
        for r in etree.fromstring(z.read(n)):
            if r.get("TargetMode") == "External":
                continue
            t = r.get("Target")
            full = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(base, t))
            if full not in names:
                problems.append(f"{n}: missing target {t}")
    for n in names:
        if n.startswith("ppt/slides/slide") and n.endswith(".xml"):
            root = etree.fromstring(z.read(n))
            ids = [e.get("id") for e in root.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")]
            if len(ids) != len(set(ids)):
                problems.append(f"duplicate shape ids in {n}")
    prs = Presentation(path)
    slides = []
    for s in prs.slides:
        info = {"layout": s.slide_layout.name, "shapes": [], "notes": None, "hidden": s._element.get("show") == "0"}
        for sh in s.shapes:
            d = {"type": str(sh.shape_type), "name": sh.name, "x": sh.left, "y": sh.top, "w": sh.width, "h": sh.height}
            if sh.is_placeholder:
                d["ph"] = str(sh.placeholder_format.type)
            if sh.has_text_frame:
                d["text"] = sh.text_frame.text
            if getattr(sh, "has_table", False) and sh.has_table:
                d["table"] = [[c.text for c in r.cells] for r in sh.table.rows]
            if sh.shape_type is not None and "PICTURE" in str(sh.shape_type):
                d["image"] = sh.image.content_type
            info["shapes"].append(d)
        if s.has_notes_slide:
            info["notes"] = s.notes_slide.notes_text_frame.text
        slides.append(info)
    out = {"ok": not problems, "problems": problems, "slides": slides,
           "size": [prs.slide_width, prs.slide_height],
           "title": prs.core_properties.title, "author": prs.core_properties.author}
    print(json.dumps(out))
    return 0 if not problems else 1

if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
