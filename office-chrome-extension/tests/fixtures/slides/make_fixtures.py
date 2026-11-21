"""Generate the .pptx fixtures used by tests/slides.test.js and tests/e2e/slides.mjs.

Run:  <python with python-pptx> tests/fixtures/slides/make_fixtures.py
python-pptx's default template is a real PowerPoint package (layouts, master, theme).
"""
import base64, copy, os, struct, zlib
from pptx import Presentation
from pptx.util import Pt, Emu, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))


def png_bytes(w, h, rgb):
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def add_transition(slide, kind):
    sld = slide._element
    tr = etree.SubElement(sld, "{%s}transition" % NS)
    tr.set("spd", "med")
    etree.SubElement(tr, "{%s}%s" % (NS, kind))
    # transition must come after clrMapOvr
    clr = sld.find("{%s}clrMapOvr" % NS)
    if clr is not None:
        clr.addnext(tr)


def basic():
    prs = Presentation()  # 4:3 default template, 720x540 pt
    prs.core_properties.title = "Fixture Deck"
    prs.core_properties.author = "Test Author"
    L = prs.slide_layouts

    s = prs.slides.add_slide(L[0])
    s.shapes.title.text = "Quarterly Review"
    s.placeholders[1].text = "Prepared by the team"
    s.notes_slide.notes_text_frame.text = "Welcome everyone\nSecond note line"
    add_transition(s, "fade")

    s = prs.slides.add_slide(L[1])
    s.shapes.title.text = "Agenda"
    tf = s.placeholders[1].text_frame
    tf.text = "First point"
    p = tf.add_paragraph(); p.text = "Sub point"; p.level = 1
    p = tf.add_paragraph(); p.text = "Deep point"; p.level = 2
    p = tf.add_paragraph(); p.text = "Second point"
    add_transition(s, "push")

    s = prs.slides.add_slide(L[3])  # two content
    s.shapes.title.text = "Two columns"
    s.placeholders[1].text = "Left side"
    s.placeholders[2].text = "Right side"

    s = prs.slides.add_slide(L[2])  # section header
    s.shapes.title.text = "Section Two"
    s.placeholders[1].text = "Details follow"

    s = prs.slides.add_slide(L[6])  # blank
    tb = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(5), Inches(1.5))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "Bold "; r.font.bold = True; r.font.size = Pt(28)
    r = p.add_run(); r.text = "italic "; r.font.italic = True; r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    r = p.add_run(); r.text = "under"; r.font.underline = True; r.font.name = "Georgia"
    p2 = tf.add_paragraph()
    r = p2.add_run(); r.text = "a link"; r.hyperlink.address = "https://example.com/"
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2.5), Inches(2), Inches(1))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0x11, 0x88, 0x33)
    sh.line.color.rgb = RGBColor(0, 0, 0); sh.line.width = Pt(2)
    sh.text_frame.text = "Box text"
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.5), Inches(2.5), Inches(1.5), Inches(1))
    sh.fill.background()
    sh = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.5), Inches(2.5), Inches(2), Inches(1))
    sh.rotation = 30
    sh = s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT, Inches(1), Inches(4), Inches(1), Inches(1))
    pic_path = os.path.join(HERE, "_tmp.png")
    with open(pic_path, "wb") as f:
        f.write(png_bytes(4, 3, (30, 90, 200)))
    s.shapes.add_picture(pic_path, Inches(3), Inches(4), Inches(2), Inches(1.5))
    os.remove(pic_path)
    s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(6), Inches(4), Inches(9), Inches(4))

    s = prs.slides.add_slide(L[5])  # title only + table
    s.shapes.title.text = "Table slide"
    tbl = s.shapes.add_table(3, 3, Inches(1), Inches(1.8), Inches(8), Inches(2)).table
    for r in range(3):
        for c in range(3):
            tbl.cell(r, c).text = f"R{r}C{c}"
    tbl.cell(1, 1).fill.solid(); tbl.cell(1, 1).fill.fore_color.rgb = RGBColor(0xFF, 0xEE, 0x00)
    s.background.fill.solid(); s.background.fill.fore_color.rgb = RGBColor(0xEE, 0xF4, 0xFF)

    s = prs.slides.add_slide(L[6])  # group + chart + hidden
    grp = s.shapes.add_group_shape()
    grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))
    grp.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3), Inches(1), Inches(1), Inches(1))
    cd = CategoryChartData(); cd.categories = ["A", "B"]; cd.add_series("S", (1, 2))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(5), Inches(3), Inches(4), Inches(3), cd)
    s._element.set("show", "0")

    prs.save(os.path.join(HERE, "basic.pptx"))


def wide():
    prs = Presentation()
    prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Wide deck"
    s.placeholders[1].text = "Only one bullet"
    prs.save(os.path.join(HERE, "wide.pptx"))


A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSDECL = 'xmlns:a="%s" xmlns:p="%s" xmlns:r="%s"' % (A, NS, R)


def frag(xml):
    """Parse a p:/a: fragment into an element."""
    return etree.fromstring('<p:wrap %s xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                            'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">%s</p:wrap>' % (NSDECL, xml))[0]


def complex_deck():
    """Real-world constructs the python-pptx API cannot make, written as raw XML."""
    import io
    prs = Presentation()
    prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)
    L = prs.slide_layouts
    # a decoration on the master (shown on slides unless the layout hides master shapes)
    master = prs.slide_master
    master.shapes._spTree.append(frag(
        '<p:sp><p:nvSpPr><p:cNvPr id="90" name="Brand bar"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="6604000"/><a:ext cx="12192000" cy="254000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>'))
    # "Title Only" layout hides master shapes and has its own background
    lay = L[5]
    lay._element.set("showMasterSp", "0")
    csld = lay._element.find("{%s}cSld" % NS)
    csld.insert(0, frag('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF5E6"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'))

    # slide 1: field + line break in the title, autofit body, theme color with lumMod
    s = prs.slides.add_slide(L[1])
    s.shapes.title.text = ""
    t = s.shapes.title._element.find(".//{%s}txBody" % NS)
    for old in t.findall("{%s}p" % A):
        t.remove(old)
    t.append(frag('<a:p><a:r><a:rPr lang="en-US"/><a:t>Results</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br>'
                  '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld></a:p>'))
    body = s.placeholders[1]
    body.text_frame.text = "Shrunk text"
    bp = body._element.find(".//{%s}bodyPr" % A)
    bp.append(frag('<a:normAutofit fontScale="50000" lnSpcReduction="20000"/>'))
    r = body._element.find(".//{%s}r" % A)
    r.insert(0, frag('<a:rPr lang="en-US"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:rPr>'))

    # slide 2: gradient shape, SVG picture (PNG fallback), SmartArt frame, merged cells, rotated group, freeform, vertical text
    s = prs.slides.add_slide(L[6])
    tree = s.shapes._spTree
    tree.append(frag(
        '<p:sp><p:nvSpPr><p:cNvPr id="20" name="Grad"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="254000" y="254000"/><a:ext cx="2540000" cy="1270000"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill></p:spPr></p:sp>'))
    pic = s.shapes.add_picture(io.BytesIO(png_bytes(2, 2, (0, 160, 80))), Inches(4), Inches(0.3), Inches(1), Inches(1))
    blip = pic._element.find(".//{%s}blip" % A)
    rid = blip.get("{%s}embed" % R)
    blip.append(frag('<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="%s"/></a:ext></a:extLst>' % rid))
    tree.append(frag(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="30" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
        '<p:xfrm><a:off x="6350000" y="254000"/><a:ext cx="2540000" cy="1905000"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId99" r:lo="rId98" r:qs="rId97" r:cs="rId96"/></a:graphicData></a:graphic></p:graphicFrame>'))
    tbl = s.shapes.add_table(2, 3, Inches(0.3), Inches(2.5), Inches(6), Inches(1.2)).table
    tbl.cell(0, 0).merge(tbl.cell(0, 1))
    tbl.cell(0, 0).text = "Merged"
    tbl.cell(1, 2).text = "C"
    tree.append(frag(
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="40" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm rot="5400000"><a:off x="7620000" y="3810000"/><a:ext cx="1270000" cy="635000"/><a:chOff x="0" y="0"/><a:chExt cx="2540000" cy="1270000"/></a:xfrm></p:grpSpPr>'
        '<p:sp><p:nvSpPr><p:cNvPr id="41" name="In group"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1270000" y="0"/><a:ext cx="1270000" cy="1270000"/></a:xfrm>'
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="336699"/></a:solidFill></p:spPr></p:sp></p:grpSp>'))
    tree.append(frag(
        '<p:sp><p:nvSpPr><p:cNvPr id="50" name="Freeform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="254000" y="4445000"/><a:ext cx="1270000" cy="1270000"/></a:xfrm>'
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>'
        '<a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></p:spPr></p:sp>'))
    tree.append(frag(
        '<p:sp><p:nvSpPr><p:cNvPr id="60" name="Vertical"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10160000" y="2540000"/><a:ext cx="508000" cy="2540000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr vert="vert" wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>Side</a:t></a:r></a:p></p:txBody></p:sp>'))

    # slide 3: Title Only layout (no master shapes, layout background), transition inside mc:AlternateContent
    s = prs.slides.add_slide(L[5])
    s.shapes.title.text = "Hidden brand bar"
    clr = s._element.find("{%s}clrMapOvr" % NS)
    clr.addnext(frag(
        '<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p14:vortex dir="r"/></p:transition></mc:Choice>'
        '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'))

    # slide 4: picture background
    s = prs.slides.add_slide(L[6])
    _, rid = s.part.get_or_add_image_part(io.BytesIO(png_bytes(3, 3, (250, 220, 120))))
    csld = s._element.find("{%s}cSld" % NS)
    csld.insert(0, frag('<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="%s"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>' % rid))
    tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text_frame.text = "On a picture"

    prs.save(os.path.join(HERE, "complex.pptx"))


if __name__ == "__main__":
    basic(); wide(); complex_deck()
    print("ok")
