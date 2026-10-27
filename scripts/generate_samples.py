#!/usr/bin/env python3
"""Create the two sample decks committed under samples/.

northwind-quarterly.pptx is a content deck (default Office theme).
harbor-company-template.pptx is a company template: custom theme, master
chrome, and the standard layout set. SlideFit's job is to move Northwind's
content onto Harbor's layouts.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"


def region_diagram() -> bytes:
    """A small original diagram so the image slide has real bytes to reflow."""
    w, h = 640, 400
    img = Image.new("RGB", (w, h), "#F7F3EA")
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 0, 18, h), fill="#0E7C7B")
    draw.rounded_rectangle((48, 36, 300, 180), radius=18, fill="#0E7C7B")
    draw.rounded_rectangle((330, 36, 592, 180), radius=18, fill="#D46A3A")
    draw.rounded_rectangle((48, 210, 210, 360), radius=18, fill="#2F4858")
    draw.rounded_rectangle((230, 210, 400, 360), radius=18, fill="#E2B15A")
    draw.rounded_rectangle((420, 210, 592, 360), radius=18, fill="#8C4A3A")
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
        small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 16)
    except OSError:
        font = ImageFont.load_default()
        small = font
    draw.text((64, 90), "Seattle", fill="white", font=font)
    draw.text((346, 90), "Portland", fill="white", font=font)
    draw.text((64, 270), "Boise", fill="white", font=small)
    draw.text((246, 270), "Coast", fill="#1C2B2A", font=small)
    draw.text((436, 270), "Inland", fill="white", font=small)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def add_title_and_bullets(slide, title, bullets):
    slide.shapes.title.text = title
    body = slide.placeholders[1]
    tf = body.text_frame
    tf.clear()
    for i, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = bullet
        p.level = 0


def build_northwind(image_bytes: bytes) -> Presentation:
    # Default 4:3 (10 × 7.5 in). The Harbor template is widescreen on purpose,
    # so a successful conversion has to reflow rather than copy coordinates.
    prs = Presentation()
    prs.core_properties.title = "Northwind Quarterly Review"
    prs.core_properties.author = "Northwind Field Team"

    title_layout = prs.slide_layouts[0]
    section_layout = prs.slide_layouts[2]
    content_layout = prs.slide_layouts[1]
    blank = prs.slide_layouts[6]

    slide = prs.slides.add_slide(title_layout)
    slide.shapes.title.text = "Northwind Quarterly Review"
    slide.placeholders[1].text = "Q3 field briefing · prepared for regional leads"

    slide = prs.slides.add_slide(section_layout)
    slide.shapes.title.text = "What changed this quarter"

    slide = prs.slides.add_slide(content_layout)
    add_title_and_bullets(
        slide,
        "Pipeline health",
        [
            "Qualified opportunities grew 18% against Q2",
            "Average deal size held at $42k",
            "Margins held after returns (net > gross & fees)",
            "Two enterprise renewals moved to verbal commit",
        ],
    )

    slide = prs.slides.add_slide(blank)
    box = slide.shapes.add_textbox(Inches(0.55), Inches(0.38), Inches(9), Inches(0.8))
    box.text_frame.paragraphs[0].text = "West region coverage"
    for paragraph in box.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(32)
            run.font.bold = True
    bullets = slide.shapes.add_textbox(Inches(0.55), Inches(1.45), Inches(4.5), Inches(4.4))
    tf = bullets.text_frame
    tf.word_wrap = True
    lines = [
        "Seattle and Portland carried the quarter",
        "A new distributor covers Boise",
        "Inventory on the coast is tight through November",
    ]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.level = 0
    slide.shapes.add_picture(io.BytesIO(image_bytes), Inches(5.2), Inches(1.55), Inches(4.3), Inches(3.5))

    slide = prs.slides.add_slide(content_layout)
    slide.shapes.title.text = "Revenue by line"
    rows, cols = 5, 3
    table_shape = slide.shapes.add_table(
        rows, cols, Inches(0.8), Inches(1.8), Inches(11.5), Inches(4.2)
    )
    table = table_shape.table
    data = [
        ["Line", "Q2 $m", "Q3 $m"],
        ["Specialty tea", "1.2", "1.5"],
        ["Field kits", "0.8", "0.9"],
        ["Wholesale", "2.1", "2.4"],
        ["Direct", "0.6", "0.7"],
    ]
    for r, row in enumerate(data):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.text = value
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    run.font.size = Pt(16)
                    run.font.bold = r == 0

    slide = prs.slides.add_slide(content_layout)
    add_title_and_bullets(
        slide,
        "Risks to watch",
        [
            "Port delays on the October sailing",
            "One key account is in procurement review",
            "Hiring for the Denver route is still open",
        ],
    )

    slide = prs.slides.add_slide(content_layout)
    add_title_and_bullets(
        slide,
        "Decisions for Thursday",
        [
            "Approve the coastal safety stock",
            "Keep the wholesale price hold",
            "Name an owner for the Denver route",
        ],
    )
    return prs


HARBOR_COLORS = {
    "dk1": "1C2B2A",
    "lt1": "F7F3EA",
    "dk2": "163330",
    "lt2": "E4D7C5",
    "accent1": "0E7C7B",
    "accent2": "D46A3A",
    "accent3": "2F4858",
    "accent4": "E2B15A",
    "accent5": "8C4A3A",
    "accent6": "3D5A4C",
    "hlink": "0E7C7B",
    "folHlink": "8C4A3A",
}


def recolor_theme(xml: str) -> str:
    """Point the Office theme at Harbor colours and fonts."""
    import re

    for name, value in HARBOR_COLORS.items():
        xml, n = re.subn(
            rf"<a:{name}>.*?</a:{name}>",
            f'<a:{name}><a:srgbClr val="{value}"/></a:{name}>',
            xml,
            count=1,
            flags=re.S,
        )
        if n != 1:
            raise SystemExit(f"theme slot {name} not updated")

    def latin(block, face):
        return re.sub(
            r'(<a:latin typeface=")[^"]*(")',
            rf"\g<1>{face}\2",
            block,
            count=1,
        )

    xml, n = re.subn(
        r"<a:majorFont>.*?</a:majorFont>",
        lambda m: latin(m.group(0), "Georgia"),
        xml,
        count=1,
        flags=re.S,
    )
    if n != 1:
        raise SystemExit("major font not updated")
    xml, n = re.subn(
        r"<a:minorFont>.*?</a:minorFont>",
        lambda m: latin(m.group(0), "Verdana"),
        xml,
        count=1,
        flags=re.S,
    )
    if n != 1:
        raise SystemExit("minor font not updated")
    return xml


# Designed for the widescreen slide (13.333 × 7.5 in) and inserted after
# placeholder coordinates are scaled, so these numbers must not be scaled again.
ACCENT_BAR = """<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="900" name="Harbor Accent"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="137160" cy="6858000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="0E7C7B"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="901" name="Harbor Wordmark"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="10212000" y="6478000"/>
      <a:ext cx="1700000" cy="280000"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="r"/>
      <a:r>
        <a:rPr lang="en-US" sz="1100" dirty="0">
          <a:solidFill><a:srgbClr val="0E7C7B"/></a:solidFill>
          <a:latin typeface="Georgia"/>
        </a:rPr>
        <a:t>HARBOR &amp; CO.</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>
"""


def insert_master_chrome(master_xml: str) -> str:
    """Non-placeholder shapes on the master show on every layout."""
    marker = "</p:spTree>"
    if marker not in master_xml:
        raise SystemExit("slide master has no spTree")
    if "Harbor Accent" in master_xml:
        return master_xml
    return master_xml.replace(marker, ACCENT_BAR + marker, 1)


# Default python-pptx layouts are drawn for a 10 × 7.5 in slide. Harbor is
# widescreen, so horizontal positions scale by exactly 4/3.
SLIDE_SX = 12192000 / 9144000


def scale_horizontal(xml: str) -> str:
    import re

    def scale_pair(tag, a, b):
        def repl(match):
            first = int(round(int(match.group(1)) * SLIDE_SX))
            return f'<{tag} {a}="{first}" {b}="{match.group(2)}"/>'

        return re.sub(rf'<{tag} {a}="(-?\d+)" {b}="(-?\d+)"/>', repl, xml)

    xml = scale_pair("a:off", "x", "y")
    xml = scale_pair("a:ext", "cx", "cy")
    xml = scale_pair("a:chOff", "x", "y")
    xml = scale_pair("a:chExt", "cx", "cy")
    return xml


def patch_template(path: Path) -> None:
    raw = path.read_bytes()
    src = zipfile.ZipFile(io.BytesIO(raw))
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            name = info.filename
            if name.endswith(".xml") and (
                name.startswith("ppt/slideMasters/") or name.startswith("ppt/slideLayouts/")
            ):
                data = scale_horizontal(data.decode("utf-8")).encode("utf-8")
            if name == "ppt/theme/theme1.xml":
                data = recolor_theme(data.decode("utf-8")).encode("utf-8")
            elif name == "ppt/slideMasters/slideMaster1.xml":
                data = insert_master_chrome(data.decode("utf-8")).encode("utf-8")
            elif name == "ppt/presentation.xml":
                text = data.decode("utf-8")
                import re

                text = re.sub(
                    r"<p:sldSz[^/]*/>",
                    '<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>',
                    text,
                    count=1,
                )
                data = text.encode("utf-8")
            dst.writestr(info, data)
    path.write_bytes(out_buf.getvalue())


def build_harbor() -> Presentation:
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    prs.core_properties.title = "Harbor & Co. template"
    prs.core_properties.subject = "SlideFit sample company template"
    prs.core_properties.author = "Harbor & Co."
    # One nearly empty title slide so the file is a normal deck a person could
    # open. SlideFit drops template content slides and keeps masters/layouts.
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Harbor & Co."
    slide.placeholders[1].text = "Company template"
    return prs


def dump_layouts(path: Path) -> None:
    prs = Presentation(str(path))
    print(f"\n== {path.name} slides={len(prs.slides)} ==")
    for i, layout in enumerate(prs.slide_layouts):
        phs = []
        for ph in layout.placeholders:
            kind = str(ph.placeholder_format.type).split(" ")[0]
            phs.append(f"{kind}:{ph.placeholder_format.idx}")
        print(f"  [{i}] {layout.name} -> {', '.join(phs)}")


def main() -> None:
    SAMPLES.mkdir(parents=True, exist_ok=True)
    image = region_diagram()

    north = SAMPLES / "northwind-quarterly.pptx"
    harbor = SAMPLES / "harbor-company-template.pptx"
    build_northwind(image).save(north)
    build_harbor().save(harbor)
    patch_template(harbor)
    print("wrote", north, north.stat().st_size)
    print("wrote", harbor, harbor.stat().st_size)
    dump_layouts(north)
    dump_layouts(harbor)


if __name__ == "__main__":
    main()
