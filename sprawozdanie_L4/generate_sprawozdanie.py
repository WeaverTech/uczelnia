import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

OUT_DIR = Path(__file__).parent

PROGRAM = """\
AC JOINT SPEED9 ACCU1 TIMER0  OX= WX= #[-86.688,-7.0655,-86.28,0.95457,-100.13,-183.58]
AC JOINT SPEED9 ACCU1 TIMER0  OX= WX= #[-115.87,-13.166,-118.3,-1.0621,-74.347,-65.354]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-115.87,-8.7489,-128.17,-1.181,-60.069,-65.05]
AC LINEAR SPEED3 ACCU1 TIMER1  OX=13 WX= #[-115.87,-8.7492,-128.16,-1.181,-60.069,-65.05]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-115.87,-14.028,-113.48,-1.0387,-79.999,-65.46]
AC JOINT SPEED9 ACCU1 TIMER0  OX=13 WX= #[-63.174,-13.609,-113.17,-0.17866,-79.296,-118.25]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-63.179,-6.5709,-126.18,-0.20307,-59.272,-118.17]
AC LINEAR SPEED3 ACCU1 TIMER1  OX= WX= #[-63.179,-6.5709,-126.18,-0.20307,-59.272,-118.17]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-63.18,-11.698,-112.82,-0.17918,-77.742,-118.24]
AC JOINT SPEED9 ACCU1 TIMER0  OX= WX= #[-113.35,-10.567,-111.94,-0.99145,-78.085,-68.048]
AC JOINT SPEED9 ACCU1 TIMER0  OX= WX= #[-113.33,-10.549,-111.91,-0.99041,-78.084,-68.057]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-113.74,-4.8209,-124.9,-1.1317,-59.375,-67.282]
AC LINEAR SPEED3 ACCU1 TIMER1  OX=13 WX= #[-113.74,-4.8212,-124.9,-1.1312,-59.375,-67.282]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-113.74,-9.8796,-112.72,-1.0003,-76.61,-67.627]
AC JOINT SPEED9 ACCU1 TIMER0  OX=13 WX= #[-65.614,-7.4757,-110.72,-0.25241,-75.68,-115.85]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-65.617,-2.4974,-122.6,-0.28564,-58.83,-115.76]
AC LINEAR SPEED3 ACCU1 TIMER1  OX= WX= #[-65.617,-2.4974,-122.6,-0.28564,-58.83,-115.76]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-65.617,-7.975,-107.47,-0.24825,-79.419,-115.86]
AC JOINT SPEED9 ACCU1 TIMER0  OX= WX= #[-110.96,-5.8703,-105.64,-0.95353,-79.651,-70.395]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-111.9,-0.060775,-120.21,-1.1015,-59.289,-69.059]
AC LINEAR SPEED3 ACCU1 TIMER1  OX=13 WX= #[-111.9,-0.060775,-120.21,-1.1015,-59.289,-69.059]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-111.9,-5.3305,-106.64,-0.96756,-78.119,-69.424]
AC JOINT SPEED9 ACCU1 TIMER0  OX=13 WX= #[-68.177,-3.6911,-105.15,-0.30019,-77.479,-113.28]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-67.63,1.6322,-118.18,-0.32823,-59.12,-113.73]
AC LINEAR SPEED3 ACCU1 TIMER1  OX= WX= #[-67.63,1.6313,-118.18,-0.32823,-59.12,-113.73]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-67.631,-3.0578,-107.19,-0.29188,-74.798,-113.82]
AC JOINT SPEED9 ACCU1 TIMER0  OX= WX= #[-110.9,-2.5803,-106.74,-0.96859,-75.258,-70.377]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-110.56,3.5247,-116.2,-1.0823,-59.69,-70.425]
AC LINEAR SPEED3 ACCU1 TIMER1  OX=13 WX= #[-110.56,3.5259,-116.2,-1.0823,-59.69,-70.425]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-110.56,-1.3533,-104.12,-0.95925,-76.646,-70.75]
AC JOINT SPEED9 ACCU1 TIMER0  OX=13 WX= #[-69.902,1.0921,-101.68,-0.33446,-76.172,-111.54]
AC LINEAR SPEED3 ACCU1 TIMER0  OX=13 WX= #[-69.357,5.1525,-114.17,-0.36407,-59.62,-111.98]
AC LINEAR SPEED3 ACCU1 TIMER1  OX= WX= #[-69.357,5.151,-114.17,-0.36407,-59.62,-111.98]
AC LINEAR SPEED3 ACCU1 TIMER0  OX= WX= #[-69.357,0.41047,-83.543,-0.31473,-94.984,-112.2]
"""

SPEED = {"9": "100%", "3": "25%"}
MOTION = {"JOINT": "Ruch osiowy", "LINEAR": "Ruch liniowy"}
ACCU = {"1": "dokł. 1mm"}
STEP_RE = re.compile(
    r"AC (JOINT|LINEAR) SPEED(\d) ACCU(\d) TIMER(\d)\s+OX=(\d*) WX= #\[([^\]]+)\]"
)


def parse_steps():
    steps = []
    gripper_closed = False
    for line in PROGRAM.strip().splitlines():
        motion, speed, accu, timer, ox, joints = STEP_RE.match(line).groups()
        parts = [f"{MOTION[motion]} {SPEED[speed]}", ACCU[accu]]
        closed = ox == "13"
        if closed != gripper_closed:
            parts.append("zamknij chwytak" if closed else "otwórz chwytak")
            gripper_closed = closed
        if timer == "1":
            parts.append("czekaj 0,3s")
        steps.append(
            {
                "code": line,
                "comment": ", ".join(parts),
                "joints": [j.strip() for j in joints.split(",")],
            }
        )
    return steps


def set_cell_shading(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def set_repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def write_cell(cell, text, size=9, bold=False, mono=False, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    if mono:
        run.font.name = "Courier New"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Courier New")


def set_col_widths(table, widths):
    table.autofit = False
    for grid_col, w in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        grid_col.set(qn("w:w"), str(int(w.twips)))
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = w


def heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)


def build():
    steps = parse_steps()
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2))
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    sec.header.paragraphs[0].text = "Wzorzec sprawozdania L4 – Podstawy Robotyki"
    sec.header.paragraphs[0].runs[0].font.size = Pt(9)

    for text, size, bold in (
        ("Laboratorium z przedmiotu", 14, False),
        ("Podstawy Robotyki", 18, True),
        ("Laboratorium nr 4", 14, True),
    ):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text)
        r.font.size = Pt(size)
        r.bold = bold

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Temat: ")
    r.bold = True
    p.add_run("Programowanie robotów Kawasaki z wykorzystaniem Teach-Pendanta")

    info = doc.add_table(rows=3, cols=2)
    info.style = "Table Grid"
    write_cell(info.cell(0, 0), "Rok akademicki: 2026/2027", 10)
    write_cell(info.cell(0, 1), "Data realizacji zajęć: 12.10.2026", 10)
    write_cell(info.cell(1, 0), "Grupa: GL04", 10)
    write_cell(info.cell(1, 1), "Program: m13x04", 10)
    merged = info.cell(2, 0).merge(info.cell(2, 1))
    merged.text = ""
    p = merged.paragraphs[0]
    p.add_run("Imiona i Nazwiska:").bold = True
    for i, resp in enumerate(("1", "2", "3", "?"), start=1):
        q = merged.add_paragraph(
            f"{i}. ……………………………………., odpowiedzialny za {resp},…."
        )
        q.paragraph_format.space_after = Pt(0)
    for row in info.rows:
        for c in row.cells:
            for par in c.paragraphs:
                for run in par.runs:
                    run.font.size = Pt(10)

    heading(
        doc,
        "1) Rysunek podanego do zaprogramowania zadania manipulacyjnego z zaznaczonym "
        "globalnym układem współrzędnych i parametrami ruchu",
    )
    doc.add_paragraph(
        "Zadanie: przeniesienie 4 elementów z miejsca A (JT1 ≈ −111…−116°) "
        "do miejsca B (JT1 ≈ −63…−70°) chwytakiem pneumatycznym."
    )
    box = doc.add_table(rows=1, cols=1)
    box.style = "Table Grid"
    box.rows[0].height = Cm(8)
    write_cell(
        box.cell(0, 0),
        "[miejsce na rysunek – może być odręczny]",
        10,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.add_run("Parametry ruchu:").bold = True
    for item in (
        "JOINT – ruch osiowy, LINEAR – ruch liniowy",
        "SPEED9 – prędkość 100%, SPEED3 – prędkość 25%",
        "ACCU1 – dokładność 1mm",
        "TIMER1 – czekanie 0,3s, TIMER0 – bez czekania",
        "OX=13 – sygnał 13: chwytak zamknięty, OX= (brak) – chwytak otwarty",
    ):
        q = doc.add_paragraph(item, style="List Bullet")
        q.paragraph_format.space_after = Pt(0)

    land = doc.add_section()
    land.orientation = WD_ORIENT.LANDSCAPE
    land.page_width, land.page_height = Cm(29.7), Cm(21)
    land.left_margin = land.right_margin = Cm(1.5)
    land.top_margin = land.bottom_margin = Cm(1.5)

    heading(
        doc,
        "2) Kod programu z Teach Pendanta wraz z komentarzem do każdego kroku",
    )
    code = doc.add_table(rows=1, cols=3)
    code.style = "Table Grid"
    code.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = (Cm(1.0), Cm(15.2), Cm(10.5))
    for cell, title in zip(code.rows[0].cells, ("Nr", "Kod programu", "Komentarz do linii programu")):
        write_cell(cell, title, 9, bold=True)
        set_cell_shading(cell, "D9D9D9")
    set_repeat_header(code.rows[0])

    rows = [("", ".PROGRAM m13x04()", "Start programu")]
    rows += [(str(i), s["code"], s["comment"]) for i, s in enumerate(steps, start=1)]
    rows.append(("", ".END", "Koniec programu"))
    for nr, line, comment in rows:
        cells = code.add_row().cells
        write_cell(cells[0], nr, 8, align=WD_ALIGN_PARAGRAPH.CENTER)
        write_cell(cells[1], line, 6.5, mono=True)
        write_cell(cells[2], comment, 8)
    set_col_widths(code, widths)

    port = doc.add_section()
    port.orientation = WD_ORIENT.PORTRAIT
    port.page_width, port.page_height = Cm(21), Cm(29.7)
    port.left_margin = port.right_margin = Cm(2)
    port.top_margin = port.bottom_margin = Cm(1.8)

    heading(
        doc,
        "3) Współrzędne zaprogramowanych punktów dla poszczególnych kroków",
    )
    doc.add_paragraph(
        "Punkty zapisane jako wartości złączowe (#[…]) – kąty osi JT1…JT6 [°]."
    ).runs[0].font.size = Pt(10)

    coords = doc.add_table(rows=1, cols=7)
    coords.style = "Table Grid"
    coords.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell, title in zip(
        coords.rows[0].cells,
        ("Krok (step)", "JT1", "JT2", "JT3", "JT4", "JT5", "JT6"),
    ):
        write_cell(cell, title, 9, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(cell, "D9D9D9")
    set_repeat_header(coords.rows[0])
    for i, s in enumerate(steps, start=1):
        cells = coords.add_row().cells
        write_cell(cells[0], str(i), 9, align=WD_ALIGN_PARAGRAPH.CENTER)
        for c, v in zip(cells[1:], s["joints"]):
            write_cell(c, v, 9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    set_col_widths(coords, [Cm(2.2)] + [Cm(2.4)] * 6)

    out = OUT_DIR / "Sprawozdanie_L4_m13x04.docx"
    doc.save(out)
    return out


if __name__ == "__main__":
    print(build())
