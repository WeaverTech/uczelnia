#!/usr/bin/env python3
"""Kalendarz zajęć 13M5: GL04, projekt gP03, angielski Majka-Pauli, specjalność SL03."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from bs4 import BeautifulSoup
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

HTML_PATH = Path("/tmp/13M5.htm")
OUT_PATH = Path("/workspace/plan-zajec-13M5.pdf")
ICS_PATH = Path("/workspace/plan-zajec-13M5.ics")
FONT_DIR = Path("/usr/share/fonts/truetype/macos")

pdfmetrics.registerFont(TTFont("Inter", str(FONT_DIR / "Inter-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Inter-Med", str(FONT_DIR / "Inter-Medium.ttf")))
pdfmetrics.registerFont(TTFont("Inter-Semi", str(FONT_DIR / "Inter-SemiBold.ttf")))
pdfmetrics.registerFont(TTFont("Inter-Bold", str(FONT_DIR / "Inter-Bold.ttf")))

PAGE_W, PAGE_H = A4  # portrait

MONTHS = {"X": 10, "XI": 11, "XII": 12, "I": 1, "II": 2}
MONTH_NAME = {
    9: "września",
    10: "października",
    11: "listopada",
    12: "grudnia",
    1: "stycznia",
    2: "lutego",
}
MONTH_SHORT = {9: "wrz", 10: "paź", 11: "lis", 12: "gru", 1: "sty", 2: "lut"}
DAY_NAME = {
    0: "poniedziałek",
    1: "wtorek",
    2: "środa",
    3: "czwartek",
    4: "piątek",
}
HOLIDAYS = {
    date(2026, 11, 11): "Niepodległości",
    date(2026, 12, 24): "Wigilia",
    date(2026, 12, 25): "Boże Narodzenie",
    date(2026, 12, 26): "Boże Narodzenie",
    date(2027, 1, 1): "Nowy Rok",
    date(2027, 1, 6): "Trzech Króli",
}

# Lesson grid used by WM PK.
SLOT_STARTS = [
    7 * 60 + 30,
    8 * 60 + 15,
    9 * 60 + 15,
    10 * 60,
    11 * 60,
    11 * 60 + 45,
    12 * 60 + 45,
    13 * 60 + 30,
    14 * 60 + 30,
    15 * 60 + 15,
    16 * 60 + 15,
    17 * 60,
    18 * 60,
    18 * 60 + 45,
]
TIME_START = 7 * 60 + 30
TIME_END = 19 * 60 + 35

SUBJECTS = {
    "ZASYM": ("MES", "Zastosowania systemu MES"),
    "POROP": ("Robotyka", "Podstawy robotyki"),
    "PONIP": ("Niezawodność", "Podstawy niezawodności"),
    "PRZAS": ("Pomiary 3D", "Programowanie systemów pomiarowych 3D"),
    "SYIND": ("Informatyczne", "Systemy informatyczne"),
    "PROBC": ("CNC", "Programowanie obrabiarek CNC"),
    "MADRI": ("Masz. drogowe", "Maszyny drogowe i budowlane"),
    "NAIST": ("Napędy", "Napędy i sterowanie maszyn"),
    "MICII": ("Miernictwo", "Miernictwo cieplne i maszynowe"),
    "MAYEK": ("Eksploatacyjne", "Materiały eksploatacyjne maszyn"),
    "KOWSB": ("KWBE", "Komp. wspomaganie badań"),
    "JEANJ": ("Angielski", "Język angielski"),
}

# Accent, fill, text. Lectures are forced to white separately.
STYLES = {
    "lab": {
        "ZASYM": ("#1D4ED8", "#DBEAFE", "#1E3A8A"),
        "POROP": ("#0F766E", "#CCFBF1", "#134E4A"),
        "PRZAS": ("#4338CA", "#E0E7FF", "#312E81"),
        "SYIND": ("#047857", "#D1FAE5", "#064E3B"),
        "PROBC": ("#BE123C", "#FFE4E6", "#881337"),
        "MADRI": ("#A16207", "#FEF3C7", "#713F12"),
        "NAIST": ("#6D28D9", "#EDE9FE", "#4C1D95"),
        "MICII": ("#0369A1", "#E0F2FE", "#0C4A6E"),
        "MAYEK": ("#3F6212", "#ECFCCB", "#365314"),
    },
    "proj": ("#C2410C", "#FFEDD5", "#7C2D12"),
    "ang": ("#166534", "#DCFCE7", "#14532D"),
    "spec02": ("#7E22CE", "#F3E8FF", "#581C87"),
    "spec03": ("#9D174D", "#FCE7F3", "#831843"),
    "wyk": ("#E5E7EB", "#FFFFFF", "#6B7280"),
}


def role_of(code: str, group: str, teacher: str) -> str | None:
    g = group.strip()
    if code == "JEANJ":
        return "ang" if "MAJKA" in teacher.upper() else None
    if "GL04" in g:
        return "lab"
    if "GK/P03" in g:
        return "proj"
    if "SL03" in g:
        return "spec03"
    if "SL02" in g:
        return None
    if any(x in g for x in ("GL02", "GL03", "GK/P02", "SL01", "SP01", "SP02")):
        return None
    if g.startswith("12") or "12A" in g or "12B" in g:
        return None
    if "13M" in g:
        return "wyk"
    return None


def parse_clock(token: str) -> int:
    hour, minute = token.strip().split(".")
    return int(hour) * 60 + int(minute)


def parse_span(text: str) -> tuple[int, int]:
    start, end = text.split("-")
    return parse_clock(start), parse_clock(end)


def merge_slots(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    if not spans:
        return []
    spans = sorted(spans)
    merged = [spans[0]]
    for start, end in spans[1:]:
        prev_s, prev_e = merged[-1]
        if start <= prev_e + 1:
            merged[-1] = (prev_s, max(prev_e, end))
        else:
            merged.append((start, end))
    return merged


def fmt_time(minutes: int) -> str:
    return f"{minutes // 60}:{minutes % 60:02d}"


def teacher_surname(raw: str) -> str:
    raw = raw.strip()
    if not raw:
        return ""
    surname = raw.split()[0]
    return "-".join(part.capitalize() for part in surname.split("-"))


def parse_events(html_path: Path) -> list[dict]:
    html = html_path.read_text(encoding="utf-8")
    html = html.replace("<td_removed>", "").replace("</td_removed>", "")
    soup = BeautifulSoup(html, "lxml")
    main = max(soup.find_all("table"), key=lambda t: len(t.find_all("tr", recursive=False)))
    rows = main.find_all("tr", recursive=False)

    def extract_one(td):
        fonts = [f.get_text(" ", strip=True) for f in td.find_all("font") if f.get_text(strip=True)]
        if not fonts:
            text = td.get_text(" ", strip=True)
            fonts = [text] if text else []
        return fonts, td.get("background") or ""

    def cell_parts(td):
        nested = td.find("table")
        if nested:
            return [extract_one(ntd) for ntd in nested.find_all("td")]
        return [extract_one(td)]

    occupied: dict[tuple[int, int], bool] = {}
    grid: dict[tuple[int, int], dict] = {}
    for r, tr in enumerate(rows):
        c = 0
        for td in tr.find_all("td", recursive=False):
            while (r, c) in occupied:
                c += 1
            colspan = int(td.get("colspan") or 1)
            rowspan = int(td.get("rowspan") or 1)
            info = {"parts": cell_parts(td), "rowspan": rowspan, "colspan": colspan}
            for dr in range(rowspan):
                for dc in range(colspan):
                    occupied[(r + dr, c + dc)] = True
                    grid[(r + dr, c + dc)] = info if dr == 0 and dc == 0 else {"cont": True}
            c += colspan

    day_headers = []
    for r in range(len(rows)):
        cell = grid.get((r, 0))
        if cell and not cell.get("cont"):
            label = " ".join(cell["parts"][0][0]).strip()
            if label:
                day_headers.append((r, label))

    events = []
    for i, (r0, day) in enumerate(day_headers):
        r1 = day_headers[i + 1][0] if i + 1 < len(day_headers) else len(rows)
        weeks = {}
        for c in range(2, 21):
            cell = grid.get((r0, c))
            if cell and not cell.get("cont"):
                weeks[c] = " ".join(cell["parts"][0][0]).strip()
        for r in range(r0 + 1, r1):
            for c in range(2, 21):
                cell = grid.get((r, c))
                if not cell or cell.get("cont"):
                    continue
                times = []
                for dr in range(cell["rowspan"]):
                    tc = grid.get((r + dr, 1))
                    if tc and not tc.get("cont"):
                        label = " ".join(tc["parts"][0][0]).strip()
                        if label and label != "e-learning" and "-" in label and label[0].isdigit():
                            times.append(label)
                if not times:
                    continue
                date_label = weeks.get(c, "")
                if not date_label or " " not in date_label:
                    continue
                day_num, mon = date_label.split()
                if mon not in MONTHS:
                    continue
                month = MONTHS[mon]
                year = 2026 if month >= 9 else 2027
                when = date(year, month, int(day_num))
                for lines, bg in cell["parts"]:
                    if not lines or bg in ("outofrange.gif", "reservation.gif"):
                        continue
                    if lines[0][:1].isdigit():
                        continue
                    code = lines[0]
                    if len(lines) == 3 and any(ch.isdigit() for ch in lines[1]):
                        teacher, group, room = "", lines[1], lines[2]
                    else:
                        teacher = lines[1] if len(lines) > 1 else ""
                        group = lines[2] if len(lines) > 2 else ""
                        room = " ".join(lines[3:]) if len(lines) > 3 else ""
                    kind = role_of(code, group, teacher)
                    if kind is None or code not in SUBJECTS:
                        continue
                    slots = []
                    for label in times:
                        try:
                            slots.append(parse_span(label))
                        except ValueError:
                            continue
                    blocks = merge_slots(slots)
                    if not blocks:
                        continue
                    short, full = SUBJECTS[code]
                    if kind == "spec03":
                        tag = "SL03"
                    elif kind == "proj":
                        tag = "projekt"
                    elif kind == "ang":
                        tag = "lektorat"
                    elif kind == "wyk":
                        tag = "wykład"
                    else:
                        tag = "lab"
                    events.append(
                        {
                            "date": when,
                            "weekday": when.weekday(),
                            "blocks": blocks,
                            "code": code,
                            "short": short,
                            "full": full,
                            "kind": kind,
                            "tag": tag,
                            "teacher": teacher_surname(teacher),
                            "room": room,
                            "group": group,
                        }
                    )
    return events


def week_mondays() -> list[date]:
    # 1 października (czwartek) ma wykład, więc pierwszy tydzień zaczyna się 28 września.
    start = date(2026, 9, 28)
    end = date(2027, 2, 1)
    days = []
    cur = start
    while cur <= end:
        days.append(cur)
        cur += timedelta(days=7)
    return days


def rhythm(monday: date) -> str:
    # 5.10.2026 is rhythm I; weeks alternate, including empty ones.
    origin = date(2026, 10, 5)
    index = (monday - origin).days // 7
    return "I" if index % 2 == 0 else "II"


def style_for(event: dict) -> tuple[str, str, str]:
    if event["kind"] == "wyk":
        return STYLES["wyk"]
    if event["kind"] == "lab":
        return STYLES["lab"][event["code"]]
    return STYLES[event["kind"]]


def hex_color(value: str) -> tuple[float, float, float]:
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) / 255 for i in (0, 2, 4))


def draw_header(c: canvas.Canvas, monday: date, page: int, pages: int) -> None:
    friday = monday + timedelta(days=4)
    c.setFillColor(hex_color("#111827"))
    c.setFont("Inter-Bold", 12.5)
    c.drawString(16, PAGE_H - 20, "Plan zajęć  ·  13M5")
    c.setFont("Inter", 8)
    c.setFillColor(hex_color("#6B7280"))
    c.drawString(138, PAGE_H - 18.5, "semestr zimowy 2026/27")

    c.setFillColor(hex_color("#9CA3AF"))
    c.setFont("Inter", 8)
    c.drawRightString(PAGE_W - 16, PAGE_H - 18.5, f"{page} / {pages}")

    c.setFillColor(hex_color("#374151"))
    c.setFont("Inter", 7.2)
    note = "GL04   ·   projekt gP03   ·   angielski Majka-Pauli   ·   specjalność SL03"
    c.drawString(16, PAGE_H - 34, note)

    if monday.month != friday.month:
        title = f"{monday.day} {MONTH_SHORT[monday.month]} – {friday.day} {MONTH_NAME[friday.month]}"
    else:
        title = f"{monday.day}–{friday.day} {MONTH_NAME[friday.month]}"
    pill = "ostatni tydzień" if monday == date(2027, 2, 1) else f"rytm {rhythm(monday)}"
    c.setFont("Inter-Semi", 8)
    pill_w = pdfmetrics.stringWidth(pill, "Inter-Semi", 8) + 14
    c.setFillColor(hex_color("#111827"))
    c.roundRect(PAGE_W - 16 - pill_w, PAGE_H - 40, pill_w, 13, 6, stroke=0, fill=1)
    c.setFillColor(hex_color("#FFFFFF"))
    c.drawCentredString(PAGE_W - 16 - pill_w / 2, PAGE_H - 36.4, pill)
    c.setFillColor(hex_color("#111827"))
    c.setFont("Inter-Semi", 8.5)
    c.drawRightString(PAGE_W - 22 - pill_w, PAGE_H - 36.2, title)


def draw_legend(c: canvas.Canvas) -> None:
    items = [
        ("#1D4ED8", "#DBEAFE", "MES lab"),
        ("#0F766E", "#CCFBF1", "Robotyka lab"),
        ("#C2410C", "#FFEDD5", "Niezawodność proj."),
        ("#4338CA", "#E0E7FF", "Pomiary 3D lab"),
        ("#047857", "#D1FAE5", "Informatyczne lab"),
        ("#BE123C", "#FFE4E6", "CNC lab"),
        ("#A16207", "#FEF3C7", "Masz. drogowe lab"),
        ("#6D28D9", "#EDE9FE", "Napędy lab"),
        ("#0369A1", "#E0F2FE", "Miernictwo lab"),
        ("#3F6212", "#ECFCCB", "Eksploatacyjne lab"),
        ("#166534", "#DCFCE7", "Angielski"),
        ("#9D174D", "#FCE7F3", "KWBE SL03"),
        ("#D1D5DB", "#FFFFFF", "wykład"),
    ]
    y = 36
    x = 16
    max_x = PAGE_W - 16
    for accent, fill, label in items:
        width = pdfmetrics.stringWidth(label, "Inter-Med", 6.3) + 16
        if x + width > max_x:
            x = 16
            y -= 11
        c.setStrokeColor(hex_color(accent))
        c.setFillColor(hex_color(fill))
        c.setLineWidth(0.8)
        c.roundRect(x, y, 8, 8, 1.5, stroke=1, fill=1)
        c.setFillColor(hex_color("#374151"))
        c.setFont("Inter-Med", 6.3)
        c.drawString(x + 11, y + 1.3, label)
        x += width + 8
    footer = (
        "Wykłady są białe, bo nie wchodzą w plan chodzenia. "
        "Specjalność KWBE to grupa SL03, wtorki 11:00, sala B206. "
        "Wykład z robotyki jest w e-learningu. Źródło: podzial.mech.pk.edu.pl, plan 13M5, aktualizacja 23.09.2026."
    )
    c.setFillColor(hex_color("#6B7280"))
    c.setFont("Inter", 6.2)
    line_y = 16
    for line in simpleSplit(footer, "Inter", 6.2, PAGE_W - 32):
        c.drawString(16, line_y, line)
        line_y -= 8


def y_of(minutes: int, grid_top: float, grid_h: float) -> float:
    frac = (minutes - TIME_START) / (TIME_END - TIME_START)
    return grid_top - frac * grid_h


def draw_block(c: canvas.Canvas, event: dict, start: int, end: int, x: float, w: float, grid_top: float, grid_h: float, col: int, cols: int) -> None:
    gap = 1.4
    inner_w = (w - 3 - gap * (cols - 1)) / cols
    bx = x + 1.5 + col * (inner_w + gap)
    top = y_of(start, grid_top, grid_h)
    bot = y_of(end, grid_top, grid_h)
    bh = top - bot
    if bh < 8:
        return
    accent, fill, text = style_for(event)
    c.saveState()
    c.setFillColor(hex_color(fill))
    c.setStrokeColor(hex_color(accent))
    c.setLineWidth(0.8 if event["kind"] != "wyk" else 0.6)
    if event["kind"] in ("spec02", "spec03"):
        c.setDash(1.5, 1.2)
    c.roundRect(bx, bot + 0.8, inner_w, bh - 1.6, 3, stroke=1, fill=1)
    c.setDash()
    # Left accent bar for classes that are actually attended.
    if event["kind"] != "wyk":
        c.setFillColor(hex_color(accent))
        c.rect(bx, bot + 2.2, 2.2, bh - 4.4, stroke=0, fill=1)
    pad_x = bx + (6 if event["kind"] != "wyk" else 4)
    text_w = inner_w - (10 if event["kind"] != "wyk" else 7)
    c.setFillColor(hex_color(text))
    name = f"{event['short']}  ·  {event['tag']}"
    when = f"{fmt_time(start)}–{fmt_time(end)}"
    detail = "  ·  ".join(part for part in (event["room"], event["teacher"]) if part)
    if bh >= 36 and text_w > 36:
        lines = [(name, "Inter-Semi", 7.1), (when, "Inter-Med", 6.4), (detail, "Inter", 6.2)]
    elif bh >= 24:
        lines = [(name, "Inter-Semi", 6.5), (when, "Inter", 5.8)]
    else:
        lines = [(event["short"], "Inter-Semi", 6)]

    clip = c.beginPath()
    clip.roundRect(bx, bot + 0.8, inner_w, bh - 1.6, 3)
    c.clipPath(clip, stroke=0, fill=0)
    cursor = top - 10
    for content, font, size in lines:
        if cursor < bot + 2:
            break
        wrapped = simpleSplit(content, font, size, text_w) or [""]
        c.setFont(font, size)
        c.drawString(pad_x, cursor, wrapped[0])
        cursor -= size + 1.3
    c.restoreState()


def assign_columns(items: list[dict]) -> None:
    """items: start, end, col, cols set in place. Cluster by overlap."""
    items.sort(key=lambda e: (e["start"], -(e["end"] - e["start"])))
    col_ends: list[int] = []
    for item in items:
        placed = False
        for i, end in enumerate(col_ends):
            if end <= item["start"]:
                item["col"] = i
                col_ends[i] = item["end"]
                placed = True
                break
        if not placed:
            item["col"] = len(col_ends)
            col_ends.append(item["end"])
    # transitive clusters
    parent = list(range(len(items)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for i, a in enumerate(items):
        for j, b in enumerate(items):
            if i >= j:
                continue
            if a["start"] < b["end"] and b["start"] < a["end"]:
                union(i, j)
    clusters: dict[int, list[int]] = defaultdict(list)
    for i in range(len(items)):
        clusters[find(i)].append(i)
    for members in clusters.values():
        cols = max(items[i]["col"] for i in members) + 1
        for i in members:
            items[i]["cols"] = cols


def draw_week(c: canvas.Canvas, monday: date, events: list[dict], page: int, pages: int) -> None:
    c.setFillColor(hex_color("#FFFFFF"))
    c.rect(0, 0, PAGE_W, PAGE_H, stroke=0, fill=1)
    draw_header(c, monday, page, pages)
    draw_legend(c)

    grid_left = 44
    grid_right = PAGE_W - 12
    grid_top = PAGE_H - 92
    grid_bottom = 62
    grid_h = grid_top - grid_bottom
    day_w = (grid_right - grid_left) / 5

    by_day: dict[date, list[dict]] = defaultdict(list)
    for event in events:
        if monday <= event["date"] <= monday + timedelta(days=4):
            for start, end in event["blocks"]:
                by_day[event["date"]].append({**event, "start": start, "end": end})

    for i in range(5):
        day = monday + timedelta(days=i)
        x = grid_left + i * day_w
        holiday = HOLIDAYS.get(day)
        if holiday:
            c.setFillColor(hex_color("#FEF2F2"))
        elif i % 2 == 0:
            c.setFillColor(hex_color("#FAFAF9"))
        else:
            c.setFillColor(hex_color("#FFFFFF"))
        c.rect(x, grid_bottom, day_w, grid_h, stroke=0, fill=1)

        c.setFillColor(hex_color("#B91C1C") if holiday else hex_color("#111827"))
        c.setFont("Inter-Semi", 7.2)
        c.drawCentredString(x + day_w / 2, grid_top + 28, DAY_NAME[i])
        c.setFont("Inter-Bold", 13)
        c.drawCentredString(x + day_w / 2, grid_top + 13, str(day.day))
        c.setFont("Inter", 6.4)
        c.setFillColor(hex_color("#B91C1C") if holiday else hex_color("#6B7280"))
        sub = holiday if holiday else MONTH_SHORT[day.month]
        c.drawCentredString(x + day_w / 2, grid_top + 3, sub)

    c.setStrokeColor(hex_color("#EEF0F3"))
    c.setLineWidth(0.4)
    for start in SLOT_STARTS:
        if not (TIME_START <= start <= TIME_END):
            continue
        y = y_of(start, grid_top, grid_h)
        c.line(grid_left, y, grid_right, y)
        c.setFillColor(hex_color("#9CA3AF"))
        c.setFont("Inter", 6)
        c.drawRightString(grid_left - 4, y - 2, fmt_time(start))

    if not by_day:
        c.setFillColor(hex_color("#9CA3AF"))
        c.setFont("Inter", 9)
        c.drawCentredString((grid_left + grid_right) / 2, (grid_top + grid_bottom) / 2, "Przerwa — brak zajęć")

    for i in range(5):
        day = monday + timedelta(days=i)
        x = grid_left + i * day_w
        items = by_day.get(day, [])
        items.sort(key=lambda e: (0 if e["kind"] == "wyk" else 1, e["start"]))
        assign_columns(items)
        for item in items:
            draw_block(
                c,
                item,
                item["start"],
                item["end"],
                x,
                day_w,
                grid_top,
                grid_h,
                item["col"],
                item["cols"],
            )

    c.setStrokeColor(hex_color("#E5E7EB"))
    c.setLineWidth(0.7)
    c.rect(grid_left, grid_bottom, grid_right - grid_left, grid_h, stroke=1, fill=0)
    for i in range(1, 5):
        x = grid_left + i * day_w
        c.line(x, grid_bottom, x, grid_top)


def ics_escape(text: str) -> str:
    return (
        text.replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\n", "\\n")
    )


def fold(line: str) -> str:
    raw = line.encode("utf-8")
    chunks = []
    while len(raw) > 73:
        cut = 73
        while cut > 0 and (raw[cut] & 0xC0) == 0x80:
            cut -= 1
        chunks.append(raw[:cut].decode("utf-8"))
        raw = raw[cut:]
    chunks.append(raw.decode("utf-8"))
    return "\r\n ".join(chunks)


def write_ics(events: list[dict]) -> int:
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//13M5//plan zajec//PL",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:Plan 13M5",
        "X-WR-TIMEZONE:Europe/Warsaw",
        "BEGIN:VTIMEZONE",
        "TZID:Europe/Warsaw",
        "X-LIC-LOCATION:Europe/Warsaw",
        "BEGIN:DAYLIGHT",
        "TZOFFSETFROM:+0100",
        "TZOFFSETTO:+0200",
        "TZNAME:CEST",
        "DTSTART:19700329T020000",
        "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
        "END:DAYLIGHT",
        "BEGIN:STANDARD",
        "TZOFFSETFROM:+0200",
        "TZOFFSETTO:+0100",
        "TZNAME:CET",
        "DTSTART:19701025T030000",
        "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
        "END:STANDARD",
        "END:VTIMEZONE",
    ]
    count = 0
    for event in events:
        for start, end in event["blocks"]:
            stamp = f"{event['date'].strftime('%Y%m%d')}T{start // 60:02d}{start % 60:02d}00"
            end_stamp = f"{event['date'].strftime('%Y%m%d')}T{end // 60:02d}{end % 60:02d}00"
            uid = f"{event['date'].isoformat()}-{start}-{event['code']}-{event['kind']}@13m5"
            summary = f"{event['full']} · {event['tag']}"
            description = "\\n".join(
                part
                for part in (
                    f"Prowadzący: {event['teacher']}" if event["teacher"] else "",
                    f"Grupa: {event['group']}" if event["group"] else "",
                    "Wykład — w planie PDF jest na biało." if event["kind"] == "wyk" else "",
                )
                if part
            )
            lines.extend(
                [
                    "BEGIN:VEVENT",
                    f"UID:{uid}",
                    "DTSTAMP:20260924T000000Z",
                    f"DTSTART;TZID=Europe/Warsaw:{stamp}",
                    f"DTEND;TZID=Europe/Warsaw:{end_stamp}",
                    fold(f"SUMMARY:{ics_escape(summary)}"),
                    fold(f"LOCATION:{ics_escape(event['room'])}"),
                    fold(f"DESCRIPTION:{description}"),
                    "TRANSP:TRANSPARENT" if event["kind"] == "wyk" else "TRANSP:OPAQUE",
                    "END:VEVENT",
                ]
            )
            count += 1
    lines.append("END:VCALENDAR")
    ICS_PATH.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
    return count


def main() -> None:
    events = parse_events(HTML_PATH)
    mondays = week_mondays()
    pages = len(mondays)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PATH), pagesize=A4)
    c.setTitle("Plan zajęć 13M5 — GL04, gP03")
    c.setAuthor("plan z podzial.mech.pk.edu.pl")
    for index, monday in enumerate(mondays, start=1):
        draw_week(c, monday, events, index, pages)
        c.showPage()
    c.save()

    attended = [e for e in events if e["kind"] != "wyk"]
    lectures = [e for e in events if e["kind"] == "wyk"]
    ics_count = write_ics(events)
    print(f"wrote {OUT_PATH}  pages={pages}  attended_blocks={len(attended)}  lectures={len(lectures)}")
    print(f"wrote {ICS_PATH}  events={ics_count}")
    counts = defaultdict(int)
    for e in events:
        counts[(e["kind"], e["code"])] += len(e["blocks"])
    for key in sorted(counts):
        print(f"  {counts[key]:3}  {key[0]:8} {key[1]}")


if __name__ == "__main__":
    main()
