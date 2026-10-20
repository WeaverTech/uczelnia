"""Generuje sprawozdanie L1 (silnik asynchroniczny) na bazie szablonu `Sprawozdanie L1.docx`.

Wszystkie elementy treści (nagłówki, akapity, podpisy, tabele, równania, listy)
są klonowane z elementów wzorcowych szablonu, dzięki czemu zachowane są jego
style, czcionki i odstępy. Strona tytułowa szablonu pozostaje bez zmian.
"""

import copy
import csv
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import FuncFormatter, MultipleLocator  # noqa: E402
from docx import Document  # noqa: E402
from docx.oxml import parse_xml  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm  # noqa: E402
from PIL import Image  # noqa: E402

HERE = Path(__file__).parent
TEMPLATE = HERE.parent / "Sprawozdanie L1.docx"
CSV_PATH = HERE / "dane" / "data_2.csv"
IMG_DIR = HERE / "img"
OUT_DOCX = HERE / "Sprawozdanie_L1_silnik_asynchroniczny.docx"

HEADER_TITLE = "Badanie właściwości silnika asynchronicznego"

# Charakterystyka mechaniczna M = f(n) – dwa zestawy pomiarów (10 punktów każdy)
# zanotowane ręcznie na stanowisku. Pozostawienie pustej listy pomija tę część
# sprawozdania, a uzupełnienie danych dodaje ją automatycznie.
# Format: {"f1": częstotliwość zasilania [Hz], "pomiary": [(M [N·m], n [obr/min]), ...]}
CHARAKTERYSTYKA_MECHANICZNA: list[dict] = []
P_BADANY = 2

FAN_MOTOR = {
    "f1": 50,
    "p": 1,
    "n_N": 2760,
    "I_N_delta": 1.10,
    "I_N_star": 0.65,
}

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NSDECL = f'xmlns:w="{W}" xmlns:m="{M}"'

PHASE_COLORS = ("#1f5fa8", "#d9822b", "#2e8b3a")


# ---------------------------------------------------------------- formatowanie


def pl(value, digits=3):
    return f"{value:.{digits}f}".replace(".", ",")


def pl_axis(digits):
    return FuncFormatter(lambda v, _: f"{v:.{digits}f}".replace(".", ","))


# ---------------------------------------------------------------- dane CSV


@dataclass
class Rozruch:
    nr: int
    t: list  # datetime
    i: list  # [(I1, I2, I3)]
    peak_idx: tuple  # indeks maksimum dla każdej fazy
    steady: range

    @property
    def i_max(self):
        return [self.i[self.peak_idx[f]][f] for f in range(3)]

    @property
    def i_ust(self):
        return [
            sum(self.i[k][f] for k in self.steady) / len(self.steady) for f in range(3)
        ]

    @property
    def i_max_sr(self):
        return sum(self.i_max) / 3

    @property
    def i_ust_sr(self):
        return sum(self.i_ust) / 3

    @property
    def k(self):
        return self.i_max_sr / self.i_ust_sr


def read_csv():
    rows = []
    with CSV_PATH.open(encoding="utf-8") as fh:
        for line in csv.reader(fh):
            parts = line[0].split("\t")
            if parts[0].startswith("Data"):
                continue
            t = datetime.strptime(parts[0], "%Y-%m-%d %H:%M:%S")
            rows.append((t, tuple(float(v.replace(",", ".")) for v in parts[1:4])))
    return rows


def find_starts(rows, peak_window=2, steady_tol=0.10):
    starts, run = [], []
    for t, cur in rows + [(None, (0, 0, 0))]:
        if any(cur):
            run.append((t, cur))
            continue
        if run:
            ts = [r[0] for r in run]
            ii = [r[1] for r in run]
            peak = tuple(
                max(range(min(peak_window, len(ii))), key=lambda k: ii[k][f])
                for f in range(3)
            )
            med = [sorted(c[f] for c in ii)[len(ii) // 2] for f in range(3)]
            ok = [
                k
                for k in range(peak_window, len(ii))
                if all(abs(ii[k][f] - med[f]) <= steady_tol * med[f] for f in range(3))
            ]
            starts.append(
                Rozruch(len(starts) + 1, ts, ii, peak, range(ok[0], ok[-1] + 1))
            )
            run = []
    return starts


# ---------------------------------------------------------------- wykresy


def setup_mpl():
    from matplotlib import font_manager

    available = {f.name for f in font_manager.fontManager.ttflist}
    family = next((f for f in ("Calibri", "Carlito") if f in available), "DejaVu Sans")
    plt.rcParams.update(
        {
            "font.family": family,
            "font.size": 10,
            "axes.grid": True,
            "grid.alpha": 0.35,
            "axes.axisbelow": True,
            "savefig.dpi": 200,
        }
    )


def plot_overview(rows, starts, path):
    t0 = rows[0][0]
    t = [(r[0] - t0).total_seconds() for r in rows]
    fig, ax = plt.subplots(figsize=(6.3, 3.0))
    for f in range(3):
        ax.plot(
            t,
            [r[1][f] for r in rows],
            color=PHASE_COLORS[f],
            lw=1.2,
            label=f"I$_{{\\mathrm{{rms{f + 1}}}}}$ (faza L{f + 1})",
        )
    for s in starts:
        x = (s.t[0] - t0).total_seconds()
        ax.annotate(f"Pomiar {s.nr}", xy=(x + 3, max(s.i_max) - 0.1), fontsize=9)
    ax.set_xlabel(f"t [s] (od {t0:%H:%M:%S})")
    ax.set_ylabel("I [A]")
    ax.set_xlim(t[0], t[-1])
    ax.set_ylim(0, 4.9)
    ax.yaxis.set_major_formatter(pl_axis(1))
    ax.legend(loc="upper center", fontsize=8.5, ncol=3, framealpha=0.95)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_start(rows, s, path, pre=3, post=3):
    idx0 = next(k for k, r in enumerate(rows) if r[0] == s.t[0])
    seg = rows[max(0, idx0 - pre) : idx0 + len(s.t) + post]
    t = [(r[0] - s.t[0]).total_seconds() for r in seg]
    fig, ax = plt.subplots(figsize=(6.3, 3.2))
    t_st0 = (s.t[s.steady.start] - s.t[0]).total_seconds()
    t_st1 = (s.t[s.steady.stop - 1] - s.t[0]).total_seconds()
    ax.axvspan(t_st0, t_st1, color="#7fbf7f", alpha=0.22, label="zakres stanu ustalonego")
    pk = sorted({(s.t[k] - s.t[0]).total_seconds() for k in s.peak_idx})
    ax.axvspan(pk[0] - 0.35, pk[-1] + 0.35, color="#e06060", alpha=0.20, label="zakres maksimum rozruchowego")
    for f in range(3):
        ax.plot(
            t,
            [r[1][f] for r in seg],
            color=PHASE_COLORS[f],
            lw=1.4,
            marker="o",
            ms=3,
            label=f"I$_{{\\mathrm{{rms{f + 1}}}}}$ (faza L{f + 1})",
        )
        k = s.peak_idx[f]
        ax.plot(
            (s.t[k] - s.t[0]).total_seconds(),
            s.i[k][f],
            marker="D",
            ms=6,
            mfc="none",
            mec=PHASE_COLORS[f],
            mew=1.4,
        )
    ax.hlines(s.i_ust_sr, t_st0, t_st1, colors="k", linestyles="--", lw=1.0, label=f"I$_{{\\mathrm{{ust}}}}$ = {pl(s.i_ust_sr)} A")
    ax.set_xlabel("t [s] (t = 0 – załączenie silnika)")
    ax.set_ylabel("I [A]")
    ax.set_xlim(t[0], t[-1])
    ax.xaxis.set_major_locator(MultipleLocator(2))
    ax.set_ylim(0, 4.4)
    ax.yaxis.set_major_formatter(pl_axis(1))
    ax.legend(loc="upper right", fontsize=8, ncol=2, framealpha=0.95)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_mechanical(sets, path):
    fig, ax = plt.subplots(figsize=(6.3, 3.6))
    for k, z in enumerate(sets):
        ax.plot(z["s"], z["M"], marker="o", ms=4, lw=1.4, color=PHASE_COLORS[k], label=f"f$_1$ = {z['f1']} Hz")
        kk = z["krytyczny"]
        ax.plot(z["s"][kk], z["M"][kk], marker="D", ms=7, mfc="none", mec="k")
        ax.annotate(
            f"M$_k$ = {pl(z['M'][kk], 2)} N·m\ns$_k$ = {pl(z['s'][kk])}",
            xy=(z["s"][kk], z["M"][kk]),
            xytext=(8, -26 - 22 * k),
            textcoords="offset points",
            fontsize=8.5,
        )
    ax.set_xlabel("s [-]")
    ax.set_ylabel("M [N·m]")
    ax.xaxis.set_major_formatter(pl_axis(2))
    ax.yaxis.set_major_formatter(pl_axis(2))
    ax.legend(fontsize=8.5)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


# ---------------------------------------------------------------- szablon

# Word wymaga kolejności elementów w:rPr zgodnej ze schematem OOXML.
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
]


def insert_rpr(rpr, el):
    rank = RPR_ORDER.index(el.tag.split("}")[1])
    for child in rpr:
        name = child.tag.split("}")[1]
        if name in RPR_ORDER and RPR_ORDER.index(name) > rank:
            child.addprevious(el)
            return
    rpr.append(el)


class Report:
    def __init__(self, template):
        self.doc = Document(template)
        body = self.doc.element.body
        children = list(body.iterchildren())
        paras = [c for c in children if c.tag == qn("w:p")]
        tables = [c for c in children if c.tag == qn("w:tbl")]

        def text(el):
            return "".join(t.text or "" for t in el.iter(qn("w:t")))

        def find(prefix):
            return next(p for p in paras if text(p).startswith(prefix))

        heading = find("Stanowisko badawcze")
        self.p_heading = copy.deepcopy(heading)
        self.p_body = copy.deepcopy(find("Należy pamiętać"))
        self.p_empty = copy.deepcopy(heading.getnext().getnext())
        fig = next(p for p in paras if p.find(".//" + qn("w:drawing")) is not None and text(p) == "" and p is not paras[12])
        self.p_figure = copy.deepcopy(fig)
        self.p_caption = copy.deepcopy(fig.getnext())
        self.p_tab_caption = copy.deepcopy(find("Tab. 3."))
        self.tbl = copy.deepcopy(tables[2])
        self.p_equation = copy.deepcopy(next(p for p in paras if p.find(".//" + qn("m:oMath")) is not None))
        self.p_where = copy.deepcopy(find("gdzie:"))
        self.p_where_item = copy.deepcopy(find("r – promień"))
        self.p_list = copy.deepcopy(find("przedmiot 1"))
        self.p_conclusion = copy.deepcopy(find("Wnioski powinny"))

        first = children.index(find("Cel i zakres"))
        self.sect = children[-1]
        for el in children[first:-1]:
            body.remove(el)
        self.body = body
        self.doc_pr_id = 100
        self.eq_no = 0

    # -- wstawianie
    def _add(self, el):
        self.sect.addprevious(el)
        return el

    @staticmethod
    def _run_template(p):
        return copy.deepcopy(p.find(qn("w:r")))

    @staticmethod
    def _clear_runs(p):
        for child in list(p):
            if child.tag != qn("w:pPr"):
                p.remove(child)

    @staticmethod
    def _rich_runs(p, proto_run, markup):
        """Dodaje do akapitu przebiegi tekstu z prostym znacznikowaniem.

        `x_{abc}` – indeks dolny, `x^{abc}` – indeks górny, `//abc//` – kursywa.
        """
        rpr0 = proto_run.find(qn("w:rPr"))
        tokens = []
        for k, chunk in enumerate(markup.split("//")):
            for token in re.split(r"(_\{[^}]*\}|\^\{[^}]*\})", chunk):
                if token:
                    tokens.append((token, k % 2 == 1))
        for token, italic in tokens:
            vert = None
            if token.startswith("_{"):
                token, vert = token[2:-1], "subscript"
            elif token.startswith("^{"):
                token, vert = token[2:-1], "superscript"
            r = parse_xml(f"<w:r {NSDECL}/>")
            rpr = copy.deepcopy(rpr0) if rpr0 is not None else parse_xml(f"<w:rPr {NSDECL}/>")
            for e in rpr.findall(qn("w:i")):
                if e.get(qn("w:val")) not in ("0", "false"):
                    rpr.remove(e)
            if italic:
                insert_rpr(rpr, parse_xml(f"<w:i {NSDECL}/>"))
            if vert:
                insert_rpr(rpr, parse_xml(f'<w:vertAlign {NSDECL} w:val="{vert}"/>'))
            r.append(rpr)
            t = parse_xml(f'<w:t {NSDECL} xml:space="preserve"/>')
            t.text = token
            r.append(t)
            p.append(r)

    def _para_from(self, proto, markup):
        p = copy.deepcopy(proto)
        run = self._run_template(proto)
        self._clear_runs(p)
        self._rich_runs(p, run, markup)
        return self._add(p)

    # -- elementy treści
    def heading(self, text):
        return self._para_from(self.p_heading, text)

    def para(self, markup):
        return self._para_from(self.p_body, markup)

    def empty(self):
        return self._add(copy.deepcopy(self.p_empty))

    def bullets(self, items):
        for k, item in enumerate(items):
            end = "." if k == len(items) - 1 else ";"
            self._para_from(self.p_list, item.rstrip(";.") + end)

    def conclusion(self, markup):
        return self._para_from(self.p_conclusion, markup)

    def figure(self, image_path, width_cm, caption):
        p = copy.deepcopy(self.p_figure)
        r_id, _ = self.doc.part.get_or_add_image(str(image_path))
        with Image.open(image_path) as im:
            w_px, h_px = im.size
        cx = int(Cm(width_cm))
        cy = int(cx * h_px / w_px)
        for el in p.iter(qn("wp:extent")):
            el.set("cx", str(cx))
            el.set("cy", str(cy))
        for el in p.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ext"):
            if el.get("cx") is not None:
                el.set("cx", str(cx))
                el.set("cy", str(cy))
        for el in p.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
            el.set(qn("r:embed"), r_id)
        for el in list(p.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect")):
            el.getparent().remove(el)
        self.doc_pr_id += 1
        for el in p.iter(qn("wp:docPr")):
            el.set("id", str(self.doc_pr_id))
            el.set("name", f"Obraz {self.doc_pr_id}")
        self.empty()
        self._add(p)
        cap = copy.deepcopy(self.p_caption)
        run = self._run_template(self.p_caption)
        self._clear_runs(cap)
        self._rich_runs(cap, run, f"Rys. {caption}")
        self._add(cap)
        self.empty()

    def table(self, caption, header, rows, widths_cm):
        self.empty()
        self._para_from(self.p_tab_caption, f"Tab. {caption}")
        tbl = copy.deepcopy(self.tbl)
        grid = tbl.find(qn("w:tblGrid"))
        for g in list(grid):
            grid.remove(g)
        widths = [int(w * 567) for w in widths_cm]
        for w_tw in widths:
            grid.append(parse_xml(f'<w:gridCol {NSDECL} w:w="{w_tw}"/>'))
        proto_tr = tbl.findall(qn("w:tr"))[0]
        for tr in tbl.findall(qn("w:tr")):
            tbl.remove(tr)
        proto_tc = proto_tr.find(qn("w:tc"))
        proto_p = proto_tc.find(qn("w:p"))
        all_rows = [header] + rows
        for row_no, values in enumerate(all_rows):
            tr = copy.deepcopy(proto_tr)
            for tc in tr.findall(qn("w:tc")):
                tr.remove(tc)
            tr_pr = tr.find(qn("w:trPr"))
            tr_pr.insert(0, parse_xml(f"<w:cantSplit {NSDECL}/>"))
            if row_no == 0:
                tr_pr.append(parse_xml(f"<w:tblHeader {NSDECL}/>"))
            for w_tw, value in zip(widths, values):
                tc = copy.deepcopy(proto_tc)
                for child in list(tc):
                    if child.tag != qn("w:tcPr"):
                        tc.remove(child)
                tc.find(qn("w:tcPr")).find(qn("w:tcW")).set(qn("w:w"), str(w_tw))
                p = copy.deepcopy(proto_p)
                run = self._run_template(proto_p)
                self._clear_runs(p)
                if row_no < len(all_rows) - 1:
                    p.find(qn("w:pPr")).insert(1, parse_xml(f"<w:keepNext {NSDECL}/>"))
                self._rich_runs(p, run, value)
                tc.append(p)
                tr.append(tc)
            tbl.append(tr)
        self._add(tbl)
        self.empty()

    def equation(self, omml):
        self.eq_no += 1
        p = copy.deepcopy(self.p_equation)
        old = p.find(qn("m:oMath"))
        new = parse_xml(f"<m:oMath {NSDECL}>{omml}</m:oMath>")
        old.addprevious(new)
        p.remove(old)
        last_t = list(p.iter(qn("w:t")))[-1]
        last_t.text = f"({self.eq_no})"
        self._add(p)
        return self.eq_no

    def where(self, items):
        self._add(copy.deepcopy(self.p_where))
        for k, (symbol, desc) in enumerate(items):
            end = "." if k == len(items) - 1 else ";"
            self._para_from(self.p_where_item, f"//{symbol}// – {desc}{end}")

    def set_header(self, text):
        for sec in self.doc.sections:
            for t in sec.header._element.iter(qn("w:t")):
                if "Tutaj podaj" in (t.text or ""):
                    t.text = text

    def save(self, path):
        w14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
        wp14 = "{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}"
        seen = set()
        for el in self.body.iter():
            for attr, extra in ((w14 + "paraId", w14 + "textId"), (wp14 + "anchorId", wp14 + "editId")):
                value = el.get(attr)
                if value is None:
                    continue
                if value in seen:
                    del el.attrib[attr]
                    el.attrib.pop(extra, None)
                else:
                    seen.add(value)
        self.doc.save(path)


# ---------------------------------------------------------------- OMML


def mr(text, italic=None, normal=False):
    if normal:
        sty = "<m:rPr><m:nor/></m:rPr>"
    elif italic is False:
        sty = '<m:rPr><m:sty m:val="p"/></m:rPr>'
    else:
        sty = ""
    return (
        f'<m:r>{sty}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'
        f"<m:t xml:space=\"preserve\">{text}</m:t></m:r>"
    )


def mfrac(num, den):
    return f"<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>"


def msub(base, sub):
    return f"<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>"


def mnary_sum(lower, upper, expr):
    return (
        '<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr>'
        f"<m:sub>{lower}</m:sub><m:sup>{upper}</m:sup><m:e>{expr}</m:e></m:nary>"
    )


def sym(base, sub=None):
    return msub(mr(base), mr(sub, italic=False)) if sub else mr(base)


def num(text):
    """Liczby z przecinkiem dziesiętnym i jednostki – bez odstępów wstawianych przez edytor równań."""
    return mr(text, normal=True)


# ---------------------------------------------------------------- treść


def build():
    setup_mpl()
    rows = read_csv()
    starts = find_starts(rows)
    assert len(starts) == 3, starts
    mech = prepare_mechanical()
    out_img = HERE / "wykresy"
    out_img.mkdir(exist_ok=True)

    r = Report(TEMPLATE)
    r.set_header(HEADER_TITLE)

    k_sr = sum(s.k for s in starts) / len(starts)
    i_ust_all = sum(s.i_ust_sr for s in starts) / len(starts)
    i_max_all = sum(s.i_max_sr for s in starts) / len(starts)
    fm = FAN_MOTOR
    n_s_fan = 60 * fm["f1"] / fm["p"]
    s_fan = (n_s_fan - fm["n_N"]) / n_s_fan
    f_obr_fan = fm["f1"] * (1 - s_fan) / fm["p"]
    asym = [100 * (1 - s.i_ust[2] / s.i_ust[f]) for s in starts for f in (0, 1)]
    asym_lo, asym_hi = min(asym), max(asym)

    # 1 ---------------------------------------------------------------------
    r.heading("Cel i zakres laboratorium")
    r.para(
        "Celem laboratorium było zapoznanie się z budową, zasadą działania oraz podstawowymi "
        "właściwościami eksploatacyjnymi trójfazowego silnika asynchronicznego klatkowego, "
        + (
            "w szczególności z jego charakterystyką mechaniczną oraz z przebiegiem prądu pobieranego "
            "z sieci podczas rozruchu."
            if mech
            else "w szczególności z przebiegiem prądu pobieranego z sieci podczas rozruchu "
            "i wyznaczenie współczynnika krotności prądu rozruchowego."
        )
    )
    r.para(
        "Silnik asynchroniczny klatkowy składa się z nieruchomego stojana, w którego żłobkach "
        "umieszczone jest trójfazowe uzwojenie o osiach przesuniętych względem siebie o 120°, "
        "oraz z wirnika, którego uzwojenie ma postać klatki ze zwartych na końcach prętów. "
        "Liczba sekcji uzwojenia stojana określa liczbę par biegunów p. Prądy fazowe płynące "
        "w uzwojeniu stojana wytwarzają pole magnetyczne wirujące z prędkością synchroniczną "
        "n_{s} = 60·f_{1}/p, gdzie f_{1} jest częstotliwością napięcia zasilającego. Pole to indukuje "
        "w prętach wirnika siłę elektromotoryczną E_{2}, pod wpływem której w zwartym uzwojeniu "
        "płynie prąd I_{2}. Oddziaływanie pola wirującego z prądem wirnika wytwarza moment obrotowy."
    )
    r.para(
        "Indukowanie się siły elektromotorycznej w wirniku jest możliwe tylko wtedy, gdy wirnik "
        "obraca się wolniej niż pole wirujące (n < n_{s}). Względną różnicę tych prędkości nazywa się "
        "poślizgiem s. Zależność momentu od poślizgu (charakterystykę mechaniczną) opisuje wzór Klossa. "
        "Dla poślizgów mniejszych od poślizgu krytycznego s_{k} praca silnika jest stabilna – wzrost "
        "obciążenia powoduje niewielki spadek prędkości i wzrost momentu. Powyżej s_{k} wzrost "
        "obciążenia powoduje spadek momentu i zatrzymanie silnika. Prędkość obrotową silnika można "
        "regulować zmieniając częstotliwość lub napięcie zasilania."
    )
    r.para(
        "W chwili załączenia wirnik jest nieruchomy (s = 1), a indukowana w nim siła "
        "elektromotoryczna i prąd osiągają największe wartości. Silnik pobiera wtedy z sieci prąd "
        "rozruchowy, wielokrotnie większy od prądu znamionowego. Stosunek tych prądów nazywany jest "
        "współczynnikiem krotności prądu rozruchowego. W celu ograniczenia prądu rozruchowego "
        "stosuje się m.in. rozruch przełącznikiem gwiazda–trójkąt: w połączeniu w gwiazdę napięcie "
        "na uzwojeniu jest √3 razy mniejsze od napięcia międzyfazowego, a prąd pobierany z sieci "
        "i moment rozruchowy są 3-krotnie mniejsze niż w połączeniu w trójkąt. Stosuje się również "
        "układy łagodnego rozruchu oraz przemienniki częstotliwości."
    )
    r.para("Zakres laboratorium obejmował:")
    scope = ["zapoznanie się z budową stanowisk badawczych i danymi znamionowymi badanych silników"]
    if mech:
        scope += [
            "pomiar momentu obciążenia i prędkości obrotowej silnika dla dwóch zestawów parametrów "
            "zasilania oraz wyznaczenie charakterystyk mechanicznych M = f(s) i parametrów krytycznych",
        ]
    scope += [
        "rejestrację wartości skutecznych prądów w trzech fazach zasilania podczas trzech kolejnych "
        "rozruchów silnika napędzającego wentylator",
        "wyznaczenie wartości prądu w stanie ustalonym, maksymalnej wartości prądu podczas rozruchu "
        "oraz współczynnika krotności prądu dla każdego pomiaru",
        "wyznaczenie średniej krotności wzrostu prądu z trzech pomiarów oraz sformułowanie wniosków",
    ]
    r.bullets(scope)
    r.empty()

    # 2 ---------------------------------------------------------------------
    r.heading("Stanowisko badawcze")
    fig_no = 0

    def fig_ref():
        return f"2.{fig_no + 1}"

    if mech:
        r.para(
            f"Charakterystykę mechaniczną wyznaczono na stanowisku przedstawionym na rysunku {fig_ref()}. "
            "Badany silnik asynchroniczny zasilany jest z przemiennika częstotliwości, a jego wał "
            "sprzęgnięty jest z alternatorem obciążonym rezystorem. Moment obciążający silnik "
            "regulowany jest zmianą prądu wzbudzenia alternatora, a prędkość obrotowa i moment "
            "odczytywane są za pomocą układu pomiarowego współpracującego z komputerem PC."
        )
        fig_no += 1
        r.figure(IMG_DIR / "stanowisko_charakterystyka.jpg", 12.5, f"2.{fig_no}. Widok stanowiska do badania charakterystyki mechanicznej silnika")
        r.para(
            f"Na rysunku {fig_ref()} przedstawiono tabliczkę znamionową badanego silnika pro-MOTOR MS 71 1-4. "
            "Jest to silnik o mocy 0,25 kW, dwóch parach biegunów (p = 2), prędkości znamionowej "
            "1350 obr/min przy 50 Hz (prędkość synchroniczna 1500 obr/min) i prądach znamionowych "
            "1,45/0,84 A dla połączenia w trójkąt/gwiazdę (230/400 V)."
        )
        fig_no += 1
        r.figure(IMG_DIR / "tabliczka_silnik_badany.jpg", 12.0, f"2.{fig_no}. Tabliczka znamionowa silnika użytego do badania charakterystyki mechanicznej")

    r.para(
        f"Pomiar prądu rozruchowego przeprowadzono na stanowisku przedstawionym na rysunku {fig_ref()}. "
        "Obiektem badań był trójfazowy silnik asynchroniczny klatkowy napędzający wentylator "
        "promieniowy, do którego wylotu przyłączony jest rurociąg. W skład stanowiska wchodzą "
        "ponadto: łącznik silnika, przemiennik częstotliwości, cyfrowy miernik parametrów sieci "
        "LUMEL oraz komputer PC z oprogramowaniem rejestrującym. Miernik mierzy wartości skuteczne "
        "prądów w trzech fazach zasilania (I_{rms1}, I_{rms2}, I_{rms3}), które są zapisywane "
        "w komputerze z okresem próbkowania 1 s."
    )
    fig_no += 1
    r.figure(IMG_DIR / "stanowisko_prad_rozruchowy.jpg", 13.0, f"2.{fig_no}. Widok stanowiska do badania prądów rozruchowych i parametrów zasilania")
    fan_fig = fig_ref()
    r.para(
        f"Tabliczkę znamionową silnika napędzającego wentylator przedstawiono na rysunku {fan_fig}, "
        "a odczytane z niej dane znamionowe zestawiono w tabeli 2.1."
    )
    fig_no += 1
    r.figure(IMG_DIR / "tabliczka_wentylator.jpg", 10.0, f"2.{fig_no}. Tabliczka znamionowa silnika napędzającego wentylator")
    r.table(
        "2.1. Dane znamionowe silnika napędzającego wentylator",
        ["Parametr", "Wartość"],
        [
            ["Producent / typ", "BESEL S.A. / SKg 63-2B1"],
            ["Moc znamionowa P_{N}", "0,25 kW"],
            ["Napięcie znamionowe (Δ/Y)", "230/400 V"],
            ["Prąd znamionowy I_{N} (Δ/Y)", "1,10/0,65 A"],
            ["Częstotliwość znamionowa f_{1}", "50 Hz"],
            ["Prędkość znamionowa n_{N}", "2760 min^{-1}"],
            ["Sprawność η", "68 %"],
            ["Klasa izolacji / stopień ochrony", "F / IP54"],
            ["Rodzaj pracy", "S1"],
        ],
        [6.5, 5.0],
    )
    r.para(
        "Prędkość znamionowa 2760 obr/min jest bliska prędkości synchronicznej 3000 obr/min, "
        "co oznacza, że jest to silnik dwubiegunowy o jednej parze biegunów (p = 1)."
    )
    r.empty()

    # 3 ---------------------------------------------------------------------
    r.heading("Badania stanowiskowe")
    tab_no = 0
    if mech:
        r.para(
            "Badanie charakterystyki mechanicznej polegało na zanotowaniu wartości momentu obciążającego "
            "silnik oraz prędkości obrotowej dla 10 ustawień obciążenia – od momentu szczątkowego do "
            "maksymalnego możliwego do uzyskania. Pomiary wykonano dla dwóch zestawów parametrów zasilania, "
            "a ich wyniki zestawiono w tabeli 3.1."
        )
        tab_no += 1
        n_pts = max(len(z["pomiary"]) for z in mech)
        rows_m = []
        for k in range(n_pts):
            row = [str(k + 1)]
            for z in mech:
                if k < len(z["pomiary"]):
                    row += [pl(z["M"][k], 2), f"{z['n'][k]:.0f}"]
                else:
                    row += ["", ""]
            rows_m.append(row)
        head = ["L.p."]
        for z in mech:
            head += [f"M [N·m] ({z['f1']} Hz)", f"n [obr/min] ({z['f1']} Hz)"]
        r.table(f"3.{tab_no}. Wyniki pomiarów momentu obciążenia i prędkości obrotowej silnika", head, rows_m, [1.3] + [3.4] * (2 * len(mech)))

    t0 = rows[0][0]
    r.para(
        "Pomiar prądu rozruchowego polegał na rejestracji wartości skutecznych prądów w trzech fazach "
        "zasilania od chwili załączenia silnika, przez okres pracy ustalonej, aż do jego wyłączenia. "
        f"Pomiar wykonano trzykrotnie w dniu {t0:%d.%m.%Y} w godzinach "
        f"{t0:%H:%M:%S}–{rows[-1][0]:%H:%M:%S}. Pełny zarejestrowany przebieg przedstawiono na rysunku 3.1 – "
        "widoczne są na nim trzy rozruchy silnika, oddzielone okresami, w których silnik był wyłączony "
        "(prąd równy zeru)."
    )
    plot_overview(rows, starts, out_img / "rys_3_1_przebieg_calkowity.png")
    r.figure(out_img / "rys_3_1_przebieg_calkowity.png", 15.5, "3.1. Przebieg wartości skutecznych prądów fazowych w całym okresie rejestracji")
    r.para(
        f"W tabelach 3.{tab_no + 1}–3.{tab_no + len(starts)} zestawiono wartości prądów zarejestrowane w czasie trwania poszczególnych "
        "rozruchów (pominięto próbki zerowe, gdy silnik był wyłączony). Czas t liczony jest od chwili "
        "załączenia silnika, tj. od pierwszej niezerowej próbki."
    )
    for s in starts:
        tab_no += 1
        trows = []
        for k, (ts, cur) in enumerate(zip(s.t, s.i)):
            trows.append(
                [str(k + 1), f"{ts:%H:%M:%S}", f"{(ts - s.t[0]).total_seconds():.0f}"]
                + [pl(c, 4) for c in cur]
            )
        r.table(
            f"3.{tab_no}. Wartości skuteczne prądów fazowych zarejestrowane podczas pomiaru {s.nr}",
            ["L.p.", "Godzina", "t [s]", "I_{rms1} [A]", "I_{rms2} [A]", "I_{rms3} [A]"],
            trows,
            [1.2, 2.4, 1.6, 2.6, 2.6, 2.6],
        )

    r.para(
        "Na rysunkach 3.2–3.4 przedstawiono przebiegi prądów fazowych w funkcji czasu dla każdego "
        "z pomiarów (trzy fazy na jednym wykresie). Kolorem czerwonym zaznaczono zakres, z którego "
        "odczytano maksymalne wartości prądu podczas rozruchu (wartości maksymalne każdej z faz "
        "wyróżniono rombami), a kolorem zielonym – zakres stanu ustalonego, z którego obliczono "
        "średnią wartość prądu pracy ustalonej. Do stanu ustalonego zaliczono próbki, w których "
        "prąd każdej fazy różnił się o nie więcej niż 10 % od mediany danego rozruchu; pominięto "
        "próbki z okresu narastania prędkości oraz z chwili wyłączania silnika."
    )
    for s in starts:
        path = out_img / f"rys_3_{s.nr + 1}_pomiar_{s.nr}.png"
        plot_start(rows, s, path)
        r.figure(path, 15.5, f"3.{s.nr + 1}. Przebieg wartości skutecznych prądów fazowych podczas rozruchu – pomiar {s.nr}")

    s1, s2, s3 = starts
    r.para(
        f"We wszystkich pomiarach prąd osiąga wartość maksymalną w ciągu pierwszych 1–2 s po "
        f"załączeniu, a następnie szybko maleje do wartości ustalonej ok. {pl(i_ust_all, 2)} A. "
        f"Najwyższe maksimum prądu zarejestrowano w pomiarze 1 ({pl(max(s1.i_max), 3)} A), "
        f"a najniższe – w pomiarze 2 ({pl(max(s2.i_max), 3)} A). W pomiarze 2 maksimum fazy L1 "
        f"zarejestrowano o jedną próbkę później niż w fazach L2 i L3, a w pomiarach 2 i 3 pierwsza "
        f"niezerowa próbka ma wartość mniejszą niż druga. Wynika to z faktu, że miernik uśrednia "
        f"wartość skuteczną w oknie 1 s, a chwila załączenia przypada w różnych miejscach tego okna – "
        f"krótkotrwały impuls prądu rozruchowego jest więc „rozmywany” na jedną lub dwie próbki."
    )
    r.empty()

    # 4 ---------------------------------------------------------------------
    r.heading("Wyniki obliczeń")
    tab_no = 0
    if mech:
        r.para("Prędkość synchroniczną silnika obliczono ze wzoru:")
        r.empty()
        r.equation(sym("n", "s") + mr("=") + mfrac(mr("60∙") + sym("f", "1"), mr("p")))
        r.where([
            ("f_{1}", "częstotliwość napięcia zasilającego [Hz]"),
            ("p", "liczba par biegunów [-]"),
        ])
        r.para("Poślizg silnika dla każdego punktu pomiarowego obliczono ze wzoru:")
        r.empty()
        r.equation(mr("s=") + mfrac(sym("n", "s") + mr("-n"), sym("n", "s")))
        r.where([
            ("n_{s}", "prędkość synchroniczna [obr/min]"),
            ("n", "zmierzona prędkość obrotowa wirnika [obr/min]"),
        ])
        for z in mech:
            tab_no += 1
            r.table(
                f"4.{tab_no}. Wyniki obliczeń poślizgu dla f_{{1}} = {z['f1']} Hz (n_{{s}} = {z['n_s']:.0f} obr/min)",
                ["L.p.", "M [N·m]", "n [obr/min]", "s [-]"],
                [[str(k + 1), pl(m_, 2), f"{n_:.0f}", pl(s_, 4)] for k, (m_, n_, s_) in enumerate(zip(z["M"], z["n"], z["s"]))],
                [1.3, 3.0, 3.0, 3.0],
            )
        path = out_img / "rys_4_1_charakterystyka_mechaniczna.png"
        plot_mechanical(mech, path)
        r.para(
            "Na rysunku 4.1 przedstawiono charakterystyki mechaniczne silnika M = f(s) dla obu zestawów "
            "parametrów zasilania w jednym układzie współrzędnych. Jako punkt krytyczny przyjęto punkt "
            "pomiarowy o największym zmierzonym momencie."
        )
        r.figure(path, 15.5, "4.1. Charakterystyki mechaniczne silnika M = f(s)")
        tab_no += 1
        r.table(
            f"4.{tab_no}. Parametry krytyczne silnika",
            ["f_{1} [Hz]", "M_{k} [N·m]", "s_{k} [-]", "n_{k} [obr/min]"],
            [[str(z["f1"]), pl(z["M"][z["krytyczny"]], 2), pl(z["s"][z["krytyczny"]], 4), f"{z['n'][z['krytyczny']]:.0f}"] for z in mech],
            [2.8, 2.8, 2.8, 2.8],
        )

    r.para(
        "Wartość prądu w stanie ustalonym dla każdej fazy obliczono jako średnią arytmetyczną "
        "próbek z zakresu stanu ustalonego (zaznaczonego na rysunkach 3.2–3.4):"
    )
    r.empty()
    eq1 = r.equation(
        msub(mr("I"), mr("ust,j", False)) + mr("=") + mfrac(mr("1"), mr("N"))
        + mnary_sum(mr("k=1"), mr("N"), msub(mr("I"), mr("j", False)) + mr("(") + msub(mr("t"), mr("k")) + mr(")"))
    )
    r.where([
        ("I_{j}(t_{k})", "wartość skuteczna prądu fazy j w k-tej próbce [A]"),
        ("N", "liczba próbek w zakresie stanu ustalonego [-]"),
        ("j", "numer fazy (1, 2, 3) [-]"),
    ])
    r.para("Średnią wartość prądu w stanie ustalonym oraz średnią maksymalną wartość prądu podczas rozruchu z trzech faz obliczono ze wzorów:")
    r.empty()
    eq2 = r.equation(
        msub(mr("I"), mr("ust", False)) + mr("=")
        + mfrac(msub(mr("I"), mr("ust,1", False)) + mr("+") + msub(mr("I"), mr("ust,2", False)) + mr("+") + msub(mr("I"), mr("ust,3", False)), mr("3"))
    )
    eq3 = r.equation(
        msub(mr("I"), mr("max", False)) + mr("=")
        + mfrac(msub(mr("I"), mr("max,1", False)) + mr("+") + msub(mr("I"), mr("max,2", False)) + mr("+") + msub(mr("I"), mr("max,3", False)), mr("3"))
    )
    r.where([
        ("I_{ust,j}", "wartość prądu fazy j w stanie ustalonym [A]"),
        ("I_{max,j}", "maksymalna zarejestrowana wartość prądu fazy j podczas rozruchu [A]"),
    ])
    r.para("Współczynnik krotności prądu dla każdego pomiaru oraz średnią krotność z trzech pomiarów obliczono ze wzorów:")
    r.empty()
    eq4 = r.equation(msub(mr("k"), mr("I", False)) + mr("=") + mfrac(msub(mr("I"), mr("max", False)), msub(mr("I"), mr("ust", False))))
    eq5 = r.equation(
        msub(mr("k"), mr("śr", False)) + mr("=")
        + mfrac(msub(mr("k"), mr("I,1", False)) + mr("+") + msub(mr("k"), mr("I,2", False)) + mr("+") + msub(mr("k"), mr("I,3", False)), mr("3"))
    )
    r.where([
        ("I_{max}", f"średnia z trzech faz maksymalna wartość prądu podczas rozruchu, wzór ({eq3}) [A]"),
        ("I_{ust}", f"średnia z trzech faz wartość prądu w stanie ustalonym, wzór ({eq2}) [A]"),
        ("k_{I,i}", "współczynnik krotności prądu w i-tym pomiarze [-]"),
    ])
    r.para(
        f"Wyniki obliczeń wartości prądu w stanie ustalonym według wzorów ({eq1}) i ({eq2}) zestawiono "
        f"w tabeli 4.{tab_no + 1}, a maksymalne wartości prądu podczas rozruchu i współczynniki krotności "
        f"prądu obliczone według wzorów ({eq3}) i ({eq4}) – w tabeli 4.{tab_no + 2}."
    )
    tab_no += 1
    r.table(
        f"4.{tab_no}. Wartości prądu w stanie ustalonym",
        ["Pomiar", "Zakres czasu", "N [-]", "I_{ust,1} [A]", "I_{ust,2} [A]", "I_{ust,3} [A]", "I_{ust} [A]"],
        [
            [
                str(s.nr),
                f"{s.t[s.steady.start]:%H:%M:%S}–{s.t[s.steady.stop - 1]:%H:%M:%S}",
                str(len(s.steady)),
            ]
            + [pl(v, 4) for v in s.i_ust]
            + [pl(s.i_ust_sr, 4)]
            for s in starts
        ],
        [1.7, 3.4, 1.2, 2.1, 2.1, 2.1, 2.1],
    )
    tab_no += 1
    r.table(
        f"4.{tab_no}. Maksymalne wartości prądu podczas rozruchu i współczynnik krotności prądu",
        ["Pomiar", "I_{max,1} [A]", "I_{max,2} [A]", "I_{max,3} [A]", "I_{max} [A]", "I_{ust} [A]", "k_{I} [-]"],
        [
            [str(s.nr)] + [pl(v, 4) for v in s.i_max] + [pl(s.i_max_sr, 4), pl(s.i_ust_sr, 4), pl(s.k, 3)]
            for s in starts
        ]
        + [["Średnio", "", "", "", pl(i_max_all, 4), pl(i_ust_all, 4), pl(k_sr, 3)]],
        [1.6, 2.2, 2.2, 2.2, 2.2, 2.2, 1.7],
    )
    ks = " + ".join(pl(s.k, 3) for s in starts)
    r.para(f"Średnia krotność wzrostu prądu z trzech pomiarów, obliczona według wzoru ({eq5}), wynosi:")
    r.empty()
    r.equation(
        msub(mr("k"), mr("śr", False)) + mr("=") + mfrac(num(ks), num("3")) + mr("≈") + num(pl(k_sr, 2))
    )
    i_n = fm["I_N_delta"]
    r.para(
        f"Średni prąd w stanie ustalonym I_{{ust}} = {pl(i_ust_all, 3)} A stanowi ok. "
        f"{i_ust_all / i_n * 100:.0f} % prądu znamionowego silnika w połączeniu w trójkąt "
        f"(I_{{N}} = {pl(i_n, 2)} A) i ok. {i_ust_all / fm['I_N_star'] * 100:.0f} % prądu znamionowego "
        f"w połączeniu w gwiazdę (I_{{N}} = {pl(fm['I_N_star'], 2)} A). Wskazuje to, że silnik pracował "
        f"w połączeniu w trójkąt przy obciążeniu zbliżonym do znamionowego. Przyjmując I_{{N}} = {pl(i_n, 2)} A, "
        f"krotność maksymalnego zarejestrowanego prądu względem prądu znamionowego wynosi "
        f"I_{{max}}/I_{{N}} = {pl(i_max_all, 3)}/{pl(i_n, 2)} ≈ {pl(i_max_all / i_n, 2)}."
    )
    r.para(
        "Na podstawie danych z tabliczki znamionowej (tab. 2.1) obliczono również prędkość synchroniczną, "
        "poślizg znamionowy oraz częstotliwość obrotów wirnika silnika napędzającego wentylator:"
    )
    r.empty()
    r.equation(
        sym("n", "s") + mr("=") + mfrac(mr("60∙") + sym("f", "1"), mr("p")) + mr("=")
        + mfrac(mr(f"60∙{fm['f1']}"), mr(str(fm["p"]))) + mr(f"={n_s_fan:.0f}") + num(" obr/min")
    )
    r.equation(
        sym("s", "N") + mr("=") + mfrac(sym("n", "s") + mr("-") + sym("n", "N"), sym("n", "s"))
        + mr("=") + mfrac(mr(f"{n_s_fan:.0f}-{fm['n_N']}"), mr(f"{n_s_fan:.0f}")) + mr("=") + num(pl(s_fan, 2))
    )
    eq8 = r.equation(
        mr("f=") + mfrac(sym("f", "1") + mr("∙(1-s)"), mr("p")) + mr("=")
        + mfrac(mr(f"{fm['f1']}∙(1-") + num(pl(s_fan, 2)) + mr(")"), mr(str(fm["p"])))
        + mr(f"={pl(f_obr_fan, 0)}") + num(" Hz")
    )
    r.where([
        ("n_{s}", "prędkość synchroniczna (prędkość wirowania pola stojana) [obr/min]"),
        ("f_{1}", "częstotliwość napięcia zasilającego [Hz]"),
        ("p", "liczba par biegunów [-]"),
        ("n_{N}", "znamionowa prędkość obrotowa wirnika [obr/min]"),
        ("s_{N}", "poślizg znamionowy [-]"),
        ("f", "częstotliwość obrotów wirnika [Hz]"),
    ])
    r.para(
        f"Poślizg znamionowy silnika wentylatora wynosi {pl(s_fan * 100, 0)} %, a częstotliwość obrotów "
        f"wirnika {pl(f_obr_fan, 0)} Hz ({fm['n_N']} obr/min), co potwierdza poprawność wzoru ({eq8}): "
        f"f = n_{{N}}/60 = {fm['n_N']}/60 = {pl(fm['n_N'] / 60, 0)} Hz."
    )
    r.empty()

    # 5 ---------------------------------------------------------------------
    r.heading("Wnioski")
    r.conclusion(
        "Celem laboratorium było poznanie budowy, zasady działania i właściwości trójfazowego silnika "
        "asynchronicznego klatkowego, "
        + ("w tym wyznaczenie jego charakterystyki mechanicznej oraz " if mech else "w tym ")
        + "określenie wartości prądu pobieranego podczas rozruchu i współczynnika krotności prądu rozruchowego."
    )
    r.conclusion(
        "W celu realizacji zadania "
        + (
            "wykonano pomiary momentu i prędkości obrotowej silnika dla dwóch zestawów parametrów zasilania, "
            "obliczono poślizg i wykreślono charakterystyki M = f(s), a następnie "
            if mech
            else ""
        )
        + "zarejestrowano wartości skuteczne prądów w trzech fazach zasilania podczas trzech rozruchów "
        "silnika napędzającego wentylator. Dla każdego pomiaru wyznaczono średni prąd w stanie ustalonym, "
        "średnią maksymalną wartość prądu podczas rozruchu oraz ich stosunek, a następnie średnią "
        "krotność wzrostu prądu."
    )
    if mech:
        parts = []
        for z in mech:
            kk = z["krytyczny"]
            parts.append(
                f"dla f_{{1}} = {z['f1']} Hz: M_{{k}} = {pl(z['M'][kk], 2)} N·m, s_{{k}} = {pl(z['s'][kk], 3)}, "
                f"n_{{k}} = {z['n'][kk]:.0f} obr/min"
            )
        r.conclusion(
            "Wyznaczone charakterystyki mechaniczne mają przebieg zgodny z teorią: w zakresie małych poślizgów "
            "moment rośnie niemal liniowo wraz z poślizgiem (zakres pracy stabilnej), a po osiągnięciu momentu "
            "krytycznego zaczyna maleć (zakres niestabilny). Wyznaczone parametry krytyczne wynoszą: "
            + "; ".join(parts) + ". Zmiana częstotliwości zasilania przesuwa charakterystykę, ponieważ "
            "zmienia się prędkość synchroniczna pola wirującego."
        )
    r.conclusion(
        f"Podczas rozruchu prąd pobierany przez silnik wzrastał od zera do wartości maksymalnej "
        f"{pl(min(s.i_max_sr for s in starts), 2)}–{pl(max(s.i_max_sr for s in starts), 2)} A (średnio z trzech faz), "
        f"a już po 1–2 s malał do wartości ustalonej ok. {pl(i_ust_all, 2)} A. Prądy w poszczególnych fazach "
        f"w stanie ustalonym różniły się nieznacznie (prąd fazy L3 był o ok. {asym_lo:.0f}–{asym_hi:.0f} % mniejszy "
        f"niż w fazach L1 i L2), "
        f"co świadczy o niewielkiej asymetrii zasilania lub uzwojeń. Współczynnik krotności prądu w kolejnych "
        f"pomiarach wyniósł {', '.join(pl(s.k, 2) for s in starts)}, a średnia krotność wzrostu prądu "
        f"k_{{śr}} ≈ {pl(k_sr, 2)}."
    )
    r.conclusion(
        "Uzyskana krotność jest mniejsza od wartości typowych dla rozruchu bezpośredniego silników klatkowych "
        "(zwykle 4–8). Wynika to przede wszystkim z rozdzielczości czasowej rejestracji – miernik zapisuje "
        "wartość skuteczną uśrednioną w oknie 1 s, natomiast prąd rozruchowy małego silnika z lekkim "
        "obciążeniem wentylatorowym trwa zaledwie kilkaset milisekund, przez co jego wartość szczytowa "
        f"jest zaniżona. Potwierdza to rozrzut krotności pomiędzy pomiarami ({pl(min(s.k for s in starts), 2)}–"
        f"{pl(max(s.k for s in starts), 2)}), wynikający z różnego położenia chwili załączenia względem okna "
        "pomiarowego. Ograniczenie prądu może wynikać również z rampy rozruchowej przemiennika "
        "częstotliwości, jeżeli silnik był przez niego zasilany."
    )
    r.conclusion(
        "W dalszych badaniach należałoby zarejestrować przebiegi prądu z większą częstotliwością próbkowania "
        "(np. oscyloskopem lub rejestratorem z okresem rzędu milisekund), co pozwoliłoby wyznaczyć "
        "rzeczywistą wartość szczytową prądu rozruchowego. Warto także porównać rozruch bezpośredni "
        "z rozruchem gwiazda–trójkąt oraz z rozruchem przez przemiennik częstotliwości z różnymi czasami "
        "rampy, aby ocenić skuteczność poszczególnych metod ograniczania prądu rozruchowego."
    )

    r.save(OUT_DOCX)
    return starts, k_sr


def prepare_mechanical():
    sets = []
    for z in CHARAKTERYSTYKA_MECHANICZNA:
        if not z.get("pomiary"):
            continue
        n_s = 60 * z["f1"] / P_BADANY
        M_ = [m_ for m_, _ in z["pomiary"]]
        n_ = [n for _, n in z["pomiary"]]
        s_ = [(n_s - n) / n_s for n in n_]
        order = sorted(range(len(s_)), key=lambda k: s_[k])
        M_, n_, s_ = [M_[k] for k in order], [n_[k] for k in order], [s_[k] for k in order]
        sets.append(
            {
                "f1": z["f1"],
                "pomiary": z["pomiary"],
                "n_s": n_s,
                "M": M_,
                "n": n_,
                "s": s_,
                "krytyczny": max(range(len(M_)), key=lambda k: M_[k]),
            }
        )
    return sets


if __name__ == "__main__":
    starts, k_sr = build()
    for s in starts:
        print(
            f"Pomiar {s.nr}: I_max={[round(v, 4) for v in s.i_max]} śr={s.i_max_sr:.4f}  "
            f"I_ust={[round(v, 4) for v in s.i_ust]} śr={s.i_ust_sr:.4f}  k={s.k:.3f}  "
            f"steady={s.t[s.steady.start]:%H:%M:%S}-{s.t[s.steady.stop - 1]:%H:%M:%S}"
        )
    print(f"k_śr = {k_sr:.3f}")
    print(f"Zapisano: {OUT_DOCX}")
