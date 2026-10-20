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
from matplotlib.ticker import FuncFormatter  # noqa: E402
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

MECH_CSV = HERE / "dane" / "charakterystyka_mechaniczna.csv"

TESTED_MOTOR = {
    "f1_N": 50,
    "p": 2,
    "n_N": 1350,
}
F_ZAS = 64

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


def read_mechanical(n_s):
    """Serie pomiarowe M = f(n) odczytane ze zdjęć ekranu stanowiska.

    Identyczne odczyty sfotografowane kilkukrotnie traktowane są jako jeden punkt.
    """
    series = {}
    with MECH_CSV.open(encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            point = (int(row["n_obr_min"]), float(row["M_Nm"]))
            pts = series.setdefault(int(row["U_zas_proc_Un"]), [])
            if point not in pts:
                pts.append(point)
    out = []
    for u, pts in sorted(series.items(), reverse=True):
        pts.sort(key=lambda p: -p[0])
        n = [p[0] for p in pts]
        M = [p[1] for p in pts]
        s = [(n_s - v) / n_s for v in n]
        k = max(range(len(M)), key=lambda i: (M[i], n[i]))
        out.append({"U": u, "n": n, "M": M, "s": s, "kk": k, "M_k": M[k], "s_k": s[k], "n_k": n[k]})
    return out


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
        x0 = (s.t[0] - t0).total_seconds()
        x_pk = (s.t[max(s.peak_idx)] - t0).total_seconds()
        ax.annotate(
            f"I$_{{\\mathrm{{max{s.nr}}}}}$",
            xy=(x_pk, max(s.i_max)),
            xytext=(x_pk + 7, max(s.i_max) - 0.35),
            fontsize=9,
            va="center",
            arrowprops={"arrowstyle": "->", "lw": 0.8},
        )
        x_st0 = (s.t[s.steady.start] - t0).total_seconds()
        x_st1 = (s.t[s.steady.stop - 1] - t0).total_seconds()
        ax.annotate("", xy=(x_st0, 1.3), xytext=(x_st1, 1.3), arrowprops={"arrowstyle": "<->", "lw": 0.8})
        ax.text(x_st1 + 2, 1.3, f"I$_{{\\mathrm{{ust{s.nr}}}}}$", ha="left", va="center", fontsize=9)
        ax.text(x0 + 2, max(s.i_max) + 0.3, f"Pomiar {s.nr}", ha="center", fontsize=9)
    ax.set_xlabel(f"t [s] (od {t0:%H:%M:%S})")
    ax.set_ylabel("I [A]")
    ax.set_xlim(t[0], t[-1])
    ax.set_ylim(0, 5.0)
    ax.yaxis.set_major_formatter(pl_axis(1))
    ax.legend(loc="upper center", fontsize=8.5, ncol=3, framealpha=0.95)
    fig.tight_layout()
    fig.savefig(path)
    plt.close(fig)


def plot_mechanical(series, path):
    fig, ax = plt.subplots(figsize=(6.3, 3.9))
    colors = ("#1f5fa8", "#c0392b")
    markers = ("D", "s")
    for k, z in enumerate(series):
        ax.plot(
            z["s"], z["M"], color=colors[k], marker=markers[k], ms=4, lw=1.4,
            label=f"M$_{{{k + 1}}}$ – U$_{{\\mathrm{{zas}}}}$ = {z['U']} % U$_{{\\mathrm{{N}}}}$",
        )
        sk, mk = z["s_k"], z["M_k"]
        ax.plot([0, sk], [mk, mk], color="0.45", lw=0.8)
        ax.plot([sk, sk], [0, mk], color=colors[k], lw=0.8)
        ax.text(0.012, mk + 0.012, f"M$_{{\\mathrm{{k{k + 1}}}}}$", ha="left", va="bottom", fontsize=9)
    sks = [z["s_k"] for z in series]
    for k, z in enumerate(series):
        offset = 0.035 if z["s_k"] == max(sks) else -0.035
        ax.annotate(
            f"s$_{{\\mathrm{{k{k + 1}}}}}$",
            xy=(z["s_k"], 0.0),
            xytext=(z["s_k"] + offset, -0.11),
            textcoords="data",
            ha="center",
            fontsize=9,
            color=colors[k],
            arrowprops={"arrowstyle": "->", "color": colors[k], "lw": 0.8},
            annotation_clip=False,
        )
    ax.set_xlabel("s [-]", labelpad=14)
    ax.set_ylabel("M [N·m]")
    ax.set_xlim(0, 1.0)
    ax.set_ylim(0, 0.8)
    ax.xaxis.set_major_formatter(pl_axis(1))
    ax.yaxis.set_major_formatter(pl_axis(1))
    ax.legend(loc="upper right", fontsize=8.5, framealpha=0.95)
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
    tm = TESTED_MOTOR
    n_s_tm = 60 * F_ZAS / tm["p"]
    n_s_N = 60 * tm["f1_N"] / tm["p"]
    s_N_tm = (n_s_N - tm["n_N"]) / n_s_N
    mech = read_mechanical(n_s_tm)
    k1, k2 = mech
    u_ratio_sq = (k2["U"] / k1["U"]) ** 2
    m_k2_teor = k1["M_k"] * u_ratio_sq
    n_rows = max(len(z["n"]) for z in mech)

    def col(z, key, fmt):
        return [fmt(v) for v in z[key]] + [""] * (n_rows - len(z[key]))

    UST_TAB = "4.3"

    # 1 ---------------------------------------------------------------------
    r.heading("Cel i zakres laboratorium")
    r.para(
        "Celem laboratorium było zapoznanie się z budową, zasadą działania oraz podstawowymi "
        "właściwościami eksploatacyjnymi trójfazowego silnika asynchronicznego klatkowego, "
        "w szczególności z jego charakterystyką mechaniczną i wpływem napięcia zasilania na parametry "
        "krytyczne, a także z przebiegiem prądu pobieranego z sieci podczas rozruchu i wyznaczenie "
        "współczynnika krotności prądu rozruchowego."
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
    scope = [
        "zapoznanie się z budową stanowisk badawczych i danymi znamionowymi badanych silników",
        "wyznaczenie charakterystyk mechanicznych M = f(s) silnika zasilanego z przemiennika "
        f"częstotliwości (f_{{1}} = {F_ZAS} Hz) przy dwóch wartościach napięcia zasilania "
        f"({k1['U']} % i {k2['U']} % U_{{N}}) oraz określenie momentu, poślizgu i prędkości krytycznej",
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

    tab_no = 0
    r.para(
        f"Charakterystykę mechaniczną wyznaczono na stanowisku przedstawionym na rysunku {fig_ref()}. "
        "Badany silnik asynchroniczny zasilany jest z przemiennika częstotliwości, który umożliwia "
        "nastawienie wartości napięcia i częstotliwości zasilania. Wał silnika sprzęgnięty jest "
        "z alternatorem obciążonym rezystorem, a moment obciążający silnik regulowany jest zmianą prądu "
        "wzbudzenia alternatora, dostarczanego z zasilacza laboratoryjnego. Moment obciążenia oraz "
        "prędkość obrotowa wału są mierzone i wyświetlane na komputerze PC."
    )
    fig_no += 1
    r.figure(IMG_DIR / "stanowisko_charakterystyka.jpg", 14.0, f"2.{fig_no}. Widok stanowiska do badania charakterystyki mechanicznej silnika")
    tab_no += 1
    tm_tab = f"2.{tab_no}"
    r.para(
        f"Na rysunku {fig_ref()} przedstawiono tabliczkę znamionową badanego silnika, a odczytane z niej "
        f"dane znamionowe zestawiono w tabeli {tm_tab}."
    )
    fig_no += 1
    r.figure(IMG_DIR / "tabliczka_silnik_badany.jpg", 12.0, f"2.{fig_no}. Tabliczka znamionowa silnika użytego do badania charakterystyki mechanicznej")
    r.table(
        f"{tm_tab}. Dane znamionowe silnika użytego do badania charakterystyki mechanicznej",
        ["Parametr", "Wartość"],
        [
            ["Producent / typ", "pro-MOTOR / MS 71 1-4"],
            ["Częstotliwość znamionowa f_{1}", "50 Hz / 60 Hz"],
            ["Moc znamionowa P_{N}", "0,25 kW / 0,3 kW"],
            ["Napięcie znamionowe (Δ/Y)", "230/400 V / 265/460 V"],
            ["Prąd znamionowy I_{N} (Δ/Y)", "1,45/0,84 A"],
            ["Prędkość znamionowa n_{N}", "1350 / 1620 obr/min"],
            ["Współczynnik mocy cos φ", "0,72"],
            ["Klasa izolacji / stopień ochrony", "F / IP55"],
            ["Rodzaj pracy / temp. otoczenia", "S1 / 50 °C"],
        ],
        [6.5, 6.0],
    )
    r.para(
        f"Prędkość znamionowa {tm['n_N']} obr/min przy częstotliwości {tm['f1_N']} Hz jest bliska prędkości "
        f"synchronicznej {n_s_N:.0f} obr/min, a więc silnik ma dwie pary biegunów (p = {tm['p']}). Podczas "
        f"badań silnik był zasilany z przemiennika napięciem o częstotliwości f_{{1}} = {F_ZAS} Hz, dla której "
        f"prędkość synchroniczna wynosi {n_s_tm:.0f} obr/min (obliczenia w rozdziale 4)."
    )
    r.empty()

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
    tab_no += 1
    fan_tab = f"2.{tab_no}"
    r.para(
        f"Tabliczkę znamionową silnika napędzającego wentylator przedstawiono na rysunku {fan_fig}, "
        f"a odczytane z niej dane znamionowe zestawiono w tabeli {fan_tab}."
    )
    fig_no += 1
    r.figure(IMG_DIR / "tabliczka_wentylator.jpg", 10.0, f"2.{fig_no}. Tabliczka znamionowa silnika napędzającego wentylator")
    r.table(
        f"{fan_tab}. Dane znamionowe silnika napędzającego wentylator",
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
    r.para(
        "Badanie charakterystyki mechanicznej polegało na określeniu wartości momentu obciążającego "
        "silnik w funkcji poślizgu. Dla każdego ustawienia obciążenia – od momentu szczątkowego do "
        "maksymalnego możliwego do uzyskania – notowano wartość momentu oraz prędkości obrotowej. "
        "Pomiary wykonano dla dwóch zestawów parametrów zasilania ustalonych przez prowadzącego: "
        f"przy częstotliwości f_{{1}} = {F_ZAS} Hz i napięciu zasilania równym {k1['U']} % oraz {k2['U']} % "
        "napięcia znamionowego U_{N}. Wartości momentu M i prędkości obrotowej n odczytywano z ekranu "
        "komputera po ustaleniu się wskazań. Przy większych obciążeniach po każdym odczycie zmniejszano "
        "obciążenie w celu ograniczenia nagrzewania się rezystora."
    )
    tab_no += 1
    mech_tab = f"3.{tab_no}"
    r.para(
        f"Wyniki pomiarów zestawiono w tabeli {mech_tab}, porządkując je według malejącej prędkości "
        f"obrotowej (powtarzające się identyczne odczyty uwzględniono jednokrotnie). Indeks 1 oznacza "
        f"pomiar przy napięciu {k1['U']} % U_{{N}}, a indeks 2 – przy napięciu {k2['U']} % U_{{N}}."
    )
    r.table(
        f"{mech_tab}. Wyniki pomiarów momentu i prędkości obrotowej silnika (f_{{1}} = {F_ZAS} Hz)",
        ["L.p.", "n_{1} [obr/min]", "M_{1} [N·m]", "n_{2} [obr/min]", "M_{2} [N·m]"],
        [
            [str(k + 1)] + list(row)
            for k, row in enumerate(zip(
                col(k1, "n", str), col(k1, "M", lambda v: pl(v, 2)),
                col(k2, "n", str), col(k2, "M", lambda v: pl(v, 2)),
            ))
        ],
        [1.3, 3.0, 3.0, 3.0, 3.0],
    )

    t0 = rows[0][0]
    r.para(
        "Pomiar prądu rozruchowego polegał na rejestracji wartości skutecznych prądów w trzech fazach "
        "zasilania od chwili załączenia silnika, przez okres pracy ustalonej, aż do jego wyłączenia. "
        f"Pomiar wykonano trzykrotnie w dniu {t0:%d.%m.%Y} w godzinach "
        f"{t0:%H:%M:%S}–{rows[-1][0]:%H:%M:%S}. Pełny zarejestrowany przebieg przedstawiono na rysunku 3.1 – "
        "widoczne są na nim trzy rozruchy silnika, oddzielone okresami, w których silnik był wyłączony "
        "(prąd równy zeru). Na wykresie zaznaczono maksymalne wartości prądu podczas rozruchu I_{max1}–I_{max3} "
        "oraz zakresy stanu ustalonego, z których wyznaczono wartości prądu I_{ust1}–I_{ust3}."
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
        "Maksymalną wartość prądu każdej fazy podczas rozruchu odczytano z dwóch pierwszych próbek "
        "po załączeniu silnika. Do zakresu stanu ustalonego zaliczono próbki, w których prąd każdej "
        "fazy różnił się o nie więcej niż 10 % od mediany danego rozruchu – pominięto próbki z okresu "
        "narastania prędkości oraz z chwili wyłączania silnika. Przyjęte zakresy stanu ustalonego "
        f"podano w tabeli {UST_TAB}."
    )

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
    r.para(
        "Prędkość synchroniczną badanego silnika (pro-MOTOR MS 71 1-4) przy częstotliwości napięcia "
        "zasilającego ustawionej na przemienniku obliczono ze wzoru:"
    )
    r.empty()
    r.equation(
        sym("n", "s") + mr("=") + mfrac(mr("60∙") + sym("f", "1"), mr("p")) + mr("=")
        + mfrac(mr(f"60∙{F_ZAS}"), mr(str(tm["p"]))) + mr(f"={n_s_tm:.0f}") + num(" obr/min")
    )
    r.where([
        ("f_{1}", "częstotliwość napięcia zasilającego [Hz]"),
        ("p", "liczba par biegunów [-]"),
    ])
    r.para(f"Poślizg dla każdego punktu pomiarowego z tabeli {mech_tab} obliczono ze wzoru:")
    r.empty()
    eq_s = r.equation(
        mr("s=") + mfrac(sym("n", "s") + mr("-n"), sym("n", "s"))
    )
    r.where([
        ("n_{s}", "prędkość synchroniczna [obr/min]"),
        ("n", "zmierzona prędkość obrotowa wirnika [obr/min]"),
    ])
    tab_no += 1
    s_tab = f"4.{tab_no}"
    r.para(
        f"Obliczone wartości poślizgu wraz z odpowiadającymi im wartościami momentu zestawiono w tabeli "
        f"{s_tab}, a wyznaczone na ich podstawie charakterystyki mechaniczne M = f(s) przedstawiono na "
        "rysunku 4.1."
    )
    r.table(
        f"{s_tab}. Poślizg i moment silnika dla obu wartości napięcia zasilania",
        ["L.p.", "s_{1} [-]", "M_{1} [N·m]", "s_{2} [-]", "M_{2} [N·m]"],
        [
            [str(k + 1)] + list(row)
            for k, row in enumerate(zip(
                col(k1, "s", lambda v: pl(v, 3)), col(k1, "M", lambda v: pl(v, 2)),
                col(k2, "s", lambda v: pl(v, 3)), col(k2, "M", lambda v: pl(v, 2)),
            ))
        ],
        [1.3, 3.0, 3.0, 3.0, 3.0],
    )
    plot_mechanical(mech, out_img / "rys_4_1_charakterystyka_mechaniczna.png")
    r.figure(
        out_img / "rys_4_1_charakterystyka_mechaniczna.png", 15.5,
        f"4.1. Charakterystyki mechaniczne M = f(s) silnika przy U_{{zas}} = {k1['U']} % "
        f"i {k2['U']} % U_{{N}} (f_{{1}} = {F_ZAS}\u00a0Hz)",
    )
    tab_no += 1
    kr_tab = f"4.{tab_no}"
    r.para(
        "Moment krytyczny M_{k} odczytano jako największą wartość momentu na charakterystyce, a poślizg "
        f"krytyczny s_{{k}} – jako odpowiadający mu poślizg obliczony według wzoru ({eq_s}) z prędkości "
        f"krytycznej n_{{k}}. Punkty krytyczne zaznaczono na rysunku 4.1, a ich parametry zestawiono "
        f"w tabeli {kr_tab}. Przykładowo dla napięcia {k1['U']} % U_{{N}}:"
    )
    r.empty()
    r.equation(
        msub(mr("s"), mr("k1", False)) + mr("=")
        + mfrac(sym("n", "s") + mr("-") + msub(mr("n"), mr("k1", False)), sym("n", "s")) + mr("=")
        + mfrac(mr(f"{n_s_tm:.0f}-{k1['n_k']}"), mr(f"{n_s_tm:.0f}"))
        + mr("=") + num(pl(k1["s_k"], 4)) + mr("≈") + num(pl(k1["s_k"], 2))
    )
    r.table(
        f"{kr_tab}. Parametry krytyczne silnika",
        ["U_{zas} [% U_{N}]", "n_{k} [obr/min]", "M_{k} [N·m]", "s_{k} [-]"],
        [[f"{z['U']} %", str(z["n_k"]), pl(z["M_k"], 2), pl(z["s_k"], 2)] for z in mech],
        [3.4, 3.0, 3.0, 3.0],
    )
    r.para(
        "Zgodnie z teorią moment krytyczny silnika asynchronicznego jest proporcjonalny do kwadratu "
        "napięcia zasilania, natomiast poślizg krytyczny od napięcia nie zależy. Teoretyczną wartość "
        f"momentu krytycznego przy napięciu {k2['U']} % U_{{N}}, wyznaczoną na podstawie pomiaru przy "
        f"{k1['U']} % U_{{N}}, obliczono ze wzoru:"
    )
    r.empty()
    r.equation(
        msub(mr("M"), mr("k2,teor", False)) + mr("=") + msub(mr("M"), mr("k1", False)) + mr("∙")
        + "<m:sSup><m:e><m:d><m:e>" + mfrac(msub(mr("U"), mr("2", False)), msub(mr("U"), mr("1", False)))
        + "</m:e></m:d></m:e><m:sup>" + mr("2") + "</m:sup></m:sSup>"
        + mr("=") + num(pl(k1["M_k"], 2)) + mr("∙")
        + "<m:sSup><m:e><m:d><m:e>" + mfrac(num(f"{k2['U']} %"), num(f"{k1['U']} %"))
        + "</m:e></m:d></m:e><m:sup>" + mr("2") + "</m:sup></m:sSup>"
        + mr("≈") + num(pl(m_k2_teor, 2)) + num(" N∙m")
    )
    r.where([
        ("M_{k1}", f"moment krytyczny zmierzony przy napięciu U_{{1}} = {k1['U']} % U_{{N}} [N·m]"),
        ("U_{1}, U_{2}", "napięcia zasilania w obu pomiarach [V]"),
    ])
    kk = k1["kk"]
    gap = min(k1["s"][kk] - k1["s"][kk - 1], k1["s"][kk + 1] - k1["s"][kk])
    r.para(
        f"Zmierzony moment krytyczny przy {k2['U']} % U_{{N}} wynosi {pl(k2['M_k'], 2)} N·m i jest o ok. "
        f"{(k2['M_k'] / m_k2_teor - 1) * 100:.0f} % większy od wartości teoretycznej {pl(m_k2_teor, 2)} N·m. "
        f"Stosunek zmierzonych momentów krytycznych wynosi M_{{k2}}/M_{{k1}} = {pl(k2['M_k'] / k1['M_k'], 3)}, "
        f"a stosunek kwadratów napięć (U_{{2}}/U_{{1}})^{{2}} = {pl(u_ratio_sq, 3)}. Poślizgi krytyczne "
        f"w obu pomiarach różnią się o {pl(abs(k1['s_k'] - k2['s_k']), 2)}, czyli mniej niż odstęp między "
        f"sąsiednimi punktami pomiarowymi w otoczeniu maksimum charakterystyki przy {k1['U']} % U_{{N}} "
        f"(co najmniej {pl(gap, 2)}). Dla porównania poślizg znamionowy silnika wynikający z tabliczki "
        f"znamionowej (n_{{N}} = {tm['n_N']} obr/min przy {tm['f1_N']} Hz) wynosi s_{{N}} = "
        f"({n_s_N:.0f} – {tm['n_N']})/{n_s_N:.0f} = {pl(s_N_tm, 2)}, a więc jest ok. "
        f"{pl(k1['s_k'] / s_N_tm, 1)}-krotnie mniejszy od poślizgu krytycznego."
    )

    r.para(
        "Wartość prądu w stanie ustalonym dla każdej fazy obliczono jako średnią arytmetyczną "
        f"próbek z zakresu stanu ustalonego (zakresy podano w tabeli {UST_TAB}):"
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
    assert f"4.{tab_no}" == UST_TAB
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
        f"Na podstawie danych z tabliczki znamionowej (tab. {fan_tab}) obliczono również prędkość synchroniczną, "
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
        "asynchronicznego klatkowego, w tym wyznaczenie jego charakterystyki mechanicznej i parametrów "
        "krytycznych oraz określenie wartości prądu pobieranego podczas rozruchu i współczynnika krotności "
        "prądu rozruchowego."
    )
    r.conclusion(
        "W celu realizacji zadania zmierzono moment i prędkość obrotową silnika pro-MOTOR MS 71 1-4 "
        f"zasilanego z przemiennika częstotliwości (f_{{1}} = {F_ZAS} Hz) przy napięciu {k1['U']} % "
        f"i {k2['U']} % U_{{N}}, wyznaczono jego charakterystyki mechaniczne M = f(s) i odczytano z nich "
        "parametry krytyczne, a następnie zarejestrowano wartości skuteczne prądów w trzech fazach zasilania podczas "
        "trzech rozruchów silnika napędzającego wentylator. Dla każdego rozruchu wyznaczono średni prąd "
        "w stanie ustalonym, średnią maksymalną wartość prądu podczas rozruchu oraz ich stosunek, "
        "a następnie średnią krotność wzrostu prądu."
    )
    r.conclusion(
        f"Wyznaczone parametry krytyczne wynoszą: przy {k1['U']} % U_{{N}} – M_{{k}} = {pl(k1['M_k'], 2)} N·m, "
        f"s_{{k}} ≈ {pl(k1['s_k'], 2)} (n_{{k}} = {k1['n_k']} obr/min), a przy {k2['U']} % U_{{N}} – "
        f"M_{{k}} = {pl(k2['M_k'], 2)} N·m, s_{{k}} ≈ {pl(k2['s_k'], 2)} (n_{{k}} = {k2['n_k']} obr/min). "
        f"Obniżenie napięcia zasilania z {k1['U']} % do {k2['U']} % U_{{N}} spowodowało spadek momentu "
        f"krytycznego o ok. {(1 - k2['M_k'] / k1['M_k']) * 100:.0f} %, przy niewielkiej zmianie poślizgu "
        "krytycznego, mieszczącej się w rozdzielczości pomiaru. Jest to zgodne z teorią: moment krytyczny "
        f"jest proporcjonalny do kwadratu napięcia (wartość teoretyczna {pl(m_k2_teor, 2)} N·m różni się od "
        f"zmierzonej o ok. {(k2['M_k'] / m_k2_teor - 1) * 100:.0f} %), a poślizg krytyczny od napięcia nie zależy. "
        "Obie charakterystyki mają kształt zgodny ze wzorem Klossa: przy małych poślizgach moment rośnie "
        "niemal liniowo, po przekroczeniu poślizgu krytycznego maleje, a przy niższym napięciu cała "
        "charakterystyka leży poniżej charakterystyki dla wyższego napięcia. Stabilna praca silnika jest "
        "możliwa dla poślizgów mniejszych od krytycznego, tj. przy prędkości większej niż "
        f"{k1['n_k']} obr/min ({k1['U']} % U_{{N}}) i {k2['n_k']} obr/min ({k2['U']} % U_{{N}}). "
        "Obniżanie napięcia zmniejsza przeciążalność silnika, a prędkość "
        "obrotową w zakresie pracy stabilnej zmienia jedynie w niewielkim stopniu, dlatego jest mało "
        "efektywną metodą regulacji prędkości."
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
        "rampy, aby ocenić skuteczność poszczególnych metod ograniczania prądu rozruchowego. W przypadku "
        "charakterystyki mechanicznej warto zagęścić punkty pomiarowe w otoczeniu momentu krytycznego, "
        "co pozwoliłoby dokładniej wyznaczyć poślizg krytyczny, oraz powtórzyć pomiary przy napięciu "
        "znamionowym i przy różnych częstotliwościach zasilania, aby porównać wpływ obu metod sterowania "
        "na kształt charakterystyki."
    )

    r.save(OUT_DOCX)
    return starts, k_sr


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
