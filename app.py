import base64
import html
import io
import os
import re

import openpyxl
import pandas as pd
import streamlit as st

import codes_data as data

st.set_page_config(page_title="Kunststoffverpackungssteuer", layout="centered")

st.markdown("""
<style>
html, body, .stApp { font-family: Arial, sans-serif; }
.header-wrap { text-align: center; width: 100%; }
.logo-wrap { width: 100%; text-align: center; margin-bottom: 0.5rem; }
.logo-wrap img { width: 200px; }
.header-wrap .titleblock { color: #004996; font-family: Arial, sans-serif; }
.header-wrap .titleblock .maintitle { display: block; font-size: 2rem; font-weight: bold; margin: 0.5rem 0 0; }
.header-wrap .titleblock .subtitle { display: block; font-size: 1rem; font-weight: normal; margin-top: 2px; }
/* Auswahl als Pills */
div[data-testid="stRadio"] > div { padding-top: 0; }
div[data-testid="stRadio"] div[role="radiogroup"] { gap: 8px; flex-wrap: wrap; }
div[data-testid="stRadio"] div[role="radiogroup"] label {
  border: 1px solid #004996; border-radius: 20px; padding: 6px 16px;
  color: #004996; background: #fff; position: relative; margin-right: 0;
  cursor: pointer;
}
div[data-testid="stRadio"] div[role="radiogroup"] label:hover { background: #eaf2fb; }
div[data-testid="stRadio"] div[role="radiogroup"] label p { color: #004996; }
div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) { background: #004996; border-color: #004996; color: #ffffff; }
div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) p,
div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) span { color: #ffffff !important; }
.hl { color: #004996; font-weight: bold; }
/* Hover-Infokasten (keine falschen Worttrennungen: Umbruch nur an Leerzeichen) */
div[data-testid="stRadio"] div[role="radiogroup"] label::after {
  display: none; content: ""; white-space: pre-wrap; text-align: left;
  position: absolute; top: 120%; left: 0; width: 340px; max-width: 80vw;
  background: #E8F2FC; border: 1px solid #b9d4ee; border-radius: 8px;
  padding: 12px 14px; color: #1a1a1a; font-size: 10pt; line-height: 1.5;
  font-weight: normal; z-index: 9999; box-shadow: 0 2px 8px rgba(0,60,120,.15);
  word-break: normal; overflow-wrap: break-word; hyphens: none;
}
div[data-testid="stRadio"] div[role="radiogroup"] label:nth-of-type(1):hover::after { display: block; content: "Sie kennen die KN-/WA-Nummer Ihres Produkts (z. B. aus der Zollanmeldung oder dem Warenverzeichnis für die Außenhandelsstatistik). Die App zeigt, ob dieser Code im untersuchten Gesetzentwurf ausgewählt ist."; }
div[data-testid="stRadio"] div[role="radiogroup"] label:nth-of-type(2):hover::after { display: block; content: "Sie kennen die GP-Nummer (Produktionsmeldenummer) Ihres Produkts aus der Produktionsstatistik. Die App zeigt die zugehörigen KN-/WA-Codes und deren Einordnung im untersuchten Gesetzentwurf."; }
div[data-testid="stRadio"] div[role="radiogroup"] label:nth-of-type(3):hover::after { display: block; content: "Alle Zeilen, die Sie in den beiden Code-Ansichten mit einem Häkchen markiert haben, werden hier zusammengeführt – als laufende Dokumentation Ihrer Auswahl. Sie können sie als CSV herunterladen."; }
.result-blue { background:#E8F2FC; border:1px solid #b9d4ee; border-radius:8px; padding:16px 18px; margin:8px 0 4px; }
.result-orange { background:#FDEBD7; border:1px solid #E8B04B; border-radius:8px; padding:16px 18px; margin:8px 0 4px; }
.result-blue h3, .result-orange h3 { margin:0 0 8px; color:#004996; font-size:12pt; }
.badge { display:inline-block; border-radius:12px; padding:3px 14px; font-weight:bold; font-size:10.5pt; }
.badge-orange { background:#E67E22; color:#fff; }
.badge-blue { background:#004996; color:#fff; }
.step { color:#004996; font-weight:bold; margin-bottom:4px; }
.footnote { font-size:9.5pt; color:#5a6b7d; line-height:1.5; }
.question { font-size:11.5pt; line-height:1.55; }
.example-box { background:#f7fbff; border:1px dashed #b9d4ee; border-radius:8px; padding:10px 14px; margin:10px 0 0; font-size:10.5pt; line-height:1.5; color:#3a4a58; }
.intro-box { background:#E8F2FC; border:1px solid #b9d4ee; border-radius:8px; padding:16px 20px; margin:4px 0 16px; font-size:11pt; line-height:1.6; color:#1a1a1a; }
.intro-box p { margin:0 0 10px; }
.intro-box p:last-child { margin-bottom:0; }
.intro-box strong { color:#004996; }
.dl-box { background:#f7fbff; border:1px dashed #b9d4ee; border-radius:8px; padding:10px 14px; margin:10px 0 0; font-size:10.5pt; line-height:1.5; color:#3a4a58; }
.term-tip { border-bottom: 1px dotted #004996; cursor: help; position: relative; }
.term-tip::after {
  content: attr(data-tip); display: none; white-space: pre-wrap; text-align: left;
  position: absolute; top: 130%; left: 0; width: 320px; max-width: 80vw;
  background: #E8F2FC; border: 1px solid #b9d4ee; border-radius: 8px;
  padding: 12px 14px; color: #1a1a1a; font-size: 9.5pt; line-height: 1.5;
  font-weight: normal; z-index: 9999; box-shadow: 0 2px 8px rgba(0,60,120,.15);
}
.term-tip:hover::after { display: block; }
.doc-table-wrap {
  overflow-x: scroll; overflow-y: auto; max-height: 600px; margin: 10px 0 4px;
  border: 1px solid #b9d4ee; border-radius: 8px;
  scrollbar-width: auto; scrollbar-color: #004996 #eaf2fb;
}
.doc-table-wrap::-webkit-scrollbar { height: 14px; width: 14px; }
.doc-table-wrap::-webkit-scrollbar-track { background: #eaf2fb; }
.doc-table-wrap::-webkit-scrollbar-thumb { background: #004996; border-radius: 7px; border: 3px solid #eaf2fb; }
.doc-table { border-collapse: collapse; width: max-content; min-width: 100%; font-size: 10pt; }
.doc-table thead th {
  background:#004996; color:#fff; padding:8px 12px; text-align:left; white-space:nowrap;
  position: sticky; top: 0; z-index: 1;
}
.doc-table td { padding:7px 12px; border-bottom:1px solid #e3ecf5; vertical-align:top; max-width:320px; white-space:pre-wrap; }
.doc-table tbody tr:nth-child(even) td { background:#f7fbff; }
.scroll-hint { color:#004996; font-size:9.5pt; margin:0 0 4px; }
</style>
""", unsafe_allow_html=True)

# ---------------- Begriffs-Tooltips ----------------
TERM_TIPS = {
    "Außenhandelscode": (
        "KN-/WA-Code: Die achtstellige Zolltarifnummer der Kombinierten Nomenklatur (KN), "
        "die bis zum 8-Steller mit der Warennummer des Warenverzeichnisses für die "
        "Außenhandelsstatistik (WA) identisch ist. Nur bei der Einfuhr aus Drittländern "
        "können bis zu drei Stellen hinzukommen (elfstellige Zolltarifnummer)."
    ),
    "Produktionscode": (
        "GP-Nummer: Die neunstellige Produktionsmeldenummer der GP-Klassifikation 2026, "
        "mit der Inlandsproduktionen in der Produktionsstatistik erfasst werden. "
        "Sie ist nicht deckungsgleich mit den KN-/WA-Nummern der Außenhandelsstatistik."
    ),
    "WA-Nummer": (
        "Warennummer des Warenverzeichnisses für die Außenhandelsstatistik (WA2026). "
        "Grundlage der KN-/Zolltarifnummern; hier als achtstellige Nummer geführt."
    ),
    "GP-Nummer": (
        "Produktionsmeldenummer der Güterklassifikation (GP2026). Erfasst die "
        "Inlandsproduktion; die App enthält die Nummern unter den Güterklassen "
        "2221, 2222, 2223 und 2226."
    ),
    "vZTA": (
        "Verbindliche Zolltarifauskunft: Für die verbindliche zolltarifliche Einordnung "
        "ist die Zollverwaltung Ansprechpartner."
    ),
    "KunststoffVStG": (
        "Der bislang untersuchte Entwurf eines Gesetzes über eine nationale "
        "Kunststoffverpackungsteuer (KunststoffVStG) sieht einen Steuersatz von 550 €/t "
        "auf bestimmte industrielle Vorprodukte und Verpackungsmaterialien vor. Die "
        "steuerpflichtigen Erzeugnisse werden anhand ausgewählter KN-/WA-Codes "
        "(§ 1 Abs. 2 Nr. 1–4) abgegrenzt."
    ),
}

_TERM_TIP_PATTERNS = [
    (r"Außenhandelscode", "Außenhandelscode"),
    (r"Produktionscode", "Produktionscode"),
    (r"WA-Nummern?", "WA-Nummer"),
    (r"GP-Nummern?", "GP-Nummer"),
    (r"vZTA", "vZTA"),
    (r"KunststoffVStG", "KunststoffVStG"),
]
_TERM_TIP_PATTERN = re.compile(
    r"(?<![A-Za-zÄÖÜäöüß])(" + "|".join(p for p, _ in _TERM_TIP_PATTERNS) + r")(?![A-Za-zÄÖÜäöüß])"
)
_TERM_TIP_LOOKUP = [(re.compile(p + "$"), key) for p, key in _TERM_TIP_PATTERNS]


def highlight_terms(text):
    def _wrap(m):
        word = m.group(1)
        key = next(key for pattern, key in _TERM_TIP_LOOKUP if pattern.match(word))
        tip = html.escape(TERM_TIPS[key], quote=True)
        return f'<span class="term-tip" data-tip="{tip}">{word}</span>'

    return _TERM_TIP_PATTERN.sub(_wrap, text)


def md_lines(text):
    """Absatzweise als Markdown ausgeben (Quelltexte enthalten Leerzeilen)."""
    for para in str(text).split("\n\n"):
        para = para.strip()
        if para:
            st.markdown(para)


# ---------------- Datenaufbereitung ----------------
def _flat(s):
    return str(s).replace("\n", " / ") if s else ""


WA_BY_CODE = {r["code"]: r for r in data.WA_ROWS}
GP_BY_CODE = {r["code"]: r for r in data.GP_ROWS}

TAB_WA = 'Übersicht Außenhandelscodes u. a. „Kunststoffe und Waren daraus“'
TAB_GP = 'Übersicht Produktionscodes u. a. „Kunststoffwaren“'
TAB_DOC = "Dokumentation"

QUELLEN_DIR = os.path.join(os.path.dirname(__file__), "quellen")
WA_SOURCE_PDF = os.path.join(QUELLEN_DIR, "WA2026-Warenverzeichnis-Aussenhandelsstatistik.pdf")
GP_SOURCE_XLSX = os.path.join(QUELLEN_DIR, "GP2026-Gueterverzeichnis-Produktionsstatistik.xlsx")
GP_SOURCE_KONKORDANZ = os.path.join(QUELLEN_DIR, "Gegenueberstellung-Warennummern-Warenklassifikationen-2026.xlsx")


@st.cache_data
def _read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def render_quelldokumente(key_suffix):
    if os.path.exists(WA_SOURCE_PDF):
        st.download_button(
            "Warenverzeichnis für die Außenhandelsstatistik 2026 (PDF) herunterladen",
            data=_read_bytes(WA_SOURCE_PDF),
            file_name="WA2026-Warenverzeichnis-Aussenhandelsstatistik.pdf",
            mime="application/pdf",
            key=f"dl_wa_pdf_{key_suffix}",
        )
    if os.path.exists(GP_SOURCE_XLSX):
        st.download_button(
            "Güterverzeichnis für die Produktionsstatistik GP2026 (Excel) herunterladen",
            data=_read_bytes(GP_SOURCE_XLSX),
            file_name="GP2026-Gueterverzeichnis-Produktionsstatistik.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_gp_xlsx_{key_suffix}",
        )
    if os.path.exists(GP_SOURCE_KONKORDANZ):
        st.download_button(
            "Gegenüberstellung Warennummern zu anderen Warenklassifikationen 2026 (Excel) herunterladen",
            data=_read_bytes(GP_SOURCE_KONKORDANZ),
            file_name="Gegenueberstellung-Warennummern-Warenklassifikationen-2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_gp_konkordanz_{key_suffix}",
        )

# ---------------- Kopf ----------------
with open(os.path.join(os.path.dirname(__file__), "logo.jpg"), "rb") as _f:
    _logo_b64 = base64.b64encode(_f.read()).decode()
st.markdown(
    '<div class="header-wrap">'
    f'<div class="logo-wrap"><img src="data:image/jpeg;base64,{_logo_b64}" alt="IK Industrievereinigung Kunststoffverpackungen e.V."></div>'
    '<div class="titleblock">'
    '<span class="maintitle">Kunststoffverpackungssteuer</span>'
    '<span class="subtitle">Fallen die von Ihnen in Deutschland produzierten Produkte, die von Ihnen importierten und exportierten Waren unter die Kunststoffverpackungssteuer?</span>'
    '</div>'
    '</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="intro-box">'
    '<p><strong>Wozu dient diese Übersicht?</strong> Der Entwurf eines Gesetzes über eine nationale '
    '<strong>Kunststoffverpackungsteuer</strong> (KunststoffVStG) sieht einen Steuersatz von '
    '<strong>550 €/t</strong> auf bestimmte industrielle Vorprodukte und Verpackungsmaterialien vor. '
    'Die Abgrenzung der steuerpflichtigen Erzeugnisse erfolgt anhand ausgewählter '
    'Zolltarifnummern der Kombinierten Nomenklatur (KN), die auch dem Warenverzeichnis für die '
    'Außenhandelsstatistik (WA) zugrunde liegt. Die vorliegende Übersicht verknüpft '
    'Außenhandelsmeldungen (KN-/WA-Nummern) mit Produktionsmeldungen (GP-Nummern), damit Sie '
    'erkennen können, ob die von Ihnen hergestellten, importierten oder exportierten Produkte '
    'in den von der Steuer erfassten Warenkreis fallen.</p>'
    '<p><strong>An wen richtet sich der Check?</strong> An alle IK-Mitglieder, die '
    'Kunststoffwaren, Verpackungsmaterialien oder verpackte Produkte herstellen, einführen oder '
    'handeln – und die ihre <strong>Außenhandelscode</strong>-Nummern (WA) oder ihre '
    '<strong>Produktionscode</strong>-Nummern (GP) in den statistischen Meldungen verwenden.</p>'
    '<p>Dieses Tool ist keine Rechtsberatung im Einzelfall. Es beschreibt die Zuordnung der '
    'Codes zum bislang untersuchten Entwurfswortlaut und ersetzt keine steuerliche oder '
    'zolltarifliche Prüfung (für verbindliche Auskünfte ist die Zollverwaltung Ansprechpartner, '
    'Stichwort <strong>vZTA</strong>). „Code ausgewählt“ beschreibt den Warenkreis, nicht '
    'sämtliche Voraussetzungen tatsächlicher Steuerpflicht.</p>'
    '</div>',
    unsafe_allow_html=True,
)

with st.expander("Hintergrund: der untersuchte Entwurf"):
    md_lines(data.DECKBLATT["hintergrund"])
    st.markdown("---")
    md_lines(data.DECKBLATT["gemeinsam"])

with st.expander("Funktionen"):
    st.markdown(
        "**Suche:** Geben Sie oben einen Code oder ein Schlagwort ein – die Suche "
        "durchsucht gleichzeitig die Außenhandelscodes (WA) und die Produktionscodes (GP) "
        "und gruppiert die Treffer nach der zugehörigen vierstelligen WA-Position bzw. "
        "GP-Güterklasse."
    )
    st.markdown(
        "**Ausgewähltes Ergebnis:** Ein Klick auf einen Treffer zeigt Code, amtliche "
        "Beschreibung und die Einordnung zur Kunststoffverpackungsteuer direkt unterhalb "
        "des Treffers an. Orange hinterlegte Ergebnisse sind im untersuchten Entwurf "
        "erfasst, blau hinterlegte nicht."
    )
    st.markdown(
        f"**Code-Verzeichnisse:** In den Bereichen *{TAB_WA}* und *{TAB_GP}* können Sie "
        "die vollständigen Verzeichnisse durchblättern und einzelne Zeilen über das "
        "Häkchen „Dokumentieren“ markieren."
    )
    st.markdown(
        f"**Dokumentation:** Alle über die Suche oder die Verzeichnisse markierten Zeilen "
        f"werden automatisch im Bereich „{TAB_DOC}“ zusammengeführt und lassen sich dort "
        "als Excel-Datei herunterladen."
    )
    st.markdown(
        "**Quelldokumente:** In den jeweiligen Bereichen sowie unter „Quellen und "
        "Hinweise“ stehen die amtlichen Originaldokumente (Warenverzeichnis WA2026, "
        "Güterverzeichnis GP2026, Gegenüberstellung der Warennummern) zur Einsicht und "
        "zum Download bereit."
    )

with st.expander("Stand und Aussagegrenzen"):
    md_lines(data.DECKBLATT["stand"])

with st.expander("Quellen und Hinweise"):
    for thema, erl in data.QUELLEN:
        st.markdown(f"**{thema}**  \n{erl}")
        st.markdown("---")
    st.markdown("**Amtliche Originaldokumente zur Einsicht**")
    render_quelldokumente("quellen")

st.divider()

def doc_table_html(doc_df):
    """Dokumentation als HTML-Tabelle mit immer sichtbarer horizontaler Scrollleiste."""
    cols = list(doc_df.columns)
    thead = "".join(f"<th>{html.escape(str(c))}</th>" for c in cols)
    body_rows = []
    for _, row in doc_df.iterrows():
        cells = "".join(f"<td>{html.escape(str(row[c]))}</td>" for c in cols)
        body_rows.append(f"<tr>{cells}</tr>")
    return (
        '<div class="doc-table-wrap"><table class="doc-table">'
        f"<thead><tr>{thead}</tr></thead><tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )


def doc_to_xlsx(doc_df):
    """Dokumentation als formatierte Excel-Datei (Tabelle mit Kopfzeile, Filter, Spaltenbreiten)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dokumentation"

    headers = list(doc_df.columns)
    ws.append(headers)
    for _, row in doc_df.iterrows():
        ws.append([row[h] for h in headers])

    n_rows, n_cols = len(doc_df) + 1, len(headers)
    ref = f"A1:{openpyxl.utils.get_column_letter(n_cols)}{n_rows}"
    tab = openpyxl.worksheet.table.Table(displayName="Dokumentation", ref=ref)
    tab.tableStyleInfo = openpyxl.worksheet.table.TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True
    )
    ws.add_table(tab)

    header_fill = openpyxl.styles.PatternFill("solid", fgColor="004996")
    header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = openpyxl.styles.Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    for i, h in enumerate(headers, 1):
        letter = openpyxl.utils.get_column_letter(i)
        if h in ("Amtliche Waren-/Güterbeschreibung", "Zuordnungshinweis"):
            ws.column_dimensions[letter].width = 60
        elif "Code" in h or h in ("Quellblatt", "Einordnung laut Gesetzentwurf",
                                 "Rechtsbezug / Zusatzkriterium", "GP-Nummern",
                                 "WA-Codes im Entwurf", "WA-Codes nicht im Entwurf",
                                 "WA-Quelle (PDF-Seite)"):
            ws.column_dimensions[letter].width = 24
        else:
            ws.column_dimensions[letter].width = 30
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(vertical="top", wrap_text=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------- Code- und Schlagwortsuche (bereichsübergreifend) ----------------
def _norm(s):
    return str(s).replace("\xa0", " ").strip()


def short_desc(text, n=100):
    t = _norm(text).replace("\n", " ")
    return t if len(t) <= n else t[: n - 1].rstrip() + "…"


def first_code(text):
    t = _norm(text)
    if not t or t.lower().startswith("keine"):
        return ""
    return t.split("\n")[0].split(" (")[0].strip()


def _matches(term, digits, code, haystack):
    if not term:
        return False
    if term in haystack:
        return True
    return bool(digits) and digits in code.replace(" ", "")


def search_all(term):
    t = term.strip().lower()
    digits = re.sub(r"\D", "", t)
    results = []
    for r in data.WA_ROWS:
        title = data.WA_POSITION_TITLES.get(r["pos"], "")
        haystack = (r["beschreibung"] + " " + title).lower()
        if _matches(t, digits, r["code"], haystack):
            results.append(("wa", r))
    for r in data.GP_ROWS:
        title = data.GP_KLASSE_TITEL.get(r["pos"], "")
        haystack = (r["beschreibung"] + " " + title).lower()
        if _matches(t, digits, r["code"], haystack):
            results.append(("gp", r))
    return results


def _is_empty_wa(text):
    t = _norm(text).lower()
    return (not t) or t in ("keine", "keine numerische angabe")


def comma_list(text):
    if _is_empty_wa(text):
        return "—"
    t = _norm(text)
    return ", ".join(p.strip() for p in t.split("\n") if p.strip())


def entwurf_cell(v):
    return "0" if _is_empty_wa(v) else _norm(v)


def wa_einordnung_text(r):
    return "erfasst" if r["einordnung"] == "Code ausgewählt" else "nicht erfasst"


def gp_einordnung_text(r):
    erfasst, nicht = r["wa_im_entwurf"], r["wa_nicht_im_entwurf"]
    has_erfasst, has_nicht = not _is_empty_wa(erfasst), not _is_empty_wa(nicht)
    if has_erfasst:
        label = "zugehörig erfassten Außenhandelscodes" if "\n" in _norm(erfasst) else "zugehörig einem erfassten Außenhandelscode"
        text = f"{label}: {comma_list(erfasst)}"
        if has_nicht:
            text += f"; nicht erfasste Außenhandelscodes: {comma_list(nicht)}"
        return text
    if has_nicht:
        return f"nicht erfasst (zugehöriger Außenhandelscode nicht erfasst: {comma_list(nicht)})"
    return "nicht erfasst (keine numerische Zuordnung zu einem Außenhandelscode)"


def html_multiline(text):
    return html.escape(_norm(text)).replace("\n", "<br>")


def select_result(kind, code):
    st.session_state["selected_kind"] = kind
    st.session_state["selected_code"] = code


def is_marked(kind, code):
    return (kind, code) in st.session_state.setdefault("doc_marks", set())


def toggle_mark(kind, code, widget_key):
    marks = st.session_state.setdefault("doc_marks", set())
    if st.session_state[widget_key]:
        marks.add((kind, code))
    else:
        marks.discard((kind, code))


def render_detail_card(kind, r, prefix="detail"):
    if kind == "wa":
        st.caption(f"Position {r['pos']}")
        st.markdown(
            f'<div class="example-box"><strong>{html.escape(r["code"])}</strong> · '
            f'{html_multiline(r["beschreibung"])}</div>',
            unsafe_allow_html=True,
        )
        if r["einordnung"] == "Code ausgewählt":
            box, badge = "result-orange", '<span class="badge badge-orange">Ausgewählt im untersuchten Entwurf</span>'
        else:
            box, badge = "result-blue", '<span class="badge badge-blue">Code nicht ausgewählt</span>'
        st.markdown(
            f'<div class="{box}"><h3>Einordnung Kunststoffverpackungssteuer</h3>{badge}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(f"**Rechtsbezug / Zusatzkriterium:**  \n{r['rechtsbezug'] or '—'}")
        st.markdown(f"**Zugehörige GP2026-Nummern:**  \n{comma_list(r['gp'])}")
        if r["hinweis"]:
            st.markdown(f"**Zuordnungshinweis:**  \n{r['hinweis']}")
        if r["quelle"] and int(r["quelle"]) in data.ORIGINAL_PAGES:
            if st.checkbox(
                f'Originaltext aus dem Warenverzeichnis 2026 anzeigen (PDF-Seite {r["quelle"]})',
                key=f"show_orig_{prefix}_{r['code']}",
            ):
                st.text(data.ORIGINAL_PAGES[int(r["quelle"])])
    else:
        st.caption(f"Güterklasse {r['pos']}")
        st.markdown(
            f'<div class="example-box"><strong>{html.escape(r["code"])}</strong> · '
            f'{html_multiline(r["beschreibung"])}</div>',
            unsafe_allow_html=True,
        )
        einordnung = r["einordnung"]
        if einordnung in ("Warenkreis vollständig erfasst", "Codezuordnung eindeutig"):
            box, badge_cls = "result-orange", "badge-orange"
        else:
            box, badge_cls = "result-blue", "badge-blue"
        st.markdown(
            f'<div class="{box}"><h3>Einordnung Kunststoffverpackungssteuer</h3>'
            f'<span class="badge {badge_cls}">{html.escape(einordnung)}</span></div>',
            unsafe_allow_html=True,
        )
        st.markdown("**Der GP-Position zugeordnete KN-/WA-Nummern, die der Entwurf erfasst:**  \n"
                    f"{comma_list(r['wa_im_entwurf'])}")
        st.markdown("**Der GP-Position zugeordnete KN-/WA-Nummern, die der Entwurf nicht erfasst:**  \n"
                    f"{comma_list(r['wa_nicht_im_entwurf'])}")
        if r["hinweis"]:
            st.markdown(f"**Erläuterung zur Zuordnung:**  \n{r['hinweis']}")


def doc_row_wa(r):
    return {
        "Quellblatt": TAB_WA,
        "Ausgewählter Code": r["code"],
        "Amtliche Waren-/Güterbeschreibung": _flat(r["beschreibung"]),
        "Einordnung laut Gesetzentwurf": wa_einordnung_text(r),
        "Rechtsbezug / Zusatzkriterium": _flat(r["rechtsbezug"]),
        "GP-Nummern": comma_list(r["gp"]),
        "WA-Codes im Entwurf": "0",
        "WA-Codes nicht im Entwurf": "0",
        "Zuordnungshinweis": r["hinweis"],
        "WA-Quelle (PDF-Seite)": r["quelle"],
    }


def doc_row_gp(r):
    return {
        "Quellblatt": TAB_GP,
        "Ausgewählter Code": r["code"],
        "Amtliche Waren-/Güterbeschreibung": _flat(r["beschreibung"]),
        "Einordnung laut Gesetzentwurf": gp_einordnung_text(r),
        "Rechtsbezug / Zusatzkriterium": "",
        "GP-Nummern": r["code"],
        "WA-Codes im Entwurf": entwurf_cell(r["wa_im_entwurf"]),
        "WA-Codes nicht im Entwurf": entwurf_cell(r["wa_nicht_im_entwurf"]),
        "Zuordnungshinweis": r["hinweis"],
        "WA-Quelle (PDF-Seite)": "",
    }


def group_label(kind, r):
    titles = data.WA_POSITION_TITLES if kind == "wa" else data.GP_KLASSE_TITEL
    title = titles.get(r["pos"], "").rstrip(":").strip()
    prefix = "WA-Position" if kind == "wa" else "GP-Güterklasse"
    return f"{prefix} {r['pos']}" + (f" – {title}" if title else "")


def result_detail_text(kind, r):
    if kind != "wa":
        return _norm(r["beschreibung"])
    title = data.WA_POSITION_TITLES.get(r["pos"], "").rstrip(":").strip()
    lines = [l.strip() for l in _norm(r["beschreibung"]).split("\n") if l.strip()]
    if lines and lines[0].rstrip(":").strip() == title:
        lines = lines[1:]
    return " / ".join(lines) if lines else _norm(r["beschreibung"])


def row_label(kind, r):
    label = f"{r['code']}  —  {short_desc(result_detail_text(kind, r))}"
    if kind == "wa":
        label += f"  ·  {wa_einordnung_text(r)}"
    else:
        label += f"  ·  {r['einordnung']}"
        example = first_code(r["wa_im_entwurf"]) or first_code(r["wa_nicht_im_entwurf"])
        if example:
            label += f"  (Beispiel WA-Code: {example})"
    return label


def render_result_row(kind, r, prefix):
    col_btn, col_mark = st.columns([6, 1])
    with col_btn:
        st.button(
            row_label(kind, r),
            key=f"{prefix}_pick_{kind}_{r['code']}",
            use_container_width=True,
            on_click=select_result,
            args=(kind, r["code"]),
        )
    with col_mark:
        mark_key = f"{prefix}_mark_{kind}_{r['code']}"
        st.checkbox(
            "Dok.",
            value=is_marked(kind, r["code"]),
            key=mark_key,
            on_change=toggle_mark,
            args=(kind, r["code"], mark_key),
            help="Für die Dokumentation markieren",
        )
    if st.session_state.get("selected_kind") == kind and st.session_state.get("selected_code") == r["code"]:
        render_detail_card(kind, r, prefix=prefix)


def render_grouped_rows(kind, rows, prefix, expanded=True):
    groups, order = {}, []
    for r in rows:
        if r["pos"] not in groups:
            groups[r["pos"]] = []
            order.append(r["pos"])
        groups[r["pos"]].append(r)
    for pos in order:
        title = group_label(kind, groups[pos][0])
        with st.expander(f"{title} ({len(groups[pos])})", expanded=expanded):
            for r in groups[pos]:
                render_result_row(kind, r, prefix)


st.markdown('<div class="step">Code oder Schlagwort suchen</div>', unsafe_allow_html=True)
st.caption(
    "Egal ob Sie Ihren Außenhandelscode (KN-/WA-Nummer) oder Ihren Produktionscode "
    "(GP-Nummer) kennen – geben Sie einfach einen Code oder ein Schlagwort ein "
    "(z. B. „Beutel“, „Folie“ oder „3923 40 10“). Die Suche durchsucht beide Verzeichnisse "
    "gleichzeitig."
)
search_term = st.text_input(
    "Code oder Schlagwort",
    placeholder="z. B. 3923 40 10, 2221 30 103 oder „Tüte“",
    label_visibility="collapsed",
    key="global_search",
)

if len(search_term.strip()) >= 2:
    results = search_all(search_term)
    MAX_RESULTS = 40
    shown = results[:MAX_RESULTS]
    st.caption(f"{len(results)} Treffer" + (f" – zeige die ersten {MAX_RESULTS}" if len(results) > MAX_RESULTS else ""))

    # Treffer in ursprünglicher Reihenfolge (WA/GP gemischt), gruppiert nach Position
    groups, order = {}, []
    for kind, r in shown:
        key = (kind, r["pos"])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(r)
    for kind, pos in order:
        title = group_label(kind, groups[(kind, pos)][0])
        with st.expander(f"{title} ({len(groups[(kind, pos)])})", expanded=True):
            for r in groups[(kind, pos)]:
                render_result_row(kind, r, "search")
elif search_term.strip():
    st.caption("Bitte mindestens 2 Zeichen eingeben.")

st.divider()

# ---------------- Auswahl & Inhalt ----------------
sel = st.radio(
    "Bereich wählen",
    [TAB_WA, TAB_GP, TAB_DOC],
    horizontal=True,
    label_visibility="collapsed",
    key="active_area",
)
st.divider()


def render_registry(kind):
    intro = data.DECKBLATT["wa_bekannt"] if kind == "wa" else data.DECKBLATT["gp_bekannt"]
    st.markdown(f'<div class="question">{highlight_terms(intro)}</div>', unsafe_allow_html=True)

    if kind == "wa" and os.path.exists(WA_SOURCE_PDF):
        st.download_button(
            "Quelldokument: Warenverzeichnis für die Außenhandelsstatistik 2026 (PDF)",
            data=_read_bytes(WA_SOURCE_PDF),
            file_name="WA2026-Warenverzeichnis-Aussenhandelsstatistik.pdf",
            mime="application/pdf",
            key="dl_wa_pdf_tab",
        )
    elif kind == "gp":
        if os.path.exists(GP_SOURCE_XLSX):
            st.download_button(
                "Quelldokument: Güterverzeichnis für die Produktionsstatistik GP2026 (Excel)",
                data=_read_bytes(GP_SOURCE_XLSX),
                file_name="GP2026-Gueterverzeichnis-Produktionsstatistik.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_gp_xlsx_tab",
            )
        if os.path.exists(GP_SOURCE_KONKORDANZ):
            st.download_button(
                "Quelldokument: Gegenüberstellung Warennummern zu anderen Warenklassifikationen 2026 (Excel)",
                data=_read_bytes(GP_SOURCE_KONKORDANZ),
                file_name="Gegenueberstellung-Warennummern-Warenklassifikationen-2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_gp_konkordanz_tab",
            )

    st.markdown('<div class="step">Code-Verzeichnis – Zeilen mit Häkchen dokumentieren</div>',
                unsafe_allow_html=True)
    rows = data.WA_ROWS if kind == "wa" else data.GP_ROWS
    render_grouped_rows(kind, rows, "reg", expanded=False)


if sel == TAB_WA:
    render_registry("wa")
elif sel == TAB_GP:
    render_registry("gp")
else:
    st.markdown('<div class="step">Dokumentation – Ihre markierten Zeilen</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<div class="dl-box">Automatische, laufende Zusammenstellung aller mit ✓ markierten '
        'Datenzeilen aus den Code-Verzeichnissen und aus der Suche oben. Originalhinweise werden '
        'übernommen. Die Auswahl ist keine Bestätigung der Steuerpflicht. Beim Entfernen eines '
        'Häkchens verschwindet die betreffende Zeile wieder aus dieser Ansicht. '
        '<strong>Archivhinweis:</strong> Diese Ansicht aktualisiert sich mit der Auswahl. Ein '
        'unveränderlicher Stand entsteht erst durch eine datierte Dateikopie bzw. einen Export '
        '(z. B. Excel-Download unten).</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Amtliche Originaldokumente zur Einsicht"):
        render_quelldokumente("doc")

    marked = {
        (kind, code): (WA_BY_CODE if kind == "wa" else GP_BY_CODE).get(code)
        for kind, code in st.session_state.get("doc_marks", set())
    }
    doc_rows = [
        doc_row_wa(r) if kind == "wa" else doc_row_gp(r)
        for (kind, code), r in marked.items() if r
    ]

    if not doc_rows:
        st.info(f"Noch keine Zeilen markiert. Öffnen Sie die Bereiche *{TAB_WA}* "
                f"oder *{TAB_GP}*, nutzen Sie die Suche oben oder setzen Sie dort "
                "das Häkchen „Dokumentieren“.")
    else:
        doc_df = pd.DataFrame(doc_rows)
        st.markdown('<p class="scroll-hint">↔ Tabelle enthält weitere Spalten – nach rechts scrollen, um alle anzuzeigen.</p>', unsafe_allow_html=True)
        st.markdown(doc_table_html(doc_df), unsafe_allow_html=True)
        st.caption(f"Ausgewählte Zeilen: {len(doc_df)}")
        st.download_button(
            "Dokumentation als Excel-Datei herunterladen",
            data=doc_to_xlsx(doc_df),
            file_name="Kunststoffverpackungssteuer_Dokumentation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.markdown("---")
st.caption(
    "Quelle: „261110_Kunststoffverpackungssteuer_Code-Übersicht und Dokumentation.xlsx“. "
    + data.DECKBLATT["stand"] +
    " Die Angaben stehen unter dem Vorbehalt angekündigter Veröffentlichungen. "
    "Dieses Tool ist keine Rechtsberatung im Einzelfall."
)
