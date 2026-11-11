"""Erzeugt codes_data.py aus der Excel-Quelle.

Aufruf:  python build_data.py
Die Excel-Datei wird relativ zum Code-Ordner erwartet unter ..\\Input.
"""

import os
import re

import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(
    BASE, "..", "Input", "261110_Kunststoffverpackungssteuer_Code-Übersicht und Dokumentation.xlsx"
)
OUT = os.path.join(BASE, "codes_data.py")

# Wenn der Zuordnungshinweis darauf hinweist, dass die Codezuordnung ein
# zusätzliches gesetzliches Merkmal nicht abbildet, ist der Zusatz "(außerhalb
# GP-Fokus)" bei den GP-Nummern irreführend und wird entfernt.
MERKMAL_HINWEIS = "Die Codezuordnung bildet das zusätzliche gesetzliche Merkmal nicht ab"

# Satz, der in der Deckblatt-Zelle "Stand und Aussagegrenzen" entfernt werden soll.
STAND_SATZ_ENTFERNEN = "Die Ergänzung der Vergleichscodes erweitert nicht die gesetzliche Auswahl."


def clean(v):
    if v is None:
        return ""
    s = str(v).replace("\xa0", " ").strip()
    return s


def pystr(s):
    return repr(s)


def main():
    wb = openpyxl.load_workbook(XLSX, data_only=True)

    # ---------- _Auswahltechnik: vereinheitlichte Datengrundlage ----------
    ws = wb["_Auswahltechnik"]
    hdr = [clean(c.value) for c in ws[1]]
    idx = {h: i for i, h in enumerate(hdr)}

    wa_rows, gp_rows = [], []
    for row in ws.iter_rows(min_row=2):
        vals = [clean(c.value) for c in row]
        if not vals[idx["Code"]]:
            continue
        rec = {
            "code": vals[idx["Code"]],
            "pos": vals[idx["Code"]][:4],
            "beschreibung": vals[idx["Amtliche Bezeichnung"]],
            "einordnung": vals[idx["Einordnung"]],
            "rechtsbezug": vals[idx["Rechtsbezug / Zusatzkriterium"]],
            "gp": vals[idx["GP-Nummern"]],
            "wa_im_entwurf": vals[idx["WA im Entwurf"]],
            "wa_nicht_im_entwurf": vals[idx["WA nicht im Entwurf"]],
            "grund": vals[idx["Aufnahmegrund"]],
            "hinweis": vals[idx["Zuordnungshinweis"]],
            "quelle": vals[idx["Quellseite"]],
        }
        if MERKMAL_HINWEIS in rec["hinweis"]:
            rec["gp"] = re.sub(r"\s*\(außerhalb GP-Fokus\)", "", rec["gp"])
        if vals[idx["Quellblatt"]] == "Außenhandelscode bekannt":
            wa_rows.append(rec)
        else:
            gp_rows.append(rec)

    # Abschnitte (Positionen / Güterklassen) in Originalreihenfolge
    wa_pos, seen = [], set()
    for r in wa_rows:
        if r["pos"] not in seen:
            seen.add(r["pos"])
            wa_pos.append(r["pos"])
    gp_cls, seen = [], set()
    for r in gp_rows:
        if r["pos"] not in seen:
            seen.add(r["pos"])
            gp_cls.append(r["pos"])

    # Positionstitel aus den Abschnittsüberschriften der Quellblätter
    wa_titles = {}
    wswa = wb["Außenhandelscode bekannt"]
    for row in wswa.iter_rows(min_row=7):
        a = clean(row[0].value)
        m = re.match(r"WA-Position (\d{4}) \| (.*)", a)
        if m:
            wa_titles[m.group(1)] = m.group(2)
    gp_titles = {}
    wsgp = wb["Produktionscode bekannt"]
    for row in list(wsgp.iter_rows(min_row=7, max_row=8)) + list(wsgp.iter_rows(min_row=9)):
        a = clean(row[0].value)
        m = re.match(r"GP-Güterklasse (\d{4}) \| (.*)", a)
        if m:
            gp_titles[m.group(1)] = m.group(2)

    # ---------- Originalseiten (Warenverzeichnis 2026) ----------
    pages = {}
    wsop = wb["Originalseiten"]
    for row in wsop.iter_rows(min_row=2):
        if row[0].value is not None:
            pages[int(row[0].value)] = clean(row[1].value)

    # ---------- Deckblatt-Texte ----------
    wsd = wb["Deckblatt"]
    deck = {
        "hintergrund": clean(wsd["B8"].value),
        "gemeinsam": clean(wsd["B11"].value),
        "wa_bekannt": clean(wsd["B14"].value),
        "gp_bekannt": clean(wsd["B18"].value),
        "dokumentation": clean(wsd["B22"].value),
        "lesart": clean(wsd["B28"].value),
        "stand": clean(wsd["B31"].value).replace(STAND_SATZ_ENTFERNEN, "").strip(),
    }

    # ---------- Quellen und Hinweise ----------
    quellen = []
    wsq = wb["Quellen und Hinweise"]
    for row in wsq.iter_rows(min_row=9, max_row=21):
        t, e = clean(row[0].value), clean(row[1].value)
        if t:
            quellen.append((t, e))

    # ---------- Schreiben ----------
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(
            '"""Automatisch erzeugt durch build_data.py – nicht manuell bearbeiten.\n'
            'Quelle: 261110_Kunststoffverpackungssteuer_Code-Übersicht und Dokumentation.xlsx\n'
            'Arbeitsfassung 10.11.2026; keine geltende Gesetzesfassung.\n"""\n\n'
        )

        def dump_list(name, rows, fields):
            f.write(f"{name} = [\n")
            for r in rows:
                f.write("    {\n")
                for k in fields:
                    f.write(f"        {k!r}: {pystr(r[k])},\n")
                f.write("    },\n")
            f.write("]\n\n")

        dump_list("WA_ROWS", wa_rows,
                  ["code", "pos", "beschreibung", "einordnung", "rechtsbezug", "gp",
                   "wa_im_entwurf", "wa_nicht_im_entwurf", "grund", "hinweis", "quelle"])
        dump_list("GP_ROWS", gp_rows,
                  ["code", "pos", "beschreibung", "einordnung", "rechtsbezug", "gp",
                   "wa_im_entwurf", "wa_nicht_im_entwurf", "grund", "hinweis", "quelle"])

        f.write("WA_POSITION_TITLES = {\n")
        for p in wa_pos:
            f.write(f"    {p!r}: {pystr(wa_titles.get(p, ''))},\n")
        f.write("}\n\n")

        f.write("GP_KLASSE_TITEL = {\n")
        for p in gp_cls:
            f.write(f"    {p!r}: {pystr(gp_titles.get(p, ''))},\n")
        f.write("}\n\n")

        f.write("ORIGINAL_PAGES = {\n")
        for p in sorted(pages):
            f.write(f"    {p}: {pystr(pages[p])},\n")
        f.write("}\n\n")

        f.write("DECKBLATT = {\n")
        for k, v in deck.items():
            f.write(f"    {k!r}: {pystr(v)},\n")
        f.write("}\n\n")

        f.write("QUELLEN = [\n")
        for t, e in quellen:
            f.write(f"    ({pystr(t)}, {pystr(e)}),\n")
        f.write("]\n")

    print(f"WA-Zeilen: {len(wa_rows)}, GP-Zeilen: {len(gp_rows)}, "
          f"Seiten: {len(pages)}, Quellen: {len(quellen)}")
    print(f"-> {OUT}")


if __name__ == "__main__":
    main()
