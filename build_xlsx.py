"""Baut den Anforderungskatalog als Excel-Datei aus der JSON-Beschreibung.

Aufruf:
    python build_xlsx.py [-i data/anforderungskatalog.json] [-o output/Anforderungskatalog.xlsx]
"""
import argparse
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- Formatvorgaben, entnommen aus der Originaldatei -------------------------
# Farben als ARGB (fuehrendes FF = deckend), sonst weicht der Alphakanal ab.
SCHRIFT = "Arial"
GROESSE = 10
KOPF_FARBE = "FF44546A"     # dunkles Blaugrau der Kopfzeile
KOPF_SCHRIFT = "FFFFFFFF"   # weisse Kopfschrift
RAHMEN_FARBE = "FFBFBFBF"   # helles Grau, duenn, rundum

FONT_DATEN = Font(name=SCHRIFT, size=GROESSE)
FONT_KOPF = Font(name=SCHRIFT, size=GROESSE, bold=True, color=KOPF_SCHRIFT)
FILL_KOPF = PatternFill(fill_type="solid", start_color=KOPF_FARBE, end_color=KOPF_FARBE)
_seite = Side(style="thin", color=RAHMEN_FARBE)
RAHMEN = Border(left=_seite, right=_seite, top=_seite, bottom=_seite)
ALIGN_KOPF = Alignment(horizontal="general", vertical="center", wrap_text=True)


def blatt_bauen(ws, spec):
    """Traegt Werte und Format eines Blattes ein."""
    zeilen = spec["rows"]
    stile = spec["column_styles"]

    for r, zeile in enumerate(zeilen, start=1):
        for c, wert in enumerate(zeile, start=1):
            zelle = ws.cell(row=r, column=c, value=wert)
            zelle.border = RAHMEN
            if r == 1:
                zelle.font = FONT_KOPF
                zelle.fill = FILL_KOPF
                zelle.alignment = ALIGN_KOPF
            else:
                stil = stile.get(get_column_letter(c), {"horizontal": "left", "wrap": False})
                zelle.font = FONT_DATEN
                zelle.alignment = Alignment(
                    horizontal=stil["horizontal"], vertical="top", wrap_text=stil["wrap"]
                )

    for spalte, breite in spec["column_widths"].items():
        ws.column_dimensions[spalte].width = breite
    for zeilennr, hoehe in spec["row_heights"].items():
        ws.row_dimensions[int(zeilennr)].height = hoehe

    if spec.get("freeze_panes"):
        ws.freeze_panes = spec["freeze_panes"]
    if spec.get("autofilter"):
        letzte = get_column_letter(len(zeilen[0]))
        ws.auto_filter.ref = f"A1:{letzte}{len(zeilen)}"


def als_vorlage(spec, zeilen=0):
    """Leert ein Blatt bis auf die Kopfzeile.

    zeilen > 0 haengt entsprechend viele leere, bereits formatierte Datenzeilen an;
    Formelspalten werden darin auf die jeweilige Zeile umgeschrieben.
    """
    kopf = spec["rows"][0]
    muster = spec["rows"][1] if len(spec["rows"]) > 1 else []
    neu = [kopf]

    for i in range(zeilen):
        zeilennr = i + 2
        neu.append(
            [
                Translator(wert, origin=f"{get_column_letter(c)}2").translate_formula(
                    f"{get_column_letter(c)}{zeilennr}"
                )
                if isinstance(wert, str) and wert.startswith("=")
                else None
                for c, wert in enumerate(muster, start=1)
            ]
        )

    kopfhoehe = spec["row_heights"].get("1") or spec["row_heights"].get(1)
    return {
        **spec,
        "rows": neu,
        "row_heights": {"1": kopfhoehe} if kopfhoehe else {},
    }


def mappe_bauen(doc, vorlage=False, zeilen=0):
    wb = Workbook()
    wb.remove(wb.active)
    for spec in doc["sheets"]:
        if vorlage:
            spec = als_vorlage(spec, zeilen)
        blatt_bauen(wb.create_sheet(title=spec["name"]), spec)
    return wb


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("-i", "--input", default="data/anforderungskatalog.json")
    p.add_argument("-o", "--output", default="output/Anforderungskatalog.xlsx")
    p.add_argument(
        "--vorlage", action="store_true", help="leere Mappe, nur die Kopfzeilen"
    )
    p.add_argument(
        "--zeilen",
        type=int,
        default=0,
        metavar="N",
        help="mit --vorlage: N leere, vorformatierte Datenzeilen samt Formeln",
    )
    args = p.parse_args()

    doc = json.loads(Path(args.input).read_text(encoding="utf-8"))
    ziel = Path(args.output)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    mappe_bauen(doc, vorlage=args.vorlage, zeilen=args.zeilen).save(ziel)

    print(f"geschrieben: {ziel}")
    for s in doc["sheets"]:
        anzahl = args.zeilen if args.vorlage else len(s["rows"]) - 1
        print(f"  {s['name']}: {anzahl} Datenzeilen x {len(s['rows'][0])} Spalten")


if __name__ == "__main__":
    main()
