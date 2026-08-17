"""Baut den Anforderungskatalog als Excel-Datei aus der JSON-Beschreibung.

Aufruf:
    python build_xlsx.py [-i data/anforderungskatalog.json] [-o output/Anforderungskatalog.xlsx]
"""
import argparse
import json
from pathlib import Path

from openpyxl import Workbook
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


def mappe_bauen(doc):
    wb = Workbook()
    wb.remove(wb.active)
    for spec in doc["sheets"]:
        blatt_bauen(wb.create_sheet(title=spec["name"]), spec)
    return wb


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("-i", "--input", default="data/anforderungskatalog.json")
    p.add_argument("-o", "--output", default="output/Anforderungskatalog.xlsx")
    args = p.parse_args()

    doc = json.loads(Path(args.input).read_text(encoding="utf-8"))
    ziel = Path(args.output)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    mappe_bauen(doc).save(ziel)

    print(f"geschrieben: {ziel}")
    for s in doc["sheets"]:
        print(f"  {s['name']}: {len(s['rows']) - 1} Datenzeilen x {len(s['rows'][0])} Spalten")


if __name__ == "__main__":
    main()
