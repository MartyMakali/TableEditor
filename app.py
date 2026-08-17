"""Lokaler Server fuer den TableEditor.

Start:
    python app.py

Danach im Browser: http://127.0.0.1:8000
"""
import io
import json
import re
import webbrowser
from pathlib import Path

from fastapi import Body, FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles

from build_xlsx import mappe_bauen
from einlesen import NichtLesbar, tabellen_aus_datei

WURZEL = Path(__file__).parent
DATEN = WURZEL / "data"
WEB = WURZEL / "web"
DATEN.mkdir(exist_ok=True)

app = FastAPI(title="TableEditor")


def pfad_zu(name: str) -> Path:
    """Erlaubt nur einfache Dateinamen und haelt alles im Datenordner."""
    if not re.fullmatch(r"[A-Za-z0-9_.-]{1,80}", name) or name in (".", ".."):
        raise HTTPException(400, "Ungueltiger Dateiname")
    return DATEN / f"{name}.json"


@app.get("/api/dateien")
def dateien():
    return sorted(p.stem for p in DATEN.glob("*.json"))


@app.get("/api/dokument/{name}")
def dokument_lesen(name: str):
    pfad = pfad_zu(name)
    if not pfad.exists():
        raise HTTPException(404, f"'{name}' nicht gefunden")
    return json.loads(pfad.read_text(encoding="utf-8"))


@app.put("/api/dokument/{name}")
def dokument_schreiben(name: str, doc: dict = Body(...)):
    pfad = pfad_zu(name)
    pfad.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    return {"gespeichert": pfad.name}


@app.post("/api/einlesen")
async def datei_einlesen(datei: UploadFile = File(...)):
    """Nimmt eine Word-, Excel- oder CSV-Datei entgegen und gibt ihre Tabellen zurueck."""
    daten = await datei.read()
    if not daten:
        raise HTTPException(400, "Die Datei ist leer")
    try:
        return {"tabellen": tabellen_aus_datei(datei.filename or "", daten)}
    except NichtLesbar as fehler:
        raise HTTPException(400, str(fehler))
    except Exception as fehler:  # beschaedigte oder unerwartet aufgebaute Datei
        raise HTTPException(400, f"Datei konnte nicht gelesen werden: {fehler}")


@app.post("/api/export")
def export(nutzlast: dict = Body(...)):
    """Baut die Mappe im Speicher und liefert sie als Download zurueck."""
    doc = nutzlast.get("dokument")
    if not doc or not doc.get("blaetter"):
        raise HTTPException(400, "Kein Blatt zum Exportieren")

    puffer = io.BytesIO()
    mappe_bauen(
        doc,
        vorlage=bool(nutzlast.get("vorlage")),
        zeilen=int(nutzlast.get("zeilen") or 0),
    ).save(puffer)

    titel = re.sub(r"[^A-Za-z0-9_-]", "_", doc.get("titel") or "Tabelle")
    if nutzlast.get("vorlage"):
        titel += "_Vorlage"
    return Response(
        content=puffer.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{titel}.xlsx"'},
    )


@app.get("/")
def start():
    return FileResponse(WEB / "index.html")


app.mount("/web", StaticFiles(directory=WEB), name="web")


if __name__ == "__main__":
    import uvicorn

    adresse = "http://127.0.0.1:8000"
    print(f"TableEditor laeuft auf {adresse}  (Beenden mit Strg+C)")
    webbrowser.open(adresse)
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")
