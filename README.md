# TableEditor

Werkzeug zum Entwerfen von Excel-Arbeitsmappen: eine Oberfläche im Browser zum
Aufbauen, Gestalten und Ausfüllen — Ausgabe ist eine `.xlsx`.

Erster Anwendungsfall: der **Anforderungskatalog** — drei Blätter, einheitlich
formatiert, mit Formelspalten.

## Loslegen

Einmalig die Abhängigkeiten installieren:

```bash
pip install -r requirements.txt
```

Oberfläche starten (öffnet den Browser auf `http://127.0.0.1:8000`):

```bash
python app.py
```

Beim Start wird `data/anforderungskatalog.json` geladen, sofern vorhanden.

## Was die Oberfläche kann

**Aufbau** — Blätter anlegen, umbenennen, löschen und per Ziehen umsortieren.
Spalten hinzufügen, umbenennen, verschieben und entfernen; ein Klick auf eine
Überschrift öffnet die Einstellungen der Spalte. Zeilen anhängen und löschen.

**Aussehen** — Schriftart und -größe, Farbe und Schriftfarbe der Kopfzeile,
Fettdruck, Rahmenstil und Rahmenfarbe, Spaltenbreiten, Ausrichtung,
Zeilenumbruch, Kopfzeilenhöhe, Fixierung und Autofilter. Die Tabelle in der
Mitte zeigt das Ergebnis direkt an, in denselben Maßen wie später in Excel.

**Inhalte** — Zellen werden an Ort und Stelle bearbeitet. Reine Zahlen werden
als Zahl gespeichert, damit Excel damit rechnen kann. Beginnt ein Eintrag mit
`=`, gilt er als Formel und wird hervorgehoben.

**Speichern und Ausgeben** — Mappen liegen als JSON unter `data/` und lassen
sich jederzeit wieder laden. *Excel* gibt die gefüllte Mappe aus,
*Excel-Vorlage* eine leere mit denselben Kopfzeilen und wahlweise einer Anzahl
vorbereiteter Zeilen.

## Aufbau des Projekts

| Datei | Zweck |
|---|---|
| `app.py` | Lokaler Server, liefert die Oberfläche und erzeugt die Excel-Datei |
| `web/` | Oberfläche: `index.html`, `style.css`, `app.js` |
| `build_xlsx.py` | Baut aus der JSON-Beschreibung die `.xlsx` — auch ohne Oberfläche nutzbar |
| `data/*.json` | Gespeicherte Mappen |
| `output/` | Erzeugte Arbeitsmappen |

## Ohne Oberfläche

```bash
python build_xlsx.py                                   # volle Mappe
python build_xlsx.py --vorlage                         # nur die Kopfzeilen
python build_xlsx.py --vorlage --zeilen 50             # mit 50 leeren Zeilen
python build_xlsx.py -i data/meine.json -o output/Meine.xlsx
```

## Aufbau der JSON-Beschreibung

```jsonc
{
  "titel": "Anforderungskatalog",
  "format": {
    "schrift": "Arial",
    "groesse": 10,
    "kopf_farbe": "FF44546A",      // ARGB oder #rrggbb
    "kopf_schrift": "FFFFFFFF",
    "kopf_fett": true,
    "rahmen_farbe": "FFBFBFBF",
    "rahmen_stil": "thin"          // thin | medium | thick | dotted | dashed | double | keiner
  },
  "blaetter": [
    {
      "name": "Kernanforderungen",
      "kopf_fixieren": true,
      "autofilter": true,
      "kopfhoehe": 30.0,
      "spalten": [
        { "name": "Nr.", "breite": 5.0, "ausrichtung": "center", "umbruch": false }
      ],
      "zeilenhoehen": { "2": 35.05 },   // Schlüssel = Zeilennummer in Excel
      "zeilen": [[1, "…"]]              // nur Datenzeilen, Werte oder Formeln
    }
  ]
}
```

## Die Blätter des Anforderungskatalogs

| Blatt | Zeilen | Spalten |
|---|---|---|
| Kernanforderungen | 45 | Nr., Kernanforderung, Stichwortgruppe, Zusammengefasste Anforderungen, Stuetzende Arbeiten, Anzahl stuetzender Arbeiten |
| Anforderungen | 144 | Nr., Anforderung, Erlaeuterung, Stichwortgruppe, Stuetzende Arbeiten, Anzahl stuetzender Arbeiten, Einzelbefund, Kernanforderung Nr. |
| Arbeiten | 56 | Kennung, Autorenschaft und Jahr, Titel, Publikationstyp, Untersuchungsraum, Destinationstyp, Raeumliche Ebene, Methodischer Zugang, Anzahl Fundstellen |

## Formeln

Zwei berechnete Spalten, als Formelstrings hinterlegt und beim Öffnen von Excel
ausgewertet:

- **Anzahl stuetzender Arbeiten** — zählt die kommagetrennten Kennungen:
  `=IF(E2="","",LEN(E2)-LEN(SUBSTITUTE(E2,",",""))+1)`
- **Einzelbefund** (nur Blatt *Anforderungen*) — `ja`, wenn genau eine Arbeit stützt:
  `=IF(F2="","",IF(F2=1,"ja","nein"))`

Beim Anhängen einer Zeile werden Formeln aus der letzten Zeile übernommen und
um eine Zeile weitergeschrieben.
