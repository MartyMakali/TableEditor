# TableEditor

Werkzeug zum Erzeugen von Excel-Arbeitsmappen aus einer JSON-Beschreibung.

Erster Anwendungsfall: der **Anforderungskatalog** — drei Blätter, einheitlich
formatiert, mit Formelspalten.

## Aufbau

| Datei | Zweck |
|---|---|
| `data/anforderungskatalog.json` | Inhalt und Layout aller Blätter (Werte, Formeln, Spaltenbreiten, Zeilenhöhen, Ausrichtung) |
| `build_xlsx.py` | Baut daraus die `.xlsx` |
| `output/Anforderungskatalog.xlsx` | Erzeugte Arbeitsmappe, vollständig gefüllt |
| `output/Anforderungskatalog_Vorlage.xlsx` | Leere Vorlage zum selbst Ausfüllen, nur Kopfzeilen |

## Benutzung

Einmalig die Abhängigkeit installieren:

```bash
pip install openpyxl
```

Mappe erzeugen:

```bash
python build_xlsx.py
```

Andere Quelle oder anderes Ziel:

```bash
python build_xlsx.py -i data/anforderungskatalog.json -o output/Anforderungskatalog.xlsx
```

Inhalte ändert man in der JSON-Datei, danach neu bauen. Die Formatierung liegt
zentral in `build_xlsx.py` und gilt für alle Blätter.

### Leere Vorlage

Gleiche Blätter, gleiches Format, nur die Kopfzeile ausgefüllt:

```bash
python build_xlsx.py --vorlage -o output/Anforderungskatalog_Vorlage.xlsx
```

Mit vorformatierten Datenzeilen — die Formelspalten sind darin schon zeilenweise
hinterlegt und rechnen mit, sobald man die Kennungen einträgt:

```bash
python build_xlsx.py --vorlage --zeilen 50 -o output/Vorlage_50.xlsx
```

## Die Blätter

| Blatt | Zeilen | Spalten |
|---|---|---|
| Kernanforderungen | 45 | Nr., Kernanforderung, Stichwortgruppe, Zusammengefasste Anforderungen, Stuetzende Arbeiten, Anzahl stuetzender Arbeiten |
| Anforderungen | 144 | Nr., Anforderung, Erlaeuterung, Stichwortgruppe, Stuetzende Arbeiten, Anzahl stuetzender Arbeiten, Einzelbefund, Kernanforderung Nr. |
| Arbeiten | 56 | Kennung, Autorenschaft und Jahr, Titel, Publikationstyp, Untersuchungsraum, Destinationstyp, Raeumliche Ebene, Methodischer Zugang, Anzahl Fundstellen |

## Format

Arial 10 durchgehend · Kopfzeile fett weiß auf `#44546A`, fixiert (`A2`) und mit
Autofilter · dünne Rahmen in `#BFBFBF` rundum · Textspalten linksbündig mit
Zeilenumbruch, oben ausgerichtet · Nummernspalten zentriert.

## Formeln

Zwei berechnete Spalten, in der JSON-Datei als Formelstrings hinterlegt und beim
Öffnen von Excel ausgewertet:

- **Anzahl stuetzender Arbeiten** — zählt die kommagetrennten Kennungen:
  `=IF(E2="","",LEN(E2)-LEN(SUBSTITUTE(E2,",",""))+1)`
- **Einzelbefund** (nur Blatt *Anforderungen*) — `ja`, wenn genau eine Arbeit stützt:
  `=IF(F2="","",IF(F2=1,"ja","nein"))`
