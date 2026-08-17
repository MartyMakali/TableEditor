/* TableEditor — Oberflaeche zum Entwerfen der Arbeitsmappe. */

const STANDARDFORMAT = {
  schrift: "Arial",
  groesse: 10,
  kopf_farbe: "FF44546A",
  kopf_schrift: "FFFFFFFF",
  kopf_fett: true,
  rahmen_farbe: "FFBFBFBF",
  rahmen_stil: "thin",
};

let dok = leeresDokument();
let aktiv = 0;          // Index des angezeigten Blattes
let spaltenindex = null; // im Spaltendialog bearbeitete Spalte

const $ = (id) => document.getElementById(id);

/* ---- Hilfsfunktionen ------------------------------------------------ */

function leeresDokument() {
  return {
    titel: "Neue Mappe",
    format: { ...STANDARDFORMAT },
    blaetter: [neuesBlatt("Tabelle1")],
  };
}

function neuesBlatt(name) {
  return {
    name,
    kopf_fixieren: true,
    autofilter: true,
    kopfhoehe: 30,
    spalten: [
      { name: "Spalte A", breite: 20, ausrichtung: "left", umbruch: true },
      { name: "Spalte B", breite: 20, ausrichtung: "left", umbruch: true },
    ],
    zeilenhoehen: {},
    zeilen: [],
  };
}

const blatt = () => dok.blaetter[aktiv];

// ARGB (Excel) <-> #rrggbb (Farbwaehler im Browser)
const zuHex = (argb) => "#" + String(argb || "FF000000").slice(-6);
const zuArgb = (hex) => "FF" + hex.replace("#", "").toUpperCase();

// Excel-Spaltenbreite und Punkt-Zeilenhoehe in Bildschirmpixel
const breiteInPx = (b) => Math.round((b || 10) * 7 + 5);
const hoeheInPx = (h) => Math.round((h || 15) * 96 / 72);

function meldung(text) {
  $("status").textContent = text;
  clearTimeout(meldung.uhr);
  meldung.uhr = setTimeout(() => ($("status").textContent = ""), 3500);
}

/** Schreibt Zellbezuege einer Formel um eine Zeilendifferenz weiter. */
function formelVersetzen(formel, versatz) {
  return formel.replace(
    /(\$?)([A-Z]{1,3})(\$?)(\d+)/g,
    (treffer, dollarSpalte, spalte, dollarZeile, zeile) =>
      dollarSpalte + spalte + dollarZeile +
      (dollarZeile ? zeile : String(Number(zeile) + versatz))
  );
}

/** Zahlen als Zahl speichern, damit Excel damit rechnen kann. */
function wertAusText(text) {
  const t = text.trim();
  if (t === "") return null;
  if (/^-?\d+$/.test(t)) return Number(t);
  if (/^-?\d+[.,]\d+$/.test(t)) return Number(t.replace(",", "."));
  return t;
}

/* ---- Darstellung ---------------------------------------------------- */

function zeichnen() {
  if (aktiv >= dok.blaetter.length) aktiv = Math.max(0, dok.blaetter.length - 1);
  $("titel").value = dok.titel || "";
  formularFuellen();
  blattlisteZeichnen();
  tabelleZeichnen();
}

function formularFuellen() {
  const f = dok.format;
  $("f-schrift").value = f.schrift;
  $("f-groesse").value = f.groesse;
  $("f-kopf-farbe").value = zuHex(f.kopf_farbe);
  $("f-kopf-schrift").value = zuHex(f.kopf_schrift);
  $("f-kopf-fett").checked = !!f.kopf_fett;
  $("f-rahmen-stil").value = f.rahmen_stil || "thin";
  $("f-rahmen-farbe").value = zuHex(f.rahmen_farbe);

  const b = blatt();
  const hat = !!b;
  ["b-name", "b-kopfhoehe", "b-fixieren", "b-autofilter", "btn-blatt-loeschen"]
    .forEach((id) => ($(id).disabled = !hat));
  if (!hat) return;
  $("b-name").value = b.name;
  $("b-kopfhoehe").value = b.kopfhoehe || 30;
  $("b-fixieren").checked = !!b.kopf_fixieren;
  $("b-autofilter").checked = !!b.autofilter;
}

function blattlisteZeichnen() {
  const liste = $("blattliste");
  liste.innerHTML = "";
  dok.blaetter.forEach((b, i) => {
    const li = document.createElement("li");
    li.className = i === aktiv ? "aktiv" : "";
    li.draggable = true;
    li.innerHTML = `<span class="griff">⠿</span><span class="benennung"></span>`;
    li.querySelector(".benennung").textContent = b.name;
    li.onclick = () => { aktiv = i; zeichnen(); };

    li.ondragstart = (e) => e.dataTransfer.setData("text/plain", String(i));
    li.ondragover = (e) => { e.preventDefault(); li.classList.add("zielt"); };
    li.ondragleave = () => li.classList.remove("zielt");
    li.ondrop = (e) => {
      e.preventDefault();
      li.classList.remove("zielt");
      const von = Number(e.dataTransfer.getData("text/plain"));
      if (von === i) return;
      const [bewegt] = dok.blaetter.splice(von, 1);
      dok.blaetter.splice(i, 0, bewegt);
      aktiv = i;
      zeichnen();
    };
    liste.appendChild(li);
  });
}

function tabelleZeichnen() {
  const tabelle = $("tabelle");
  tabelle.innerHTML = "";
  const b = blatt();
  if (!b) {
    $("masse").textContent = "";
    $("tabellenrahmen").innerHTML = '<p class="leerhinweis">Kein Blatt vorhanden.</p>';
    return;
  }
  if (!$("tabelle")) $("tabellenrahmen").innerHTML = '<table id="tabelle"></table>';

  const f = dok.format;
  const rahmen = f.rahmen_stil === "keiner"
    ? "none"
    : `${f.rahmen_stil === "thick" ? 2 : 1}px ${
        { dotted: "dotted", dashed: "dashed", double: "double" }[f.rahmen_stil] || "solid"
      } ${zuHex(f.rahmen_farbe)}`;

  tabelle.style.font = `${f.groesse}pt ${f.schrift}`;

  // Kopfzeile
  const kopf = tabelle.createTHead().insertRow();
  const ecke = document.createElement("th");
  ecke.className = "ecke";
  kopf.appendChild(ecke);

  b.spalten.forEach((s, i) => {
    const th = document.createElement("th");
    th.style.cssText = `width:${breiteInPx(s.breite)}px;background:${zuHex(f.kopf_farbe)};` +
      `color:${zuHex(f.kopf_schrift)};font-weight:${f.kopf_fett ? 700 : 400};border:${rahmen};` +
      `height:${hoeheInPx(b.kopfhoehe)}px;white-space:normal`;
    th.innerHTML = `<span class="zahnrad">⚙</span>`;
    th.prepend(document.createTextNode(s.name));
    th.title = "Anklicken zum Einstellen der Spalte";
    th.onclick = () => spaltendialogOeffnen(i);
    kopf.appendChild(th);
  });

  // Datenzeilen
  const koerper = tabelle.createTBody();
  b.zeilen.forEach((zeile, r) => {
    const tr = koerper.insertRow();
    const nr = tr.insertCell();
    nr.className = "zeilennr";
    nr.innerHTML = `<span>${r + 2}</span> <button title="Zeile löschen">✕</button>`;
    nr.querySelector("button").onclick = () => {
      b.zeilen.splice(r, 1);
      tabelleZeichnen();
    };

    b.spalten.forEach((s, c) => {
      const wert = zeile[c];
      const td = tr.insertCell();
      const istFormel = typeof wert === "string" && wert.startsWith("=");
      td.className = istFormel ? "formel" : "";
      td.contentEditable = "true";
      td.textContent = wert === null || wert === undefined ? "" : String(wert);
      td.style.cssText = `text-align:${s.ausrichtung};border:${rahmen};` +
        `white-space:${s.umbruch ? "pre-wrap" : "nowrap"}`;
      td.onblur = () => {
        zeile[c] = wertAusText(td.textContent);
        td.className = typeof zeile[c] === "string" && zeile[c].startsWith("=") ? "formel" : "";
      };
      td.onkeydown = (e) => {
        if (e.key === "Escape") td.blur();
      };
    });
  });

  $("masse").textContent =
    `${b.zeilen.length} Zeilen · ${b.spalten.length} Spalten`;
}

/* ---- Spaltendialog --------------------------------------------------- */

function spaltendialogOeffnen(i) {
  spaltenindex = i;
  const s = blatt().spalten[i];
  $("s-name").value = s.name || "";
  $("s-breite").value = s.breite || 20;
  $("s-ausrichtung").value = s.ausrichtung || "left";
  $("s-umbruch").checked = !!s.umbruch;
  $("spaltendialog").classList.remove("verborgen");
  $("s-name").focus();
}

function spaltendialogSchliessen() {
  $("spaltendialog").classList.add("verborgen");
  spaltenindex = null;
}

function spalteUebernehmen() {
  if (spaltenindex === null) return;
  const s = blatt().spalten[spaltenindex];
  s.name = $("s-name").value;
  s.breite = Number($("s-breite").value) || 20;
  s.ausrichtung = $("s-ausrichtung").value;
  s.umbruch = $("s-umbruch").checked;
}

function spalteVerschieben(richtung) {
  spalteUebernehmen();
  const b = blatt();
  const ziel = spaltenindex + richtung;
  if (ziel < 0 || ziel >= b.spalten.length) return;
  [b.spalten[spaltenindex], b.spalten[ziel]] = [b.spalten[ziel], b.spalten[spaltenindex]];
  b.zeilen.forEach((z) => {
    [z[spaltenindex], z[ziel]] = [z[ziel], z[spaltenindex]];
  });
  spaltenindex = ziel;
  tabelleZeichnen();
}

/* ---- Server ---------------------------------------------------------- */

async function dateilisteHolen() {
  const namen = await (await fetch("/api/dateien")).json();
  const wahl = $("dateiwahl");
  wahl.innerHTML = namen.map((n) => `<option>${n}</option>`).join("");
}

async function laden() {
  const name = $("dateiwahl").value;
  if (!name) return meldung("Keine gespeicherte Mappe vorhanden");
  const antwort = await fetch(`/api/dokument/${name}`);
  if (!antwort.ok) return meldung("Laden fehlgeschlagen");
  dok = await antwort.json();
  aktiv = 0;
  zeichnen();
  meldung(`„${name}“ geladen`);
}

async function speichern() {
  const vorschlag = (dok.titel || "mappe").toLowerCase().replace(/[^a-z0-9_-]+/g, "_");
  const name = prompt("Unter welchem Namen speichern?", vorschlag);
  if (!name) return;
  const antwort = await fetch(`/api/dokument/${name}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(dok),
  });
  if (!antwort.ok) return meldung("Speichern fehlgeschlagen");
  await dateilisteHolen();
  $("dateiwahl").value = name;
  meldung(`als „${name}“ gespeichert`);
}

async function exportieren(alsVorlage) {
  let zeilen = 0;
  if (alsVorlage) {
    const eingabe = prompt("Wie viele leere Zeilen vorbereiten? (0 = nur Kopfzeile)", "0");
    if (eingabe === null) return;
    zeilen = Number(eingabe) || 0;
  }
  const antwort = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ dokument: dok, vorlage: alsVorlage, zeilen }),
  });
  if (!antwort.ok) return meldung("Export fehlgeschlagen");

  const blob = await antwort.blob();
  const verweis = document.createElement("a");
  verweis.href = URL.createObjectURL(blob);
  verweis.download =
    (dok.titel || "Tabelle").replace(/[^A-Za-z0-9_-]/g, "_") +
    (alsVorlage ? "_Vorlage" : "") + ".xlsx";
  verweis.click();
  URL.revokeObjectURL(verweis.href);
  meldung("Excel-Datei erzeugt");
}

/* ---- Verdrahtung ----------------------------------------------------- */

function verdrahten() {
  $("titel").oninput = () => (dok.titel = $("titel").value);

  // Aussehen
  const fmt = {
    "f-schrift": (e) => (dok.format.schrift = e.value),
    "f-groesse": (e) => (dok.format.groesse = Number(e.value) || 10),
    "f-kopf-farbe": (e) => (dok.format.kopf_farbe = zuArgb(e.value)),
    "f-kopf-schrift": (e) => (dok.format.kopf_schrift = zuArgb(e.value)),
    "f-kopf-fett": (e) => (dok.format.kopf_fett = e.checked),
    "f-rahmen-stil": (e) => (dok.format.rahmen_stil = e.value),
    "f-rahmen-farbe": (e) => (dok.format.rahmen_farbe = zuArgb(e.value)),
  };
  Object.entries(fmt).forEach(([id, setzen]) => {
    $(id).oninput = () => { setzen($(id)); tabelleZeichnen(); };
  });

  // Blatt
  $("b-name").oninput = () => { blatt().name = $("b-name").value; blattlisteZeichnen(); };
  $("b-kopfhoehe").oninput = () => {
    blatt().kopfhoehe = Number($("b-kopfhoehe").value) || 30;
    tabelleZeichnen();
  };
  $("b-fixieren").onchange = () => (blatt().kopf_fixieren = $("b-fixieren").checked);
  $("b-autofilter").onchange = () => (blatt().autofilter = $("b-autofilter").checked);

  $("btn-blatt-neu").onclick = () => {
    const namen = dok.blaetter.map((b) => b.name);
    let n = dok.blaetter.length + 1;
    while (namen.includes(`Tabelle${n}`)) n++;
    dok.blaetter.push(neuesBlatt(`Tabelle${n}`));
    aktiv = dok.blaetter.length - 1;
    zeichnen();
  };

  $("btn-blatt-loeschen").onclick = () => {
    const b = blatt();
    if (!b || !confirm(`Blatt „${b.name}“ mit ${b.zeilen.length} Zeilen löschen?`)) return;
    dok.blaetter.splice(aktiv, 1);
    zeichnen();
  };

  // Spalten und Zeilen
  $("btn-spalte-neu").onclick = () => {
    const b = blatt();
    b.spalten.push({ name: `Spalte ${b.spalten.length + 1}`, breite: 20, ausrichtung: "left", umbruch: true });
    b.zeilen.forEach((z) => z.push(null));
    tabelleZeichnen();
  };

  $("btn-zeile-neu").onclick = () => {
    const b = blatt();
    const letzte = b.zeilen[b.zeilen.length - 1];
    // Formeln aus der letzten Zeile uebernehmen, um eine Zeile weitergeschrieben
    const neue = b.spalten.map((_, c) => {
      const wert = letzte ? letzte[c] : null;
      return typeof wert === "string" && wert.startsWith("=")
        ? formelVersetzen(wert, 1)
        : null;
    });
    b.zeilen.push(neue);
    tabelleZeichnen();
    const rahmen = $("tabellenrahmen");
    rahmen.scrollTop = rahmen.scrollHeight;
  };

  // Spaltendialog
  $("s-fertig").onclick = () => { spalteUebernehmen(); spaltendialogSchliessen(); zeichnen(); };
  $("s-links").onclick = () => spalteVerschieben(-1);
  $("s-rechts").onclick = () => spalteVerschieben(1);
  $("s-loeschen").onclick = () => {
    const b = blatt();
    if (b.spalten.length === 1) return alert("Das letzte Feld lässt sich nicht löschen.");
    if (!confirm(`Spalte „${b.spalten[spaltenindex].name}“ mit allen Werten löschen?`)) return;
    b.spalten.splice(spaltenindex, 1);
    b.zeilen.forEach((z) => z.splice(spaltenindex, 1));
    spaltendialogSchliessen();
    zeichnen();
  };
  $("spaltendialog").onclick = (e) => {
    if (e.target === $("spaltendialog")) { spalteUebernehmen(); spaltendialogSchliessen(); zeichnen(); }
  };

  // Dateien
  $("btn-laden").onclick = laden;
  $("btn-speichern").onclick = speichern;
  $("btn-neu").onclick = () => {
    if (!confirm("Neue, leere Mappe anlegen? Nicht Gespeichertes geht verloren.")) return;
    dok = leeresDokument();
    aktiv = 0;
    zeichnen();
  };
  $("btn-export").onclick = () => exportieren(false);
  $("btn-export-vorlage").onclick = () => exportieren(true);
}

/* ---- Start ----------------------------------------------------------- */

(async function start() {
  verdrahten();
  await dateilisteHolen();
  const namen = [...$("dateiwahl").options].map((o) => o.value);
  if (namen.includes("anforderungskatalog")) {
    $("dateiwahl").value = "anforderungskatalog";
    await laden();
  } else {
    zeichnen();
  }
})();
