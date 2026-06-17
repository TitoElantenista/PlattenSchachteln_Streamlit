from __future__ import annotations

import csv
import io
import re
import tempfile
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from queue import Empty, Queue
from threading import Event

import streamlit as st

import SchachtelnScript as schachtel


PFLICHTSPALTEN = ("lfd-Nr.", "Anzahl", "Profil", "Güte", "Länge", "Breite")
STANDARD_TAFELTYPEN = [
    {"Name": "2,0 x 12,0 m", "Breite [mm]": 2000, "Länge [mm]": 12000},
    {"Name": "2,5 x 12,0 m", "Breite [mm]": 2500, "Länge [mm]": 12000},
]


@dataclass(frozen=True)
class CSVPruefung:
    text: str
    encoding: str
    spalten: list[str]
    daten_zeilen: int
    bl_zeilen: int
    einzelteile: int
    staerken: Counter[str]
    gueten: Counter[str]
    gruppen_zeilen: Counter[tuple[str, str]]
    gruppen_teile: Counter[tuple[str, str]]
    gruppen_flaeche_mm2: Counter[tuple[str, str]]
    fehler: list[str]
    hinweise: list[str]

    @property
    def ist_gueltig(self) -> bool:
        return not self.fehler


class BerechnungAbgebrochen(RuntimeError):
    pass


def dateiname_sicher(wert: str) -> str:
    name = Path(wert or "eingabe.csv").name
    name = re.sub(r"[^A-Za-z0-9_.+-]+", "_", name)
    return name.strip("_.") or "eingabe.csv"


def csv_dekodieren(daten: bytes) -> tuple[str, str]:
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return daten.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return daten.decode("cp1252", errors="replace"), "cp1252 mit Ersatzzeichen"


def zeile_hat_inhalt(row: dict[str | None, object]) -> bool:
    for key, wert in row.items():
        if key is None:
            if any(str(eintrag).strip() for eintrag in wert or []):
                return True
            continue
        if str(wert or "").strip():
            return True
    return False


def csv_pruefen(daten: bytes) -> CSVPruefung:
    text, encoding = csv_dekodieren(daten)
    reader = csv.DictReader(io.StringIO(text), delimiter=";")
    spalten = list(reader.fieldnames or [])
    daten_zeilen = 0
    bl_zeilen = 0
    einzelteile = 0
    staerken: Counter[str] = Counter()
    gueten: Counter[str] = Counter()
    gruppen_zeilen: Counter[tuple[str, str]] = Counter()
    gruppen_teile: Counter[tuple[str, str]] = Counter()
    gruppen_flaeche_mm2: Counter[tuple[str, str]] = Counter()
    fehler: list[str] = []
    hinweise: list[str] = []

    fehlende_spalten = [spalte for spalte in PFLICHTSPALTEN if spalte not in spalten]
    if fehlende_spalten:
        fehler.append("Fehlende Pflichtspalten: " + ", ".join(fehlende_spalten))

    for zeilennummer, row in enumerate(reader, start=2):
        if not zeile_hat_inhalt(row):
            continue
        daten_zeilen += 1
        profil = (row.get("Profil") or "").strip()
        if not profil.upper().startswith("BL"):
            continue
        bl_zeilen += 1

        staerke = schachtel.staerke_aus_profil(profil)
        if staerke is None:
            hinweise.append(f"Zeile {zeilennummer}: Profil kann nicht gelesen werden: {profil}")
        else:
            staerken[staerke] += 1

        guete = schachtel.materialguete_bereinigen(row.get("Güte", ""))
        gueten[guete] += 1
        gruppe = (staerke, guete) if staerke is not None else None
        if gruppe is not None:
            gruppen_zeilen[gruppe] += 1

        try:
            anzahl = schachtel.parse_int_wert(row.get("Anzahl", "1"))
            laenge_mm = schachtel.parse_int_wert(row.get("Länge", ""))
            breite_mm = schachtel.parse_int_wert(row.get("Breite", ""))
            einzelteile += max(anzahl, 0)
            if gruppe is not None:
                gruppen_teile[gruppe] += max(anzahl, 0)
                gruppen_flaeche_mm2[gruppe] += max(anzahl, 0) * laenge_mm * breite_mm
        except ValueError as exc:
            hinweise.append(f"Zeile {zeilennummer}: Zahlenwert prüfen ({exc}).")

    if daten_zeilen == 0:
        fehler.append("Die CSV enthält keine Datenzeilen.")
    if not fehlende_spalten and bl_zeilen == 0:
        fehler.append("Es wurden keine Profile gefunden, die mit 'BL' beginnen.")

    return CSVPruefung(
        text=text,
        encoding=encoding,
        spalten=spalten,
        daten_zeilen=daten_zeilen,
        bl_zeilen=bl_zeilen,
        einzelteile=einzelteile,
        staerken=staerken,
        gueten=gueten,
        gruppen_zeilen=gruppen_zeilen,
        gruppen_teile=gruppen_teile,
        gruppen_flaeche_mm2=gruppen_flaeche_mm2,
        fehler=fehler,
        hinweise=hinweise,
    )


def counter_als_tabelle(counter: Counter[str], spaltenname: str) -> list[dict[str, object]]:
    return [
        {spaltenname: key, "CSV-Zeilen": wert}
        for key, wert in sorted(counter.items(), key=lambda item: item[0])
    ]


def dauer_text(sekunden: float) -> str:
    sekunden = max(0, int(round(sekunden)))
    minuten, rest = divmod(sekunden, 60)
    stunden, minuten = divmod(minuten, 60)
    if stunden:
        return f"{stunden} h {minuten:02d} min"
    if minuten:
        return f"{minuten} min {rest:02d} s"
    return f"{rest} s"


def berechnung_zeit_text(startzeit: float, percent: float) -> str:
    verstrichen = time.monotonic() - startzeit
    if percent > 0.03 and percent < 1.0:
        rest = max(0.0, verstrichen * (1.0 - percent) / percent)
        return (
            f"Verstrichen: {dauer_text(verstrichen)} · "
            f"Restzeit ca.: {dauer_text(rest)} · "
            f"Fortschritt: {percent * 100:.0f} %"
        )
    return (
        f"Verstrichen: {dauer_text(verstrichen)} · "
        f"Fortschritt: {percent * 100:.0f} %"
    )


def staerken_als_tabelle(counter: Counter[str]) -> list[dict[str, object]]:
    return [
        {"Blechstärke": f"BL{key}", "CSV-Zeilen": wert}
        for key, wert in sorted(counter.items(), key=lambda item: schachtel.staerke_sortierwert(item[0]))
    ]


def gruppen_sortiert(pruefung: CSVPruefung) -> list[tuple[str, str]]:
    return sorted(
        pruefung.gruppen_zeilen,
        key=lambda gruppe: (schachtel.staerke_sortierwert(gruppe[0]), gruppe[1]),
    )


def tafelzuordnung_key(index: int, tafelname: str) -> str:
    return f"tafelzuordnung_gruppen_v2_{index}_{schachtel.dateiname_sicher(tafelname)}"


def tafel_alle_key(index: int, tafelname: str) -> str:
    return f"tafelzuordnung_alle_{index}_{schachtel.dateiname_sicher(tafelname)}"


def tafel_keine_key(index: int, tafelname: str) -> str:
    return f"tafelzuordnung_keine_{index}_{schachtel.dateiname_sicher(tafelname)}"


def gruppen_auswahl_label(pruefung: CSVPruefung, gruppe: tuple[str, str]) -> str:
    staerke, guete = gruppe
    teile = pruefung.gruppen_teile.get(gruppe, 0)
    flaeche_m2 = schachtel.fmt_zahl(schachtel.mm2_zu_m2(pruefung.gruppen_flaeche_mm2.get(gruppe, 0)), 1)
    return f"BL{staerke} / {guete} - {teile} Teile - {flaeche_m2} m²"


def tafeltypen_aus_editor(rows: object) -> tuple[tuple[str, int, int], ...]:
    tafeltypen: list[tuple[str, int, int]] = []
    if hasattr(rows, "to_dict"):
        rows = rows.to_dict("records")
    if not isinstance(rows, list):
        return tuple(tafeltypen)

    for row in rows:
        if not isinstance(row, dict):
            continue
        name = str(row.get("Name") or "").strip()
        breite = row.get("Breite [mm]")
        laenge = row.get("Länge [mm]")
        if not name and (breite in (None, "") and laenge in (None, "")):
            continue
        try:
            breite_int = int(float(breite))
            laenge_int = int(float(laenge))
        except (TypeError, ValueError):
            tafeltypen.append((name, 0, 0))
            continue
        tafeltypen.append((name, breite_int, laenge_int))
    return tuple(tafeltypen)


def parameter_fehler(
    tafeltypen: tuple[tuple[str, int, int], ...],
    rand_mm: int,
    abstand_mm: int,
) -> list[str]:
    fehler: list[str] = []
    if not tafeltypen:
        fehler.append("Mindestens ein Tafeltyp ist erforderlich.")
        return fehler

    namen = [name.strip() for name, _breite, _laenge in tafeltypen]
    if any(not name for name in namen):
        fehler.append("Jeder Tafeltyp braucht einen Namen.")
    if len(set(namen)) != len(namen):
        fehler.append("Die Namen der Tafeltypen müssen unterschiedlich sein.")

    for name, breite, laenge in tafeltypen:
        if breite <= 0 or laenge <= 0:
            fehler.append(f"{name}: Breite und Länge müssen größer als 0 mm sein.")
        if 2 * rand_mm >= min(breite, laenge):
            fehler.append(f"{name}: Der Rand ist zu groß für diese Tafelabmessung.")

    if abstand_mm < 0:
        fehler.append("Der Abstand darf nicht negativ sein.")
    return fehler


def konfiguration_text(
    projektname: str,
    original_dateiname: str,
    tafeltypen: tuple[tuple[str, int, int], ...],
    rand_mm: int,
    abstand_mm: int,
    mindest_staerke_mm: int,
    stahl_kg_pro_m2_und_mm: float,
    pruefung: CSVPruefung,
    gruppen_tafelkonfiguration: schachtel.GruppenTafelKonfiguration,
) -> str:
    zeilen = [
        f"Projekt: {projektname}",
        f"CSV-Datei: {original_dateiname}",
        f"Erkannte Kodierung: {pruefung.encoding}",
        "",
        "Parameter:",
    ]
    for name, breite, laenge in tafeltypen:
        zeilen.append(f"- Tafeltyp: {name}, Breite {breite} mm, Länge {laenge} mm")
    zeilen.extend(
        [
            f"- Rand zur ersten Kontur: {rand_mm} mm",
            f"- Abstand zwischen zwei Teilen: {abstand_mm} mm",
            f"- Mindeststärke: BL{mindest_staerke_mm}",
            f"- Stahlgewicht: {stahl_kg_pro_m2_und_mm:.3f} kg/m²/mm",
            "",
            "CSV-Prüfung:",
            f"- Datenzeilen: {pruefung.daten_zeilen}",
            f"- BL-Zeilen: {pruefung.bl_zeilen}",
            f"- Einzelteile aus Anzahl: {pruefung.einzelteile}",
            f"- Blechstärken: {', '.join(f'BL{k} ({v})' for k, v in pruefung.staerken.items())}",
            f"- Güten: {', '.join(f'{k} ({v})' for k, v in pruefung.gueten.items())}",
        ]
    )
    if gruppen_tafelkonfiguration:
        zeilen.extend(["", "Abweichende Tafeleinstellungen:"])
        for (staerke, guete), (gruppe_tafeltypen, misch_erlaubt) in sorted(
            gruppen_tafelkonfiguration.items(),
            key=lambda item: (schachtel.staerke_sortierwert(item[0][0]), item[0][1]),
        ):
            tafeltypen_text = ", ".join(
                f"{name} ({breite} x {laenge} mm)"
                for name, breite, laenge in gruppe_tafeltypen
            )
            misch_text = "ja" if misch_erlaubt and len(gruppe_tafeltypen) >= 2 else "nein"
            zeilen.append(f"- BL{staerke} / {guete}: {tafeltypen_text}; Mischkombinationen: {misch_text}")
    else:
        zeilen.extend(["", "Abweichende Tafeleinstellungen: keine"])
    return "\n".join(zeilen).rstrip() + "\n"


def ausgaben_zippen(ausgabe_dir: Path, projektname: str, konfiguration: str) -> bytes:
    zip_ordner = schachtel.dateiname_sicher(projektname)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zip_datei:
        zip_datei.writestr(f"{zip_ordner}/Konfiguration.txt", konfiguration)
        for pfad in sorted(ausgabe_dir.rglob("*")):
            if not pfad.is_file():
                continue
            rel_pfad = pfad.relative_to(ausgabe_dir).as_posix()
            zip_datei.write(pfad, f"{zip_ordner}/{rel_pfad}")
    return buffer.getvalue()


def schachtelung_berechnen(
    pruefung: CSVPruefung,
    original_dateiname: str,
    projektname: str,
    tafeltypen: tuple[tuple[str, int, int], ...],
    rand_mm: int,
    abstand_mm: int,
    mindest_staerke_mm: int,
    stahl_kg_pro_m2_und_mm: float,
    gruppen_tafelkonfiguration: schachtel.GruppenTafelKonfiguration,
    progress_callback: schachtel.FortschrittCallback | None = None,
) -> tuple[bytes, schachtel.LaufErgebnis]:
    with tempfile.TemporaryDirectory(prefix="platten_schachtelung_") as temp_name:
        temp_dir = Path(temp_name)
        eingabe_dir = temp_dir / "eingabe"
        ausgabe_dir = temp_dir / "ausgabe"
        eingabe_dir.mkdir()
        ausgabe_dir.mkdir()

        csv_pfad = eingabe_dir / dateiname_sicher(original_dateiname)
        csv_pfad.write_text(pruefung.text, encoding="cp1252", errors="replace", newline="")

        ergebnis = schachtel.schachtelung_ausfuehren(
            csv_datei=csv_pfad,
            ausgabe_dir=ausgabe_dir,
            projektname=projektname,
            tafeltypen=tafeltypen,
            rand_mm=rand_mm,
            abstand_mm=abstand_mm,
            mindest_staerke_mm=mindest_staerke_mm,
            stahl_kg_pro_m2_und_mm=stahl_kg_pro_m2_und_mm,
            fortschritt_ausgeben=False,
            progress_callback=progress_callback,
            gruppen_tafelkonfiguration=gruppen_tafelkonfiguration,
        )
        konfiguration = konfiguration_text(
            projektname=projektname,
            original_dateiname=original_dateiname,
            tafeltypen=tafeltypen,
            rand_mm=rand_mm,
            abstand_mm=abstand_mm,
            mindest_staerke_mm=mindest_staerke_mm,
            stahl_kg_pro_m2_und_mm=stahl_kg_pro_m2_und_mm,
            pruefung=pruefung,
            gruppen_tafelkonfiguration=gruppen_tafelkonfiguration,
        )
        return ausgaben_zippen(ausgabe_dir, projektname, konfiguration), ergebnis


def berechnung_abbrechen_wenn_noetig(cancel_event: Event) -> None:
    if cancel_event.is_set():
        raise BerechnungAbgebrochen("Die Berechnung wurde abgebrochen.")


def fortschritt_aus_job_lesen(job: dict[str, object]) -> dict[str, object]:
    progress_queue = job["progress_queue"]
    assert isinstance(progress_queue, Queue)
    letztes_event = job.get(
        "letztes_event",
        {"percent": 0.0, "message": "Berechnung wird vorbereitet..."},
    )
    while True:
        try:
            letztes_event = progress_queue.get_nowait()
        except Empty:
            break

    percent = float(letztes_event.get("percent", 0.0) or 0.0)
    percent = max(float(job.get("letzter_percent", 0.0) or 0.0), min(percent, 1.0))
    letztes_event = {
        **letztes_event,
        "percent": percent,
        "message": str(letztes_event.get("message") or "Schachtelung wird berechnet..."),
    }
    job["letztes_event"] = letztes_event
    job["letzter_percent"] = percent
    return letztes_event


def berechnung_job_aufräumen(job: dict[str, object]) -> None:
    executor = job.get("executor")
    if isinstance(executor, ThreadPoolExecutor):
        executor.shutdown(wait=False, cancel_futures=True)
    st.session_state.pop("berechnung_job", None)


def laufende_berechnung_anzeigen() -> bool:
    job = st.session_state.get("berechnung_job")
    if not isinstance(job, dict):
        return False

    event = fortschritt_aus_job_lesen(job)
    future = job["future"]
    startzeit = float(job.get("startzeit", time.monotonic()))
    percent = float(event.get("percent", 0.0) or 0.0)
    nachricht = str(event.get("message") or "Schachtelung wird berechnet...")

    if future.done():
        try:
            zip_bytes, ergebnis = future.result()
        except BerechnungAbgebrochen:
            st.warning("Die Berechnung wurde abgebrochen.")
        except Exception as exc:
            st.error("Die Berechnung konnte nicht abgeschlossen werden.")
            st.code(str(exc))
        else:
            st.session_state["zip_bytes"] = zip_bytes
            st.session_state["zip_name"] = f"{schachtel.dateiname_sicher(str(job['projektname']))}.zip"
            st.session_state["lauf_ergebnis"] = ergebnis
            st.success("Fertig. Die ZIP-Datei wurde erzeugt. Klicken Sie unten auf 'ZIP herunterladen'.")
        finally:
            berechnung_job_aufräumen(job)
        return False

    st.subheader("Berechnungsfortschritt")
    st.caption("Die Restzeit ist eine grobe Schätzung und wird während der Berechnung angepasst.")
    st.progress(percent, text=nachricht)
    st.info(nachricht)
    st.caption(berechnung_zeit_text(startzeit, percent))

    cancel_event = job["cancel_event"]
    assert isinstance(cancel_event, Event)
    if cancel_event.is_set():
        st.warning("Abbruch angefordert. Die laufende Variante wird noch beendet.")
    elif st.button("Berechnung abbrechen", type="secondary"):
        cancel_event.set()
        st.warning("Abbruch angefordert. Die laufende Variante wird noch beendet.")

    time.sleep(0.5)
    st.rerun()
    return True


def app() -> None:
    st.set_page_config(page_title="Plattenschachtelung", page_icon=None, layout="wide")

    st.title("Plattenschachtelung")
    st.caption("CSV prüfen, Parameter festlegen und Ergebnis als ZIP herunterladen.")

    with st.expander("CSV-Aufbau", expanded=True):
        st.markdown(
            """
Die Datei muss eine CSV-Datei mit Semikolon als Trennzeichen sein. Eine Zeile beschreibt eine Position. Nur Profile, deren `Profil` mit `BL` beginnt, werden geschachtelt.

Längen und Breiten werden in Millimeter gelesen. Die Spalte `Anzahl` gibt an, wie oft diese Platte vorkommt.
"""
        )
        st.table(
            [
                {
                    "Spalte": "lfd-Nr.",
                    "Pflicht": "ja",
                    "Bedeutung": "Positionsnummer oder laufende Nummer",
                    "Beispiel": "2005",
                },
                {
                    "Spalte": "Anzahl",
                    "Pflicht": "ja",
                    "Bedeutung": "Stückzahl dieser Position",
                    "Beispiel": "4",
                },
                {
                    "Spalte": "Profil",
                    "Pflicht": "ja",
                    "Bedeutung": "Blechstärke, muss mit BL beginnen",
                    "Beispiel": "BL10",
                },
                {
                    "Spalte": "Güte",
                    "Pflicht": "ja",
                    "Bedeutung": "Materialgüte, wird unverändert übernommen",
                    "Beispiel": "S355J2+Z15",
                },
                {
                    "Spalte": "Länge",
                    "Pflicht": "ja",
                    "Bedeutung": "Plattenlänge in mm",
                    "Beispiel": "350",
                },
                {
                    "Spalte": "Breite",
                    "Pflicht": "ja",
                    "Bedeutung": "Plattenbreite in mm",
                    "Beispiel": "200",
                },
            ]
        )

    st.subheader("Projekt und Parameter")
    projektname = st.text_input(
        "Projektname",
        value="Plattenschachtelung",
        help="Dieser Name erscheint in den Ergebnisdateien und im ZIP-Dateinamen.",
    ).strip() or "Plattenschachtelung"

    st.markdown("**Tafeltypen**")
    st.caption(
        "Breite und Länge werden in Millimeter angegeben. Fügen Sie bei Bedarf weitere verfügbare Tafelmaße hinzu."
    )
    tafeltypen_editor = st.data_editor(
        STANDARD_TAFELTYPEN,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        column_config={
            "Name": st.column_config.TextColumn(
                "Name",
                help="Bezeichnung, die später in den Ergebnissen erscheint.",
                required=True,
            ),
            "Breite [mm]": st.column_config.NumberColumn(
                "Breite [mm]",
                min_value=1,
                step=50,
                required=True,
            ),
            "Länge [mm]": st.column_config.NumberColumn(
                "Länge [mm]",
                min_value=1,
                step=50,
                required=True,
            ),
        },
        key="tafeltypen_editor",
    )
    tafeltypen = tafeltypen_aus_editor(tafeltypen_editor)

    param_spalten = st.columns(4)
    with param_spalten[0]:
        rand_mm = st.number_input(
            "Rand [mm]",
            min_value=0,
            value=40,
            step=5,
            help="Freier Sicherheitsrand vom Tafelrand bis zur ersten Kontur.",
        )
    with param_spalten[1]:
        abstand_mm = st.number_input(
            "Abstand [mm]",
            min_value=0,
            value=20,
            step=5,
            help="Mindestabstand zwischen zwei geschnittenen Teilen.",
        )
    with param_spalten[2]:
        mindest_staerke_mm = st.number_input(
            "Mindeststärke [mm]",
            min_value=0,
            value=6,
            step=1,
            help="BL-Teile unter dieser Stärke werden separat gelistet und nicht geschachtelt.",
        )
    with param_spalten[3]:
        stahl_kg_pro_m2_und_mm = st.number_input(
            "Stahlgewicht [kg/m²/mm]",
            min_value=0.001,
            value=7.85,
            step=0.05,
            format="%.3f",
            help="Rechenwert für die Gewichtsangaben in Excel, HTML und Markdown.",
        )

    parameterprobleme = parameter_fehler(
        tafeltypen=tafeltypen,
        rand_mm=int(rand_mm),
        abstand_mm=int(abstand_mm),
    )
    for problem in parameterprobleme:
        st.error(problem)

    st.subheader("CSV hochladen")
    hochgeladen = st.file_uploader("CSV-Datei auswählen", type=["csv"])

    pruefung: CSVPruefung | None = None
    gruppen_tafelkonfiguration: schachtel.GruppenTafelKonfiguration = {}
    gruppen_ohne_tafel: list[tuple[str, str]] = []
    if hochgeladen is not None:
        daten = hochgeladen.getvalue()
        pruefung = csv_pruefen(daten)

        if pruefung.fehler:
            for fehler in pruefung.fehler:
                st.error(fehler)
        else:
            st.success("Die CSV-Struktur ist verwendbar.")

        st.caption(f"Erkannte Kodierung: {pruefung.encoding}")
        kennzahlen = st.columns(4)
        kennzahlen[0].metric("Datenzeilen", pruefung.daten_zeilen)
        kennzahlen[1].metric("BL-Zeilen", pruefung.bl_zeilen)
        kennzahlen[2].metric("Einzelteile", pruefung.einzelteile)
        kennzahlen[3].metric("Spalten", len(pruefung.spalten))

        listen = st.columns(2)
        with listen[0]:
            st.markdown("**Gefundene Blechstärken**")
            st.dataframe(staerken_als_tabelle(pruefung.staerken), width="stretch")
        with listen[1]:
            st.markdown("**Gefundene Güten**")
            st.dataframe(counter_als_tabelle(pruefung.gueten, "Güte"), width="stretch")

        if pruefung.hinweise:
            with st.expander("Hinweise aus der Prüfung"):
                for hinweis in pruefung.hinweise[:30]:
                    st.warning(hinweis)
                if len(pruefung.hinweise) > 30:
                    st.info(f"{len(pruefung.hinweise) - 30} weitere Hinweise ausgeblendet.")

        if pruefung.ist_gueltig:
            with st.expander("Erweiterte Tafelzuordnung", expanded=False):
                st.markdown(
                    """
Standardmäßig kann jede Blechgruppe alle oben definierten Tafeltypen verwenden. Pro Tafeltyp können Gruppen entfernt werden, die auf dieser Tafel nicht laufen sollen.
"""
                )
                st.caption(
                    "Alle Gruppen sind zuerst ausgewählt. Einschränkungen werden im ZIP in Konfiguration.txt dokumentiert."
                )
                gruppen = gruppen_sortiert(pruefung)
                gruppen_labels = {
                    gruppe: gruppen_auswahl_label(pruefung, gruppe)
                    for gruppe in gruppen
                }
                erlaubte_tafeln_pro_gruppe: dict[
                    tuple[str, str],
                    list[tuple[str, int, int]],
                ] = {gruppe: [] for gruppe in gruppen}
                eingeschraenkte_tafeln = 0

                for index, tafeltyp in enumerate(tafeltypen, start=1):
                    name, breite, laenge = tafeltyp
                    st.markdown(f"**{name}** · {breite} x {laenge} mm")
                    auswahl_key = tafelzuordnung_key(index, name)
                    gruppen_fingerprint_key = f"{auswahl_key}_gruppen"
                    gruppen_fingerprint = tuple(gruppen)
                    if st.session_state.get(gruppen_fingerprint_key) != gruppen_fingerprint:
                        st.session_state[auswahl_key] = list(gruppen)
                        st.session_state[gruppen_fingerprint_key] = gruppen_fingerprint
                    else:
                        st.session_state[auswahl_key] = [
                            gruppe
                            for gruppe in st.session_state.get(auswahl_key, list(gruppen))
                            if gruppe in gruppen
                        ]
                    auswahl_spalten = st.columns([1, 1, 5])
                    if auswahl_spalten[0].button(
                        "Alle",
                        key=tafel_alle_key(index, name),
                        help="Alle Blechgruppen für diesen Tafeltyp auswählen.",
                    ):
                        st.session_state[auswahl_key] = list(gruppen)
                    if auswahl_spalten[1].button(
                        "Keine",
                        key=tafel_keine_key(index, name),
                        help="Alle Blechgruppen für diesen Tafeltyp abwählen.",
                    ):
                        st.session_state[auswahl_key] = []
                    gewaehlte_gruppen = st.multiselect(
                        "Zulässige Blechgruppen",
                        gruppen,
                        format_func=lambda gruppe: gruppen_labels[gruppe],
                        key=auswahl_key,
                        help="Entfernen Sie die Gruppen, die nicht auf diesem Tafeltyp geschachtelt werden sollen.",
                    )
                    if len(gewaehlte_gruppen) != len(gruppen):
                        eingeschraenkte_tafeln += 1
                    for gruppe in gewaehlte_gruppen:
                        erlaubte_tafeln_pro_gruppe[gruppe].append(tafeltyp)

                gruppen_ohne_tafel = [
                    gruppe for gruppe, erlaubte_tafeln in erlaubte_tafeln_pro_gruppe.items()
                    if not erlaubte_tafeln
                ]
                if gruppen_ohne_tafel:
                    st.warning(
                        f"{len(gruppen_ohne_tafel)} Gruppe"
                        f"{'n haben' if len(gruppen_ohne_tafel) != 1 else ' hat'} "
                        "keine erlaubte Tafel."
                    )
                    st.dataframe(
                        [
                            {"Gruppe": gruppen_labels[gruppe]}
                            for gruppe in gruppen_ohne_tafel
                        ],
                        width="stretch",
                        hide_index=True,
                    )

                standard_tafeltypen = tuple(tafeltypen)
                for gruppe, erlaubte_tafeln in erlaubte_tafeln_pro_gruppe.items():
                    erlaubte_tuple = tuple(erlaubte_tafeln)
                    if erlaubte_tuple != standard_tafeltypen:
                        gruppen_tafelkonfiguration[gruppe] = (erlaubte_tuple, True)

                if eingeschraenkte_tafeln:
                    st.info(
                        f"{eingeschraenkte_tafeln} Tafeltyp(en) wurden auf bestimmte Gruppen eingeschränkt."
                    )

                with st.expander("Vorschau der wirksamen Tafeltypen pro Gruppe", expanded=False):
                    st.dataframe(
                        [
                            {
                                "Gruppe": gruppen_labels[(staerke, guete)],
                                "Teile": pruefung.gruppen_teile.get((staerke, guete), 0),
                                "Erlaubte Tafeltypen": ", ".join(
                                    name for name, _breite, _laenge
                                    in erlaubte_tafeln_pro_gruppe[(staerke, guete)]
                                ),
                            }
                            for staerke, guete in gruppen
                        ],
                        width="stretch",
                    )

            if gruppen_ohne_tafel:
                anzahl_gruppen = len(gruppen_ohne_tafel)
                st.error(
                    f"{anzahl_gruppen} Gruppe{'n haben' if anzahl_gruppen != 1 else ' hat'} "
                    "keine erlaubte Tafel."
                )

    kann_berechnen = (
        pruefung is not None
        and pruefung.ist_gueltig
        and not parameterprobleme
        and not gruppen_ohne_tafel
    )

    berechnung_laeuft = laufende_berechnung_anzeigen()

    if st.button("Schachtelung berechnen", type="primary", disabled=not kann_berechnen or berechnung_laeuft):
        assert pruefung is not None
        st.session_state.pop("zip_bytes", None)
        st.session_state.pop("zip_name", None)
        st.session_state.pop("lauf_ergebnis", None)

        cancel_event = Event()
        progress_queue: Queue[dict[str, object]] = Queue()

        def job_fortschritt(event: dict[str, object]) -> None:
            berechnung_abbrechen_wenn_noetig(cancel_event)
            progress_queue.put(event)
            berechnung_abbrechen_wenn_noetig(cancel_event)

        executor = ThreadPoolExecutor(max_workers=1)
        future = executor.submit(
            schachtelung_berechnen,
            pruefung=pruefung,
            original_dateiname=hochgeladen.name if hochgeladen is not None else "eingabe.csv",
            projektname=projektname,
            tafeltypen=tafeltypen,
            rand_mm=int(rand_mm),
            abstand_mm=int(abstand_mm),
            mindest_staerke_mm=int(mindest_staerke_mm),
            stahl_kg_pro_m2_und_mm=float(stahl_kg_pro_m2_und_mm),
            gruppen_tafelkonfiguration={
                gruppe: (tuple(tafeltypen_gruppe), bool(misch_erlaubt))
                for gruppe, (tafeltypen_gruppe, misch_erlaubt)
                in gruppen_tafelkonfiguration.items()
            },
            progress_callback=job_fortschritt,
        )
        st.session_state["berechnung_job"] = {
            "future": future,
            "executor": executor,
            "cancel_event": cancel_event,
            "progress_queue": progress_queue,
            "projektname": projektname,
            "startzeit": time.monotonic(),
            "letzter_percent": 0.0,
            "letztes_event": {
                "percent": 0.0,
                "message": "Berechnung wird vorbereitet...",
            },
        }
        st.rerun()

    if "zip_bytes" in st.session_state and "lauf_ergebnis" in st.session_state:
        ergebnis = st.session_state["lauf_ergebnis"]
        st.subheader("Ergebnis")
        ergebnis_spalten = st.columns(5)
        ergebnis_spalten[0].metric("Gruppen", ergebnis.gruppen)
        ergebnis_spalten[1].metric("Tafeln", ergebnis.tafeln_gesamt)
        ergebnis_spalten[2].metric("BL-Teile", ergebnis.teile_ab_mindest)
        ergebnis_spalten[3].metric("Unter Mindeststärke", ergebnis.teile_unter_mindest)
        ergebnis_spalten[4].metric("Nicht platzierbar", ergebnis.nicht_platzierbar)
        st.download_button(
            "ZIP herunterladen",
            data=st.session_state["zip_bytes"],
            file_name=st.session_state["zip_name"],
            mime="application/zip",
        )


if __name__ == "__main__":
    app()
