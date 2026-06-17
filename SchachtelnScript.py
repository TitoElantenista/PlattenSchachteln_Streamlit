#!/usr/bin/env python3
from __future__ import annotations

import csv
import hashlib
import html
import math
import re
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable
from urllib.parse import quote
from xml.sax.saxutils import escape as xml_escape


SCRIPT_DIR = Path(__file__).resolve().parent
PROJEKTNAME = "Plattenschachtelung"
CSV_DATEI = SCRIPT_DIR / "eingabe.csv"
AUSGABE_MD = SCRIPT_DIR / "Schachtel Ergebnis.md"
AUSGABE_CSV = SCRIPT_DIR / "Schachtel Ergebnis.csv"
AUSGABE_UNTER_MINDEST_CSV = SCRIPT_DIR / "Platten unter Mindeststaerke.csv"
AUSGABE_HTML = SCRIPT_DIR / "Schachtel Kontrolle.html"
AUSGABE_XLSX = SCRIPT_DIR / "Materialliste Bestellung.xlsx"
AUSGABE_BILDER_DIR = SCRIPT_DIR / "Schachtel Bilder"

TAFELTYPEN = (
    ("2,0 x 12,0 m", 2000, 12000),
    ("2,5 x 12,0 m", 2500, 12000),
)
RAND_MM = 40
ABSTAND_MM = 20
MINDEST_STAERKE_MM = 6
STAHL_KG_PRO_M2_UND_MM = 7.85
SORTIERSTRATEGIEN = (
    "kurze_seite",
    "flaeche",
    "lange_seite",
    "umfang",
    "breite",
    "hoehe",
    "seitenverhaeltnis",
    "lfd_mix",
)
PLATZIERSTRATEGIEN = (
    "best_short_side",
    "best_area",
    "bottom_left",
    "contact",
)
SUCHSTRATEGIEN = (
    ("kurze_seite", "best_short_side"),
    ("flaeche", "best_short_side"),
    ("lange_seite", "best_short_side"),
    ("umfang", "best_area"),
    ("kurze_seite", "contact"),
    ("flaeche", "contact"),
)
KOMBI_SUCHSTRATEGIEN = (
    ("kurze_seite", "best_short_side"),
    ("flaeche", "best_short_side"),
    ("kurze_seite", "contact"),
    ("flaeche", "contact"),
    ("lange_seite", "bottom_left"),
)
FORTSCHRITT_AUSGEBEN = True
FortschrittCallback = Callable[[dict[str, object]], None]
GruppenTafelKonfiguration = dict[
    tuple[str, str],
    tuple[tuple[tuple[str, int, int], ...], bool],
]
GRUPPEN_TAFELKONFIGURATION: GruppenTafelKonfiguration = {}

TEIL_FARBEN = (
    "#6B8EED",
    "#F2A65A",
    "#4C956C",
    "#D96459",
    "#8E6C88",
    "#2A9D8F",
    "#E0B43D",
    "#5C80BC",
    "#C77D72",
    "#6A994E",
    "#B56576",
    "#457B9D",
)


@dataclass(frozen=True)
class Teil:
    nummer: int
    lfd_nr: str
    profil: str
    staerke: str
    guete: str
    laenge_mm: int
    breite_mm: int
    mengen_index: int

    @property
    def flaeche_mm2(self) -> int:
        return self.laenge_mm * self.breite_mm


@dataclass
class Platzierung:
    teil: Teil
    tafel_nr: int
    tafeltyp: str
    x_mm: int
    y_mm: int
    laenge_mm: int
    breite_mm: int
    gedreht: bool


@dataclass
class FreierBereich:
    x_mm: int
    y_mm: int
    laenge_mm: int
    breite_mm: int

    @property
    def rechts_mm(self) -> int:
        return self.x_mm + self.laenge_mm

    @property
    def unten_mm(self) -> int:
        return self.y_mm + self.breite_mm


@dataclass
class Tafel:
    nr: int
    typ: str
    breite_mm: int
    laenge_mm: int
    strategie: str = "best_short_side"
    freie_bereiche: list[FreierBereich] = field(default_factory=list)
    platzierungen: list[Platzierung] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.freie_bereiche:
            return
        nutzlaenge = self.laenge_mm - 2 * RAND_MM
        nutzbreite = self.breite_mm - 2 * RAND_MM
        if nutzlaenge <= 0 or nutzbreite <= 0:
            return
        self.freie_bereiche.append(
            FreierBereich(
                x_mm=RAND_MM,
                y_mm=RAND_MM,
                laenge_mm=nutzlaenge + ABSTAND_MM,
                breite_mm=nutzbreite + ABSTAND_MM,
            )
        )

    @property
    def flaeche_mm2(self) -> int:
        return self.breite_mm * self.laenge_mm

    @property
    def belegte_flaeche_mm2(self) -> int:
        return sum(p.teil.flaeche_mm2 for p in self.platzierungen)

    def passt_auf_leere_tafel(self, teil: Teil) -> bool:
        return any(
            laenge <= self.laenge_mm - 2 * RAND_MM
            and breite <= self.breite_mm - 2 * RAND_MM
            for laenge, breite, _ in orientierungen(teil)
        )

    def platziere(self, teil: Teil) -> bool:
        kandidat = self._finde_position_in_freien_bereichen(teil)
        if kandidat is None:
            return False
        self.platziere_kandidat(teil, kandidat)
        return True

    def platziere_kandidat(
        self,
        teil: Teil,
        kandidat: tuple[tuple[int, int, int, int, int], FreierBereich, int, int, bool],
    ) -> None:
        _score, bereich, laenge, breite, gedreht = kandidat
        platzierung = Platzierung(
            teil=teil,
            tafel_nr=self.nr,
            tafeltyp=self.typ,
            x_mm=bereich.x_mm,
            y_mm=bereich.y_mm,
            laenge_mm=laenge,
            breite_mm=breite,
            gedreht=gedreht,
        )
        self.platzierungen.append(platzierung)
        belegter_bereich = FreierBereich(
            x_mm=bereich.x_mm,
            y_mm=bereich.y_mm,
            laenge_mm=laenge + ABSTAND_MM,
            breite_mm=breite + ABSTAND_MM,
        )
        self._belege_freien_bereich(belegter_bereich)

    def _finde_position_in_freien_bereichen(
        self, teil: Teil
    ) -> tuple[tuple[int, int, int, int, int], FreierBereich, int, int, bool] | None:
        kandidaten: list[tuple[tuple[int, int, int, int, int], FreierBereich, int, int, bool]] = []
        for bereich in self.freie_bereiche:
            for laenge, breite, gedreht in orientierungen(teil):
                belegte_laenge = laenge + ABSTAND_MM
                belegte_breite = breite + ABSTAND_MM
                if belegte_laenge > bereich.laenge_mm or belegte_breite > bereich.breite_mm:
                    continue
                restlaenge = bereich.laenge_mm - belegte_laenge
                restbreite = bereich.breite_mm - belegte_breite
                score = self._score_fuer_position(
                    bereich=bereich,
                    belegte_laenge=belegte_laenge,
                    belegte_breite=belegte_breite,
                    restlaenge=restlaenge,
                    restbreite=restbreite,
                )
                kandidaten.append((score, bereich, laenge, breite, gedreht))

        if not kandidaten:
            return None
        return min(kandidaten, key=lambda item: item[0])

    def _score_fuer_position(
        self,
        bereich: FreierBereich,
        belegte_laenge: int,
        belegte_breite: int,
        restlaenge: int,
        restbreite: int,
    ) -> tuple[int, int, int, int, int]:
        restflaeche = restlaenge * restbreite
        if self.strategie == "best_area":
            return (restflaeche, min(restlaenge, restbreite), max(restlaenge, restbreite), bereich.y_mm, bereich.x_mm)
        if self.strategie == "bottom_left":
            return (bereich.y_mm + belegte_breite, bereich.x_mm, restflaeche, min(restlaenge, restbreite), 0)
        if self.strategie == "contact":
            kontakt = self._kontaktlaenge(bereich, belegte_laenge, belegte_breite)
            return (-kontakt, restflaeche, min(restlaenge, restbreite), bereich.y_mm, bereich.x_mm)
        return (min(restlaenge, restbreite), max(restlaenge, restbreite), restflaeche, bereich.y_mm, bereich.x_mm)

    def _kontaktlaenge(self, bereich: FreierBereich, belegte_laenge: int, belegte_breite: int) -> int:
        kontakt = 0
        nutz_rechts = RAND_MM + (self.laenge_mm - 2 * RAND_MM) + ABSTAND_MM
        nutz_unten = RAND_MM + (self.breite_mm - 2 * RAND_MM) + ABSTAND_MM

        if bereich.x_mm == RAND_MM:
            kontakt += belegte_breite
        if bereich.y_mm == RAND_MM:
            kontakt += belegte_laenge
        if bereich.x_mm + belegte_laenge == nutz_rechts:
            kontakt += belegte_breite
        if bereich.y_mm + belegte_breite == nutz_unten:
            kontakt += belegte_laenge

        for platzierung in self.platzierungen:
            platz_x = platzierung.x_mm
            platz_y = platzierung.y_mm
            platz_rechts = platz_x + platzierung.laenge_mm + ABSTAND_MM
            platz_unten = platz_y + platzierung.breite_mm + ABSTAND_MM
            bereich_rechts = bereich.x_mm + belegte_laenge
            bereich_unten = bereich.y_mm + belegte_breite

            if platz_rechts == bereich.x_mm or bereich_rechts == platz_x:
                kontakt += ueberlappende_laenge(
                    platz_y, platz_unten, bereich.y_mm, bereich_unten
                )
            if platz_unten == bereich.y_mm or bereich_unten == platz_y:
                kontakt += ueberlappende_laenge(
                    platz_x, platz_rechts, bereich.x_mm, bereich_rechts
                )
        return kontakt

    def _belege_freien_bereich(self, belegter_bereich: FreierBereich) -> None:
        neue_bereiche: list[FreierBereich] = []
        for bereich in self.freie_bereiche:
            if not bereiche_ueberlappen(bereich, belegter_bereich):
                neue_bereiche.append(bereich)
                continue
            neue_bereiche.extend(splitte_freien_bereich(bereich, belegter_bereich))

        self.freie_bereiche = bereiche_bereinigen(neue_bereiche)


def bereiche_ueberlappen(a: FreierBereich, b: FreierBereich) -> bool:
    return a.x_mm < b.rechts_mm and a.rechts_mm > b.x_mm and a.y_mm < b.unten_mm and a.unten_mm > b.y_mm


def ueberlappende_laenge(a_start: int, a_ende: int, b_start: int, b_ende: int) -> int:
    return max(0, min(a_ende, b_ende) - max(a_start, b_start))


def splitte_freien_bereich(
    freier_bereich: FreierBereich, belegter_bereich: FreierBereich
) -> list[FreierBereich]:
    neue_bereiche: list[FreierBereich] = []

    if belegter_bereich.x_mm < freier_bereich.rechts_mm and belegter_bereich.rechts_mm > freier_bereich.x_mm:
        if belegter_bereich.y_mm > freier_bereich.y_mm:
            neue_bereiche.append(
                FreierBereich(
                    x_mm=freier_bereich.x_mm,
                    y_mm=freier_bereich.y_mm,
                    laenge_mm=freier_bereich.laenge_mm,
                    breite_mm=belegter_bereich.y_mm - freier_bereich.y_mm,
                )
            )
        if belegter_bereich.unten_mm < freier_bereich.unten_mm:
            neue_bereiche.append(
                FreierBereich(
                    x_mm=freier_bereich.x_mm,
                    y_mm=belegter_bereich.unten_mm,
                    laenge_mm=freier_bereich.laenge_mm,
                    breite_mm=freier_bereich.unten_mm - belegter_bereich.unten_mm,
                )
            )

    if belegter_bereich.y_mm < freier_bereich.unten_mm and belegter_bereich.unten_mm > freier_bereich.y_mm:
        if belegter_bereich.x_mm > freier_bereich.x_mm:
            neue_bereiche.append(
                FreierBereich(
                    x_mm=freier_bereich.x_mm,
                    y_mm=freier_bereich.y_mm,
                    laenge_mm=belegter_bereich.x_mm - freier_bereich.x_mm,
                    breite_mm=freier_bereich.breite_mm,
                )
            )
        if belegter_bereich.rechts_mm < freier_bereich.rechts_mm:
            neue_bereiche.append(
                FreierBereich(
                    x_mm=belegter_bereich.rechts_mm,
                    y_mm=freier_bereich.y_mm,
                    laenge_mm=freier_bereich.rechts_mm - belegter_bereich.rechts_mm,
                    breite_mm=freier_bereich.breite_mm,
                )
            )

    return [
        bereich
        for bereich in neue_bereiche
        if bereich.laenge_mm > 0 and bereich.breite_mm > 0
    ]


def bereich_enthaelt(a: FreierBereich, b: FreierBereich) -> bool:
    return (
        a.x_mm <= b.x_mm
        and a.y_mm <= b.y_mm
        and a.rechts_mm >= b.rechts_mm
        and a.unten_mm >= b.unten_mm
    )


def bereiche_bereinigen(bereiche: list[FreierBereich]) -> list[FreierBereich]:
    bereinigte: list[FreierBereich] = []
    for index, bereich in enumerate(bereiche):
        if any(
            anderer_index != index and bereich_enthaelt(anderer, bereich)
            for anderer_index, anderer in enumerate(bereiche)
        ):
            continue
        bereinigte.append(bereich)
    return bereinigte


@dataclass
class SchachtelErgebnis:
    staerke: str
    guete: str
    tafeltyp: str
    tafel_breite_mm: int
    tafel_laenge_mm: int
    tafeln: list[Tafel]
    nicht_platzierbar: list[Teil]

    @property
    def platzierungen(self) -> list[Platzierung]:
        return [p for tafel in self.tafeln for p in tafel.platzierungen]

    @property
    def teile_gesamt(self) -> int:
        return len(self.platzierungen) + len(self.nicht_platzierbar)

    @property
    def belegte_flaeche_mm2(self) -> int:
        return sum(p.teil.flaeche_mm2 for p in self.platzierungen)

    @property
    def tafel_flaeche_mm2(self) -> int:
        return sum(tafel.flaeche_mm2 for tafel in self.tafeln)

    @property
    def verschnitt_prozent(self) -> float:
        if not self.tafel_flaeche_mm2:
            return 0.0
        return 100.0 * (1.0 - self.belegte_flaeche_mm2 / self.tafel_flaeche_mm2)


@dataclass(frozen=True)
class MaterialPosition:
    staerke: str
    guete: str
    tafeltyp: str
    breite_mm: int
    laenge_mm: int
    anzahl_tafeln: int
    teile: int
    belegte_flaeche_mm2: int
    tafel_flaeche_mm2: int
    nicht_platzierbar: int

    @property
    def verschnitt_prozent(self) -> float:
        if not self.tafel_flaeche_mm2:
            return 0.0
        return 100.0 * (1.0 - self.belegte_flaeche_mm2 / self.tafel_flaeche_mm2)


@dataclass(frozen=True)
class UnterMindestGruppe:
    staerke: str
    guete: str
    profil: str
    laenge_mm: int
    breite_mm: int
    anzahl: int
    lfd_nr_liste: tuple[str, ...]
    flaeche_mm2: int
    gewicht_kg: float | None


@dataclass(frozen=True)
class LaufErgebnis:
    projektname: str
    teile_ab_mindest: int
    teile_unter_mindest: int
    gruppen: int
    tafeln_gesamt: int
    nicht_platzierbar: int
    ausgabe_dateien: tuple[Path, ...]
    bilder_dir: Path
    bilder_anzahl: int


def parse_int_wert(wert: str) -> int:
    text = (wert or "").strip().replace(".", "").replace(",", ".")
    if not text:
        raise ValueError("leerer Zahlenwert")
    return int(round(float(text)))


def eingabedatei_vorhanden() -> bool:
    if CSV_DATEI.exists():
        return True
    print(f"Fehler: Eingabedatei nicht gefunden: {CSV_DATEI}")
    return False


def materialguete_bereinigen(wert: str) -> str:
    guete = (wert or "").strip()
    if not guete:
        return "ohne Angabe"
    return guete


def staerke_aus_profil(profil: str) -> str | None:
    treffer = re.match(r"^BL\s*([0-9]+(?:[.,][0-9]+)?)", profil.strip(), flags=re.IGNORECASE)
    if not treffer:
        return None
    staerke = treffer.group(1).replace(",", ".")
    if staerke.endswith(".0"):
        staerke = staerke[:-2]
    return staerke


def orientierungen(teil: Teil) -> list[tuple[int, int, bool]]:
    basis = [
        (teil.laenge_mm, teil.breite_mm, False),
        (teil.breite_mm, teil.laenge_mm, True),
    ]
    eindeutig: list[tuple[int, int, bool]] = []
    gesehen: set[tuple[int, int]] = set()
    for laenge, breite, gedreht in basis:
        if (laenge, breite) in gesehen:
            continue
        gesehen.add((laenge, breite))
        eindeutig.append((laenge, breite, gedreht))
    return sorted(eindeutig, key=lambda item: (item[1], item[0]))


def teile_laden() -> tuple[list[Teil], list[Teil]]:
    teile: list[Teil] = []
    teile_unter_mindest: list[Teil] = []
    naechste_nummer = 1
    naechste_unter_mindest_nummer = 1
    with CSV_DATEI.open("r", encoding="cp1252", newline="") as handle:
        reader = csv.DictReader(handle, delimiter=";")
        for zeilennummer, row in enumerate(reader, start=2):
            profil = (row.get("Profil") or "").strip()
            if not profil.upper().startswith("BL"):
                continue

            staerke = staerke_aus_profil(profil)
            if staerke is None:
                print(f"Warnung: Profil in Zeile {zeilennummer} wird uebersprungen: {profil}")
                continue
            staerke_mm = staerke_als_float(staerke)

            try:
                anzahl = parse_int_wert(row.get("Anzahl", "1"))
                laenge_mm = parse_int_wert(row.get("Länge", ""))
                breite_mm = parse_int_wert(row.get("Breite", ""))
            except ValueError as exc:
                print(f"Warnung: Zeile {zeilennummer} wird uebersprungen: {exc}")
                continue

            guete = materialguete_bereinigen(row.get("Güte", ""))
            lfd_nr = (row.get("lfd-Nr.") or "").strip() or str(zeilennummer)
            ziel_liste = teile
            if staerke_mm is None or staerke_mm < MINDEST_STAERKE_MM:
                ziel_liste = teile_unter_mindest

            for mengen_index in range(1, anzahl + 1):
                nummer = (
                    naechste_nummer
                    if ziel_liste is teile
                    else naechste_unter_mindest_nummer
                )
                ziel_liste.append(
                    Teil(
                        nummer=nummer,
                        lfd_nr=lfd_nr,
                        profil=profil,
                        staerke=staerke,
                        guete=guete,
                        laenge_mm=laenge_mm,
                        breite_mm=breite_mm,
                        mengen_index=mengen_index,
                    )
                )
                if ziel_liste is teile:
                    naechste_nummer += 1
                else:
                    naechste_unter_mindest_nummer += 1
    return teile, teile_unter_mindest


def teile_sortierschluessel(teil: Teil) -> tuple[int, int, int]:
    kurze_seite = min(teil.laenge_mm, teil.breite_mm)
    lange_seite = max(teil.laenge_mm, teil.breite_mm)
    return kurze_seite, lange_seite, teil.flaeche_mm2


def stabiler_teil_hash(teil: Teil) -> int:
    text = f"{teil.lfd_nr}-{teil.mengen_index}-{teil.laenge_mm}-{teil.breite_mm}"
    wert = 0
    for zeichen in text:
        wert = (wert * 131 + ord(zeichen)) % 1_000_003
    return wert


def teile_sortieren(teile: list[Teil], strategie: str) -> list[Teil]:
    def schluessel(teil: Teil) -> tuple[float, ...]:
        kurze_seite = min(teil.laenge_mm, teil.breite_mm)
        lange_seite = max(teil.laenge_mm, teil.breite_mm)
        flaeche = teil.flaeche_mm2
        umfang = teil.laenge_mm + teil.breite_mm
        seitenverhaeltnis = lange_seite / max(kurze_seite, 1)

        if strategie == "flaeche":
            return (flaeche, lange_seite, kurze_seite, -stabiler_teil_hash(teil))
        if strategie == "lange_seite":
            return (lange_seite, flaeche, kurze_seite, -stabiler_teil_hash(teil))
        if strategie == "umfang":
            return (umfang, flaeche, lange_seite, -stabiler_teil_hash(teil))
        if strategie == "breite":
            return (teil.breite_mm, teil.laenge_mm, flaeche, -stabiler_teil_hash(teil))
        if strategie == "hoehe":
            return (teil.laenge_mm, teil.breite_mm, flaeche, -stabiler_teil_hash(teil))
        if strategie == "seitenverhaeltnis":
            return (seitenverhaeltnis, flaeche, lange_seite, -stabiler_teil_hash(teil))
        if strategie == "lfd_mix":
            return (flaeche, lange_seite, stabiler_teil_hash(teil))
        return (kurze_seite, lange_seite, flaeche, -stabiler_teil_hash(teil))

    return sorted(teile, key=schluessel, reverse=True)


def schachteln_fuer_tafeltyp(
    staerke: str,
    guete: str,
    teile: list[Teil],
    tafeltyp: str,
    breite_mm: int,
    laenge_mm: int,
    sortierstrategie: str = "kurze_seite",
    platzierstrategie: str = "best_short_side",
    progress_callback: FortschrittCallback | None = None,
    progress_label: str = "",
) -> SchachtelErgebnis:
    tafeln: list[Tafel] = []
    nicht_platzierbar: list[Teil] = []
    sortierte_teile = teile_sortieren(teile, sortierstrategie)
    melde_intervall = max(1, len(sortierte_teile) // 25)
    zuletzt_gemeldet = 0

    for teil_index, teil in enumerate(sortierte_teile, start=1):
        probe = Tafel(
            nr=0,
            typ=tafeltyp,
            breite_mm=breite_mm,
            laenge_mm=laenge_mm,
            strategie=platzierstrategie,
        )
        if not probe.passt_auf_leere_tafel(teil):
            nicht_platzierbar.append(teil)
            continue

        tafel_kandidaten = []
        for tafel in tafeln:
            kandidat = tafel._finde_position_in_freien_bereichen(teil)
            if kandidat is not None:
                tafel_kandidaten.append(
                    (
                        kandidat[0],
                        -tafel.belegte_flaeche_mm2,
                        tafel.nr,
                        id(tafel),
                        kandidat,
                        tafel,
                    )
                )
        if tafel_kandidaten:
            (
                _score,
                _belegte_flaeche_negativ,
                _tafel_nr,
                _tafel_id,
                kandidat,
                tafel,
            ) = min(tafel_kandidaten)
            tafel.platziere_kandidat(teil, kandidat)
        else:
            tafel = Tafel(
                nr=len(tafeln) + 1,
                typ=tafeltyp,
                breite_mm=breite_mm,
                laenge_mm=laenge_mm,
                strategie=platzierstrategie,
            )
            if tafel.platziere(teil):
                tafeln.append(tafel)
            else:
                nicht_platzierbar.append(teil)

        if progress_callback and (
            teil_index == len(sortierte_teile)
            or teil_index - zuletzt_gemeldet >= melde_intervall
        ):
            progress_callback(
                {
                    "event": "work",
                    "increment": teil_index - zuletzt_gemeldet,
                    "message": (
                        f"{progress_label} - {teil_index}/{len(sortierte_teile)} Teile"
                    ),
                }
            )
            zuletzt_gemeldet = teil_index

    return SchachtelErgebnis(
        staerke=staerke,
        guete=guete,
        tafeltyp=tafeltyp,
        tafel_breite_mm=breite_mm,
        tafel_laenge_mm=laenge_mm,
        tafeln=tafeln,
        nicht_platzierbar=nicht_platzierbar,
    )


def schachteln_fuer_feste_tafeln(
    staerke: str,
    guete: str,
    teile: list[Teil],
    tafelspezifikationen: list[tuple[str, int, int]],
    sortierstrategie: str = "kurze_seite",
    platzierstrategie: str = "best_short_side",
    progress_callback: FortschrittCallback | None = None,
    progress_label: str = "",
) -> SchachtelErgebnis:
    tafeln = [
        Tafel(
            nr=index,
            typ=typ,
            breite_mm=breite,
            laenge_mm=laenge,
            strategie=platzierstrategie,
        )
        for index, (typ, breite, laenge) in enumerate(tafelspezifikationen, start=1)
    ]
    nicht_platzierbar: list[Teil] = []
    sortierte_teile = teile_sortieren(teile, sortierstrategie)
    melde_intervall = max(1, len(sortierte_teile) // 25)
    zuletzt_gemeldet = 0

    for teil_index, teil in enumerate(sortierte_teile, start=1):
        tafel_kandidaten = []
        for tafel in tafeln:
            kandidat = tafel._finde_position_in_freien_bereichen(teil)
            if kandidat is not None:
                tafel_kandidaten.append(
                    (
                        kandidat[0],
                        tafel.flaeche_mm2,
                        -tafel.belegte_flaeche_mm2,
                        tafel.nr,
                        id(tafel),
                        kandidat,
                        tafel,
                    )
                )
        if not tafel_kandidaten:
            nicht_platzierbar.append(teil)
            if progress_callback and (
                teil_index == len(sortierte_teile)
                or teil_index - zuletzt_gemeldet >= melde_intervall
            ):
                progress_callback(
                    {
                        "event": "work",
                        "increment": teil_index - zuletzt_gemeldet,
                        "message": (
                            f"{progress_label} - {teil_index}/{len(sortierte_teile)} Teile"
                        ),
                    }
                )
                zuletzt_gemeldet = teil_index
            continue

        (
            _score,
            _tafel_flaeche,
            _belegte_flaeche_negativ,
            _tafel_nr,
            _tafel_id,
            kandidat,
            tafel,
        ) = min(tafel_kandidaten)
        tafel.platziere_kandidat(teil, kandidat)

        if progress_callback and (
            teil_index == len(sortierte_teile)
            or teil_index - zuletzt_gemeldet >= melde_intervall
        ):
            progress_callback(
                {
                    "event": "work",
                    "increment": teil_index - zuletzt_gemeldet,
                    "message": (
                        f"{progress_label} - {teil_index}/{len(sortierte_teile)} Teile"
                    ),
                }
            )
            zuletzt_gemeldet = teil_index

    verwendete_tafeln = [tafel for tafel in tafeln if tafel.platzierungen]
    for neue_nr, tafel in enumerate(verwendete_tafeln, start=1):
        tafel.nr = neue_nr
        for platzierung in tafel.platzierungen:
            platzierung.tafel_nr = neue_nr
    return SchachtelErgebnis(
        staerke=staerke,
        guete=guete,
        tafeltyp=tafeltypen_text_aus_tafeln(verwendete_tafeln),
        tafel_breite_mm=0,
        tafel_laenge_mm=0,
        tafeln=verwendete_tafeln,
        nicht_platzierbar=nicht_platzierbar,
    )


def bestes_ergebnis(
    staerke: str,
    guete: str,
    teile: list[Teil],
    tafeltypen: tuple[tuple[str, int, int], ...] | None = None,
    misch_kombinationen_erlaubt: bool = True,
    progress_callback: FortschrittCallback | None = None,
    progress_label: str = "",
) -> SchachtelErgebnis:
    tafeltypen = tafeltypen or TAFELTYPEN
    varianten = homogene_tafelvarianten(
        staerke,
        guete,
        teile,
        tafeltypen=tafeltypen,
        progress_callback=progress_callback,
        progress_label=progress_label,
    )
    beste_homogene_variante = min(varianten, key=ergebnis_sortierschluessel)
    varianten.extend(
        gemischte_tafelvarianten(
            staerke,
            guete,
            teile,
            beste_homogene_variante,
            tafeltypen=tafeltypen,
            misch_kombinationen_erlaubt=misch_kombinationen_erlaubt,
            progress_callback=progress_callback,
            progress_label=progress_label,
        )
    )
    return min(
        varianten,
        key=ergebnis_sortierschluessel,
    )


def homogene_tafelvarianten(
    staerke: str,
    guete: str,
    teile: list[Teil],
    tafeltypen: tuple[tuple[str, int, int], ...] | None = None,
    progress_callback: FortschrittCallback | None = None,
    progress_label: str = "",
) -> list[SchachtelErgebnis]:
    tafeltypen = tafeltypen or TAFELTYPEN
    varianten: list[SchachtelErgebnis] = []
    varianten_gesamt = len(tafeltypen) * len(SUCHSTRATEGIEN)
    varianten_index = 0
    for typ, breite, laenge in tafeltypen:
        for sortierstrategie, platzierstrategie in SUCHSTRATEGIEN:
            varianten_index += 1
            label = (
                f"{progress_label} - homogene Variante {varianten_index}/{varianten_gesamt}: "
                f"{typ}, {sortierstrategie}/{platzierstrategie}"
            ).strip(" -")
            varianten.append(
                schachteln_fuer_tafeltyp(
                    staerke,
                    guete,
                    teile,
                    typ,
                    breite,
                    laenge,
                    sortierstrategie=sortierstrategie,
                    platzierstrategie=platzierstrategie,
                    progress_callback=progress_callback,
                    progress_label=label,
                )
            )
    return varianten


def ergebnis_sortierschluessel(ergebnis: SchachtelErgebnis) -> tuple[int, int, int, float]:
    return (
        len(ergebnis.nicht_platzierbar),
        ergebnis.tafel_flaeche_mm2,
        len(ergebnis.tafeln),
        ergebnis.verschnitt_prozent,
    )


def zaehlkombinationen(summe: int, laenge: int) -> list[tuple[int, ...]]:
    if laenge <= 0:
        return []
    if laenge == 1:
        return [(summe,)]

    kombinationen: list[tuple[int, ...]] = []
    for wert in range(summe + 1):
        for rest in zaehlkombinationen(summe - wert, laenge - 1):
            kombinationen.append((wert, *rest))
    return kombinationen


def gemischte_tafelvarianten(
    staerke: str,
    guete: str,
    teile: list[Teil],
    beste_homogene_variante: SchachtelErgebnis,
    tafeltypen: tuple[tuple[str, int, int], ...] | None = None,
    misch_kombinationen_erlaubt: bool = True,
    progress_callback: FortschrittCallback | None = None,
    progress_label: str = "",
) -> list[SchachtelErgebnis]:
    tafeltypen = tafeltypen or TAFELTYPEN
    if beste_homogene_variante.nicht_platzierbar:
        return []
    if not misch_kombinationen_erlaubt or len(tafeltypen) < 2:
        return []

    varianten: list[SchachtelErgebnis] = []
    beste_flaeche = beste_homogene_variante.tafel_flaeche_mm2
    belegte_flaeche = sum(teil.flaeche_mm2 for teil in teile)
    max_tafeln = len(beste_homogene_variante.tafeln)
    kombinationen_nach_flaeche: dict[int, list[list[tuple[str, int, int]]]] = defaultdict(list)

    for tafelanzahl in range(1, max_tafeln + 1):
        for zaehlung in zaehlkombinationen(tafelanzahl, len(tafeltypen)):
            verwendete_typen = sum(1 for anzahl in zaehlung if anzahl > 0)
            if verwendete_typen < 2:
                continue
            tafelspezifikationen: list[tuple[str, int, int]] = []
            for anzahl, tafelspezifikation in zip(zaehlung, tafeltypen):
                tafelspezifikationen.extend([tafelspezifikation] * anzahl)
            tafelflaeche = sum(breite * laenge for _typ, breite, laenge in tafelspezifikationen)
            if tafelflaeche >= beste_flaeche:
                continue
            if tafelflaeche < belegte_flaeche:
                continue
            nutzflaeche = sum(
                (breite - 2 * RAND_MM) * (laenge - 2 * RAND_MM)
                for _typ, breite, laenge in tafelspezifikationen
            )
            if nutzflaeche < belegte_flaeche:
                continue

            tafelspezifikationen = sorted(
                tafelspezifikationen, key=lambda item: item[1] * item[2], reverse=True
            )
            kombinationen_nach_flaeche[tafelflaeche].append(tafelspezifikationen)

    for tafelflaeche in sorted(kombinationen_nach_flaeche):
        erfolgreiche_varianten: list[SchachtelErgebnis] = []
        kombinationen = kombinationen_nach_flaeche[tafelflaeche]
        gemischt_gesamt = max(1, len(kombinationen) * len(KOMBI_SUCHSTRATEGIEN))
        gemischt_index = 0
        for tafelspezifikationen in kombinationen:
            for sortierstrategie, platzierstrategie in KOMBI_SUCHSTRATEGIEN:
                gemischt_index += 1
                label = (
                    f"{progress_label} - gemischte Variante {gemischt_index}/{gemischt_gesamt}: "
                    f"{sortierstrategie}/{platzierstrategie}"
                ).strip(" -")
                ergebnis = schachteln_fuer_feste_tafeln(
                    staerke,
                    guete,
                    teile,
                    tafelspezifikationen,
                    sortierstrategie=sortierstrategie,
                    platzierstrategie=platzierstrategie,
                    progress_callback=progress_callback,
                    progress_label=label,
                )
                if not ergebnis.nicht_platzierbar:
                    erfolgreiche_varianten.append(ergebnis)
        if erfolgreiche_varianten:
            varianten.extend(erfolgreiche_varianten)
            break

    return varianten


def tafeltypen_text_aus_tafeln(tafeln: list[Tafel]) -> str:
    zaehler = Counter(tafel.typ for tafel in tafeln)
    teile: list[str] = []
    for typ, _breite, _laenge in TAFELTYPEN:
        anzahl = zaehler.get(typ, 0)
        if anzahl:
            teile.append(f"{anzahl} x {typ}")
    return " + ".join(teile) if teile else "keine Tafeln"


def gruppieren(teile: list[Teil]) -> dict[tuple[str, str], list[Teil]]:
    gruppen: dict[tuple[str, str], list[Teil]] = defaultdict(list)
    for teil in teile:
        gruppen[(teil.staerke, teil.guete)].append(teil)
    return gruppen


def staerke_sortierwert(staerke: str) -> tuple[float, str]:
    try:
        return float(staerke), staerke
    except ValueError:
        return math.inf, staerke


def mm2_zu_m2(wert: int | float) -> float:
    return wert / 1_000_000.0


def fmt_zahl(wert: int | float, nachkommastellen: int = 1) -> str:
    return f"{wert:.{nachkommastellen}f}".replace(".", ",")


def fmt_prozent(wert: float) -> str:
    return f"{wert:.1f} %".replace(".", ",")


def fmt_cm_von_mm(wert_mm: int | float) -> str:
    return fmt_zahl(wert_mm / 10.0, 1)


def fmt_gewicht_kg(wert: float | None) -> str:
    if wert is None:
        return "nicht berechenbar"
    if abs(wert) >= 1000:
        return f"{fmt_zahl(wert / 1000, 2)} t"
    return f"{fmt_zahl(wert, 1)} kg"


def html_esc(wert: object) -> str:
    return html.escape(str(wert), quote=True)


def dateiname_sicher(wert: str) -> str:
    text = wert.replace(",", ".").replace(" ", "_")
    text = re.sub(r"[^A-Za-z0-9_.+-]+", "_", text)
    return text.strip("_.") or "wert"


def tafeltypen_konfiguration_text() -> str:
    return ", ".join(
        f"`{typ}` ({breite} x {laenge} mm)" for typ, breite, laenge in TAFELTYPEN
    )


def tafelkonfiguration_fuer_gruppe(
    staerke: str,
    guete: str,
) -> tuple[tuple[tuple[str, int, int], ...], bool]:
    return GRUPPEN_TAFELKONFIGURATION.get((staerke, guete), (TAFELTYPEN, True))


def gruppen_tafelkonfiguration_zeilen() -> list[str]:
    zeilen: list[str] = []
    for (staerke, guete), (tafeltypen, misch_erlaubt) in sorted(
        GRUPPEN_TAFELKONFIGURATION.items(),
        key=lambda item: (staerke_sortierwert(item[0][0]), item[0][1]),
    ):
        tafeltypen_text = ", ".join(
            f"{typ} ({breite} x {laenge} mm)" for typ, breite, laenge in tafeltypen
        )
        misch_text = "ja" if misch_erlaubt and len(tafeltypen) >= 2 else "nein"
        zeilen.append(f"- BL{staerke} / {guete}: {tafeltypen_text}; Mischkombinationen: {misch_text}.")
    return zeilen


def pfad_als_html_url(pfad: Path, basis: Path) -> str:
    relativ = pfad.relative_to(basis).as_posix()
    return quote(relativ, safe="/")


def staerke_als_float(staerke: str) -> float | None:
    try:
        return float(staerke.replace(",", "."))
    except ValueError:
        return None


def gewicht_kg_fuer_flaeche_mm2(flaeche_mm2: int | float, staerke: str) -> float | None:
    staerke_mm = staerke_als_float(staerke)
    if staerke_mm is None:
        return None
    return mm2_zu_m2(flaeche_mm2) * staerke_mm * STAHL_KG_PRO_M2_UND_MM


def summe_kg(werte: list[float | None]) -> float | None:
    gesamt = 0.0
    for wert in werte:
        if wert is None:
            return None
        gesamt += wert
    return gesamt


def gewicht_kg_fuer_ergebnis(ergebnis: SchachtelErgebnis) -> float | None:
    return gewicht_kg_fuer_flaeche_mm2(ergebnis.tafel_flaeche_mm2, ergebnis.staerke)


def bestelltext_position(position: MaterialPosition) -> str:
    return (
        f"Blech BL{position.staerke} {position.guete}, "
        f"{position.breite_mm} x {position.laenge_mm} mm"
    )


def materialpositionen(ergebnisse: list[SchachtelErgebnis]) -> list[MaterialPosition]:
    positionen: list[MaterialPosition] = []
    for ergebnis in ergebnisse:
        tafeln_nach_typ: dict[tuple[str, int, int], list[Tafel]] = defaultdict(list)
        for tafel in ergebnis.tafeln:
            tafeln_nach_typ[(tafel.typ, tafel.breite_mm, tafel.laenge_mm)].append(tafel)

        for (tafeltyp, breite_mm, laenge_mm), tafeln in tafeln_nach_typ.items():
            positionen.append(
                MaterialPosition(
                    staerke=ergebnis.staerke,
                    guete=ergebnis.guete,
                    tafeltyp=tafeltyp,
                    breite_mm=breite_mm,
                    laenge_mm=laenge_mm,
                    anzahl_tafeln=len(tafeln),
                    teile=sum(len(tafel.platzierungen) for tafel in tafeln),
                    belegte_flaeche_mm2=sum(tafel.belegte_flaeche_mm2 for tafel in tafeln),
                    tafel_flaeche_mm2=sum(tafel.flaeche_mm2 for tafel in tafeln),
                    nicht_platzierbar=len(ergebnis.nicht_platzierbar),
                )
            )

    return sorted(
        positionen,
        key=lambda position: (
            staerke_sortierwert(position.staerke),
            position.guete,
            position.breite_mm,
            position.laenge_mm,
        ),
    )


def gruppiere_teile_unter_mindest(teile_unter_mindest: list[Teil]) -> list[UnterMindestGruppe]:
    gruppen: dict[tuple[str, str, str, int, int], list[Teil]] = defaultdict(list)
    for teil in teile_unter_mindest:
        gruppen[(teil.staerke, teil.guete, teil.profil, teil.laenge_mm, teil.breite_mm)].append(
            teil
        )

    ergebnis: list[UnterMindestGruppe] = []
    for (staerke, guete, profil, laenge_mm, breite_mm), teile in gruppen.items():
        flaeche_mm2 = sum(teil.flaeche_mm2 for teil in teile)
        gewicht_kg = gewicht_kg_fuer_flaeche_mm2(flaeche_mm2, staerke)
        ergebnis.append(
            UnterMindestGruppe(
                staerke=staerke,
                guete=guete,
                profil=profil,
                laenge_mm=laenge_mm,
                breite_mm=breite_mm,
                anzahl=len(teile),
                lfd_nr_liste=tuple(sorted({teil.lfd_nr for teil in teile}, key=lfd_nr_sortierschluessel)),
                flaeche_mm2=flaeche_mm2,
                gewicht_kg=gewicht_kg,
            )
        )

    return sorted(
        ergebnis,
        key=lambda gruppe: (
            staerke_sortierwert(gruppe.staerke),
            gruppe.guete,
            gruppe.profil,
            gruppe.laenge_mm,
            gruppe.breite_mm,
        ),
    )


def lfd_nr_sortierschluessel(lfd_nr: str) -> tuple[float, str]:
    try:
        return float(int(lfd_nr)), lfd_nr
    except ValueError:
        return math.inf, lfd_nr


def hinweis_fuer_ergebnis(ergebnis: SchachtelErgebnis) -> str:
    if ergebnis.nicht_platzierbar:
        return f"{len(ergebnis.nicht_platzierbar)} Teile nicht platzierbar - manuell pruefen"
    return "OK"


def hinweis_fuer_materialposition(position: MaterialPosition) -> str:
    if position.nicht_platzierbar:
        return f"{position.nicht_platzierbar} Teile nicht platzierbar - manuell pruefen"
    return "OK"


def belegung_prozent_tafel(tafel: Tafel) -> float:
    if not tafel.flaeche_mm2:
        return 0.0
    return 100.0 * tafel.belegte_flaeche_mm2 / tafel.flaeche_mm2


def farbe_fuer_platzierung(platz: Platzierung) -> str:
    schluessel = f"{platz.teil.lfd_nr}-{platz.teil.profil}"
    index = int.from_bytes(
        hashlib.blake2s(schluessel.encode("utf-8"), digest_size=4).digest(),
        "big",
    )
    return TEIL_FARBEN[index % len(TEIL_FARBEN)]


def svg_tooltip(pos: int, platz: Platzierung) -> str:
    drehung = "ja" if platz.gedreht else "nein"
    return "\n".join(
        [
            f"Position auf Tafel: {pos}",
            f"Eindeutige ID: {platz.teil.nummer}",
            f"lfd-Nr.: {platz.teil.lfd_nr}",
            f"Profil: {platz.teil.profil}",
            f"Material: BL{platz.teil.staerke} / {platz.teil.guete}",
            f"Originalgeometrie: {platz.teil.laenge_mm} x {platz.teil.breite_mm} mm",
            f"Platzierte Geometrie: {platz.laenge_mm} x {platz.breite_mm} mm",
            f"Position: X {platz.x_mm} mm / Y {platz.y_mm} mm",
            f"Gedreht: {drehung}",
        ]
    )


def svg_text_breite_schaetzen(text: str, font_size: float) -> float:
    return len(text) * font_size * 0.62


def svg_text_kuerzen(text: str, max_breite: float, font_size: float) -> str:
    if svg_text_breite_schaetzen(text, font_size) <= max_breite:
        return text
    max_zeichen = int(max_breite / max(font_size * 0.62, 1))
    if max_zeichen <= 0:
        return ""
    if max_zeichen <= 3:
        return text[:max_zeichen]
    return text[: max_zeichen - 3] + "..."


def svg_font_size_fuer_text(
    text: str,
    max_breite: float,
    max_hoehe: float,
    min_font_size: float,
    max_font_size: float,
) -> float | None:
    if not text or max_breite <= 0 or max_hoehe <= 0:
        return None

    font_size = min(max_font_size, max_hoehe, max_breite / max(len(text) * 0.62, 1))
    if font_size >= min_font_size:
        return font_size

    if max_breite >= min_font_size * 1.25 and max_hoehe >= min_font_size:
        return min_font_size
    return None


def schreibe_svg_tafel(ergebnis: SchachtelErgebnis, tafel: Tafel, pfad: Path) -> None:
    pfad.parent.mkdir(parents=True, exist_ok=True)
    titel = f"BL{ergebnis.staerke} {ergebnis.guete} - Tafel {tafel.nr} ({tafel.typ})"
    zeilen: list[str] = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        (
            f'<svg xmlns="http://www.w3.org/2000/svg" '
            f'viewBox="0 0 {tafel.laenge_mm} {tafel.breite_mm}" '
            f'class="tafel-svg" role="img" aria-label="{html_esc(titel)}">'
        ),
        f"<title>{html_esc(titel)}</title>",
        (
            f"<desc>Tafel {tafel.laenge_mm} x {tafel.breite_mm} mm, "
            f"{len(tafel.platzierungen)} Teile, Rand {RAND_MM} mm, Abstand {ABSTAND_MM} mm.</desc>"
        ),
        "<style>",
        ".tafel{fill:#fffdf7;stroke:#172033;stroke-width:14}",
        ".rand{fill:none;stroke:#7d8899;stroke-width:8;stroke-dasharray:42 34}",
        ".grid{stroke:#d9dee7;stroke-width:4}",
        ".teil{stroke:#1f2937;stroke-width:5}",
        ".teil-gruppe{cursor:pointer}",
        ".label{font-family:Arial,sans-serif;font-weight:700;fill:#111827;pointer-events:none}",
        ".id-label{font-family:Arial,sans-serif;font-weight:700;fill:#111827;opacity:.78;pointer-events:none}",
        "</style>",
        f'<rect class="tafel" x="0" y="0" width="{tafel.laenge_mm}" height="{tafel.breite_mm}"/>',
    ]

    for x_mm in range(1000, tafel.laenge_mm, 1000):
        zeilen.append(f'<line class="grid" x1="{x_mm}" y1="0" x2="{x_mm}" y2="{tafel.breite_mm}"/>')
    for y_mm in range(1000, tafel.breite_mm, 1000):
        zeilen.append(f'<line class="grid" x1="0" y1="{y_mm}" x2="{tafel.laenge_mm}" y2="{y_mm}"/>')

    zeilen.append(
        (
            f'<rect class="rand" x="{RAND_MM}" y="{RAND_MM}" '
            f'width="{tafel.laenge_mm - 2 * RAND_MM}" '
            f'height="{tafel.breite_mm - 2 * RAND_MM}"/>'
        )
    )

    for pos, platz in enumerate(tafel.platzierungen, start=1):
        farbe = farbe_fuer_platzierung(platz)
        tooltip_text = svg_tooltip(pos, platz)
        tooltip = html_esc(tooltip_text)
        tooltip_attr = tooltip.replace("\n", "&#10;")
        label = f"{platz.teil.lfd_nr}-{platz.teil.mengen_index}"
        id_label = f"ID {platz.teil.nummer}"
        clip_id = f"teil-clip-{tafel.nr}-{pos}-{platz.teil.nummer}"
        zeilen.extend(
            [
                (
                    f'<g class="teil-gruppe" data-tooltip="{tooltip_attr}" '
                    f'data-teil-id="{platz.teil.nummer}">'
                ),
                f"<title>{tooltip}</title>",
                (
                    f'<clipPath id="{clip_id}"><rect x="{platz.x_mm}" y="{platz.y_mm}" '
                    f'width="{platz.laenge_mm}" height="{platz.breite_mm}" '
                    f'rx="10" ry="10"/></clipPath>'
                ),
                (
                    f'<rect class="teil" x="{platz.x_mm}" y="{platz.y_mm}" '
                    f'width="{platz.laenge_mm}" height="{platz.breite_mm}" '
                    f'rx="10" ry="10" fill="{farbe}" fill-opacity="0.82"/>'
                ),
            ]
        )

        if platz.laenge_mm >= 120 and platz.breite_mm >= 70:
            text_x = platz.x_mm + 14
            text_y_max = platz.y_mm + platz.breite_mm - 14
            text_width = max(24.0, platz.laenge_mm - 28)
            font_size = svg_font_size_fuer_text(
                label,
                text_width,
                platz.breite_mm * 0.42,
                min_font_size=24.0,
                max_font_size=105.0,
            )
            if font_size is not None:
                sichtbares_label = svg_text_kuerzen(label, text_width, font_size)
                text_y = min(text_y_max, platz.y_mm + font_size + 18)
                if sichtbares_label:
                    zeilen.append(
                        (
                            f'<text class="label" x="{text_x}" y="{text_y}" '
                            f'font-size="{font_size:.1f}" clip-path="url(#{clip_id})">'
                            f'{html_esc(sichtbares_label)}</text>'
                        )
                    )
            if platz.breite_mm >= 150 and font_size is not None:
                id_font_size = svg_font_size_fuer_text(
                    id_label,
                    text_width,
                    font_size * 0.62,
                    min_font_size=20.0,
                    max_font_size=72.0,
                )
                if id_font_size is not None:
                    sichtbares_id_label = svg_text_kuerzen(id_label, text_width, id_font_size)
                    id_text_y = text_y + id_font_size + 14
                    if sichtbares_id_label and id_text_y < platz.y_mm + platz.breite_mm - 10:
                        zeilen.append(
                            (
                                f'<text class="id-label" x="{text_x}" y="{id_text_y}" '
                                f'font-size="{id_font_size:.1f}" clip-path="url(#{clip_id})">'
                                f'{html_esc(sichtbares_id_label)}</text>'
                            )
                        )

        zeilen.append("</g>")

    zeilen.append("</svg>")
    pfad.write_text("\n".join(zeilen) + "\n", encoding="utf-8")


def schreibe_tafelbilder(ergebnisse: list[SchachtelErgebnis]) -> list[tuple[SchachtelErgebnis, Tafel, Path]]:
    AUSGABE_BILDER_DIR.mkdir(parents=True, exist_ok=True)
    for alte_datei in AUSGABE_BILDER_DIR.glob("tafel_*.svg"):
        alte_datei.unlink()

    bilder: list[tuple[SchachtelErgebnis, Tafel, Path]] = []
    laufende_nr = 1
    for ergebnis in ergebnisse:
        for tafel in ergebnis.tafeln:
            dateiname = (
                f"tafel_{laufende_nr:03d}_"
                f"BL{dateiname_sicher(ergebnis.staerke)}_"
                f"{dateiname_sicher(ergebnis.guete)}_"
                f"T{tafel.nr:02d}.svg"
            )
            pfad = AUSGABE_BILDER_DIR / dateiname
            schreibe_svg_tafel(ergebnis, tafel, pfad)
            bilder.append((ergebnis, tafel, pfad))
            laufende_nr += 1
    return bilder


def svg_inline_fuer_html(pfad: Path) -> str:
    svg = pfad.read_text(encoding="utf-8")
    if svg.startswith("<?xml"):
        svg = svg.split("\n", 1)[1]
    return svg


def schreibe_markdown(
    ergebnisse: list[SchachtelErgebnis],
    unter_mindest_gruppen: list[UnterMindestGruppe],
) -> None:
    jetzt = datetime.now().strftime("%d.%m.%Y %H:%M")
    zeilen: list[str] = [
        f"# {PROJEKTNAME}",
        "",
        "Schachtelergebnis Platten",
        "",
        f"Erstellt am: {jetzt}",
        "",
        "## Grundlage",
        "",
        f"- Projekt: `{PROJEKTNAME}`",
        f"- Eingabedatei: `{CSV_DATEI}`",
        f"- Berücksichtigt werden nur Profile, deren Bezeichnung mit `BL` beginnt und deren Stärke mindestens {MINDEST_STAERKE_MM} mm beträgt.",
        "- Gruppierung: Blechstärke aus `Profil` und `Güte` aus dem CSV.",
        f"- Verfügbare Tafeln: {tafeltypen_konfiguration_text()}.",
        f"- Rand zur ersten Kontur: {RAND_MM} mm ({fmt_cm_von_mm(RAND_MM)} cm).",
        f"- Abstand zwischen zwei Teilen: {ABSTAND_MM} mm ({fmt_cm_von_mm(ABSTAND_MM)} cm).",
        f"- Schachtelverfahren: deterministische MaxRects-Freiraumheuristik mit {len(SUCHSTRATEGIEN)} Suchstrategien fuer homogene Tafeln und {len(KOMBI_SUCHSTRATEGIEN)} Suchstrategien fuer gemischte Tafel-Kombinationen.",
        "- Es werden auch gemischte Tafel-Kombinationen geprüft, solange sie weniger Gesamtfläche als die beste homogene Variante benötigen.",
        "- Auswahl je Gruppe: zuerst keine unplatzierbaren Teile, danach geringere Gesamtfläche, danach möglichst wenige Tafeln.",
        "",
    ]

    if GRUPPEN_TAFELKONFIGURATION:
        zeilen.extend(
            [
                "## Abweichende Tafeleinstellungen",
                "",
                *gruppen_tafelkonfiguration_zeilen(),
                "",
            ]
        )

    zeilen.extend(
        [
        "## Zusammenfassung",
        "",
        "| Stärke | Güte | Teile | Tafeltyp | Tafeln | Belegte Fläche | Tafel-Fläche | Verschnitt | Hinweis |",
        "|---:|---|---:|---|---:|---:|---:|---:|---|",
        ]
    )

    for ergebnis in ergebnisse:
        hinweis = ""
        if ergebnis.nicht_platzierbar:
            hinweis = f"{len(ergebnis.nicht_platzierbar)} Teile nicht platzierbar"
        zeilen.append(
            "| "
            + " | ".join(
                [
                    ergebnis.staerke,
                    ergebnis.guete,
                    str(ergebnis.teile_gesamt),
                    tafeltypen_text_aus_tafeln(ergebnis.tafeln),
                    str(len(ergebnis.tafeln)),
                    f"{fmt_zahl(mm2_zu_m2(ergebnis.belegte_flaeche_mm2), 2)} m²",
                    f"{fmt_zahl(mm2_zu_m2(ergebnis.tafel_flaeche_mm2), 2)} m²",
                    fmt_prozent(ergebnis.verschnitt_prozent),
                    hinweis,
                ]
            )
            + " |"
        )

    zeilen.extend(
        [
            "",
            f"## Platten unter Mindeststaerke BL{MINDEST_STAERKE_MM}",
            "",
            f"- Diese Platten werden nicht geschachtelt und nicht in der Tafel-Materialliste bestellt.",
            f"- Anzahl Einzelteile: `{sum(gruppe.anzahl for gruppe in unter_mindest_gruppen)}`",
            f"- Gruppierte Positionen: `{len(unter_mindest_gruppen)}`",
            "",
            "| Stärke | Güte | Profil | Anzahl | Länge x Breite | lfd-Nr. | Fläche | Gewicht ca. |",
            "|---:|---|---|---:|---:|---|---:|---:|",
        ]
    )
    for gruppe in unter_mindest_gruppen:
        gewicht = fmt_gewicht_kg(gruppe.gewicht_kg)
        zeilen.append(
            "| "
            + " | ".join(
                [
                    f"BL{gruppe.staerke}",
                    gruppe.guete,
                    gruppe.profil,
                    str(gruppe.anzahl),
                    f"{gruppe.laenge_mm} x {gruppe.breite_mm} mm",
                    ", ".join(gruppe.lfd_nr_liste),
                    f"{fmt_zahl(mm2_zu_m2(gruppe.flaeche_mm2), 3)} m²",
                    gewicht,
                ]
            )
            + " |"
        )

    for ergebnis in ergebnisse:
        zeilen.extend(
            [
                "",
                f"## BL{ergebnis.staerke} / {ergebnis.guete}",
                "",
                f"- Gewählte Tafeln: `{tafeltypen_text_aus_tafeln(ergebnis.tafeln)}`",
                f"- Benötigte Tafeln: `{len(ergebnis.tafeln)}`",
                f"- Teile: `{ergebnis.teile_gesamt}`",
                f"- Belegte Fläche: `{fmt_zahl(mm2_zu_m2(ergebnis.belegte_flaeche_mm2), 2)} m²`",
                f"- Verschnitt rechnerisch: `{fmt_prozent(ergebnis.verschnitt_prozent)}`",
                "",
            ]
        )

        if ergebnis.nicht_platzierbar:
            zeilen.extend(
                [
                    "### Nicht platzierbare Teile",
                    "",
                    "| lfd-Nr. | Profil | Länge x Breite | Hinweis |",
                    "|---:|---|---:|---|",
                ]
            )
            for teil in ergebnis.nicht_platzierbar:
                zeilen.append(
                    f"| {teil.lfd_nr} | {teil.profil} | {teil.laenge_mm} x {teil.breite_mm} mm | passt mit Rand nicht auf den Tafeltyp |"
                )
            zeilen.append("")

        for tafel in ergebnis.tafeln:
            zeilen.extend(
                [
                    f"### Tafel {tafel.nr} ({tafel.typ})",
                    "",
                    f"- Belegte Fläche: `{fmt_zahl(mm2_zu_m2(tafel.belegte_flaeche_mm2), 2)} m²`",
                    f"- Teile auf dieser Tafel: `{len(tafel.platzierungen)}`",
                    "",
                    "| Pos. | lfd-Nr. | Profil | Teil | Länge x Breite | X | Y | Drehung |",
                    "|---:|---:|---|---:|---:|---:|---:|---|",
                ]
            )
            for pos, platz in enumerate(tafel.platzierungen, start=1):
                drehung = "ja" if platz.gedreht else "nein"
                zeilen.append(
                    f"| {pos} | {platz.teil.lfd_nr} | {platz.teil.profil} | {platz.teil.mengen_index} | "
                    f"{platz.laenge_mm} x {platz.breite_mm} mm | {platz.x_mm} mm | {platz.y_mm} mm | {drehung} |"
                )
            zeilen.append("")

    AUSGABE_MD.write_text("\n".join(zeilen).rstrip() + "\n", encoding="utf-8")


def schreibe_csv(ergebnisse: list[SchachtelErgebnis]) -> None:
    with AUSGABE_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(
            [
                "Stärke",
                "Güte",
                "Tafel",
                "Tafeltyp",
                "lfd-Nr.",
                "Profil",
                "Teil",
                "Länge_mm",
                "Breite_mm",
                "X_mm",
                "Y_mm",
                "Gedreht",
            ]
        )
        for ergebnis in ergebnisse:
            for tafel in ergebnis.tafeln:
                for platz in tafel.platzierungen:
                    writer.writerow(
                        [
                            ergebnis.staerke,
                            ergebnis.guete,
                            tafel.nr,
                            tafel.typ,
                            platz.teil.lfd_nr,
                            platz.teil.profil,
                            platz.teil.mengen_index,
                            platz.laenge_mm,
                            platz.breite_mm,
                            platz.x_mm,
                            platz.y_mm,
                            "ja" if platz.gedreht else "nein",
                        ]
                    )


def schreibe_unter_mindest_csv(unter_mindest_gruppen: list[UnterMindestGruppe]) -> None:
    with AUSGABE_UNTER_MINDEST_CSV.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(
            [
                "Stärke",
                "Güte",
                "Profil",
                "Anzahl",
                "Länge_mm",
                "Breite_mm",
                "lfd-Nr.",
                "Fläche_m2",
                "Gewicht_ca_kg",
                "Hinweis",
            ]
        )
        for gruppe in unter_mindest_gruppen:
            writer.writerow(
                [
                    f"BL{gruppe.staerke}",
                    gruppe.guete,
                    gruppe.profil,
                    gruppe.anzahl,
                    gruppe.laenge_mm,
                    gruppe.breite_mm,
                    ", ".join(gruppe.lfd_nr_liste),
                    fmt_zahl(mm2_zu_m2(gruppe.flaeche_mm2), 3),
                    "" if gruppe.gewicht_kg is None else fmt_zahl(gruppe.gewicht_kg, 1),
                    f"Unter Mindeststaerke BL{MINDEST_STAERKE_MM}, nicht geschachtelt",
                ]
            )


def schreibe_html(
    ergebnisse: list[SchachtelErgebnis],
    bilder: list[tuple[SchachtelErgebnis, Tafel, Path]],
    unter_mindest_gruppen: list[UnterMindestGruppe],
) -> None:
    jetzt = datetime.now().strftime("%d.%m.%Y %H:%M")
    teile_gesamt = sum(ergebnis.teile_gesamt for ergebnis in ergebnisse)
    tafeln_gesamt = sum(len(ergebnis.tafeln) for ergebnis in ergebnisse)
    nicht_platzierbar = sum(len(ergebnis.nicht_platzierbar) for ergebnis in ergebnisse)
    belegte_flaeche = sum(ergebnis.belegte_flaeche_mm2 for ergebnis in ergebnisse)
    tafelflaeche = sum(ergebnis.tafel_flaeche_mm2 for ergebnis in ergebnisse)
    verschnitt = 0.0 if not tafelflaeche else 100.0 * (1.0 - belegte_flaeche / tafelflaeche)
    flaechen_nutzung = 0.0 if not tafelflaeche else 100.0 * belegte_flaeche / tafelflaeche
    restflaeche = tafelflaeche - belegte_flaeche
    csv_gewicht = summe_kg(
        [
            gewicht_kg_fuer_flaeche_mm2(ergebnis.belegte_flaeche_mm2, ergebnis.staerke)
            for ergebnis in ergebnisse
        ]
    )
    tafeln_gewicht = summe_kg([gewicht_kg_fuer_ergebnis(ergebnis) for ergebnis in ergebnisse])
    gewicht_reserve = (
        None if csv_gewicht is None or tafeln_gewicht is None else tafeln_gewicht - csv_gewicht
    )
    gewicht_nutzung = (
        None
        if csv_gewicht is None or tafeln_gewicht is None or tafeln_gewicht == 0
        else 100.0 * csv_gewicht / tafeln_gewicht
    )
    unter_mindest_anzahl = sum(gruppe.anzahl for gruppe in unter_mindest_gruppen)
    unter_mindest_flaeche = sum(gruppe.flaeche_mm2 for gruppe in unter_mindest_gruppen)
    unter_mindest_gewicht = summe_kg([gruppe.gewicht_kg for gruppe in unter_mindest_gruppen])

    csv_url = pfad_als_html_url(AUSGABE_CSV, AUSGABE_HTML.parent)
    unter_mindest_csv_url = pfad_als_html_url(AUSGABE_UNTER_MINDEST_CSV, AUSGABE_HTML.parent)
    xlsx_url = pfad_als_html_url(AUSGABE_XLSX, AUSGABE_HTML.parent)
    csv_gewicht_text = fmt_gewicht_kg(csv_gewicht)
    tafeln_gewicht_text = fmt_gewicht_kg(tafeln_gewicht)
    gewicht_reserve_text = fmt_gewicht_kg(gewicht_reserve)
    unter_mindest_gewicht_text = fmt_gewicht_kg(unter_mindest_gewicht)
    gewicht_nutzung_text = (
        "nicht berechenbar" if gewicht_nutzung is None else fmt_prozent(gewicht_nutzung)
    )

    zeilen: list[str] = [
        "<!doctype html>",
        '<html lang="de">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        f"<title>{html_esc(PROJEKTNAME)} - Kontrolle der Plattenschachtelung</title>",
        "<style>",
        ":root{color-scheme:light;--ink:#162033;--muted:#647184;--line:#d7dde7;--bg:#f7f7f3;--panel:#ffffff;--accent:#2f6f7e;--accent2:#c46a4a;--warn:#ffe7a3}",
        "*{box-sizing:border-box}",
        "body{margin:0;font-family:Inter,Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--ink);line-height:1.45}",
        "header{display:flex;justify-content:space-between;gap:24px;align-items:flex-start;padding:28px 32px 18px;border-bottom:1px solid var(--line);background:#fbfbf8;position:sticky;top:0;z-index:4}",
        "h1{font-size:28px;margin:0 0 6px}",
        "h2{font-size:20px;margin:30px 0 12px}",
        "h3{font-size:17px;margin:0}",
        ".eyebrow{margin:0 0 4px;color:var(--accent);font-weight:700;font-size:12px;text-transform:uppercase;letter-spacing:.08em}",
        ".muted{color:var(--muted)}",
        ".wrap{max-width:1480px;margin:0 auto;padding:0 24px 42px}",
        ".actions{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}",
        ".button{display:inline-flex;align-items:center;gap:8px;padding:9px 12px;border:1px solid var(--line);border-radius:6px;background:#fff;color:var(--ink);text-decoration:none;font-weight:700;font-size:14px}",
        ".button.primary{background:var(--accent);border-color:var(--accent);color:#fff}",
        ".metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:22px 0}",
        ".metric{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:14px}",
        ".metric span{display:block;color:var(--muted);font-size:12px}",
        ".metric strong{display:block;margin-top:4px;font-size:22px}",
        ".compare{display:grid;grid-template-columns:1.2fr 1fr;gap:14px;margin:18px 0 26px;align-items:start}",
        ".note-box{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px 16px}",
        ".note-box p{margin:8px 0 0}",
        ".table-wrap{overflow:auto;border:1px solid var(--line);background:#fff;border-radius:8px}",
        "table{width:100%;border-collapse:collapse;font-size:13px}",
        "th,td{padding:9px 10px;border-bottom:1px solid #e9edf3;text-align:left;white-space:nowrap}",
        "th{background:#233047;color:#fff;position:sticky;top:0;z-index:1}",
        "td.num,th.num{text-align:right}",
        "tr.warn td{background:var(--warn)}",
        ".toolbar{display:flex;justify-content:space-between;gap:12px;align-items:center;margin:28px 0 12px}",
        "#filter{width:min(520px,100%);padding:11px 12px;border:1px solid var(--line);border-radius:6px;background:#fff;font-size:15px}",
        ".sheet-list{display:grid;gap:18px}",
        ".sheet-card{background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden}",
        ".sheet-head{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;padding:14px 16px;border-bottom:1px solid var(--line);background:#fbfcfd}",
        ".sheet-meta{display:flex;gap:14px;flex-wrap:wrap;margin-top:6px;color:var(--muted);font-size:13px}",
        ".sheet-actions{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end}",
        ".image-frame{padding:12px;background:#eef2f6;overflow:auto}",
        ".image-frame svg{display:block;min-width:980px;width:100%;height:auto;background:#fff;border:1px solid #cfd6df}",
        ".hover-tip{position:fixed;z-index:20;max-width:380px;padding:10px 12px;border:1px solid #101827;border-radius:6px;background:#fff;color:#111827;box-shadow:0 10px 30px rgba(15,23,42,.22);font-size:13px;line-height:1.35;white-space:pre-line;pointer-events:none;opacity:0;transform:translateY(4px);transition:opacity .08s ease,transform .08s ease}",
        ".hover-tip.visible{opacity:1;transform:translateY(0)}",
        "details{padding:0 16px 16px}",
        "summary{cursor:pointer;padding:12px 0;font-weight:700;color:var(--accent2)}",
        ".empty{padding:18px;border:1px dashed var(--line);border-radius:8px;background:#fff}",
        "@media (max-width:900px){header{position:static;display:block;padding:22px 18px}.actions{justify-content:flex-start;margin-top:14px}.wrap{padding:0 14px 32px}.metrics{grid-template-columns:repeat(2,1fr)}.compare{grid-template-columns:1fr}.toolbar{display:block}.toolbar .muted{margin-top:8px}.sheet-head{display:block}.sheet-actions{justify-content:flex-start;margin-top:10px}.image-frame svg{min-width:780px}}",
        "</style>",
        "</head>",
        "<body>",
        "<header>",
        "<div>",
        '<p class="eyebrow">Ergebnis der Plattenschachtelung</p>',
        f"<h1>{html_esc(PROJEKTNAME)}</h1>",
        f'<div class="muted">Erstellt am {html_esc(jetzt)} aus {html_esc(CSV_DATEI.name)}</div>',
        "</div>",
        '<div class="actions">',
        f'<a class="button primary" href="{xlsx_url}">Materialliste Excel</a>',
        f'<a class="button" href="{csv_url}">CSV Platzierungen</a>',
        f'<a class="button" href="{unter_mindest_csv_url}">CSV unter Mindeststärke</a>',
        "</div>",
        "</header>",
        '<main class="wrap">',
        '<section class="metrics" aria-label="Kennzahlen">',
        f'<div class="metric"><span>BL-Teile ab {MINDEST_STAERKE_MM} mm</span><strong>{teile_gesamt}</strong></div>',
        f'<div class="metric"><span>BL-Teile unter {MINDEST_STAERKE_MM} mm</span><strong>{unter_mindest_anzahl}</strong></div>',
        f'<div class="metric"><span>Gruppen</span><strong>{len(ergebnisse)}</strong></div>',
        f'<div class="metric"><span>Tafeln</span><strong>{tafeln_gesamt}</strong></div>',
        f'<div class="metric"><span>Nicht platzierbar</span><strong>{nicht_platzierbar}</strong></div>',
        f'<div class="metric"><span>CSV-Gewicht netto</span><strong>{html_esc(csv_gewicht_text)}</strong></div>',
        f'<div class="metric"><span>Tafel-Gewicht brutto</span><strong>{html_esc(tafeln_gewicht_text)}</strong></div>',
        f'<div class="metric"><span>CSV-Fläche netto</span><strong>{html_esc(fmt_zahl(mm2_zu_m2(belegte_flaeche), 2))} m²</strong></div>',
        f'<div class="metric"><span>Tafel-Fläche brutto</span><strong>{html_esc(fmt_zahl(mm2_zu_m2(tafelflaeche), 2))} m²</strong></div>',
        f'<div class="metric"><span>Verschnitt gesamt</span><strong>{html_esc(fmt_prozent(verschnitt))}</strong></div>',
        "</section>",
        '<section class="compare">',
        '<div class="table-wrap"><table>',
        "<thead><tr><th>Vergleich</th><th class=\"num\">Fläche</th><th class=\"num\">Gewicht</th></tr></thead>",
        "<tbody>",
        (
            f'<tr><td>CSV-Liste netto (BL-Teile ab {MINDEST_STAERKE_MM} mm)</td>'
            f'<td class="num">{html_esc(fmt_zahl(mm2_zu_m2(belegte_flaeche), 2))} m²</td>'
            f'<td class="num">{html_esc(csv_gewicht_text)}</td></tr>'
        ),
        (
            f'<tr><td>Bestellte Tafeln brutto</td>'
            f'<td class="num">{html_esc(fmt_zahl(mm2_zu_m2(tafelflaeche), 2))} m²</td>'
            f'<td class="num">{html_esc(tafeln_gewicht_text)}</td></tr>'
        ),
        (
            f'<tr><td>Differenz / Reserve</td>'
            f'<td class="num">{html_esc(fmt_zahl(mm2_zu_m2(restflaeche), 2))} m²</td>'
            f'<td class="num">{html_esc(gewicht_reserve_text)}</td></tr>'
        ),
        (
            f'<tr><td>Nutzung</td>'
            f'<td class="num">{html_esc(fmt_prozent(flaechen_nutzung))}</td>'
            f'<td class="num">{html_esc(gewicht_nutzung_text)}</td></tr>'
        ),
        "</tbody></table></div>",
        '<div class="note-box">',
        "<h2>Gewicht und Fläche</h2>",
        (
            f"<p><strong>CSV-Gewicht netto</strong> ist das berechnete Gewicht der tatsächlich "
            f"aufgenommenen BL-Teile ab {MINDEST_STAERKE_MM} mm aus der CSV-Datei.</p>"
        ),
        (
            f"<p><strong>BL-Teile unter {MINDEST_STAERKE_MM} mm</strong> werden separat gelistet "
            f"und nicht in die Plattenschachtelung aufgenommen.</p>"
        ),
        (
            f"<p><strong>Tafel-Gewicht brutto</strong> ist das Gewicht der kompletten bestellten "
            f"Tafeln. Berechnungsbasis: Stahl {html_esc(fmt_zahl(STAHL_KG_PRO_M2_UND_MM, 2))} kg/m²/mm.</p>"
        ),
        "<p>Die Differenz entspricht rechnerisch der Reserve beziehungsweise dem Verschnitt der bestellten Tafeln.</p>",
        "</div>",
        "</section>",
        "<section>",
        f"<h2>Platten unter Mindeststärke BL{MINDEST_STAERKE_MM}</h2>",
        (
            f'<p class="muted">{unter_mindest_anzahl} Einzelteile in '
            f"{len(unter_mindest_gruppen)} gruppierten Positionen, "
            f"{html_esc(fmt_zahl(mm2_zu_m2(unter_mindest_flaeche), 3))} m², "
            f"{html_esc(unter_mindest_gewicht_text)}. Diese Teile werden nicht in die "
            f"Plattenschachtelung aufgenommen.</p>"
        ),
        '<div class="table-wrap"><table>',
        "<thead><tr><th>Stärke</th><th>Güte</th><th>Profil</th><th class=\"num\">Anzahl</th><th class=\"num\">Länge</th><th class=\"num\">Breite</th><th>lfd-Nr.</th><th class=\"num\">Fläche</th><th class=\"num\">Gewicht ca.</th></tr></thead>",
        "<tbody>",
    ]

    if unter_mindest_gruppen:
        for gruppe in unter_mindest_gruppen:
            zeilen.append(
                "<tr>"
                f'<td>BL{html_esc(gruppe.staerke)}</td>'
                f'<td>{html_esc(gruppe.guete)}</td>'
                f'<td>{html_esc(gruppe.profil)}</td>'
                f'<td class="num">{gruppe.anzahl}</td>'
                f'<td class="num">{gruppe.laenge_mm} mm</td>'
                f'<td class="num">{gruppe.breite_mm} mm</td>'
                f'<td>{html_esc(", ".join(gruppe.lfd_nr_liste))}</td>'
                f'<td class="num">{html_esc(fmt_zahl(mm2_zu_m2(gruppe.flaeche_mm2), 3))} m²</td>'
                f'<td class="num">{html_esc(fmt_gewicht_kg(gruppe.gewicht_kg))}</td>'
                "</tr>"
            )
    else:
        zeilen.append(
            '<tr><td colspan="9" class="muted">Keine BL-Teile unter Mindeststärke gefunden.</td></tr>'
        )

    zeilen.extend(
        [
            "</tbody></table></div>",
            "</section>",
            "<section>",
            "<h2>Materialliste Bestellung</h2>",
            '<div class="table-wrap"><table>',
            "<thead><tr><th>Pos.</th><th>Bestelltext</th><th>Stärke</th><th>Güte</th><th>Tafeltyp</th><th class=\"num\">Tafeln</th><th class=\"num\">Gesamtfläche</th><th class=\"num\">Gewicht ca.</th><th>Hinweis</th></tr></thead>",
            "<tbody>",
        ]
    )

    for pos, position in enumerate(materialpositionen(ergebnisse), start=1):
        gewicht = gewicht_kg_fuer_flaeche_mm2(position.tafel_flaeche_mm2, position.staerke)
        gewicht_text = "" if gewicht is None else f"{fmt_zahl(gewicht, 1)} kg"
        klasse = ' class="warn"' if position.nicht_platzierbar else ""
        zeilen.append(
            f"<tr{klasse}>"
            f'<td>{pos}</td>'
            f'<td>{html_esc(bestelltext_position(position))}</td>'
            f'<td>BL{html_esc(position.staerke)}</td>'
            f'<td>{html_esc(position.guete)}</td>'
            f'<td>{html_esc(position.tafeltyp)}</td>'
            f'<td class="num">{position.anzahl_tafeln}</td>'
            f'<td class="num">{html_esc(fmt_zahl(mm2_zu_m2(position.tafel_flaeche_mm2), 2))} m²</td>'
            f'<td class="num">{html_esc(gewicht_text)}</td>'
            f'<td>{html_esc(hinweis_fuer_materialposition(position))}</td>'
            "</tr>"
        )

    zeilen.extend(
        [
            "</tbody></table></div>",
            "</section>",
            "<section>",
            "<h2>Zusammenfassung Gruppen</h2>",
            '<div class="table-wrap"><table>',
            "<thead><tr><th>Stärke</th><th>Güte</th><th class=\"num\">Teile</th><th>Tafeltyp</th><th class=\"num\">Tafeln</th><th class=\"num\">Belegt</th><th class=\"num\">Tafelfläche</th><th class=\"num\">Verschnitt</th><th>Hinweis</th></tr></thead>",
            "<tbody>",
        ]
    )

    for ergebnis in ergebnisse:
        klasse = ' class="warn"' if ergebnis.nicht_platzierbar else ""
        zeilen.append(
            f"<tr{klasse}>"
            f'<td>BL{html_esc(ergebnis.staerke)}</td>'
            f'<td>{html_esc(ergebnis.guete)}</td>'
            f'<td class="num">{ergebnis.teile_gesamt}</td>'
            f'<td>{html_esc(tafeltypen_text_aus_tafeln(ergebnis.tafeln))}</td>'
            f'<td class="num">{len(ergebnis.tafeln)}</td>'
            f'<td class="num">{html_esc(fmt_zahl(mm2_zu_m2(ergebnis.belegte_flaeche_mm2), 2))} m²</td>'
            f'<td class="num">{html_esc(fmt_zahl(mm2_zu_m2(ergebnis.tafel_flaeche_mm2), 2))} m²</td>'
            f'<td class="num">{html_esc(fmt_prozent(ergebnis.verschnitt_prozent))}</td>'
            f'<td>{html_esc(hinweis_fuer_ergebnis(ergebnis))}</td>'
            "</tr>"
        )

    zeilen.extend(
        [
            "</tbody></table></div>",
            "</section>",
            '<section class="toolbar">',
            "<h2>Tafeln mit Platzierungen</h2>",
            '<input id="filter" type="search" placeholder="Suchen: BL, Güte, Tafel, lfd-Nr.">',
            "</section>",
            f'<p class="muted">{len(bilder)} SVG-Bilder in {html_esc(AUSGABE_BILDER_DIR.name)}. Jede Platte hat einen Tooltip mit Position, Maß und Drehung.</p>',
            '<section class="sheet-list" id="sheets">',
        ]
    )

    if not bilder:
        zeilen.append('<div class="empty">Keine Tafeln erzeugt.</div>')

    for ergebnis, tafel, pfad in bilder:
        bild_url = pfad_als_html_url(pfad, AUSGABE_HTML.parent)
        lfd_filter = " ".join(sorted({platz.teil.lfd_nr for platz in tafel.platzierungen}))
        plattengewicht = gewicht_kg_fuer_flaeche_mm2(tafel.belegte_flaeche_mm2, ergebnis.staerke)
        verschnittgewicht = gewicht_kg_fuer_flaeche_mm2(
            tafel.flaeche_mm2 - tafel.belegte_flaeche_mm2, ergebnis.staerke
        )
        filter_text = (
            f"BL{ergebnis.staerke} {ergebnis.guete} {tafel.typ} tafel {tafel.nr} {lfd_filter}"
        ).lower()
        alt = f"BL{ergebnis.staerke} {ergebnis.guete}, Tafel {tafel.nr}, {tafel.typ}"
        zeilen.extend(
            [
                f'<article class="sheet-card" data-filter="{html_esc(filter_text)}">',
                '<div class="sheet-head">',
                "<div>",
                f"<h3>{html_esc(alt)}</h3>",
                '<div class="sheet-meta">',
                f"<span>{tafel.laenge_mm} x {tafel.breite_mm} mm</span>",
                f"<span>{len(tafel.platzierungen)} Teile</span>",
                f"<span>{html_esc(fmt_zahl(mm2_zu_m2(tafel.belegte_flaeche_mm2), 2))} m² belegt</span>",
                f"<span>{html_esc(fmt_prozent(100.0 - belegung_prozent_tafel(tafel)))} Restfläche</span>",
                f"<span>{html_esc(fmt_gewicht_kg(plattengewicht))} Plattengewicht</span>",
                f"<span>{html_esc(fmt_gewicht_kg(verschnittgewicht))} Verschnittgewicht</span>",
                "</div>",
                "</div>",
                '<div class="sheet-actions">',
                f'<a class="button" href="{bild_url}" target="_blank" rel="noreferrer">SVG öffnen</a>',
                "</div>",
                "</div>",
                '<div class="image-frame">',
                svg_inline_fuer_html(pfad),
                "</div>",
                "<details>",
                f"<summary>Teile auf Tafel {tafel.nr}</summary>",
                '<div class="table-wrap"><table>',
                "<thead><tr><th class=\"num\">Pos.</th><th>lfd-Nr.</th><th>Profil</th><th class=\"num\">Teil</th><th class=\"num\">Länge</th><th class=\"num\">Breite</th><th class=\"num\">X</th><th class=\"num\">Y</th><th>Drehung</th></tr></thead>",
                "<tbody>",
            ]
        )
        for pos, platz in enumerate(tafel.platzierungen, start=1):
            zeilen.append(
                "<tr>"
                f'<td class="num">{pos}</td>'
                f'<td>{html_esc(platz.teil.lfd_nr)}</td>'
                f'<td>{html_esc(platz.teil.profil)}</td>'
                f'<td class="num">{platz.teil.mengen_index}</td>'
                f'<td class="num">{platz.laenge_mm} mm</td>'
                f'<td class="num">{platz.breite_mm} mm</td>'
                f'<td class="num">{platz.x_mm} mm</td>'
                f'<td class="num">{platz.y_mm} mm</td>'
                f'<td>{"ja" if platz.gedreht else "nein"}</td>'
                "</tr>"
            )
        zeilen.extend(
            [
                "</tbody></table></div>",
                "</details>",
                "</article>",
            ]
        )

    zeilen.extend(
        [
            "</section>",
            "</main>",
            "<script>",
            "const filter=document.getElementById('filter');",
            "const cards=[...document.querySelectorAll('.sheet-card')];",
            "filter?.addEventListener('input',()=>{const q=filter.value.trim().toLowerCase();cards.forEach(card=>{card.hidden=q&&!card.dataset.filter.includes(q);});});",
            "const tip=document.createElement('div');tip.className='hover-tip';document.body.appendChild(tip);",
            "let tipVisible=false;",
            "function moveTip(e){const pad=14;let x=e.clientX+pad;let y=e.clientY+pad;if(x+tip.offsetWidth>window.innerWidth-8)x=e.clientX-tip.offsetWidth-pad;if(y+tip.offsetHeight>window.innerHeight-8)y=e.clientY-tip.offsetHeight-pad;tip.style.left=Math.max(8,x)+'px';tip.style.top=Math.max(8,y)+'px';}",
            "document.addEventListener('pointerover',e=>{const plate=e.target.closest?.('.teil-gruppe');if(!plate)return;tip.textContent=plate.dataset.tooltip||'';tip.classList.add('visible');tipVisible=true;moveTip(e);});",
            "document.addEventListener('pointermove',e=>{if(tipVisible)moveTip(e);});",
            "document.addEventListener('pointerout',e=>{const plate=e.target.closest?.('.teil-gruppe');if(!plate)return;if(e.relatedTarget&&plate.contains(e.relatedTarget))return;tip.classList.remove('visible');tipVisible=false;});",
            "</script>",
            "</body>",
            "</html>",
        ]
    )

    AUSGABE_HTML.write_text("\n".join(zeilen) + "\n", encoding="utf-8")


def xml_attr(wert: object) -> str:
    return xml_escape(str(wert), {'"': "&quot;", "'": "&apos;"})


def xlsx_spaltenname(index: int) -> str:
    name = ""
    while index:
        index, rest = divmod(index - 1, 26)
        name = chr(65 + rest) + name
    return name


def xlsx_zahl(wert: int | float) -> str:
    if isinstance(wert, float):
        return f"{wert:.12g}"
    return str(wert)


def xlsx_cell(wert: object, style: int = 0) -> tuple[object, int]:
    return wert, style


def xlsx_zelle_xml(wert: object, style: int, zeile: int, spalte: int) -> str:
    ref = f"{xlsx_spaltenname(spalte)}{zeile}"
    style_attr = f' s="{style}"' if style else ""
    if wert is None or wert == "":
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(wert, (int, float)) and not isinstance(wert, bool) and math.isfinite(float(wert)):
        return f'<c r="{ref}"{style_attr}><v>{xlsx_zahl(wert)}</v></c>'
    return (
        f'<c r="{ref}" t="inlineStr"{style_attr}>'
        f"<is><t>{xml_escape(str(wert))}</t></is></c>"
    )


def xlsx_worksheet_xml(
    rows: list[list[tuple[object, int]]],
    column_widths: list[float],
    autofilter: bool = True,
) -> str:
    max_cols = max((len(row) for row in rows), default=1)
    max_rows = max(len(rows), 1)
    dimension = f"A1:{xlsx_spaltenname(max_cols)}{max_rows}"
    zeilen: list[str] = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        f'<dimension ref="{dimension}"/>',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
    ]

    if column_widths:
        zeilen.append("<cols>")
        for index, width in enumerate(column_widths, start=1):
            zeilen.append(f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>')
        zeilen.append("</cols>")

    zeilen.append("<sheetData>")
    for zeilen_nr, row in enumerate(rows, start=1):
        zeilen.append(f'<row r="{zeilen_nr}">')
        for spalten_nr, (wert, style) in enumerate(row, start=1):
            zeilen.append(xlsx_zelle_xml(wert, style, zeilen_nr, spalten_nr))
        zeilen.append("</row>")
    zeilen.append("</sheetData>")

    if autofilter and rows:
        zeilen.append(f'<autoFilter ref="A1:{xlsx_spaltenname(max_cols)}{max_rows}"/>')

    zeilen.extend(
        [
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
            "</worksheet>",
        ]
    )
    return "\n".join(zeilen)


def xlsx_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
    <numFmt numFmtId="165" formatCode="0.0%"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><color rgb="FF162033"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF233047"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFE7A3"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD7DDE7"/></left>
      <right style="thin"><color rgb="FFD7DDE7"/></right>
      <top style="thin"><color rgb="FFD7DDE7"/></top>
      <bottom style="thin"><color rgb="FFD7DDE7"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def xlsx_workbook_xml(sheet_names: list[str]) -> str:
    zeilen = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheets>",
    ]
    for index, name in enumerate(sheet_names, start=1):
        zeilen.append(f'<sheet name="{xml_attr(name)}" sheetId="{index}" r:id="rId{index}"/>')
    zeilen.extend(["</sheets>", "</workbook>"])
    return "\n".join(zeilen)


def xlsx_workbook_rels_xml(sheet_count: int) -> str:
    zeilen = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for index in range(1, sheet_count + 1):
        zeilen.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    zeilen.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    zeilen.append("</Relationships>")
    return "\n".join(zeilen)


def xlsx_content_types_xml(sheet_count: int) -> str:
    zeilen = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        zeilen.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    zeilen.append("</Types>")
    return "\n".join(zeilen)


def xlsx_root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def schreibe_materialliste_xlsx(
    ergebnisse: list[SchachtelErgebnis],
    unter_mindest_gruppen: list[UnterMindestGruppe],
) -> None:
    HEADER = 1
    INT = 2
    DEC = 3
    PCT = 4
    WARN = 5

    bestellung: list[list[tuple[object, int]]] = [
        [
            xlsx_cell("Pos.", HEADER),
            xlsx_cell("Bestelltext", HEADER),
            xlsx_cell("Staerke_mm", HEADER),
            xlsx_cell("Guete", HEADER),
            xlsx_cell("Tafeltyp", HEADER),
            xlsx_cell("Breite_mm", HEADER),
            xlsx_cell("Laenge_mm", HEADER),
            xlsx_cell("Tafeln", HEADER),
            xlsx_cell("Gesamtflaeche_m2", HEADER),
            xlsx_cell("Gewicht_ca_kg", HEADER),
            xlsx_cell("Teile", HEADER),
            xlsx_cell("Belegte_Flaeche_m2", HEADER),
            xlsx_cell("Verschnitt", HEADER),
            xlsx_cell("Hinweis", HEADER),
        ]
    ]

    for pos, position in enumerate(materialpositionen(ergebnisse), start=1):
        gewicht = gewicht_kg_fuer_flaeche_mm2(position.tafel_flaeche_mm2, position.staerke)
        staerke_mm = staerke_als_float(position.staerke)
        bestellung.append(
            [
                xlsx_cell(pos, INT),
                xlsx_cell(bestelltext_position(position)),
                xlsx_cell(staerke_mm if staerke_mm is not None else position.staerke, DEC),
                xlsx_cell(position.guete),
                xlsx_cell(position.tafeltyp),
                xlsx_cell(position.breite_mm, INT),
                xlsx_cell(position.laenge_mm, INT),
                xlsx_cell(position.anzahl_tafeln, INT),
                xlsx_cell(mm2_zu_m2(position.tafel_flaeche_mm2), DEC),
                xlsx_cell(gewicht if gewicht is not None else "", DEC),
                xlsx_cell(position.teile, INT),
                xlsx_cell(mm2_zu_m2(position.belegte_flaeche_mm2), DEC),
                xlsx_cell(position.verschnitt_prozent / 100.0, PCT),
                xlsx_cell(
                    hinweis_fuer_materialposition(position),
                    WARN if position.nicht_platzierbar else 0,
                ),
            ]
        )

    zusammenfassung: list[list[tuple[object, int]]] = [
        [
            xlsx_cell("Staerke", HEADER),
            xlsx_cell("Guete", HEADER),
            xlsx_cell("Teile", HEADER),
            xlsx_cell("Tafeltyp", HEADER),
            xlsx_cell("Tafeln", HEADER),
            xlsx_cell("Belegte_Flaeche_m2", HEADER),
            xlsx_cell("Tafelflaeche_m2", HEADER),
            xlsx_cell("Verschnitt", HEADER),
            xlsx_cell("Nicht_platzierbar", HEADER),
        ]
    ]
    for ergebnis in ergebnisse:
        zusammenfassung.append(
            [
                xlsx_cell(f"BL{ergebnis.staerke}"),
                xlsx_cell(ergebnis.guete),
                xlsx_cell(ergebnis.teile_gesamt, INT),
                xlsx_cell(tafeltypen_text_aus_tafeln(ergebnis.tafeln)),
                xlsx_cell(len(ergebnis.tafeln), INT),
                xlsx_cell(mm2_zu_m2(ergebnis.belegte_flaeche_mm2), DEC),
                xlsx_cell(mm2_zu_m2(ergebnis.tafel_flaeche_mm2), DEC),
                xlsx_cell(ergebnis.verschnitt_prozent / 100.0, PCT),
                xlsx_cell(len(ergebnis.nicht_platzierbar), INT),
            ]
        )

    platzierungen: list[list[tuple[object, int]]] = [
        [
            xlsx_cell("Staerke", HEADER),
            xlsx_cell("Guete", HEADER),
            xlsx_cell("Tafel", HEADER),
            xlsx_cell("Tafeltyp", HEADER),
            xlsx_cell("Tafel_Breite_mm", HEADER),
            xlsx_cell("Tafel_Laenge_mm", HEADER),
            xlsx_cell("Pos_auf_Tafel", HEADER),
            xlsx_cell("lfd-Nr.", HEADER),
            xlsx_cell("Profil", HEADER),
            xlsx_cell("Teil", HEADER),
            xlsx_cell("Original_Laenge_mm", HEADER),
            xlsx_cell("Original_Breite_mm", HEADER),
            xlsx_cell("Platz_Laenge_mm", HEADER),
            xlsx_cell("Platz_Breite_mm", HEADER),
            xlsx_cell("X_mm", HEADER),
            xlsx_cell("Y_mm", HEADER),
            xlsx_cell("Gedreht", HEADER),
        ]
    ]
    for ergebnis in ergebnisse:
        for tafel in ergebnis.tafeln:
            for pos, platz in enumerate(tafel.platzierungen, start=1):
                platzierungen.append(
                    [
                        xlsx_cell(f"BL{ergebnis.staerke}"),
                        xlsx_cell(ergebnis.guete),
                        xlsx_cell(tafel.nr, INT),
                        xlsx_cell(tafel.typ),
                        xlsx_cell(tafel.breite_mm, INT),
                        xlsx_cell(tafel.laenge_mm, INT),
                        xlsx_cell(pos, INT),
                        xlsx_cell(platz.teil.lfd_nr),
                        xlsx_cell(platz.teil.profil),
                        xlsx_cell(platz.teil.mengen_index, INT),
                        xlsx_cell(platz.teil.laenge_mm, INT),
                        xlsx_cell(platz.teil.breite_mm, INT),
                        xlsx_cell(platz.laenge_mm, INT),
                        xlsx_cell(platz.breite_mm, INT),
                        xlsx_cell(platz.x_mm, INT),
                        xlsx_cell(platz.y_mm, INT),
                        xlsx_cell("ja" if platz.gedreht else "nein"),
                    ]
                )

    nicht_platzierbar: list[list[tuple[object, int]]] = [
        [
            xlsx_cell("Staerke", HEADER),
            xlsx_cell("Guete", HEADER),
            xlsx_cell("lfd-Nr.", HEADER),
            xlsx_cell("Profil", HEADER),
            xlsx_cell("Teil", HEADER),
            xlsx_cell("Laenge_mm", HEADER),
            xlsx_cell("Breite_mm", HEADER),
            xlsx_cell("Hinweis", HEADER),
        ]
    ]
    for ergebnis in ergebnisse:
        for teil in ergebnis.nicht_platzierbar:
            nicht_platzierbar.append(
                [
                    xlsx_cell(f"BL{ergebnis.staerke}"),
                    xlsx_cell(ergebnis.guete),
                    xlsx_cell(teil.lfd_nr),
                    xlsx_cell(teil.profil),
                    xlsx_cell(teil.mengen_index, INT),
                    xlsx_cell(teil.laenge_mm, INT),
                    xlsx_cell(teil.breite_mm, INT),
                    xlsx_cell("Passt mit Rand nicht auf den Tafeltyp", WARN),
                ]
            )

    unter_mindest: list[list[tuple[object, int]]] = [
        [
            xlsx_cell("Staerke", HEADER),
            xlsx_cell("Staerke_mm", HEADER),
            xlsx_cell("Guete", HEADER),
            xlsx_cell("Profil", HEADER),
            xlsx_cell("Anzahl", HEADER),
            xlsx_cell("Laenge_mm", HEADER),
            xlsx_cell("Breite_mm", HEADER),
            xlsx_cell("lfd-Nr.", HEADER),
            xlsx_cell("Flaeche_m2", HEADER),
            xlsx_cell("Gewicht_ca_kg", HEADER),
            xlsx_cell("Hinweis", HEADER),
        ]
    ]
    for gruppe in unter_mindest_gruppen:
        staerke_mm = staerke_als_float(gruppe.staerke)
        unter_mindest.append(
            [
                xlsx_cell(f"BL{gruppe.staerke}"),
                xlsx_cell(staerke_mm if staerke_mm is not None else gruppe.staerke, DEC),
                xlsx_cell(gruppe.guete),
                xlsx_cell(gruppe.profil),
                xlsx_cell(gruppe.anzahl, INT),
                xlsx_cell(gruppe.laenge_mm, INT),
                xlsx_cell(gruppe.breite_mm, INT),
                xlsx_cell(", ".join(gruppe.lfd_nr_liste)),
                xlsx_cell(mm2_zu_m2(gruppe.flaeche_mm2), DEC),
                xlsx_cell(gruppe.gewicht_kg if gruppe.gewicht_kg is not None else "", DEC),
                xlsx_cell(f"Unter Mindeststaerke BL{MINDEST_STAERKE_MM}, nicht geschachtelt", WARN),
            ]
        )

    sheets = [
        ("Bestellung", bestellung, [7, 42, 11, 16, 15, 11, 11, 9, 15, 15, 10, 18, 12, 34]),
        ("Zusammenfassung", zusammenfassung, [11, 16, 10, 15, 9, 18, 16, 12, 17]),
        ("Platzierungen", platzierungen, [10, 16, 8, 15, 15, 15, 13, 12, 12, 8, 18, 18, 16, 16, 10, 10, 10]),
        ("Nicht platzierbar", nicht_platzierbar, [10, 16, 12, 12, 8, 12, 12, 34]),
        ("Unter Mindeststaerke", unter_mindest, [10, 11, 16, 12, 9, 12, 12, 22, 13, 15, 42]),
    ]

    with zipfile.ZipFile(AUSGABE_XLSX, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", xlsx_content_types_xml(len(sheets)))
        xlsx.writestr("_rels/.rels", xlsx_root_rels_xml())
        xlsx.writestr("xl/workbook.xml", xlsx_workbook_xml([name for name, _, _ in sheets]))
        xlsx.writestr("xl/_rels/workbook.xml.rels", xlsx_workbook_rels_xml(len(sheets)))
        xlsx.writestr("xl/styles.xml", xlsx_styles_xml())
        for index, (_name, rows, widths) in enumerate(sheets, start=1):
            xlsx.writestr(f"xl/worksheets/sheet{index}.xml", xlsx_worksheet_xml(rows, widths))


def konfiguriere_schachtelung(
    *,
    csv_datei: Path | str | None = None,
    ausgabe_dir: Path | str | None = None,
    projektname: str | None = None,
    tafeltypen: tuple[tuple[str, int, int], ...] | None = None,
    rand_mm: int | None = None,
    abstand_mm: int | None = None,
    mindest_staerke_mm: int | None = None,
    stahl_kg_pro_m2_und_mm: float | None = None,
    fortschritt_ausgeben: bool | None = None,
    gruppen_tafelkonfiguration: GruppenTafelKonfiguration | None = None,
) -> None:
    global PROJEKTNAME
    global CSV_DATEI
    global AUSGABE_MD
    global AUSGABE_CSV
    global AUSGABE_UNTER_MINDEST_CSV
    global AUSGABE_HTML
    global AUSGABE_XLSX
    global AUSGABE_BILDER_DIR
    global TAFELTYPEN
    global RAND_MM
    global ABSTAND_MM
    global MINDEST_STAERKE_MM
    global STAHL_KG_PRO_M2_UND_MM
    global FORTSCHRITT_AUSGEBEN
    global GRUPPEN_TAFELKONFIGURATION

    if csv_datei is not None:
        CSV_DATEI = Path(csv_datei)

    if ausgabe_dir is not None:
        ziel = Path(ausgabe_dir)
        ziel.mkdir(parents=True, exist_ok=True)
        AUSGABE_MD = ziel / "Schachtel Ergebnis.md"
        AUSGABE_CSV = ziel / "Schachtel Ergebnis.csv"
        AUSGABE_UNTER_MINDEST_CSV = ziel / "Platten unter Mindeststaerke.csv"
        AUSGABE_HTML = ziel / "Schachtel Kontrolle.html"
        AUSGABE_XLSX = ziel / "Materialliste Bestellung.xlsx"
        AUSGABE_BILDER_DIR = ziel / "Schachtel Bilder"

    if projektname is not None:
        PROJEKTNAME = projektname.strip() or "Plattenschachtelung"

    if tafeltypen is not None:
        if not tafeltypen:
            raise ValueError("Mindestens ein Tafeltyp ist erforderlich.")
        TAFELTYPEN = tuple((name, int(breite), int(laenge)) for name, breite, laenge in tafeltypen)

    if rand_mm is not None:
        RAND_MM = int(rand_mm)
    if abstand_mm is not None:
        ABSTAND_MM = int(abstand_mm)
    if mindest_staerke_mm is not None:
        MINDEST_STAERKE_MM = int(mindest_staerke_mm)
    if stahl_kg_pro_m2_und_mm is not None:
        STAHL_KG_PRO_M2_UND_MM = float(stahl_kg_pro_m2_und_mm)
    if fortschritt_ausgeben is not None:
        FORTSCHRITT_AUSGEBEN = bool(fortschritt_ausgeben)
    if gruppen_tafelkonfiguration is not None:
        GRUPPEN_TAFELKONFIGURATION = {
            (staerke, guete): (tuple(tafeltypen), bool(misch_erlaubt))
            for (staerke, guete), (tafeltypen, misch_erlaubt) in gruppen_tafelkonfiguration.items()
        }


def berechne_schachtelung(
    progress_callback: FortschrittCallback | None = None,
) -> LaufErgebnis:
    if not CSV_DATEI.exists():
        raise FileNotFoundError(f"Eingabedatei nicht gefunden: {CSV_DATEI}")

    def melde_fortschritt(phase: str, percent: float, message: str) -> None:
        if progress_callback is None:
            return
        progress_callback(
            {
                "phase": phase,
                "percent": max(0.0, min(percent, 1.0)),
                "message": message,
            }
        )

    melde_fortschritt("vorbereitung", 0.01, "CSV-Datei wird gelesen...")
    teile, teile_unter_mindest = teile_laden()
    unter_mindest_gruppen = gruppiere_teile_unter_mindest(teile_unter_mindest)
    gruppen = gruppieren(teile)

    ergebnisse: list[SchachtelErgebnis] = []
    gruppen_sortiert = sorted(gruppen, key=lambda item: (staerke_sortierwert(item[0]), item[1]))
    basis_schritte = 0
    for staerke, guete in gruppen_sortiert:
        tafeltypen_gruppe, _misch_erlaubt = tafelkonfiguration_fuer_gruppe(staerke, guete)
        basis_schritte += len(gruppen[(staerke, guete)]) * len(tafeltypen_gruppe) * len(SUCHSTRATEGIEN)
    geschaetzte_schritte = max(1, int(basis_schritte * 1.25))
    erledigte_schritte = 0

    def melde_arbeit(event: dict[str, object]) -> None:
        nonlocal erledigte_schritte
        inkrement = int(event.get("increment", 0) or 0)
        if inkrement > 0:
            erledigte_schritte += inkrement
        berechnungsfortschritt = min(erledigte_schritte / geschaetzte_schritte, 1.0)
        percent = 0.02 + berechnungsfortschritt * 0.88
        message = str(event.get("message") or "Schachtelung wird berechnet...")
        melde_fortschritt("berechnung", percent, message)

    melde_fortschritt("berechnung", 0.02, "Gruppen werden vorbereitet...")
    for index, (staerke, guete) in enumerate(gruppen_sortiert, start=1):
        gruppen_label = (
            f"Gruppe {index}/{len(gruppen_sortiert)}: "
            f"BL{staerke} / {guete} ({len(gruppen[(staerke, guete)])} Teile)"
        )
        tafeltypen_gruppe, misch_erlaubt = tafelkonfiguration_fuer_gruppe(staerke, guete)
        aktueller_percent = 0.02 + min(erledigte_schritte / geschaetzte_schritte, 1.0) * 0.88
        melde_fortschritt("berechnung", aktueller_percent, f"Starte {gruppen_label}")
        if FORTSCHRITT_AUSGEBEN:
            print(
                f"Berechne Gruppe {index}/{len(gruppen_sortiert)}: "
                f"BL{staerke} / {guete} ({len(gruppen[(staerke, guete)])} Teile)",
                flush=True,
            )
        ergebnisse.append(
            bestes_ergebnis(
                staerke,
                guete,
                gruppen[(staerke, guete)],
                tafeltypen=tafeltypen_gruppe,
                misch_kombinationen_erlaubt=misch_erlaubt,
                progress_callback=melde_arbeit,
                progress_label=gruppen_label,
            )
        )

    melde_fortschritt("ausgabe", 0.91, "Markdown wird geschrieben...")
    schreibe_markdown(ergebnisse, unter_mindest_gruppen)
    melde_fortschritt("ausgabe", 0.93, "CSV-Dateien werden geschrieben...")
    schreibe_csv(ergebnisse)
    schreibe_unter_mindest_csv(unter_mindest_gruppen)
    melde_fortschritt("ausgabe", 0.95, "Excel-Materialliste wird geschrieben...")
    schreibe_materialliste_xlsx(ergebnisse, unter_mindest_gruppen)
    melde_fortschritt("ausgabe", 0.97, "Tafelbilder werden erzeugt...")
    bilder = schreibe_tafelbilder(ergebnisse)
    melde_fortschritt("ausgabe", 0.99, "Kontroll-HTML wird geschrieben...")
    schreibe_html(ergebnisse, bilder, unter_mindest_gruppen)
    melde_fortschritt("fertig", 1.0, "Schachtelung abgeschlossen.")

    tafeln_gesamt = sum(len(ergebnis.tafeln) for ergebnis in ergebnisse)
    nicht_platzierbar = sum(len(ergebnis.nicht_platzierbar) for ergebnis in ergebnisse)
    ausgabe_dateien = tuple(
        pfad
        for pfad in (
            AUSGABE_MD,
            AUSGABE_CSV,
            AUSGABE_UNTER_MINDEST_CSV,
            AUSGABE_XLSX,
            AUSGABE_HTML,
        )
        if pfad.exists()
    )
    return LaufErgebnis(
        projektname=PROJEKTNAME,
        teile_ab_mindest=len(teile),
        teile_unter_mindest=len(teile_unter_mindest),
        gruppen=len(ergebnisse),
        tafeln_gesamt=tafeln_gesamt,
        nicht_platzierbar=nicht_platzierbar,
        ausgabe_dateien=ausgabe_dateien,
        bilder_dir=AUSGABE_BILDER_DIR,
        bilder_anzahl=len(bilder),
    )


def schachtelung_ausfuehren(
    *,
    csv_datei: Path | str | None = None,
    ausgabe_dir: Path | str | None = None,
    projektname: str | None = None,
    tafeltypen: tuple[tuple[str, int, int], ...] | None = None,
    rand_mm: int | None = None,
    abstand_mm: int | None = None,
    mindest_staerke_mm: int | None = None,
    stahl_kg_pro_m2_und_mm: float | None = None,
    fortschritt_ausgeben: bool = False,
    progress_callback: FortschrittCallback | None = None,
    gruppen_tafelkonfiguration: GruppenTafelKonfiguration | None = None,
) -> LaufErgebnis:
    konfiguriere_schachtelung(
        csv_datei=csv_datei,
        ausgabe_dir=ausgabe_dir,
        projektname=projektname,
        tafeltypen=tafeltypen,
        rand_mm=rand_mm,
        abstand_mm=abstand_mm,
        mindest_staerke_mm=mindest_staerke_mm,
        stahl_kg_pro_m2_und_mm=stahl_kg_pro_m2_und_mm,
        fortschritt_ausgeben=fortschritt_ausgeben,
        gruppen_tafelkonfiguration=gruppen_tafelkonfiguration,
    )
    return berechne_schachtelung(progress_callback=progress_callback)


def main() -> None:
    if not eingabedatei_vorhanden():
        return

    ergebnis = berechne_schachtelung()
    print("Schachteln abgeschlossen.")
    print(f"BL-Teile ab {MINDEST_STAERKE_MM} mm: {ergebnis.teile_ab_mindest}")
    print(f"BL-Teile unter {MINDEST_STAERKE_MM} mm: {ergebnis.teile_unter_mindest}")
    print(f"Gruppen: {ergebnis.gruppen}")
    print(f"Tafeln gesamt: {ergebnis.tafeln_gesamt}")
    print(f"Nicht platzierbare Teile: {ergebnis.nicht_platzierbar}")
    print(f"Markdown: {AUSGABE_MD}")
    print(f"CSV: {AUSGABE_CSV}")
    print(f"CSV unter Mindeststaerke: {AUSGABE_UNTER_MINDEST_CSV}")
    print(f"Materialliste Excel: {AUSGABE_XLSX}")
    print(f"Kontroll-HTML: {AUSGABE_HTML}")
    print(f"Tafelbilder: {AUSGABE_BILDER_DIR}")


if __name__ == "__main__":
    main()
