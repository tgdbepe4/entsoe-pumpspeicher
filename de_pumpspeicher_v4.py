#!/usr/bin/env python3
"""
Deutschland – Pumpspeicher: Erzeugung UND Pumpverbrauch, via ENTSO-E
================================================================================
Version: v4.1 - Spaltenbreite timestamp (A) automatisch gesetzt, kein manuelles Verbreitern mehr noetig

Schwester-Skript zu at_pumpspeicher.py / ch_pumpspeicher.py, hier für
Deutschland (DE). Fokus: Wie viel Strom wird zum Hochpumpen verbraucht
(direction "out"), nicht nur wie viel beim Turbinieren erzeugt
(direction "in")?

WARUM DEUTSCHLAND: Wie Österreich ist Deutschland EU-Mitglied und fällt
unter die verbindliche EU-Transparenzverordnung 543/2013 - die
Datenabdeckung auf der ENTSO-E Transparency Platform sollte deshalb
ähnlich vollständig sein wie bei AT, und im Gegensatz zur Schweiz auch
die "out"-Zeitreihe (Pumpverbrauch) enthalten.

EINSCHRAENKUNG: Deutschland hat mehrere TSO-Regelzonen (50Hertz,
Amprion, TenneT DE, TransnetBW). Der hier verwendete Domain-Code "DE"
sollte bei entsoe-py auf die gesamtdeutsche Bidding Zone (DE-LU oder
DE/AT/LU je nach Zeitraum) abgebildet werden - bei Bedarf mit --raw
prüfen, ob die Anlagen, die dich interessieren (z.B. Goldisthal,
Markersbach), tatsächlich enthalten sind.

KEINE AKW-LOGIK: Diese Variante enthält bewusst keine Pro-Block-AKW-
Auswertung - reiner Fokus auf Pumpspeicher, zum direkten Vergleich mit
AT/CH. (Deutschland hat ohnehin keine in Betrieb befindlichen AKW mehr.)

Drei Modi:
  (Standard)   Live-Snapshot in der Konsole (letzter verfügbarer Wert).
  --week       Holt die letzten N Tage (Default 7) und schreibt ein Excel-
               File mit Turbiniert_MW und Hochgepumpt_MW pro Stunde.
  --raw        Rohe XML-Antwort statt Parsing zeigen (Debug, nur Live-Modus).

Datenquelle:
  - documentType A75 "Actual Generation per Production Type", PSR-Type B10
    (Hydro Pumped Storage), unterschieden über inBiddingZone_Domain
    (Erzeugung) / outBiddingZone_Domain (Pumpverbrauch).

Installation:
    pip install entsoe-py pandas openpyxl --break-system-packages

HINWEIS: Ungetestet (kein eigener Token verfügbar). Bei Problemen zuerst
mit --raw die Rohantwort ansehen.
"""

import os
from dotenv import load_dotenv
load_dotenv()  # liest .env im aktuellen Verzeichnis
import re
import sys
import time
import argparse
import xml.etree.ElementTree as ET

import requests
import pandas as pd
from entsoe import EntsoeRawClient
from entsoe.exceptions import NoMatchingDataError
from openpyxl.chart import LineChart, Reference
from boxwhisker_injector import inject_box_whisker_chart

PSR_PUMPED_STORAGE = "B10"

# =============================================================================
# ENTSO-E API-TOKEN - hier eintragen (gleicher Token wie im CH-Skript moeglich,
# es ist derselbe Account/dieselbe Plattform).
#
# WARNUNG: Echtes, funktionsfaehiges Credential im Klartext. Vor einem Git-
# Commit unbedingt entfernen oder .gitignore verwenden.
# =============================================================================
ENTSOE_TOKEN_HARDCODED = ""  # Token wird aus .env-Datei gelesen (ENTSOE_TOKEN=...)


def strip_ns(tag: str) -> str:
    """Entfernt das XML-Namespace-Präfix '{...}' von einem Tag-Namen."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def fix_mojibake(text: str) -> str:
    """entsoe-py/requests dekodiert die HTTP-Antwort manchmal mit der
    falschen Zeichenkodierung, wenn der Server keinen expliziten
    charset=utf-8-Header schickt. Fix: als Latin-1 zurueck in Bytes
    verwandeln, dann als UTF-8 lesen."""
    try:
        return text.encode("latin-1").decode("utf-8")
    except (UnicodeDecodeError, UnicodeEncodeError):
        return text


def iso8601_duration_to_timedelta(duration: str) -> pd.Timedelta:
    """Parst einfache ISO8601-Dauern wie 'PT15M', 'PT60M', 'P1D'."""
    m = re.fullmatch(r"P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?", duration or "")
    if not m or not any(m.groups()):
        raise ValueError(f"Unbekanntes resolution-Format: {duration!r}")
    days, hours, minutes = (int(g) if g else 0 for g in m.groups())
    return pd.Timedelta(days=days, hours=hours, minutes=minutes)


def parse_generation_timeseries(xml_text: str) -> pd.DataFrame:
    """Parst ALLE Punkte aus einer ENTSO-E-Antwort. Liefert Long-Format:
    timestamp | psr_type | direction | mw

    direction: "in" = Erzeugung (Turbine), "out" = Verbrauch (Pumpe) -
    bestimmt über das vorhandene Domain-Element (inBiddingZone_Domain vs.
    outBiddingZone_Domain), NICHT über businessType (siehe Chat-Diskussion
    zum CH-Skript, wo sich die businessType-Annahme als falsch erwies)."""
    xml_text = fix_mojibake(xml_text)
    root = ET.fromstring(xml_text)
    rows = []

    for ts in root:
        if strip_ns(ts.tag) != "TimeSeries":
            continue

        psr_type = direction = None
        for el in ts:
            tag = strip_ns(el.tag)
            if tag == "inBiddingZone_Domain.mRID":
                direction = "in"
            elif tag == "outBiddingZone_Domain.mRID":
                direction = "out"
            elif tag == "MktPSRType":
                for sub in el.iter():
                    if strip_ns(sub.tag) == "psrType":
                        psr_type = sub.text

        for period in ts:
            if strip_ns(period.tag) != "Period":
                continue

            start_text = resolution_text = None
            for el in period.iter():
                tag = strip_ns(el.tag)
                if tag == "start" and start_text is None:
                    start_text = el.text
                elif tag == "resolution" and resolution_text is None:
                    resolution_text = el.text

            if start_text is None or resolution_text is None:
                continue

            period_start = pd.Timestamp(start_text)
            delta = iso8601_duration_to_timedelta(resolution_text)

            for point in period:
                if strip_ns(point.tag) != "Point":
                    continue
                position = quantity = None
                for c in point:
                    ctag = strip_ns(c.tag)
                    if ctag == "position":
                        position = int(c.text)
                    elif ctag == "quantity":
                        quantity = float(c.text)
                if position is None or quantity is None:
                    continue

                rows.append({
                    "timestamp": period_start + (position - 1) * delta,
                    "psr_type": psr_type or "?",
                    "direction": direction,
                    "mw": quantity,
                })

    return pd.DataFrame(rows)


def fetch_with_retry(query_fn, day_start, day_end, retries: int = 3, base_delay: int = 3):
    """Ruft query_fn(day_start, day_end) auf. Bei transienten HTTP-Fehlern
    (401/429/5xx) wird mit steigender Pause erneut versucht. Bei
    NoMatchingDataError (legitime "keine Daten"-Antwort) sofort weiterreichen."""
    for attempt in range(retries):
        try:
            return query_fn(day_start, day_end)
        except NoMatchingDataError:
            raise
        except requests.exceptions.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else None
            if status in (401, 429, 500, 502, 503, 504) and attempt < retries - 1:
                wait = base_delay * (attempt + 1)
                print(f"    HTTP {status} - warte {wait}s und versuche erneut "
                      f"(Versuch {attempt + 2}/{retries})...")
                time.sleep(wait)
                continue
            raise


def fetch_days(query_fn, days: int) -> pd.DataFrame:
    """Ruft query_fn(start, end) tagesweise auf (UTC-Tagesgrenzen) und
    haengt die geparsten Resultate zusammen. Kurze Pause zwischen den
    Tagen, um Rate-Limiting vorzubeugen."""
    end_day = pd.Timestamp.now(tz="UTC").normalize() + pd.Timedelta(days=1)
    frames = []
    for i in range(days):
        day_end = end_day - i * pd.Timedelta(days=1)
        day_start = day_end - pd.Timedelta(days=1)
        try:
            xml_text = fetch_with_retry(query_fn, day_start, day_end)
        except NoMatchingDataError:
            print(f"  Hinweis: keine Daten fuer {day_start.date()} - uebersprungen")
            time.sleep(1)
            continue
        except Exception as exc:
            print(f"  Warnung: Anfrage fuer {day_start.date()} fehlgeschlagen ({exc}) - uebersprungen")
            time.sleep(1)
            continue
        df = parse_generation_timeseries(xml_text)
        if not df.empty:
            frames.append(df)
        time.sleep(3)
    if not frames:
        return pd.DataFrame()
    result = pd.concat(frames, ignore_index=True).sort_values("timestamp")
    # Excel kann keine zeitzonen-bewussten Datetimes - nach Europe/Berlin
    # (lokale Zeit) konvertieren und dann die TZ-Info entfernen.
    result["timestamp"] = (
        result["timestamp"].dt.tz_convert("Europe/Berlin").dt.tz_localize(None)
    )
    return result


def build_pumped_wide(df: pd.DataFrame) -> pd.DataFrame:
    """Pivotiert auf Turbiniert_MW / Hochgepumpt_MW. Eine Pumpspeicher-
    anlage ist zu einem Zeitpunkt entweder am Turbinieren oder am Pumpen,
    nie beides gleichzeitig - die jeweils inaktive Seite liefert deshalb
    KEINEN Messwert fuer diesen Zeitstempel (kein Messfehler, sondern
    "Modus nicht aktiv"). Wir ersetzen das deshalb bewusst durch 0 statt
    NaN, damit die Tabelle sich intuitiv liest ("0 MW Pumpverbrauch"
    statt eine vermeintliche Datenluecke)."""
    if df.empty:
        return pd.DataFrame()
    df = df[df["direction"].isin(["in", "out"])]
    if df.empty:
        return pd.DataFrame()
    wide = df.pivot_table(index="timestamp", columns="direction", values="mw", aggfunc="mean")
    wide = wide.rename(columns={
        "in": "Turbiniert_MW",
        "out": "Hochgepumpt_MW",
    })
    # Beide Spalten MUESSEN existieren (auch wenn z.B. bei der Schweiz nie
    # "out"-Daten/Pumpverbrauch gemeldet werden) - sonst erwartet der
    # Box-Whisker-Chart-Injektor (fixe Spalten A=Zeit, B=Turbiniert,
    # C=Hochgepumpt) eine Spalte, die gar nicht da ist. Eine durchgehend
    # 0-Spalte macht zudem explizit sichtbar "hier wird nichts gemeldet"
    # statt die Spalte stillschweigend wegzulassen.
    for col in ("Turbiniert_MW", "Hochgepumpt_MW"):
        if col not in wide.columns:
            wide[col] = 0.0
        wide[col] = wide[col].fillna(0.0)
    wide = wide[["Turbiniert_MW", "Hochgepumpt_MW"]]  # feste Spaltenreihenfolge B, C
    return wide.reset_index()


def _set_x_axis_label_rotation(chart, angle_deg: float = -45) -> None:
    """Dreht die Beschriftungen der X-Achse schraeg, damit sie bei vielen
    Zeitpunkten nicht uebereinander liegen. Rein kosmetisch - falls sich
    die openpyxl-API mal aendert, lieber ohne Rotation weiterfahren als
    das ganze Skript abstuerzen zu lassen."""
    try:
        from openpyxl.chart.text import RichText
        from openpyxl.drawing.text import (
            RichTextProperties, Paragraph, ParagraphProperties, CharacterProperties
        )
        rotation_units = int(angle_deg * 60000)  # Grad -> 1/60000-Grad-Einheiten
        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(rot=rotation_units, vert="horz"),
            p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties()),
                         endParaRPr=CharacterProperties())],
        )
    except Exception:
        pass


def add_line_chart(writer, sheet_name: str, df: pd.DataFrame) -> None:
    """Fuegt dem Excel-Sheet ein eingebettetes Liniendiagramm hinzu, das
    Turbiniert_MW und Hochgepumpt_MW ueber die Zeit darstellt. Wird nach
    df.to_excel(...) aufgerufen, da das Diagramm sich auf bereits
    geschriebene Zellen bezieht (Spalte A = timestamp, B/C = Werte)."""
    if df.empty or "timestamp" not in df.columns:
        return
    ws = writer.sheets[sheet_name]
    n_rows = len(df)

    value_cols = [c for c in df.columns if c != "timestamp"]
    if not value_cols:
        return

    chart = LineChart()
    chart.title = f"{sheet_name}: Turbiniert vs. Hochgepumpt (MW)"
    chart.y_axis.title = "MW"
    chart.x_axis.title = "Zeit"
    chart.width = 28
    chart.height = 12

    # Spaltenindizes (1-basiert, Excel-Style): timestamp ist immer Spalte 1
    for col in value_cols:
        col_idx = df.columns.get_loc(col) + 1  # 1-basiert
        data_ref = Reference(ws, min_col=col_idx, min_row=1, max_row=n_rows + 1)
        chart.add_data(data_ref, titles_from_data=True)

    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=n_rows + 1)
    chart.set_categories(cats_ref)

    # Lesbarkeit der Zeit-Achse: Bei 15-Min-Aufloesung (AT/DE) oder auch
    # stuendlich ueber eine Woche (CH) waeren sonst 100+ Beschriftungen
    # uebereinander - nur jedes n-te Label anzeigen (Ziel: ~20 sichtbare
    # Labels) und schraeg drehen, damit sie nicht ueberlappen.
    chart.x_axis.tickLblSkip = max(1, n_rows // 20)
    chart.x_axis.number_format = "dd.mm HH:mm"
    _set_x_axis_label_rotation(chart, angle_deg=-45)

    # Diagramm rechts neben den Datenspalten platzieren
    anchor_col_letter = chr(ord("A") + len(df.columns) + 1)
    ws.add_chart(chart, f"{anchor_col_letter}2")


def main():
    parser = argparse.ArgumentParser(
        description="DE Pumpspeicher Erzeugung + Pumpverbrauch via entsoe-py")
    parser.add_argument("--raw", action="store_true",
                         help="Rohe XML-Antwort zeigen (nur Live-Modus)")
    parser.add_argument("--week", action="store_true",
                         help="Letzte N Tage als Excel exportieren")
    parser.add_argument("--days", type=int, default=7,
                         help="Anzahl Tage fuer --week (Default 7)")
    parser.add_argument("--out", default=None,
                         help="Output-Pfad fuer --week (Default: "
                              "de_pumpspeicher_woche_JJJJMMTT_HHMM.xlsx)")
    args = parser.parse_args()

    if args.out is None:
        zeitstempel = pd.Timestamp.now(tz="Europe/Berlin").strftime("%Y%m%d_%H%M")
        args.out = f"de_pumpspeicher_woche_{zeitstempel}.xlsx"

    token = os.environ.get("ENTSOE_TOKEN", "") or ENTSOE_TOKEN_HARDCODED
    if not token or token == "dein-token-hier-eintragen":
        sys.exit("Bitte ENTSOE_TOKEN_HARDCODED oben im Skript eintragen "
                  "(oder die Umgebungsvariable ENTSOE_TOKEN setzen).")

    client = EntsoeRawClient(api_key=token)

    if args.week:
        print(f"Hole {args.days} Tage Daten für Deutschland (tagesweise, dauert evtl. etwas)...")

        pumped_df = fetch_days(
            lambda s, e: client.query_generation(
                country_code="DE", start=s, end=e, psr_type=PSR_PUMPED_STORAGE),
            args.days,
        )

        pumped_wide = build_pumped_wide(pumped_df)

        with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
            (pumped_wide if not pumped_wide.empty
             else pd.DataFrame({"Hinweis": ["Keine Daten erhalten - mit --raw pruefen"]})
             ).to_excel(writer, sheet_name="Pumpspeicher_DE", index=False)
            writer.sheets["Pumpspeicher_DE"].column_dimensions["A"].width = 20

        # Box-Whisker-Diagramm wird NACH dem Speichern eingefuegt, da die
        # Injektion die bereits geschriebene xlsx-Datei als ZIP neu oeffnet.
        if not pumped_wide.empty:
            datum_von = pumped_wide["timestamp"].min().strftime("%d.%m.%Y")
            datum_bis = pumped_wide["timestamp"].max().strftime("%d.%m.%Y")
            chart_title = (
                f"Pumpspeicher Deutschland ({datum_von} – {datum_bis}): "
                f"Turbiniert vs. Hochgepumpt (MW)"
            )
            inject_box_whisker_chart(
                args.out,
                sheet_name="Pumpspeicher_DE",
                n_data_rows=len(pumped_wide),
                chart_title=chart_title,
            )

        print(f"Fertig: {args.out}")
        print(f"  Pumpspeicher-Zeilen: {len(pumped_wide)}")
        if not pumped_wide.empty:
            hat_pumpverbrauch = "Hochgepumpt_MW" in pumped_wide.columns
            print(f"  Pumpverbrauch-Spalte vorhanden: {'JA' if hat_pumpverbrauch else 'NEIN'}")
        return

    # --- Live-Snapshot -------------------------------------------------------
    now = pd.Timestamp.now(tz="UTC")
    start = now - pd.Timedelta(hours=3)
    end = now + pd.Timedelta(hours=1)

    try:
        xml_pumped = client.query_generation(
            country_code="DE", start=start, end=end, psr_type=PSR_PUMPED_STORAGE)
    except Exception as exc:
        sys.exit(f"Anfrage an ENTSO-E fehlgeschlagen: {exc}")

    if args.raw:
        print(xml_pumped)
        return

    pumped_df = parse_generation_timeseries(xml_pumped)

    print(f"Pumpspeicher Deutschland (Stand ca. {now.strftime('%d.%m.%Y %H:%M UTC')}):\n")
    if pumped_df.empty:
        print("  Keine Daten erhalten - mit --raw die Rohantwort pruefen.")
        return

    latest = pumped_df.sort_values("timestamp").groupby("direction").tail(1)
    for _, row in latest.iterrows():
        label = {
            "in": "Erzeugung (Turbine läuft, Strom raus)",
            "out": "Pumpverbrauch (Pumpe läuft, Strom rein)",
        }.get(row["direction"], str(row["direction"]))
        sign = "-" if row["direction"] == "out" else ""
        print(f"  {label:42s} {sign}{row['mw']:7.1f} MW")


if __name__ == "__main__":
    main()
