#!/usr/bin/env python3
"""
Schweizer Strom – AKW pro Block + Pumpspeicher, mit Wochen-Export nach Excel
================================================================================
Version: v5.1 - Box-Whisker-Diagramme fuer AKW_Grafik- und Pumpspeicher-Sheet; timestamp-Spaltenbreite automatisch

Drei Modi:
  (Standard)   Live-Snapshot in der Konsole (letzter verfügbarer Wert).
  --week       Holt die letzten N Tage (Default 7) und schreibt ein Excel-
               File mit zwei Sheets: "AKW" (Leistung pro Block + eine
               errechnete Gösgen-Spalte, siehe unten) und "Pumpspeicher"
               (Erzeugung vs. Pumpverbrauch).
  --raw        Rohe XML-Antworten statt Parsing zeigen (Debug, nur Live-Modus).

Datenquellen:
  - documentType A73 "Actual Generation per Generation Unit"  -> AKW pro Block
  - documentType A75 "Actual Generation per Production Type"  -> Pumpspeicher
    (PSR-Type B10) UND Nuklear-Gesamtsumme (PSR-Type B14), unterschieden
    über inBiddingZone_Domain (Erzeugung) / outBiddingZone_Domain
    (Verbrauch/Pumpen).

GÖSGEN-BESONDERHEIT: Die A73-Pro-Block-Meldung für CH liefert für Gösgen
fast nie Daten (siehe Chat-Diskussion - vermutlich ein meldetechnisches
Loch nur für diese Anlage), während die A75-Nuklear-Gesamtsumme für CH
Gösgen offenbar korrekt mit einschliesst. Das Skript errechnet deshalb
zusätzlich:
    Goesgen_errechnet_MW = Total_Nuklear_A75 - Beznau1 - Beznau2 - Leibstadt
Das ist eine ABLEITUNG, kein direkt gemeldeter Wert - und wird bewusst zu
NaN, sobald auch nur einer der anderen drei Blöcke für den Zeitpunkt fehlt
(sonst wäre die Rechnung falsch, nicht nur ungenau).

Für --week wird TAGEWEISE abgefragt statt eines einzigen mehrtägigen Calls,
weil unklar ist, ob/wie entsoe-py bzw. die API mehrtägige Anfragen für diese
Dokumenttypen intern behandeln - mit Tagesabfragen + eigenem Zusammenführen
entfällt dieses Risiko. Kostet ein paar API-Calls mehr (z.B. 14 für eine
Woche: 7 Tage x 2 Abfragen), das ENTSO-E-Ratenlimit (mehrere hundert
Calls/Minute) ist damit nicht ansatzweise ausgereizt.

Installation:
    pip install entsoe-py pandas openpyxl --break-system-packages

Voraussetzung: ENTSO-E API-Token als Umgebungsvariable:
    export ENTSOE_TOKEN="dein-token"

HINWEIS: Ungetestet (kein eigener Token verfügbar). Bei Problemen zuerst
mit --raw die Rohantwort einer einzelnen Abfrage ansehen.
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
from boxwhisker_injector import inject_box_whisker_chart
from openpyxl.chart import LineChart, Reference

PSR_NUCLEAR = "B14"
PSR_PUMPED_STORAGE = "B10"
AKW_NAMEN = ["beznau", "gösgen", "goesgen", "leibstadt"]

# =============================================================================
# ENTSO-E API-TOKEN - hier eintragen.
#
# WARNUNG: Das ist ein echtes, funktionsfaehiges Credential im Klartext.
# Falls dieses Skript jemals in ein Git-Repo (z.B. GitHub) kommt: VORHER
# entweder den Token wieder rausnehmen, oder die Datei in .gitignore
# aufnehmen. Sonst landet der Token oeffentlich im Commit-Verlauf - das
# laesst sich auch durch ein spaeteres Loeschen nicht mehr rueckgaengig
# machen (alte Commits bleiben in der Git-Historie sichtbar).
# =============================================================================
ENTSOE_TOKEN_HARDCODED = "dein-token-hier-eintragen"


def strip_ns(tag: str) -> str:
    """Entfernt das XML-Namespace-Präfix '{...}' von einem Tag-Namen."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def iso8601_duration_to_timedelta(duration: str) -> pd.Timedelta:
    """Parst einfache ISO8601-Dauern wie 'PT15M', 'PT60M', 'P1D'."""
    m = re.fullmatch(r"P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?", duration or "")
    if not m or not any(m.groups()):
        raise ValueError(f"Unbekanntes resolution-Format: {duration!r}")
    days, hours, minutes = (int(g) if g else 0 for g in m.groups())
    return pd.Timedelta(days=days, hours=hours, minutes=minutes)


def fix_mojibake(text: str) -> str:
    """entsoe-py/requests dekodiert die HTTP-Antwort manchmal mit der
    falschen Zeichenkodierung (z.B. 'GÃ¶sgen' statt 'Gösgen'), wenn der
    Server keinen expliziten charset=utf-8-Header schickt. Klassischer
    Fix: als Latin-1 zurueck in Bytes verwandeln, dann als UTF-8 lesen."""
    try:
        return text.encode("latin-1").decode("utf-8")
    except (UnicodeDecodeError, UnicodeEncodeError):
        return text  # war schon richtig kodiert - unveraendert lassen


def parse_generation_timeseries(xml_text: str) -> pd.DataFrame:
    """Parst ALLE Punkte (nicht nur den letzten) aus einer ENTSO-E-Antwort.
    Liefert Long-Format: timestamp | unit | psr_type | business_type |
    direction | mw

    'direction' unterscheidet bei "Actual Generation per Production Type"
    Erzeugung von Pumpverbrauch - NICHT ueber businessType (A01/A04), wie
    ich urspruenglich angenommen hatte, sondern ueber das vorhandene
    Domain-Element: eine TimeSeries mit <inBiddingZone_Domain.mRID>
    enthaelt Erzeugungswerte, eine mit <outBiddingZone_Domain.mRID>
    enthaelt Verbrauchswerte (Pumpen). businessType wird trotzdem mit
    extrahiert, falls es fuer andere Auswertungen mal nuetzlich ist."""
    xml_text = fix_mojibake(xml_text)
    root = ET.fromstring(xml_text)
    rows = []

    for ts in root:
        if strip_ns(ts.tag) != "TimeSeries":
            continue

        psr_type = unit_name = business_type = direction = None
        for el in ts:
            tag = strip_ns(el.tag)
            if tag == "businessType":
                business_type = el.text
            elif tag == "inBiddingZone_Domain.mRID":
                direction = "in"    # Erzeugung
            elif tag == "outBiddingZone_Domain.mRID":
                direction = "out"   # Verbrauch (Pumpen)
            elif tag == "MktPSRType":
                for sub in el.iter():
                    stag = strip_ns(sub.tag)
                    if stag == "psrType":
                        psr_type = sub.text
                    elif stag == "name":
                        unit_name = sub.text

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
                    "unit": unit_name or "Unbekannt",
                    "psr_type": psr_type or "?",
                    "business_type": business_type,
                    "direction": direction,
                    "mw": quantity,
                })

    return pd.DataFrame(rows)


def fetch_with_retry(query_fn, day_start, day_end, retries: int = 3, base_delay: int = 3):
    """Ruft query_fn(day_start, day_end) auf. Bei transienten HTTP-Fehlern
    (401/429/5xx - typisch fuer Rate-Limiting oder kurze Serverhaenger)
    wird mit steigender Pause automatisch erneut versucht. Bei
    NoMatchingDataError (= legitime "keine Daten"-Antwort) wird NICHT
    erneut versucht, sondern sofort weitergereicht."""
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
    haengt die geparsten Resultate zusammen. Tage ohne Daten (z.B. der
    laufende, noch nicht abgeschlossene heutige Tag - ENTSO-E liefert fuer
    Per-Generation-Unit-Daten oft erst mit Verzoegerung) werden ueberspungen
    statt das ganze Skript abzubrechen. Zwischen den Tagen wird eine kurze
    Pause eingelegt, um moegliches Rate-Limiting (das ENTSO-E offenbar mit
    401 statt 429 quittiert) von vornherein zu vermeiden."""
    end_day = pd.Timestamp.now(tz="UTC").normalize() + pd.Timedelta(days=1)
    frames = []
    for i in range(days):
        day_end = end_day - i * pd.Timedelta(days=1)
        day_start = day_end - pd.Timedelta(days=1)
        try:
            xml_text = fetch_with_retry(query_fn, day_start, day_end)
        except NoMatchingDataError:
            print(f"  Hinweis: keine Daten fuer {day_start.date()} - uebersprungen")
            time.sleep(3)
            continue
        except Exception as exc:
            print(f"  Warnung: Anfrage fuer {day_start.date()} fehlgeschlagen ({exc}) - uebersprungen")
            time.sleep(3)
            continue
        df = parse_generation_timeseries(xml_text)
        if not df.empty:
            frames.append(df)
        time.sleep(3)
    if not frames:
        return pd.DataFrame()
    result = pd.concat(frames, ignore_index=True).sort_values("timestamp")
    # Excel kann keine zeitzonen-bewussten Datetimes - nach Europe/Zurich
    # (lokale Zeit) konvertieren und dann die TZ-Info entfernen.
    result["timestamp"] = (
        result["timestamp"].dt.tz_convert("Europe/Zurich").dt.tz_localize(None)
    )
    return result


def build_akw_wide(df: pd.DataFrame, total_nuclear_df: pd.DataFrame = None) -> pd.DataFrame:
    """Baut die AKW-Tabelle. Wenn total_nuclear_df (die A75-Summe ueber
    alle CH-Kernkraftwerke) mitgegeben wird, wird zusaetzlich eine Spalte
    'Goesgen_errechnet_MW' = Total_Nuklear_A75 - (Summe der anderen
    gemeldeten Bloecke) gebildet. Das ist KEIN direkt gemeldeter Wert,
    sondern eine Ableitung - sinnvoll, weil die A75-Summe fuer CH
    offenbar vollstaendig ist, waehrend die A73-Pro-Block-Meldung fuer
    Goesgen oft leer bleibt (siehe Chat-Diskussion). Fehlt fuer einen
    Zeitpunkt einer der anderen Bloecke, wird absichtlich NaN
    zurueckgegeben statt einer falschen Zahl (skipna=False).

    WICHTIG (Bugfix): Der Join mit total_nuclear_df muss "outer" sein,
    nicht "left". Bei "left" haette ein Tag, fuer den ENTSO-E GAR KEINE
    A73-Pro-Block-Daten liefert (auch nicht fuer Beznau/Leibstadt), in
    der Wide-Tabelle ueberhaupt keine Zeile - selbst wenn fuer genau
    diesen Tag die A75-Summe (mit echter Goesgen-Produktion drin)
    verfuegbar waere. Damit kaeme die Berechnung fuer solche Tage nie
    zum Zug. Mit "outer" werden auch reine A75-Zeitstempel als Zeile
    aufgenommen (die A73-Spalten sind dann dort NaN, was korrekt ist -
    wir haben fuer diesen Zeitpunkt schlicht keine Pro-Block-Daten)."""
    has_units = df is not None and not df.empty
    akw = pd.DataFrame()
    if has_units:
        akw = df[
            (df["psr_type"] == PSR_NUCLEAR)
            | df["unit"].str.lower().apply(lambda u: any(n in u for n in AKW_NAMEN))
        ]

    if akw.empty and (total_nuclear_df is None or total_nuclear_df.empty):
        return pd.DataFrame()

    if not akw.empty:
        wide = akw.pivot_table(index="timestamp", columns="unit", values="mw", aggfunc="mean")
    else:
        wide = pd.DataFrame()

    if total_nuclear_df is not None and not total_nuclear_df.empty:
        total_in = total_nuclear_df[total_nuclear_df["direction"] == "in"]
        if not total_in.empty:
            total_series = total_in.groupby("timestamp")["mw"].mean()
            wide = wide.join(total_series.rename("Total_Nuklear_A75_MW"), how="outer")

            block_cols = [c for c in wide.columns if c != "Total_Nuklear_A75_MW"]
            goesgen_cols = [c for c in block_cols if "g\u00f6sgen" in c.lower() or "goesgen" in c.lower()]
            andere_cols = [c for c in block_cols if c not in goesgen_cols]
            if andere_cols:
                andere_summe = wide[andere_cols].sum(axis=1, skipna=False)
                wide["Goesgen_errechnet_MW"] = wide["Total_Nuklear_A75_MW"] - andere_summe

    if wide.empty:
        return pd.DataFrame()

    block_value_cols = [
        c for c in wide.columns
        if c not in ("Total_Nuklear_A75_MW", "Goesgen_errechnet_MW")
    ]
    wide["Total_MW"] = (
        wide[block_value_cols].sum(axis=1, skipna=False) if block_value_cols else None
    )

    return wide.sort_index().reset_index()


def build_pumped_wide(df: pd.DataFrame) -> pd.DataFrame:
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
    # Beide Spalten muessen existieren (CH meldet Hochgepumpt nicht,
    # aber der Box-Whisker-Injektor erwartet immer genau 2 Datenspalten)
    for col in ("Turbiniert_MW", "Hochgepumpt_MW"):
        if col not in wide.columns:
            wide[col] = 0.0
        wide[col] = wide[col].fillna(0.0)
    wide = wide[["Turbiniert_MW", "Hochgepumpt_MW"]]
    return wide.reset_index()


def main():
    parser = argparse.ArgumentParser(description="CH-AKW + Pumpspeicher via entsoe-py")
    parser.add_argument("--raw", action="store_true",
                         help="Rohe XML-Antworten zeigen (nur Live-Modus)")
    parser.add_argument("--week", action="store_true",
                         help="Letzte N Tage als Excel exportieren")
    parser.add_argument("--days", type=int, default=7,
                         help="Anzahl Tage fuer --week (Default 7)")
    parser.add_argument("--out", default=None,
                         help="Output-Pfad fuer --week (Default: "
                              "ch_strom_woche_JJJJMMTT_HHMM.xlsx mit aktuellem Zeitstempel)")
    args = parser.parse_args()

    if args.out is None:
        zeitstempel = pd.Timestamp.now(tz="Europe/Zurich").strftime("%Y%m%d_%H%M")
        args.out = f"ch_strom_woche_{zeitstempel}.xlsx"

    token = os.environ.get("ENTSOE_TOKEN", "") or ENTSOE_TOKEN_HARDCODED
    if not token or token == "dein-token-hier-eintragen":
        sys.exit("Bitte ENTSOE_TOKEN_HARDCODED oben im Skript eintragen "
                  "(oder die Umgebungsvariable ENTSOE_TOKEN setzen).")

    client = EntsoeRawClient(api_key=token)

    if args.week:
        print(f"Hole {args.days} Tage Daten (tagesweise, dauert evtl. ein paar Sekunden)...")

        akw_df = fetch_days(
            lambda s, e: client.query_generation_per_plant(country_code="CH", start=s, end=e),
            args.days,
        )
        time.sleep(5)
        total_nuclear_df = fetch_days(
            lambda s, e: client.query_generation(
                country_code="CH", start=s, end=e, psr_type=PSR_NUCLEAR),
            args.days,
        )
        time.sleep(5)
        pumped_df = fetch_days(
            lambda s, e: client.query_generation(
                country_code="CH", start=s, end=e, psr_type=PSR_PUMPED_STORAGE),
            args.days,
        )

        akw_wide = build_akw_wide(akw_df, total_nuclear_df)
        pumped_wide = build_pumped_wide(pumped_df)

        with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
            (akw_wide if not akw_wide.empty
             else pd.DataFrame({"Hinweis": ["Keine AKW-Daten erhalten - mit --raw pruefen"]})
             ).to_excel(writer, sheet_name="AKW", index=False)

            if not akw_wide.empty:
                ws = writer.sheets["AKW"]
                ws.column_dimensions["A"].width = 20
                n_rows = len(akw_wide)
                col_indices = {col: i + 1 for i, col in enumerate(akw_wide.columns)}

                # Liniendiagramm: Total_Nuklear_A75_MW als Hauptlinie (vollstaendig),
                # plus Einzelbloecke. NaN-Werte erscheinen als Luecken in der Linie –
                # bewusst so, da die Luecken echte Datenfehler widerspiegeln.
                chart = LineChart()
                datum_von = akw_wide["timestamp"].dropna().min().strftime("%d.%m.%Y")
                datum_bis = akw_wide["timestamp"].dropna().max().strftime("%d.%m.%Y")
                chart.title = f"AKW Schweiz ({datum_von} – {datum_bis}): Nuklearproduktion (MW)"
                chart.y_axis.title = "MW"
                chart.width = 28
                chart.height = 14

                # Serien in gewuenschter Reihenfolge: zuerst Gesamtsumme, dann Einzelbloecke
                for col_name in ["Total_Nuklear_A75_MW", "Beznau 1", "Beznau 2",
                                  "Leibstadt", "Goesgen_errechnet_MW"]:
                    if col_name in col_indices:
                        idx = col_indices[col_name]
                        data = Reference(ws, min_col=idx, min_row=1, max_row=n_rows + 1)
                        chart.add_data(data, titles_from_data=True)

                # Zeitachse (Spalte A)
                cats = Reference(ws, min_col=1, min_row=2, max_row=n_rows + 1)
                chart.set_categories(cats)
                chart.x_axis.tickLblSkip = max(1, n_rows // 20)
                chart.x_axis.number_format = "dd.MM HH:mm"

                # Diagramm rechts neben den Datenspalten platzieren (ab Spalte J)
                chart.anchor = "J2"
                ws.add_chart(chart)

            (pumped_wide if not pumped_wide.empty
             else pd.DataFrame({"Hinweis": ["Keine Pumpspeicher-Daten erhalten - mit --raw pruefen"]})
             ).to_excel(writer, sheet_name="Pumpspeicher", index=False)
            if not pumped_wide.empty:
                writer.sheets["Pumpspeicher"].column_dimensions["A"].width = 20

        # Pumpspeicher: Box-Whisker-Diagramm via Injektor (sheet_index=2)
        if not pumped_wide.empty:
            datum_von = pumped_wide["timestamp"].min().strftime("%d.%m.%Y")
            datum_bis = pumped_wide["timestamp"].max().strftime("%d.%m.%Y")
            inject_box_whisker_chart(
                args.out,
                sheet_name="Pumpspeicher",
                n_data_rows=len(pumped_wide),
                chart_title=f"Pumpspeicher Schweiz ({datum_von} – {datum_bis}): Turbiniert vs. Hochgepumpt (MW)",
                series1_name="Turbiniert_MW",
                series2_name="Hochgepumpt_MW",
                sheet_index=2,
            )

        print(f"Fertig: {args.out}")
        print(f"  AKW-Zeilen: {len(akw_wide)}, Pumpspeicher-Zeilen: {len(pumped_wide)}")
        return

    # --- Live-Snapshot -------------------------------------------------------
    now = pd.Timestamp.now(tz="UTC")
    start = now - pd.Timedelta(hours=3)
    end = now + pd.Timedelta(hours=1)

    try:
        xml_units = client.query_generation_per_plant(country_code="CH", start=start, end=end)
        xml_pumped = client.query_generation(
            country_code="CH", start=start, end=end, psr_type=PSR_PUMPED_STORAGE)
    except Exception as exc:
        sys.exit(f"Anfrage an ENTSO-E fehlgeschlagen: {exc}")

    if args.raw:
        print("=== Per Generation Unit (A73) ===\n", xml_units)
        print("\n=== Per Production Type, Pumpspeicher B10 (A75) ===\n", xml_pumped)
        return

    unit_df = parse_generation_timeseries(xml_units)
    pumped_df = parse_generation_timeseries(xml_pumped)

    akw = pd.DataFrame()
    if not unit_df.empty:
        akw = unit_df[
            (unit_df["psr_type"] == PSR_NUCLEAR)
            | unit_df["unit"].str.lower().apply(lambda u: any(n in u for n in AKW_NAMEN))
        ]

    print(f"Aktuelle Leistung Schweizer AKW (Stand ca. {now.strftime('%d.%m.%Y %H:%M UTC')}):\n")
    if akw.empty:
        print("  Keine Kernkraft-Einheiten gefunden (--raw zum Pruefen).")
    else:
        latest = akw.sort_values("timestamp").groupby("unit").tail(1)
        total = 0.0
        for _, row in latest.sort_values("unit").iterrows():
            print(f"  {row['unit']:25s} {row['mw']:8.1f} MW")
            total += row["mw"]
        print(f"\n  {'TOTAL':25s} {total:8.1f} MW")

    print("\nPumpspeicher Schweiz, aggregiert (Art. 16.1.B&C):\n")
    if pumped_df.empty:
        print("  Keine Daten erhalten - mit --raw die Rohantwort pruefen.")
        return

    latest_p = pumped_df.sort_values("timestamp").groupby("direction").tail(1)
    for _, row in latest_p.iterrows():
        label = {
            "in": "Erzeugung (Turbine laeuft, Strom raus)",
            "out": "Pumpverbrauch (Pumpe laeuft, Strom rein)",
        }.get(row["direction"], str(row["direction"]))
        sign = "-" if row["direction"] == "out" else ""
        print(f"  {label:42s} {sign}{row['mw']:7.1f} MW")


if __name__ == "__main__":
    main()
