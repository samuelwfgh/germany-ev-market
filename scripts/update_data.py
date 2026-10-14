#!/usr/bin/env python3
"""
Update the Germany EV/car-market datasets used by index.html.

Design goals
------------
* Prefer authoritative/public German sources where practical.
* Never replace a valid local dataset with an empty/obviously broken parse.
* Keep the site working when an upstream website changes or is temporarily down.
* Preserve missing values as missing. Do not interpolate or manufacture data.
* Record source/check status in data/metadata.json.

The script is intentionally defensive because KBA/BAFA Excel layouts can change.
When a parser cannot validate a new source, it logs the error and keeps the last
known-good JSON unchanged.

Run:
    python scripts/update_data.py --data-dir data
    python scripts/update_data.py --data-dir data --verbose
    python scripts/update_data.py --data-dir data --no-network   # validation only
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import math
import re
import sys
import tempfile
import time
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import pdfplumber


USER_AGENT = "germany-ev-market-data-updater/1.0 (+GitHub Pages research project)"
TIMEOUT = 15
GERMANY_CSV = "https://raw.githubusercontent.com/robbieandrew/robbieandrew.github.io/master/carsales/data/all_carsales_monthly.csv"
ADAC_NEW = "https://www.adac.de/news/neuzulassungen-kba/"
BNETZA = "https://www.bundesnetzagentur.de/DE/Fachthemen/ElektrizitaetundGas/E-Mobilitaet/start.html"
IPHE_GERMANY = "https://www.iphe.net/archives/germany"
BAFA_BASE = "https://www.bafa.de/SharedDocs/Downloads/DE/Energie/Mineraloel/"
KBA_BASE = "https://www.kba.de/SharedDocs/Downloads/DE/Statistik/Fahrzeuge"

CANONICAL = ("Diesel", "Petrol", "NonPluginHybrid", "PluginHybrid", "BatteryElectric", "Hydrogen", "Others")
MONTH_EN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MONTH_DE = {
    "januar":1,"februar":2,"märz":3,"maerz":3,"april":4,"mai":5,"juni":6,
    "juli":7,"august":8,"september":9,"oktober":10,"november":11,"dezember":12,
}

session = requests.Session()
session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "de-DE,de;q=0.9,en;q=0.7"})


@dataclass
class Outcome:
    ok: bool
    changed: bool = False
    latest: str | None = None
    note: str = ""
    source: str = ""
    error: str | None = None


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def log(msg: str, verbose: bool = True) -> None:
    if verbose:
        print(msg, flush=True)


def load_json(path: Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return deepcopy(default)


def write_json_if_changed(path: Path, obj: Any) -> bool:
    new = json.dumps(obj, ensure_ascii=False, indent=2, sort_keys=False) + "\n"
    old = path.read_text(encoding="utf-8") if path.exists() else None
    if old == new:
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(new, encoding="utf-8")
    tmp.replace(path)
    return True


def get(url: str, *, binary: bool = False) -> bytes | str:
    last = None
    for attempt in range(2):
        try:
            r = session.get(url, timeout=TIMEOUT)
            # Missing monthly files are normal; do not waste retries on 4xx
            # except rate limiting.
            if 400 <= r.status_code < 500 and r.status_code != 429:
                raise RuntimeError(f"HTTP {r.status_code}")
            r.raise_for_status()
            if binary:
                return r.content
            r.encoding = r.apparent_encoding or r.encoding or "utf-8"
            return r.text
        except Exception as e:
            last = e
            if "HTTP 4" in str(e) and "429" not in str(e):
                break
            if attempt < 1:
                time.sleep(1.2)
    raise RuntimeError(f"GET failed: {url} ({last})")


def parse_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)) and math.isfinite(float(value)):
        return float(value)
    s = str(value).strip().replace("\xa0", " ")
    if not s or s in {"-", "–", "—", "n.v.", "n/a"}:
        return None
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s:
        return None
    # German-formatted integers/decimals.
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        tail = s.rsplit(",", 1)[-1]
        s = s.replace(".", "")
        s = s.replace(",", "." if len(tail) != 3 else "")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    elif s.count(".") == 1 and len(s.rsplit(".",1)[-1]) == 3:
        s = s.replace(".", "")
    try:
        v = float(s)
        return v if math.isfinite(v) else None
    except ValueError:
        return None


def german_int(s: str) -> int:
    return int(re.sub(r"[^\d]", "", s))


def ym_key(year: int, month: int) -> str:
    return f"{year:04d}{month:02d}"


def canonical_fuel(name: str) -> str | None:
    n = re.sub(r"[^a-z0-9]+", "", name.lower())
    if n in {"diesel"}: return "Diesel"
    if n in {"petrol","gasoline","benzin"}: return "Petrol"
    if n in {"nonpluginhybrid","hev","hybridwithoutplugin","hybridohneplugin"}: return "NonPluginHybrid"
    if n in {"pluginhybrid","phev"}: return "PluginHybrid"
    if n in {"batteryelectric","bev","electric","elektro"}: return "BatteryElectric"
    if n in {"hydrogen","fuelcell","fcev","wasserstoff"}: return "Hydrogen"
    if n in {"others","other","sonstige","lpg","lpgblend","cng"}: return "Others"
    return None


def validate_new_rows(rows: list[dict]) -> None:
    if len(rows) < 100:
        raise ValueError(f"new-registration series too short ({len(rows)} rows)")
    by_month: dict[str, list[dict]] = {}
    for r in rows:
        if r.get("fuel") not in CANONICAL:
            continue
        if not isinstance(r.get("value"), (int, float)) or r["value"] < 0:
            raise ValueError("invalid new-registration value")
        by_month.setdefault(str(r["ym"]), []).append(r)
    if not by_month:
        raise ValueError("no months in new-registration data")
    for k, vals in list(sorted(by_month.items()))[-12:]:
        total = sum(float(x["value"]) for x in vals)
        if not (50_000 <= total <= 500_000):
            raise ValueError(f"implausible new-registration total {total:g} for {k}")


def update_new_registrations(path: Path, verbose: bool) -> Outcome:
    existing = load_json(path, {"schema_version": 1, "rows": []})
    old_rows = existing.get("rows", [])
    merged = {(str(r["ym"]), r["fuel"]): r for r in old_rows if r.get("fuel") in CANONICAL}

    # Long-run consolidated series. This is very useful for history and is
    # cross-checked/overlaid with a current KBA-derived monthly summary below.
    text = get(GERMANY_CSV)
    reader = csv.DictReader(io.StringIO(text))
    source_rows = 0
    source_agg: dict[tuple[str,str], int] = {}
    for row in reader:
        if str(row.get("Country", "")).strip().lower() != "germany":
            continue
        fuel = canonical_fuel(str(row.get("Fuel", "")))
        if not fuel:
            continue
        raw_ym = str(row.get("YYYYMM", "")).replace(".0", "")
        if not re.fullmatch(r"\d{6}", raw_ym):
            continue
        value = parse_number(row.get("Value"))
        if value is None or value < 0:
            continue
        y, m = int(raw_ym[:4]), int(raw_ym[4:])
        if not 1 <= m <= 12:
            continue
        # Multiple minor-fuel categories can map to "Others"; add them rather
        # than allowing the last row to overwrite the earlier one.
        source_agg[(raw_ym, fuel)] = source_agg.get((raw_ym, fuel), 0) + int(round(value))
        source_rows += 1

    if source_rows < 100:
        raise ValueError(f"historical source returned only {source_rows} Germany rows")
    for (raw_ym, fuel), value in source_agg.items():
        y, m = int(raw_ym[:4]), int(raw_ym[4:])
        merged[(raw_ym, fuel)] = {"ym": raw_ym, "year": y, "month": m, "fuel": fuel, "value": value}

    # Try to overlay the latest public KBA-derived figures from ADAC's rolling
    # KBA page. This avoids waiting for the consolidated history file when a
    # fresh German month has just been published.
    overlay_note = ""
    try:
        html = get(ADAC_NEW)
        text_clean = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)
        lower = text_clean.lower()
        # Locate the most recent month/year explicitly mentioned around "Neuzulassungen".
        month_year = None
        for name, num in MONTH_DE.items():
            matches = list(re.finditer(rf"\b{name}\s+(20\d{{2}})\b", lower))
            if matches:
                yr = int(matches[-1].group(1))
                cand = (yr, num)
                if month_year is None or cand > month_year:
                    month_year = cand
        # Current ADAC wording has stable category phrases. Keep regexes narrow;
        # if any required category is missing we simply do not overlay.
        def first_int(patterns: Iterable[str]) -> int | None:
            for p in patterns:
                m = re.search(p, text_clean, flags=re.I)
                if m:
                    return german_int(m.group(1))
            return None

        bev = first_int([r"wurden\s+([\d.\s]+)\s+E-Autos", r"([\d.\s]+)\s+E-Autos.*?neu", r"Elektro.*?([\d.\s]+)\s+(?:Neuzulassungen|E-Autos)"])
        hev = first_int([r"Hybride\s+mit\s+([\d.\s]+)\s+Neuzulassungen", r"Hybrid(?:e)?.{0,30}?([\d.\s]+)\s+Neuzulassungen"])
        petrol = first_int([r"Benziner\s+mit\s+([\d.\s]+)\s+Zulassungen", r"Benziner.{0,30}?([\d.\s]+)\s+(?:Zulassungen|Neuanmeldungen)"])
        phev = first_int([r"Plug.?in.Hybrid(?:e)?.{0,40}?([\d.\s]+)\s+(?:Neuanmeldungen|Neuzulassungen)", r"([\d.\s]+)\s+Plug.?in.Hybrid"])
        diesel = first_int([r"Diesel(?:fahrzeuge)?.{0,50}?([\d.\s]+)\s+(?:Anmeldungen|Zulassungen)", r"Diesel.{0,30}?([\d.\s]+)\s+(?:Anmeldungen|Zulassungen)"])
        total = first_int([r"insgesamt\s+([\d.\s]+)\s+(?:neue\s+)?Pkw", r"([\d.\s]+)\s+(?:Pkw-)?Neuzulassungen"])
        if month_year and all(v is not None for v in (bev, hev, petrol, phev, diesel, total)):
            y, m = month_year
            main = {"BatteryElectric":bev,"NonPluginHybrid":hev,"Petrol":petrol,"PluginHybrid":phev,"Diesel":diesel}
            other = total - sum(main.values())
            if 0 <= other <= max(15_000, int(total * 0.08)):
                main["Hydrogen"] = 0
                main["Others"] = other
                ym = ym_key(y,m)
                for fuel, value in main.items():
                    merged[(ym, fuel)] = {"ym":ym,"year":y,"month":m,"fuel":fuel,"value":int(value)}
                overlay_note = f" Current month overlaid from ADAC's KBA-derived page: {y}-{m:02d}."
    except Exception as e:
        overlay_note = f" Current-month overlay skipped: {e}"

    rows = sorted(merged.values(), key=lambda r: (r["ym"], CANONICAL.index(r["fuel"])))
    validate_new_rows(rows)
    latest = max(r["ym"] for r in rows)
    obj = deepcopy(existing)
    obj.update({
        "schema_version": 1,
        "updated_at": now_iso(),
        "latest": f"{latest[:4]}-{latest[4:]}",
        "source": {
            "history": "Robbie Andrew consolidated monthly car-sales dataset (German rows; official registration sources collated)",
            "current_overlay": "ADAC rolling KBA summary when a complete latest-month parse is available",
            "history_url": GERMANY_CSV,
            "current_url": ADAC_NEW,
        },
        "rows": rows,
    })
    changed = write_json_if_changed(path, obj)
    return Outcome(True, changed, obj["latest"], f"{len(rows)} rows.{overlay_note}", GERMANY_CSV)


# ---------- Generic KBA workbook parsing ----------

def workbook_from_bytes(content: bytes):
    return load_workbook(io.BytesIO(content), data_only=True, read_only=True)


def row_text(values: Iterable[Any]) -> str:
    return " | ".join("" if v is None else str(v) for v in values)


def numeric_cells(row: list[Any]) -> list[tuple[int, float]]:
    out = []
    for i, v in enumerate(row):
        n = parse_number(v)
        if n is not None and n >= 0:
            out.append((i, n))
    return out


def fuel_label_from_text(s: str) -> str | None:
    t = s.lower()
    if "plug" in t and "hybrid" in t:
        return "PluginHybrid"
    if ("elektro" in t or "battery" in t) and "hybrid" not in t:
        return "BatteryElectric"
    if re.search(r"\bdiesel\b", t):
        return "Diesel"
    if "benzin" in t or "ottomotor" in t or "ottokraft" in t:
        return "Petrol"
    if "wasserstoff" in t or "brennstoffzelle" in t:
        return "Hydrogen"
    if "hybrid" in t and "plug" not in t:
        return "HybridTotal"
    if "sonstig" in t or "flüssiggas" in t or "fluessiggas" in t or "erdgas" in t:
        return "OtherComponent"
    return None


def extract_kba_fuels(content: bytes) -> tuple[dict[str,int], int | None]:
    """
    Best-effort parser for KBA FZ8/FZ9/FZ10 workbooks.

    KBA workbooks evolve. We deliberately require the parsed five main
    drivetrain categories to form a plausible passenger-car total before
    accepting a month. If validation fails, the caller retains old data.
    """
    wb = workbook_from_bytes(content)
    best: tuple[float, dict[str,int], int | None] | None = None

    for ws in wb.worksheets:
        matrix = []
        for ridx, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = list(row)
            if any(v is not None for v in vals):
                matrix.append((ridx, vals))
            if ridx > 450:
                break

        # Approach 1: categories are rows; compare values in common numeric columns.
        category_rows: dict[str, list[Any]] = {}
        total_rows: list[list[Any]] = []
        for _, vals in matrix:
            txt = row_text(vals)
            lab = fuel_label_from_text(txt)
            if lab and lab not in category_rows:
                category_rows[lab] = vals
            lo = txt.lower()
            if ("personenkraftwagen" in lo or re.search(r"\bpkw\b", lo)) and ("insgesamt" in lo or "gesamt" in lo):
                total_rows.append(vals)

        required = {"Diesel","Petrol","PluginHybrid","BatteryElectric"}
        if required.issubset(category_rows) and ("HybridTotal" in category_rows):
            # Candidate current-month columns are those numeric in all important rows.
            common_cols = None
            for k in required | {"HybridTotal"}:
                cols = {i for i,n in numeric_cells(category_rows[k]) if 10 <= n <= 5_000_000}
                common_cols = cols if common_cols is None else common_cols & cols
            for col in sorted(common_cols or []):
                vals = {}
                ok = True
                for k in required | {"HybridTotal"}:
                    n = parse_number(category_rows[k][col] if col < len(category_rows[k]) else None)
                    if n is None:
                        ok = False
                        break
                    vals[k] = int(round(n))
                if not ok or vals["HybridTotal"] < vals["PluginHybrid"]:
                    continue
                parsed = {
                    "Diesel":vals["Diesel"],
                    "Petrol":vals["Petrol"],
                    "PluginHybrid":vals["PluginHybrid"],
                    "NonPluginHybrid":vals["HybridTotal"]-vals["PluginHybrid"],
                    "BatteryElectric":vals["BatteryElectric"],
                }
                main_sum = sum(parsed.values())
                if not 50_000 <= main_sum <= 5_000_000:
                    continue

                total = None
                for tr in total_rows:
                    n = parse_number(tr[col] if col < len(tr) else None)
                    if n and n >= main_sum:
                        total = int(round(n))
                        break
                # Prefer a column that produces a normal monthly Pkw total.
                if total is not None:
                    ratio = main_sum / total
                    if not 0.85 <= ratio <= 1.01:
                        continue
                    score = 1000 - abs(total - 250_000) / 1000
                else:
                    # Monthly main sums are typically <500k; cumulative columns often millions.
                    if main_sum > 500_000:
                        continue
                    score = 500 - abs(main_sum - 230_000) / 1000
                if best is None or score > best[0]:
                    best = (score, parsed, total)

        # Approach 2: category labels and values are laid out as columns in a row.
        for pos, (_, header) in enumerate(matrix):
            labels = {}
            for c, v in enumerate(header):
                if v is None: continue
                lab = fuel_label_from_text(str(v))
                if lab:
                    labels[lab] = c
            if not required.issubset(labels) or "HybridTotal" not in labels:
                continue
            for _, vals in matrix[pos+1:pos+10]:
                txt = row_text(vals).lower()
                if not ("personenkraftwagen" in txt or re.search(r"\bpkw\b", txt) or "insgesamt" in txt):
                    continue
                nums = {}
                try:
                    for k,c in labels.items():
                        if k in required | {"HybridTotal"}:
                            n=parse_number(vals[c] if c<len(vals) else None)
                            if n is None: raise ValueError
                            nums[k]=int(round(n))
                    if nums["HybridTotal"] < nums["PluginHybrid"]: continue
                    parsed={"Diesel":nums["Diesel"],"Petrol":nums["Petrol"],"PluginHybrid":nums["PluginHybrid"],
                            "NonPluginHybrid":nums["HybridTotal"]-nums["PluginHybrid"],"BatteryElectric":nums["BatteryElectric"]}
                    main_sum=sum(parsed.values())
                    if 50_000 <= main_sum <= 500_000:
                        best=(900,parsed,None) if best is None or 900>best[0] else best
                except Exception:
                    pass

    if not best:
        raise ValueError("could not find a validated Pkw drivetrain table in workbook")
    parsed, total = best[1], best[2]
    if total is None:
        # We can only derive 'Other' safely when the total is known.
        parsed["Hydrogen"] = 0
        parsed["Others"] = 0
    else:
        other = total - sum(parsed.values())
        if other < 0 or other > max(20_000, int(total*0.12)):
            raise ValueError("KBA workbook residual 'Other' is implausible")
        parsed["Hydrogen"] = 0
        parsed["Others"] = other
    return parsed, total


def kba_candidate_urls(series: str, year: int, month: int) -> list[str]:
    stem = f"{series.lower()}_{year}{month:02d}.xlsx"
    folder = series.upper()
    return [
        f"{KBA_BASE}/{folder}/{stem}?__blob=publicationFile&v={v}"
        for v in (4,3,2,1)
    ]


def download_first(urls: Iterable[str]) -> tuple[bytes, str]:
    last = None
    for url in urls:
        try:
            b = get(url, binary=True)
            if isinstance(b, bytes) and len(b) > 10_000 and b[:2] == b"PK":
                return b, url
        except Exception as e:
            last = e
    raise RuntimeError(str(last or "no candidate URL succeeded"))


def update_used_transfers(path: Path, new_path: Path, verbose: bool) -> Outcome:
    obj = load_json(path, {"schema_version":1,"drivetrain_shares":[],"market_totals":[]})
    shares = {str(r["ym"]): r for r in obj.get("drivetrain_shares", [])}
    totals = {str(r["ym"]): r for r in obj.get("market_totals", [])}
    new_obj = load_json(new_path, {"rows":[]})
    new_month_total = {}
    for r in new_obj.get("rows", []):
        new_month_total.setdefault(str(r["ym"]), 0)
        new_month_total[str(r["ym"])] += int(r.get("value",0))

    now = datetime.now(timezone.utc)
    # Recheck recent months; KBA can revise files.
    months = []
    y, m = now.year, now.month
    for delta in range(0, 7):
        mm = m - delta
        yy = y
        while mm <= 0:
            mm += 12; yy -= 1
        months.append((yy,mm))
    changed_source = False
    notes = []
    for yy,mm in reversed(months):
        ym = ym_key(yy,mm)
        try:
            content,url = download_first(kba_candidate_urls("FZ9", yy, mm))
            parsed,total = extract_kba_fuels(content)
            if not total or not 150_000 <= total <= 900_000:
                raise ValueError(f"implausible/missing used-car total {total}")
            main_sum = sum(parsed.get(k,0) for k in ("Diesel","Petrol","NonPluginHybrid","PluginHybrid","BatteryElectric","Hydrogen","Others"))
            if abs(main_sum-total) > max(100, total*0.01):
                raise ValueError("drivetrain counts do not reconcile with used total")
            share = {"ym":ym,"year":yy,"month":mm,"total":100.0}
            for k in ("Petrol","Diesel","NonPluginHybrid","PluginHybrid","BatteryElectric","Others"):
                share[k] = round(parsed.get(k,0) / total * 100, 4)
            shares[ym] = share
            totals[ym] = {"ym":ym,"label":MONTH_EN[mm-1],"year":yy,"month":mm,"Used":total,
                          "New":new_month_total.get(ym)}
            changed_source = True
            notes.append(ym)
        except Exception as e:
            log(f"[used] {ym}: {e}", verbose)

    out = deepcopy(obj)
    out["schema_version"]=1
    out["updated_at"]=now_iso()
    out["source"]={"name":"KBA FZ 9","url":"https://www.govdata.de/suche/daten/fz-9-2026-besitzumschreibungen-von-kraftfahrzeugen-und-kraftfahrzeuganhangern"}
    out["drivetrain_shares"]=sorted(shares.values(), key=lambda r:r["ym"])
    # Backfill new totals for already-known used months from fresh new series.
    for ym,r in list(totals.items()):
        if ym in new_month_total:
            r["New"]=new_month_total[ym]
    out["market_totals"]=sorted(totals.values(), key=lambda r:r["ym"])
    if out["drivetrain_shares"]:
        out["latest_detailed"]=max(r["ym"] for r in out["drivetrain_shares"])
        out["latest_detailed"]=out["latest_detailed"][:4]+"-"+out["latest_detailed"][4:]
    if out["market_totals"]:
        out["latest_total"]=max(r["ym"] for r in out["market_totals"])
        out["latest_total"]=out["latest_total"][:4]+"-"+out["latest_total"][4:]
    # Validate existing+new.
    for r in out["drivetrain_shares"]:
        s=sum(float(r.get(k,0)) for k in ("Petrol","Diesel","NonPluginHybrid","PluginHybrid","BatteryElectric","Others"))
        if not 98.5 <= s <= 101.5:
            raise ValueError(f"used share sum invalid for {r['ym']}: {s}")
    changed=write_json_if_changed(path,out)
    latest=f"{out.get('latest_total','?')} total / {out.get('latest_detailed','?')} drivetrain"
    note=("Parsed KBA months: "+", ".join(notes)) if notes else "No newer validated KBA workbook found; kept last known-good detailed series."
    return Outcome(True,changed,latest,note,out["source"]["url"])


# ---------- BAFA road fuels ----------

def extract_bafa_xlsx(content: bytes) -> tuple[float|None,float|None]:
    wb = workbook_from_bytes(content)
    diesel = petrol = None
    # Prefer sheets/tables containing "Inlandsablieferungen".
    sheets = sorted(wb.worksheets, key=lambda w: ("inlandsabliefer" not in w.title.lower(), w.title))
    for ws in sheets:
        rows=list(ws.iter_rows(values_only=True))
        context=" ".join(row_text(r) for r in rows[:35]).lower()
        for vals in rows:
            txt=row_text(vals)
            lo=txt.lower()
            if "dieselkraftstoff" in lo and diesel is None:
                nums=[n for _,n in numeric_cells(list(vals)) if 100_000 <= n <= 10_000_000]
                if nums: diesel=nums[0]/1000.0  # t -> 1000 t
            if ("ottokraftstoff" in lo or re.search(r"\bbenzin\b",lo)) and petrol is None:
                nums=[n for _,n in numeric_cells(list(vals)) if 100_000 <= n <= 10_000_000]
                if nums: petrol=nums[0]/1000.0
        if diesel is not None and petrol is not None:
            break
    return diesel,petrol


def extract_bafa_pdf(content: bytes) -> tuple[float|None,float|None]:
    with tempfile.NamedTemporaryFile(suffix=".pdf") as f:
        f.write(content); f.flush()
        texts=[]
        with pdfplumber.open(f.name) as pdf:
            for page in pdf.pages:
                t=page.extract_text() or ""
                if "Inlandsabliefer" in t or "Dieselkraftstoff" in t or "Ottokraftstoff" in t:
                    texts.append(t)
    alltext="\n".join(texts)
    def find(label_re):
        m=re.search(label_re + r".{0,100}?([\d.]{4,})", alltext, flags=re.I|re.S)
        if not m: return None
        n=parse_number(m.group(1))
        if n is None: return None
        # Published tables normally use tonnes, occasionally 1000 t. Normalise.
        return n/1000.0 if n > 100_000 else n
    return find(r"Dieselkraftstoff"), find(r"(?:Ottokraftstoff|Benzin)")


def bafa_month_file(year:int,month:int) -> tuple[bytes,str,str]:
    stem=f"moel_amtliche_daten_{year}_{month:02d}"
    landing=BAFA_BASE+stem+".html"
    html=get(landing)
    soup=BeautifulSoup(html,"html.parser")
    candidates=[]
    for a in soup.find_all("a",href=True):
        href=a["href"]
        lo=href.lower()
        if stem in lo and any(ext in lo for ext in (".xlsx",".xls",".pdf")):
            candidates.append(urljoin(landing,href))
    # Direct fallback patterns used on BAFA.
    candidates += [
        BAFA_BASE+stem+".xlsx?__blob=publicationFile&v=2",
        BAFA_BASE+stem+".pdf?__blob=publicationFile&v=2",
    ]
    seen=set()
    for url in candidates:
        if url in seen: continue
        seen.add(url)
        try:
            b=get(url,binary=True)
            if not isinstance(b,bytes) or len(b)<5000: continue
            typ="xlsx" if b[:2]==b"PK" else ("pdf" if b[:4]==b"%PDF" else "")
            if typ: return b,typ,url
        except Exception:
            pass
    raise RuntimeError("no downloadable BAFA monthly file found")


def update_road_fuels(path:Path,verbose:bool) -> Outcome:
    obj=load_json(path,{"schema_version":1,"rows":[]})
    rows={str(r["ym"]):r for r in obj.get("rows",[])}
    now=datetime.now(timezone.utc)
    found=[]
    # Recheck 20 recent report months; BAFA can have long publication lag/revisions.
    months=[]
    yy,mm=now.year,now.month
    for d in range(16):
        m=mm-d; y=yy
        while m<=0: m+=12; y-=1
        months.append((y,m))
    for y,m in reversed(months):
        ym=ym_key(y,m)
        try:
            content,typ,url=bafa_month_file(y,m)
            diesel,petrol=(extract_bafa_xlsx(content) if typ=="xlsx" else extract_bafa_pdf(content))
            # Require plausible monthly road-fuel quantities in 1000 t.
            if diesel is not None and not 800 <= diesel <= 5000: diesel=None
            if petrol is not None and not 400 <= petrol <= 3000: petrol=None
            if diesel is None and petrol is None:
                raise ValueError("could not identify a plausible diesel/petrol domestic-delivery row")
            old=rows.get(ym,{"ym":ym,"year":y,"month":m,"Diesel":None,"Petrol":None})
            rec=dict(old)
            if diesel is not None: rec["Diesel"]=round(diesel,1)
            if petrol is not None: rec["Petrol"]=round(petrol,1)
            rows[ym]=rec
            found.append(ym)
        except Exception as e:
            log(f"[road fuels] {ym}: {e}",verbose)
    out=deepcopy(obj)
    out.update({
        "schema_version":1,"updated_at":now_iso(),"unit":"1000 tonnes",
        "source":{"name":"BAFA Amtliche Mineralöldaten","url":"https://www.bafa.de/DE/Energie/Rohstoffe/Mineraloelstatistik/mineraloel_node.html"},
        "definition":"Domestic deliveries/consumption of diesel and petrol. This is not literal filling-station till sales; published monthly values may be revised.",
        "rows":sorted(rows.values(),key=lambda r:r["ym"])
    })
    changed=write_json_if_changed(path,out)
    complete=[r["ym"] for r in out["rows"] if r.get("Diesel") is not None and r.get("Petrol") is not None]
    latest=(max(complete)[:4]+"-"+max(complete)[4:]) if complete else None
    note=("Validated BAFA reports: "+", ".join(found[-6:])) if found else "No newer parsable BAFA monthly report; kept last known-good data."
    return Outcome(True,changed,latest,note,out["source"]["url"])


# ---------- Bundesnetzagentur charging ----------

def parse_bnetza_current(html:str) -> dict[str,Any]:
    txt=BeautifulSoup(html,"html.parser").get_text(" ",strip=True)
    # "157.762 Normalladepunkte und 56.572 Schnellladepunkte, die am 1. September 2026..."
    m=re.search(r"([\d.\s]+)\s+Normalladepunkte\s+und\s+([\d.\s]+)\s+Schnellladepunkte.*?am\s+(\d{1,2})\.\s+([A-Za-zÄÖÜäöü]+)\s+(20\d{2}).*?insgesamt\s+([\d,.]+)\s*GW",txt,re.I|re.S)
    if not m:
        raise ValueError("current BNetzA headline figures not found")
    normal=german_int(m.group(1)); fast=german_int(m.group(2))
    day=int(m.group(3)); month=MONTH_DE.get(m.group(4).lower()); year=int(m.group(5))
    if not month: raise ValueError("BNetzA month name not understood")
    power=parse_number(m.group(6))
    if power is None: raise ValueError("BNetzA rated power not understood")
    if not (50_000<=normal<=500_000 and 5_000<=fast<=300_000 and 1<=power<=100):
        raise ValueError("BNetzA figures outside plausibility bounds")
    return {"date":f"{year:04d}-{month:02d}-{day:02d}","label":f"{MONTH_EN[month-1]} {year}",
            "normal":normal,"fast":fast,"power_mw":int(round(power*1000))}


def update_charging(path:Path,verbose:bool) -> Outcome:
    obj=load_json(path,{"schema_version":1,"rows":[]})
    rows={str(r["date"]):r for r in obj.get("rows",[])}
    html=get(BNETZA)
    cur=parse_bnetza_current(html)
    rows[cur["date"]]=cur

    # Also try the small "Ladeinfrastruktur in Zahlen" workbook for historical
    # snapshots. If its layout is not recognised, the current headline still updates.
    workbook_note=""
    try:
        soup=BeautifulSoup(html,"html.parser")
        href=None
        for a in soup.find_all("a",href=True):
            t=(a.get_text(" ",strip=True)+" "+a["href"]).lower()
            if "ladeinfrastruktur" in t and "zahlen" in t and "xlsx" in t:
                href=urljoin(BNETZA,a["href"]); break
        if href:
            content=get(href,binary=True)
            wb=workbook_from_bytes(content)
            added=0
            for ws in wb.worksheets:
                rows_ws=list(ws.iter_rows(values_only=True))
                # Find a header containing normal/schnell; then locate date + values below.
                for hi,header in enumerate(rows_ws[:60]):
                    hs=[str(x or "").lower() for x in header]
                    ncol=next((i for i,x in enumerate(hs) if "normal" in x and "lade" in x),None)
                    fcol=next((i for i,x in enumerate(hs) if "schnell" in x and "lade" in x),None)
                    if ncol is None or fcol is None: continue
                    for vals in rows_ws[hi+1:]:
                        n=parse_number(vals[ncol] if ncol<len(vals) else None)
                        f=parse_number(vals[fcol] if fcol<len(vals) else None)
                        if n is None or f is None or not (500<=n<=500000 and 50<=f<=300000): continue
                        # Find a date/year in the same row.
                        date=None
                        for v in vals:
                            if isinstance(v,datetime):
                                date=v.strftime("%Y-%m-%d");break
                            s=str(v or "")
                            md=re.search(r"(20\d{2})[-./](\d{1,2})(?:[-./](\d{1,2}))?",s)
                            if md:
                                date=f"{int(md.group(1)):04d}-{int(md.group(2)):02d}-{int(md.group(3) or 1):02d}";break
                            if re.fullmatch(r"20\d{2}",s.strip()):
                                date=f"{s.strip()}-01-01";break
                        if date:
                            prev=rows.get(date,{})
                            rows[date]={**prev,"date":date,"label":prev.get("label", date[:7]),"normal":int(round(n)),"fast":int(round(f))}
                            added+=1
                    break
            workbook_note=f" Historical workbook rows recognised: {added}."
    except Exception as e:
        workbook_note=f" Historical workbook parse skipped: {e}"

    out=deepcopy(obj)
    out.update({"schema_version":1,"updated_at":now_iso(),"latest":cur["date"],
                "source":{"name":"Bundesnetzagentur Ladesäulenregister","url":BNETZA},
                "rows":sorted(rows.values(),key=lambda r:r["date"])})
    changed=write_json_if_changed(path,out)
    return Outcome(True,changed,cur["date"],f"{cur['normal']:,} normal + {cur['fast']:,} fast; {cur['power_mw']/1000:.2f} GW.{workbook_note}",BNETZA)


# ---------- Hydrogen documented snapshots ----------

def update_hydrogen(path:Path,verbose:bool) -> Outcome:
    obj=load_json(path,{"schema_version":1,"rows":[]})
    rows={str(r["date"]):r for r in obj.get("rows",[])}
    html=get(IPHE_GERMANY)
    txt=BeautifulSoup(html,"html.parser").get_text(" ",strip=True)
    m=re.search(r"As of\s+([A-Za-z]+)\s+(20\d{2}).{0,120}?(\d+)\s+stations?\s+in\s+operation",txt,re.I|re.S)
    if not m:
        raise ValueError("IPHE operational-station snapshot not found")
    mon_name=m.group(1).lower()
    month={name:num for name,num in MONTH_DE.items()}.get(mon_name)
    if month is None:
        en={x.lower():i+1 for i,x in enumerate(["January","February","March","April","May","June","July","August","September","October","November","December"])}
        month=en.get(mon_name)
    if month is None: raise ValueError("IPHE snapshot month not understood")
    year=int(m.group(2));count=int(m.group(3))
    if not 20<=count<=300: raise ValueError("IPHE H2 station count implausible")
    date=f"{year:04d}-{month:02d}-01"
    rows[date]={"date":date,"label":f"{MONTH_EN[month-1]} {year}","count":count,
                "note":"Documented IPHE operational-station snapshot; not a real-time availability count."}
    out=deepcopy(obj)
    out.update({"schema_version":1,"updated_at":now_iso(),"latest":f"{year:04d}-{month:02d}",
                "cadence":"irregular documented snapshots; not interpolated monthly",
                "source":{"name":"IPHE Germany country update","url":IPHE_GERMANY},
                "rows":sorted(rows.values(),key=lambda r:r["date"])})
    changed=write_json_if_changed(path,out)
    return Outcome(True,changed,out["latest"],f"{count} stations in operation in documented snapshot.",IPHE_GERMANY)


def validate_all(data_dir:Path) -> None:
    new=load_json(data_dir/"new_registrations.json",{})
    validate_new_rows(new.get("rows",[]))
    used=load_json(data_dir/"used_transfers.json",{})
    for r in used.get("drivetrain_shares",[]):
        s=sum(float(r.get(k,0)) for k in ("Petrol","Diesel","NonPluginHybrid","PluginHybrid","BatteryElectric","Others"))
        if not 98.5<=s<=101.5:
            raise ValueError(f"used shares do not sum to ~100 for {r.get('ym')}: {s}")
    fuel=load_json(data_dir/"road_fuels.json",{})
    if len(fuel.get("rows",[]))<60: raise ValueError("road-fuel history unexpectedly short")
    for r in fuel["rows"]:
        for k,lo,hi in (("Diesel",500,5000),("Petrol",300,3000)):
            v=r.get(k)
            if v is not None and not lo<=float(v)<=hi:
                raise ValueError(f"{k} value implausible for {r.get('ym')}: {v}")
    charging=load_json(data_dir/"charging_infrastructure.json",{})
    if len(charging.get("rows",[]))<5: raise ValueError("charging history unexpectedly short")
    h2=load_json(data_dir/"hydrogen_stations.json",{})
    if len(h2.get("rows",[]))<2: raise ValueError("hydrogen history unexpectedly short")


def main() -> int:
    ap=argparse.ArgumentParser()
    ap.add_argument("--data-dir",default="data")
    ap.add_argument("--no-network",action="store_true")
    ap.add_argument("--verbose",action="store_true")
    args=ap.parse_args()
    data_dir=Path(args.data_dir)
    data_dir.mkdir(parents=True,exist_ok=True)

    if args.no_network:
        validate_all(data_dir)
        print("Local datasets validated; network update skipped.")
        return 0

    jobs=[
        ("new_registrations", lambda: update_new_registrations(data_dir/"new_registrations.json", args.verbose)),
        ("used_transfers", lambda: update_used_transfers(data_dir/"used_transfers.json", data_dir/"new_registrations.json", args.verbose)),
        ("road_fuels", lambda: update_road_fuels(data_dir/"road_fuels.json", args.verbose)),
        ("charging", lambda: update_charging(data_dir/"charging_infrastructure.json", args.verbose)),
        ("hydrogen", lambda: update_hydrogen(data_dir/"hydrogen_stations.json", args.verbose)),
    ]
    outcomes={}
    for name,fn in jobs:
        print(f"Updating {name}...")
        try:
            out=fn()
            outcomes[name]=out
            print(f"  OK · latest {out.latest or 'n/a'} · {'changed' if out.changed else 'no file change'}")
        except Exception as e:
            outcomes[name]=Outcome(False,error=str(e))
            print(f"  WARNING · {e}",file=sys.stderr)

    # Validate the complete repository state before metadata is written.
    validate_all(data_dir)

    meta_path=data_dir/"metadata.json"
    old_meta=load_json(meta_path,{"schema_version":1,"datasets":{}})
    meta={"schema_version":1,"checked_at":now_iso(),"datasets":deepcopy(old_meta.get("datasets",{}))}
    labels={
        "new_registrations":("New registrations","monthly","KBA / VDA + consolidated history"),
        "used_transfers":("Used transfers","monthly","KBA FZ 9"),
        "road_fuels":("Road fuels","monthly with publication lag","BAFA"),
        "charging":("Charging infrastructure","monthly","Bundesnetzagentur"),
        "hydrogen":("Hydrogen stations","irregular snapshots","IPHE / H2 Mobility"),
    }
    for key,out in outcomes.items():
        label,cadence,source=labels[key]
        previous=meta["datasets"].get(key,{})
        meta["datasets"][key]={
            "label":label,
            "status":"ok" if out.ok else "warning",
            "latest":out.latest or previous.get("latest","unknown"),
            "cadence":cadence,
            "source":source,
            "note":out.note if out.ok else f"Update failed; last known-good data retained. {out.error}",
        }
    write_json_if_changed(meta_path,meta)

    warnings=[k for k,o in outcomes.items() if not o.ok]
    if warnings:
        print("Completed with source warnings (last known-good data retained): "+", ".join(warnings))
    else:
        print("All source checks completed successfully.")
    return 0


if __name__=="__main__":
    raise SystemExit(main())
