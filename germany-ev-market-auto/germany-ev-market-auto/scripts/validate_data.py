#!/usr/bin/env python3
"""Fail the CI job if a committed dataset is structurally or numerically implausible."""
from __future__ import annotations
import json, sys
from pathlib import Path

FUEL_KEYS=("Diesel","Petrol","NonPluginHybrid","PluginHybrid","BatteryElectric","Hydrogen","Others")

def load(p:Path):
    return json.loads(p.read_text(encoding="utf-8"))

def main():
    data=Path(sys.argv[1] if len(sys.argv)>1 else "data")

    new=load(data/"new_registrations.json")
    rows=new.get("rows",[])
    assert len(rows)>=100, "new registrations unexpectedly short"
    months={}
    for r in rows:
        assert r["fuel"] in FUEL_KEYS
        assert int(r["value"])>=0
        months.setdefault(r["ym"],0)
        months[r["ym"]]+=int(r["value"])
    for ym,total in sorted(months.items())[-12:]:
        assert 50_000<=total<=500_000, f"implausible new-registration total for {ym}: {total}"

    used=load(data/"used_transfers.json")
    for r in used.get("drivetrain_shares",[]):
        s=sum(float(r.get(k,0)) for k in ("Petrol","Diesel","NonPluginHybrid","PluginHybrid","BatteryElectric","Others"))
        assert 98.5<=s<=101.5, f"used shares do not sum to ~100 for {r['ym']}: {s}"
    for r in used.get("market_totals",[]):
        assert 100_000<=int(r["Used"])<=1_000_000
        if r.get("New") is not None:
            assert 50_000<=int(r["New"])<=500_000

    fuels=load(data/"road_fuels.json")
    assert len(fuels.get("rows",[]))>=60
    for r in fuels["rows"]:
        d=r.get("Diesel"); p=r.get("Petrol")
        if d is not None: assert 500<=float(d)<=5000, (r["ym"],d)
        if p is not None: assert 300<=float(p)<=3000, (r["ym"],p)

    charging=load(data/"charging_infrastructure.json")
    assert len(charging.get("rows",[]))>=5
    for r in charging["rows"]:
        assert 0<int(r["normal"])<1_000_000
        assert 0<int(r["fast"])<1_000_000

    h2=load(data/"hydrogen_stations.json")
    assert len(h2.get("rows",[]))>=2
    for r in h2["rows"]:
        assert 10<=int(r["count"])<=500

    meta=load(data/"metadata.json")
    assert "checked_at" in meta and "datasets" in meta
    print("All committed datasets validated.")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
