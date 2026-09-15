"""
Regenerates data/<year>.json for rookie-draft-guide from the master Excel workbook
(Final_Rookie_Markers.xlsx). Excel formulas already compute Missing/Markers/Final.Grade/etc;
this script only transcribes those computed values into the site's JSON schema — it does not
reimplement any grading logic. It reads the QB/RB/WR/TE sheets and nothing else: not the
pre-draft IMPORT staging sheets, not COLLEGE IDS, and not Final Rankings (see SOURCE_SHEETS).

Usage:
    python export_from_workbook.py [--year 2025] [--out-dir data] [--check]

--check compares the freshly exported records against the existing data/<year>.json
(matched by cfr_id) and prints any field where the values differ by more than a small
tolerance, without writing anything.
"""
import argparse
import json
import sys
from pathlib import Path

import openpyxl

WORKBOOK_PATH = Path(r"C:\Users\hayes\Documents\Fantasy Database\Final_Rookie_Markers.xlsx")
REPO_DATA_DIR = Path(__file__).resolve().parent.parent / "data"

# Columns A-T are laid out identically across the QB/RB/WR/TE sheets.
COMMON_MAP = {
    "Position": "pos",
    "Player": "name",
    "College": "school",
    "Conference": "conf",
    "CFR_ID": "cfr_id",
    "PFR_ID": "pfr_id",
    "Age": "age",
    "Round": "round",
    "Overall": "overall",
    "Years": "years",
}

POSITION_MAP = {
    "QB": {
        "Career.PA": "careerPA", "Career.PYd": "careerPYd", "Career.PTD": "careerPTD",
        "Career.INT": "careerINT", "Career.CMP%": "careerCMP", "Career.PAPG": "careerPAPG",
        "Career.YPA": "careerYPA", "Career.YPG": "careerYPG", "Career.TDPG": "careerTDPG",
        "Career.INTPG": "careerINTPG", "Career.RuYd": "careerRuYd", "Career.RuYPG": "careerRuYPG",
        "Best.PYd": "bestPYd", "Best.PTD": "bestPTD", "Best.INT": "bestINT",
        "Best.CMP%": "bestCMP", "Best.RuYd": "bestRuYd", "Last.PPG": "lastPPG",
        "Best.PPG": "bestPPG", "Weight": "wt", "Height": "ht", "BMI": "bmi",
        "Career.Acc%": "careerAcc", "Career.BTT": "careerBTTcount", "Career.BTT%": "careerBTT",
        "Career.TWP%": "careerTWP", "Career.Sack%": "careerSackPct", "Best.Acc%": "bestAcc",
        "Best.BTT": "bestBTTcount", "Best.BTT%": "bestBTT", "Best.TWP%": "bestTWP",
        "Best.Sack%": "bestSackPct", "40-Dash": "dash", "WaSS": "wass", "Burst Score": "burst",
    },
    "RB": {
        "YPC": "ypc", "YPTP": "yptp", "Best Receptions": "bestRec",
        "Best Total Yards Y1/Y2": "bestTotalYds", "First PPG": "firstPPG", "Last PPG": "lastPPG",
        "Best PPG": "bestPPG", "Career.BA/A": "careerBA", "Career.MTF/A": "careerMTF",
        "Career.1D/A": "career1D", "Career.YACO/A": "careerYACO", "Career.YPRR": "careerYPRR",
        "Career.TPRR": "careerTPRR", "Best.BA/A": "bestBA", "Best.MTF/A": "bestMTF",
        "Best.1D/A": "best1D", "Best.YACO/A": "bestYACO", "Best.YPRR": "bestYPRR",
        "Best.TPRR": "bestTPRR", "Weight": "wt", "Height": "ht", "BMI": "bmi",
        "40-Dash": "dash", "WaSS": "wass", "Burst Score": "burst",
    },
    "WR": {
        "BOY20": "boy20", "BOY30": "boy30", "BOA20": "boa20", "BOA30": "boa30",
        "First.YPTPA": "firstYPTPA", "Last.YPTPA": "lastYPTPA", "Best.YPTPA": "bestYPTPA",
        "First.MSYards": "firstMSYards", "Best.MSYards": "bestMSYards",
        "Career.YPRR": "careerYPRR", "Career.YPRR.Zone": "careerYPRRzone",
        "Career.TPRR": "careerTPRR", "Career.1DRR": "career1DRR", "Career.YACR": "careerYACR",
        "Career.CTR": "careerCTR", "Career.CR": "careerCR", "Career.MTF": "careerMTF",
        "Career.Slot Rate": "careerSlot", "Best.YPRR": "bestYPRR", "Best.TPRR": "bestTPRR",
        "Best.1DRR": "best1DRR", "Best.YACR": "bestYACR", "High.CTR": "highCTR",
        "High.CR": "highCR", "High.MTF": "highMTF", "High.Slot Rate": "highSlot",
        "Last.PPG": "lastPPG", "Best.PPG": "bestPPG", "Weight": "wt", "Height": "ht",
        "BMI": "bmi", "40-Dash": "dash", "WaSS": "wass", "Burst Score": "burst",
    },
    "TE": {
        "BOY15": "boy15", "BOY20": "boy20", "BOA20": "boa20", "YPR": "careerYPR",
        "YPG": "careerYPG", "First.YPTPA": "firstYPTPA", "Last.YPTPA": "lastYPTPA",
        "Best.YPTPA": "bestYPTPA", "Best.MSYards": "bestMSYards", "Career.YPRR": "careerYPRR",
        "Career.TPRR": "careerTPRR", "Career.1DRR": "career1DRR", "Career.YACR": "careerYACR",
        "Career.CTR": "careerCTR", "Career.CR": "careerCR", "Career.MTF": "careerMTF",
        "Career.Slot Rate": "careerSlot", "Average.PassBlock": "avgPassBlock",
        "Average.RunBlock": "avgRunBlock", "Best.YPRR": "bestYPRR", "Best.TPRR": "bestTPRR",
        "Best.1DRR": "best1DRR", "Best.YACR": "bestYACR", "High.MTF": "highMTF",
        "High.Slot Rate": "highSlot", "Last.PPG": "lastPPG", "Best.PPG": "bestPPG",
        "Weight": "wt", "Height": "ht", "BMI": "bmi", "40-Dash": "dash", "WaSS": "wass",
        "Burst Score": "burst",
    },
}

GRADE_ONE_DECIMAL = {"grade", "finalGrade"}

# The workbook renamed this column Advanced.Markers -> Adv.Markers on 2026-09-14. Accept
# either spelling, resolved once per sheet against the header row, and raise if neither is
# present: the old code read it with raw.get(), so a rename silently dropped "pff" from every
# record and the site published blank Advanced Markers with nothing anywhere saying why.
ADV_MARKER_COLS = ("Adv.Markers", "Advanced.Markers")

# The ONLY sheets this script may read. The workbook also has a "Final Rankings" tab carrying
# an Adj.Markers column, and this script used to source adjMarkers from it -- but that tab is a
# stale snapshot: its Final.Grade disagrees with the position sheets on 1,242 of 1,257 shared
# rows (Ashton Jeanty 100.0 there vs 97.1 on the RB sheet), so nothing on it can be trusted.
# adjMarkers has no home on the position sheets, so it is now carried forward from the
# published JSON via PRESERVED_FIELDS and is no longer refreshed from the workbook at all.
SOURCE_SHEETS = ("QB", "RB", "WR", "TE")


def round_value(json_key, value):
    if value is None or value == "":
        return None
    if not isinstance(value, (int, float)):
        return value
    if json_key in GRADE_ONE_DECIMAL:
        return round(value, 1)
    if json_key in ("markers", "missing"):
        return int(round(value))
    if isinstance(value, float):
        return round(value, 4)
    return value


def export_position(wb, pos):
    ws = wb[pos]
    headers = [c.value for c in ws[2]]
    adv_col = next((c for c in ADV_MARKER_COLS if c in headers), None)
    if adv_col is None:
        raise KeyError(
            "%s sheet has none of %s -- Advanced Markers column renamed again?"
            % (pos, ", ".join(ADV_MARKER_COLS))
        )
    records_by_year = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        raw = {headers[i]: v for i, v in enumerate(row) if headers[i]}
        if not raw.get("Player") or raw.get("Year") is None:
            continue
        p = {}
        p["year"] = raw["Year"]
        for excel_col, json_key in COMMON_MAP.items():
            p[json_key] = round_value(json_key, raw.get(excel_col))
        p["missing"] = round_value("missing", raw.get("Missing"))
        p["markers"] = round_value("markers", raw.get("Markers"))
        p["grade"] = round_value("grade", raw.get("Final.Grade"))
        p["finalGrade"] = p["grade"]
        p["ath"] = round_value("ath", raw.get("Athleticism%"))
        p["prod"] = round_value("prod", raw.get("Production%"))
        p["pff"] = round_value("pff", raw.get(adv_col))
        for excel_col, json_key in POSITION_MAP[pos].items():
            val = round_value(json_key, raw.get(excel_col))
            if val is not None:
                p[json_key] = val
        # drop None-valued common fields to match existing sparse style
        p = {k: v for k, v in p.items() if v is not None}
        records_by_year.setdefault(raw["Year"], []).append(p)
    return records_by_year


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, help="Only export this draft year")
    ap.add_argument("--out-dir", default=str(REPO_DATA_DIR))
    ap.add_argument("--check", action="store_true", help="Diff against existing JSON instead of writing")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(WORKBOOK_PATH, data_only=True)

    all_years = {}
    for pos in SOURCE_SHEETS:
        for year, records in export_position(wb, pos).items():
            all_years.setdefault(year, []).extend(records)

    years = [args.year] if args.year else sorted(all_years)
    out_dir = Path(args.out_dir)

    for year in years:
        records = all_years.get(year, [])
        if not records:
            print(f"{year}: no rows found in workbook, skipping")
            continue
        out_path = out_dir / f"{year}.json"
        if not out_path.exists() and not args.year:
            # The workbook gets next year's prospects long before they are a published class
            # (an ungraded placeholder row is enough to create a year). Adding a class to the
            # site is a deliberate act: pass --year to do it.
            print(f"{year}: no existing {out_path.name}, skipping (pass --year {year} to create it)")
            continue
        if args.check:
            check_against_existing(year, records, out_dir)
        else:
            preserve_unexported_fields(records, out_path)
            # Minified, matching what is already published: these files are fetched by the
            # browser (the grades page pulls ~18 of them), and indent=2 inflates the set
            # from 1.63 MB to 2.40 MB for no reader benefit.
            out_path.write_text(json.dumps(records, separators=(",", ":")), encoding="utf-8")
            print(f"{year}: wrote {len(records)} records to {out_path}")


# Fields the site carries that this exporter cannot produce from the workbook. Without this,
# every regeneration silently deletes them:
#   adjMarkers        hand-curated on the Final Rankings tab, covers only a subset of prospects
#   projected_overall pre-draft projected draft slot, maintained outside the workbook entirely
# Only these are carried forward -- a blanket carry-forward would also resurrect fields that
# were removed on purpose.
PRESERVED_FIELDS = ("adjMarkers", "projected_overall")


def preserve_unexported_fields(records, existing_path):
    """Never overwrite an existing site value with nothing. If the workbook has no value for
    a PRESERVED_FIELDS key, keep whatever is already published for that player."""
    if not existing_path.exists():
        return
    existing = json.loads(existing_path.read_text(encoding="utf-8"))
    existing_by_key = {(p.get("cfr_id") or f"{p.get('pos')}:{p.get('name')}"): p for p in existing}
    for p in records:
        key = p.get("cfr_id") or f"{p.get('pos')}:{p.get('name')}"
        old = existing_by_key.get(key)
        if not old:
            continue
        for field in PRESERVED_FIELDS:
            if field not in p and field in old:
                p[field] = old[field]


def check_against_existing(year, records, out_dir):
    existing_path = out_dir / f"{year}.json"
    if not existing_path.exists():
        print(f"{year}: no existing file to compare ({len(records)} new records)")
        return
    def key(p):
        return p.get("cfr_id") or f"{p.get('pos')}:{p.get('name')}"

    existing = {key(p): p for p in json.loads(existing_path.read_text(encoding="utf-8"))}
    fresh = {key(p): p for p in records}
    missing_in_fresh = set(existing) - set(fresh)
    missing_in_existing = set(fresh) - set(existing)
    if missing_in_fresh:
        print(f"{year}: {len(missing_in_fresh)} players in existing file but not regenerated: {sorted(missing_in_fresh)[:5]}...")
    if missing_in_existing:
        print(f"{year}: {len(missing_in_existing)} players newly regenerated but not in existing file: {sorted(missing_in_existing)[:5]}...")

    diff_count = 0
    for cfr_id in sorted(set(existing) & set(fresh)):
        old, new = existing[cfr_id], fresh[cfr_id]
        for key in set(old) | set(new):
            ov, nv = old.get(key), new.get(key)
            if isinstance(ov, (int, float)) and isinstance(nv, (int, float)):
                if abs(ov - nv) > max(0.05, abs(ov) * 0.005):
                    print(f"{year} {old.get('name', cfr_id)}: {key} existing={ov} fresh={nv}")
                    diff_count += 1
            elif ov != nv:
                print(f"{year} {old.get('name', cfr_id)}: {key} existing={ov!r} fresh={nv!r}")
                diff_count += 1
    print(f"{year}: {len(existing & fresh if False else set(existing)&set(fresh))} players compared, {diff_count} field diffs")


if __name__ == "__main__":
    sys.exit(main())
