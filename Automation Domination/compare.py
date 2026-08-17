"""
compare.py — CAMA vs MLS comparison logic
Extracted from CAMAvsMLSv6-main/streamlit_app_V6.2.py (Streamlit removed).
Produces 4 Excel files in the output folder.
"""

import io
import re
import numpy as np
import pandas as pd
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment

# ── Column mapping (MLS → CAMA) ────────────────────────────────────────────────
UNIQUE_ID_COLUMN = {"mls_col": "Parcel Number", "cama_col": "PARID"}

COLUMNS_TO_COMPARE = [
    {"mls_col": "Above Grade Finished Area", "cama_col": "SFLA"},
    {"mls_col": "Bedrooms Total",            "cama_col": "RMBED"},
    {"mls_col": "Bathrooms Full",            "cama_col": "FIXBATH"},
    {"mls_col": "Bathrooms Half",            "cama_col": "FIXHALF"},
]

COLUMNS_TO_COMPARE_SUM = [
    {"mls_col": "Below Grade Finished Area",
     "cama_cols": ["RECROMAREA", "FINBSMTAREA", "UFEATAREA"]},
]

COLUMNS_TO_COMPARE_CATEGORICAL = [
    {
        "mls_col":             "Cooling",
        "cama_col":            "HEAT",
        "mls_check_contains":  "Central Air",
        "cama_expected_if_true":  1,
        "cama_expected_if_false": 0,
        "case_sensitive":      False,
    },
]

NUMERIC_TOLERANCE = 0.01
SKIP_ZERO_VALUES  = True

# ── Appraiser info (used for SaleTab / MassEntrance output) ───────────────────
APPRAISER_INITIALS  = "JMJ"
APPRAISER_FULL_NAME = "Jason Jeffries"

ADDRESS_COLUMNS = {
    "address": "Address",
    "city":    "City",
    "state":   "State or Province",
    "zip":     "Postal Code",
}

ZILLOW_URL_BASE = "https://www.zillow.com/homes/"

# ── TAXDIST → Zone mapping ─────────────────────────────────────────────────────
TAXDIST_TO_ZONE = {
    "00010": "North East",   "00020": "Canton",       "00025": "Canton",
    "00030": "Canton",       "00035": "Canton",       "00040": "Western",
    "00050": "Western",      "00060": "Western",      "00065": "Western",
    "00070": "Southern",     "00080": "Southern",     "00090": "Southern",
    "00100": "Southern",     "00110": "Southern",     "00112": "Southern",
    "00115": "Southern",     "00120": "Southern",     "00130": "North Western",
    "00140": "North Western","00150": "North Western","00160": "Northern",
    "00170": "Northern",     "00180": "Northern",     "00190": "Northern",
    "00200": "Northern",     "00210": "North Western","00220": "North Western",
    "00230": "North Western","00240": "North Western","00245": "North Western",
    "00250": "North East",   "00260": "North East",   "00270": "North East",
    "00280": "North East",   "00290": "North East",   "00300": "North East",
    "00305": "North East",   "00310": "North East",   "00320": "North East",
    "00330": "North East",   "00340": "Southern",     "00345": "Southern",
    "00350": "Southern",     "00355": "Southern",     "00360": "Southern",
    "00370": "Southern",     "00380": "Southern",     "00390": "Southern",
    "00400": "Western",      "00410": "Western",      "00415": "Southern",
    "00420": "Western",      "00430": "Western",      "00440": "Western",
    "00445": "Southern",     "00450": "Southern",     "00460": "Southern",
    "00470": "Southern",     "00480": "Southern",     "00490": "Northern",
    "00520": "Northern",     "00530": "Northern",     "00535": "Northern",
    "00545": "Northern",     "00555": "Northern",     "00560": "Northern",
    "00565": "Canton",       "00570": "Southern",     "00580": "Southern",
    "00590": "Southern",     "00600": "Southern",     "00610": "Southern",
    "00620": "Southern",     "00630": "Southern",     "00640": "Southern",
    "00660": "Southern",     "00670": "Southern",     "00680": "Southern",
    "00690": "Western",      "00700": "Western",      "00710": "Western",
    "00715": "North East",   "00720": "North East",   "00725": "North East",
    "00730": "North East",   "00740": "North East",   "00750": "North East",
}


# ── Helper functions ───────────────────────────────────────────────────────────

def lookup_zone(taxdist):
    if pd.isna(taxdist) or str(taxdist).strip() == "":
        return "Unknown"
    try:
        key = str(int(float(str(taxdist)))).zfill(5)
    except (ValueError, TypeError):
        key = str(taxdist).strip().zfill(5)
    return TAXDIST_TO_ZONE.get(key, "Unknown")


def format_zillow_url(address, city, state, zip_code):
    if pd.isna(address) or pd.isna(city) or pd.isna(zip_code):
        return ""
    address_clean = re.sub(r"\s+(Apt|Unit|#|Suite)\s*[\w-]*$", "",
                           str(address).strip(), flags=re.IGNORECASE)
    addr_fmt = re.sub(r"\s+", "-", re.sub(r"[^\w\s-]", "", address_clean))
    city_fmt  = re.sub(r"\s+", "-", re.sub(r"[^\w\s-]", "", str(city).strip()))
    zip_match = re.search(r"^\d{5}", str(zip_code).strip())
    zip_fmt   = zip_match.group() if zip_match else str(zip_code).strip().split("-")[0]
    return f"{ZILLOW_URL_BASE}{addr_fmt}-{city_fmt}-OH-{zip_fmt}_rb/"


def values_equal(val1, val2):
    try:
        n1 = pd.to_numeric(val1, errors="raise")
        n2 = pd.to_numeric(val2, errors="raise")
        if pd.isna(n1) and pd.isna(n2):
            return True
        if pd.isna(n1) != pd.isna(n2):
            return False
        return np.isclose(n1, n2, equal_nan=False, rtol=1e-9, atol=NUMERIC_TOLERANCE)
    except (ValueError, TypeError):
        s1 = str(val1).strip().lower() if pd.notna(val1) else ""
        s2 = str(val2).strip().lower() if pd.notna(val2) else ""
        return s1 == s2


def categorical_match(mls_val, cama_val, mapping):
    check_text   = mapping.get("mls_check_contains", "")
    exp_true     = mapping.get("cama_expected_if_true")
    exp_false    = mapping.get("cama_expected_if_false")
    case_sens    = mapping.get("case_sensitive", False)
    mls_str      = str(mls_val).strip() if pd.notna(mls_val) else ""
    if not case_sens:
        mls_str    = mls_str.lower()
        check_text = check_text.lower()
    expected = exp_true if (check_text in mls_str) else exp_false
    try:
        cn = pd.to_numeric(cama_val, errors="coerce")
        en = pd.to_numeric(expected,  errors="coerce")
        if pd.isna(cn) and pd.isna(en):
            return True
        if pd.isna(cn) or pd.isna(en):
            return False
        return np.isclose(cn, en, equal_nan=False, rtol=1e-9, atol=NUMERIC_TOLERANCE)
    except Exception:
        return str(cama_val).strip().lower() == str(expected).strip().lower()


def calculate_difference(val1, val2):
    try:
        n1 = pd.to_numeric(val1, errors="raise")
        n2 = pd.to_numeric(val2, errors="raise")
        if pd.isna(n1) or pd.isna(n2):
            return "N/A"
        return f"{n1 - n2:,.2f}"
    except (ValueError, TypeError):
        return "Text difference"


# ── Batch search file for missing_in_mls ──────────────────────────────────────

def _build_batch_excel(df_missing, batch_size=35) -> bytes:
    """Format missing-in-MLS parcels into batched tabs for MLS paste search."""
    output = io.BytesIO()
    parcel_ids = df_missing["Parcel_ID"].tolist()
    formatted  = []
    for p in parcel_ids:
        try:
            formatted.append(str(int(float(str(p)))).zfill(8))
        except (ValueError, TypeError):
            formatted.append(str(p).zfill(8))

    total      = len(formatted)
    num_batches = max(1, (total + batch_size - 1) // batch_size)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary = {
            "Total Parcels":    [total],
            "Batch Size":       [batch_size],
            "Number of Batches":[num_batches],
            "Generated":        [datetime.now().strftime("%m/%d/%Y %H:%M")],
        }
        pd.DataFrame(summary).to_excel(writer, index=False, sheet_name="Summary")
        ws = writer.sheets["Summary"]
        for col in ["A", "B", "C", "D"]:
            ws.column_dimensions[col].width = 20

        for i in range(num_batches):
            start  = i * batch_size
            end    = min(start + batch_size, total)
            batch  = formatted[start:end]
            comma  = ",".join(batch)
            sname  = f"Batch {i+1} ({start+1}-{end})"
            pd.DataFrame([[]]).to_excel(writer, index=False, header=False, sheet_name=sname)
            ws2 = writer.sheets[sname]
            ws2["A1"] = (f"Batch {i+1} — {len(batch)} parcels — "
                         f"copy cell A2 and paste into MLS parcel search")
            ws2["A2"] = comma
            ws2.column_dimensions["A"].width = 120
            ws2["A1"].font  = Font(bold=True, color="FFFFFF", size=11)
            ws2["A1"].fill  = PatternFill(fill_type="solid", fgColor="1F4E79")
            ws2["A2"].alignment = Alignment(wrap_text=False)
            ws2.row_dimensions[2].height = 20

    output.seek(0)
    return output.getvalue()


# ── SaleTab / MassEntrance generation ────────────────────────────────────────

def _generate_mass_update_files(combined_df, output_folder, date_str):
    """
    Generate SALETAB_MassUpdate_MMDDYY.xlsx and MassEntrance_MMDDYY.csv
    from the combined perfect_matches + value_mismatches DataFrame.
    Mirrors the logic in streamlit_app_V6.2.py::generate_mass_update_files().
    """
    timestamp = datetime.strptime(date_str, "%Y-%m-%d").strftime("%m%d%y")
    user_initials  = APPRAISER_INITIALS
    user_full_name = APPRAISER_FULL_NAME
    now_str = datetime.now().strftime("%m/%d/%Y %H:%M")
    today_str = datetime.now().strftime("%Y-%m-%d")

    listing_col = ("Listing_Number" if "Listing_Number" in combined_df.columns
                   else ("Listing #" if "Listing #" in combined_df.columns else None))

    unique_df = combined_df.drop_duplicates(subset=["Parcel_ID"], keep="first").copy()

    saletab_rows  = []
    entrance_rows = []

    for _, row in unique_df.iterrows():
        main_parcel            = str(row["Parcel_ID"]).strip()
        salekey_str              = str(row.get("SALEKEY", "")).strip()
        additional_parcels_str   = str(row.get("ADDITIONAL_PARCELS", "")).strip()
        additional_salekeys_str  = str(row.get("ADDITIONAL_SALEKEYS", "")).strip()

        listing_number = 0
        if listing_col:
            listing_str = str(row.get(listing_col, "")).strip()
            if listing_str and listing_str != "nan":
                first = listing_str.rstrip(",").split(",")[0].strip()
                try:
                    listing_number = int(float(first))
                except (ValueError, TypeError):
                    listing_number = 0

        # Use only the first salekey / main parcel (matches manual behavior)
        first_salekey = ""
        if salekey_str and salekey_str != "nan":
            first_salekey = salekey_str.rstrip(",").split(",")[0].strip()

        try:
            salekey_int = int(first_salekey)
            parcel_int  = int(main_parcel)
            saletab_rows.append({
                "PARID":   parcel_int,
                "SALEKEY": salekey_int,
                "USER11":  listing_number,
                "SOURCE":  0,
                "SALEVAL": 0,
                "USER1":   user_initials,
                "USER2":   today_str,
            })
            entrance_rows.append({
                "Change Type":            "existing",
                "appraiser":              user_full_name,
                "parcelnum":              parcel_int,
                "comment":                "",
                "Review Status":          "Reviewed",
                "Determination":          "",
                "Est. Value Change":      "",
                "Last Changed Date/Time": now_str,
                "Last Changed By":        user_full_name,
            })
            # Add rows for any additional parcels in the same multi-part sale.
            # Use ADDITIONAL_SALEKEYS (same positional order as ADDITIONAL_PARCELS)
            # so each parcel gets its own correct salekey, not the lead parcel's key.
            if additional_parcels_str and additional_parcels_str != "nan":
                add_salekeys = []
                if additional_salekeys_str and additional_salekeys_str != "nan":
                    add_salekeys = [s.strip() for s in additional_salekeys_str.split(",") if s.strip()]
                for idx, add_parcel_str in enumerate(additional_parcels_str.split(",")):
                    add_parcel_str = add_parcel_str.strip()
                    if not add_parcel_str:
                        continue
                    try:
                        add_salekey_int = int(add_salekeys[idx]) if idx < len(add_salekeys) else salekey_int
                        add_parcel_int = int(float(add_parcel_str))
                        saletab_rows.append({
                            "PARID":   add_parcel_int,
                            "SALEKEY": add_salekey_int,
                            "USER11":  listing_number,
                            "SOURCE":  0,
                            "SALEVAL": 0,
                            "USER1":   user_initials,
                            "USER2":   today_str,
                        })
                        entrance_rows.append({
                            "Change Type":            "existing",
                            "appraiser":              user_full_name,
                            "parcelnum":              add_parcel_int,
                            "comment":                "",
                            "Review Status":          "Reviewed",
                            "Determination":          "",
                            "Est. Value Change":      "",
                            "Last Changed Date/Time": now_str,
                            "Last Changed By":        user_full_name,
                        })
                    except (ValueError, TypeError):
                        pass
        except (ValueError, TypeError):
            pass

    saletab_df = pd.DataFrame(saletab_rows)
    if not saletab_df.empty:
        saletab_df = (saletab_df
                      .drop_duplicates(subset=["PARID"], keep="first")
                      .sort_values("SALEKEY")
                      .reset_index(drop=True))

    entrance_df = pd.DataFrame(entrance_rows)

    saletab_path  = output_folder / f"SALETAB_MassUpdate_{timestamp}.xlsx"
    entrance_path = output_folder / f"MassEntrance{timestamp}.csv"

    with pd.ExcelWriter(saletab_path, engine="openpyxl") as writer:
        saletab_df.to_excel(writer, index=False, sheet_name="Sheet1")

    entrance_df.to_csv(entrance_path, index=False)

    print(f"  ✓ SALETAB_MassUpdate_{timestamp}.xlsx  ({len(saletab_df)} rows)")
    print(f"  ✓ MassEntrance{timestamp}.csv  ({len(entrance_df)} rows)")


# ── Fuzzy parcel-typo fallback ────────────────────────────────────────────────
#
# Why: MLS agents occasionally type the parcel number wrong (transposed digits,
# off-by-one). The exact-parcel merge then drops the sale into BOTH
# missing_in_cama (MLS-side, bogus parcel) and missing_in_mls (CAMA-side, real
# parcel), and the sale never reaches the portal.
#
# This function re-scans those two residuals and surfaces likely re-pairings
# where city + close price + closed date all line up. It writes a separate
# "likely_parcel_mismatches" workbook — it does NOT auto-merge — so a human
# can confirm and correct the parcel before it flows through the rest of the
# pipeline.
#
# Match rule: same city (case-insensitive), close-price equal within $1,
# closed-date within ±FUZZY_DATE_WINDOW_DAYS.

FUZZY_DATE_WINDOW_DAYS = 7
FUZZY_PRICE_TOLERANCE  = 1.0   # dollars


def _find_likely_parcel_mismatches(merged: pd.DataFrame,
                                   output_folder: Path,
                                   date_str: str) -> int:
    """
    Scan the outer-merge residuals (left_only = missing in CAMA,
    right_only = missing in MLS) and write any (city + price + date±7) hits
    to likely_parcel_mismatches_{date_str}.xlsx for human review.

    Returns the number of likely-match rows written.
    """
    left_only  = merged[merged["_merge"] == "left_only"].copy()
    right_only = merged[merged["_merge"] == "right_only"].copy()

    if left_only.empty or right_only.empty:
        return 0

    # MLS side — normalise the fuzzy keys
    left_only["_mls_city"]  = left_only.get("City", "").astype(str).str.strip().str.upper()
    left_only["_mls_price"] = pd.to_numeric(left_only.get("Close Price"), errors="coerce")
    left_only["_mls_date"]  = pd.to_datetime(left_only.get("Closed Date"), errors="coerce")

    # CAMA side
    right_only["_cama_city"]  = right_only.get("CITYNAME", "").astype(str).str.strip().str.upper()
    right_only["_cama_price"] = pd.to_numeric(right_only.get("PRICE"), errors="coerce")
    right_only["_cama_date"]  = pd.to_datetime(right_only.get("SALEDT"), errors="coerce")

    # Drop rows we can't compare (missing any of city / price / date)
    left_only  = left_only.dropna(subset=["_mls_price", "_mls_date"])
    right_only = right_only.dropna(subset=["_cama_price", "_cama_date"])
    left_only  = left_only[left_only["_mls_city"]  != ""]
    right_only = right_only[right_only["_cama_city"] != ""]
    if left_only.empty or right_only.empty:
        return 0

    # Cross-join on city (cheap filter), then apply price + date tolerances.
    candidates = left_only.merge(
        right_only, left_on="_mls_city", right_on="_cama_city",
        how="inner", suffixes=("_mls", "_cama"),
    )
    if candidates.empty:
        return 0

    price_ok = (candidates["_mls_price"] - candidates["_cama_price"]).abs() <= FUZZY_PRICE_TOLERANCE
    date_diff = (candidates["_mls_date"] - candidates["_cama_date"]).dt.days.abs()
    date_ok  = date_diff <= FUZZY_DATE_WINDOW_DAYS
    candidates = candidates[price_ok & date_ok].copy()
    if candidates.empty:
        return 0

    candidates["Date_Diff_Days"] = date_diff[price_ok & date_ok]

    # left_only and right_only share every column (both slices of the outer
    # merge), so the cross-merge with suffixes=("_mls","_cama") renames them
    # all — including PARID. MLS-origin fields are real on the _mls side only;
    # CAMA-origin fields on the _cama side only.
    def _disp(p):
        try:
            return str(int(float(str(p))))
        except (ValueError, TypeError):
            return str(p).lstrip("0") or str(p)

    def _parcel_digit_diff(a: str, b: str) -> int:
        """Count digit positions that differ when both parcels are zero-padded
        to the same length. A real transposed-digit typo scores 2; unrelated
        parcels score 4+. Reviewers can sort by this — small = likely typo."""
        a, b = str(a), str(b)
        w = max(len(a), len(b))
        return sum(1 for x, y in zip(a.zfill(w), b.zfill(w)) if x != y)

    out_rows = []
    for _, r in candidates.iterrows():
        mls_p  = _disp(r.get("PARID_mls", ""))
        cama_p = _disp(r.get("PARID_cama", ""))
        out_rows.append({
            "MLS_Parcel_(typo?)":  mls_p,
            "CAMA_Parcel_(real)":  cama_p,
            "Parcel_Digit_Diff":   _parcel_digit_diff(mls_p, cama_p),
            "MLS_Listing":         r.get("Listing #_mls", ""),
            "MLS_Address":         r.get("Address_mls", ""),
            "City":                r.get("_mls_city", ""),
            "MLS_Closed_Date":     r["_mls_date"].strftime("%Y-%m-%d") if pd.notna(r["_mls_date"]) else "",
            "CAMA_SALEDT":         r["_cama_date"].strftime("%Y-%m-%d") if pd.notna(r["_cama_date"]) else "",
            "Date_Diff_Days":      int(r["Date_Diff_Days"]),
            "MLS_Close_Price":     r["_mls_price"],
            "CAMA_PRICE":          r["_cama_price"],
            "CAMA_SALEKEY":        r.get("SALEKEY_cama", ""),
            "CAMA_OWN1":           r.get("OWN1_cama", ""),
            "CAMA_OLDOWN":         r.get("OLDOWN_cama", ""),
        })
    # Sort: most-likely typos first (smallest digit-diff, then smallest date gap).
    out_rows.sort(key=lambda r: (r["Parcel_Digit_Diff"], r["Date_Diff_Days"]))

    path = output_folder / f"likely_parcel_mismatches_{date_str}.xlsx"
    pd.DataFrame(out_rows).to_excel(path, index=False)
    print(f"  ✓ likely_parcel_mismatches_{date_str}.xlsx  ({len(out_rows)} rows)  "
          f"← review these; parcel # in MLS may be wrong")
    return len(out_rows)


# ── Main comparison ────────────────────────────────────────────────────────────

def run_comparison(df_mls: pd.DataFrame, df_cama: pd.DataFrame,
                   output_folder: Path, date_str: str) -> dict:
    """
    Compare MLS and CAMA DataFrames, write 4 Excel files to output_folder.
    date_str format: 'YYYY-MM-DD'
    Returns dict with counts: mismatches, perfects, missing_cama, missing_mls.
    """
    mls_id  = UNIQUE_ID_COLUMN["mls_col"]
    cama_id = UNIQUE_ID_COLUMN["cama_col"]

    if mls_id not in df_mls.columns:
        raise ValueError(f"Column '{mls_id}' not found in MLS data. "
                         f"Available: {list(df_mls.columns)}")
    if cama_id not in df_cama.columns:
        raise ValueError(f"Column '{cama_id}' not found in CAMA data. "
                         f"Available: {list(df_cama.columns)}")

    # Fill MLS blanks with 0 — replicates the original app's auto-fill step
    # (replaces the manual CTRL+H blank-to-zero the user did each week).
    df_mls  = df_mls.copy().fillna(0).replace('', 0)
    df_cama = df_cama.copy()

    # Normalise parcel ID: strip whitespace, zero-pad to 8 digits
    # MLS exports 8-digit zero-padded IDs (e.g. 00101803);
    # CAMA stores variable-length IDs (e.g. 101803) — zfill(8) aligns them.
    def _norm_parcel(s):
        s = str(s).strip()
        try:
            return str(int(float(s))).zfill(8)   # handles '106948.0' → '00106948'
        except (ValueError, TypeError):
            return s.zfill(8)

    df_mls[mls_id]   = df_mls[mls_id].apply(_norm_parcel)
    df_cama[cama_id] = df_cama[cama_id].apply(_norm_parcel)

    # Rename MLS parcel column to match CAMA for merge
    df_mls_r = df_mls.rename(columns={mls_id: cama_id})

    merged = pd.merge(df_mls_r, df_cama, on=cama_id, how="outer", indicator=True)

    missing_in_cama = []
    missing_in_mls  = []
    value_mismatches = []
    perfect_matches  = []
    city_stats = {}   # city -> {total, matched}

    def _display_parcel(padded):
        """Strip leading zeros for display, iasWorld links, and photo lookups."""
        try:
            return str(int(float(str(padded))))
        except (ValueError, TypeError):
            return str(padded).lstrip("0") or str(padded)

    for _, row in merged.iterrows():
        record_id    = _display_parcel(row.get(cama_id))
        merge_status = row.get("_merge")

        if merge_status == "left_only":
            missing_in_cama.append({
                "Parcel_ID":      record_id,
                "Listing_Number": row.get("Listing #", ""),
                "Closed_Date":    row.get("Closed Date", ""),
            })

        elif merge_status == "right_only":
            missing_in_mls.append({"Parcel_ID": record_id})
            cityname = str(row.get("CITYNAME", "")).strip().upper()
            if cityname and cityname != "NAN":
                cs = city_stats.setdefault(cityname, {"total": 0, "matched": 0})
                cs["total"] += 1

        elif merge_status == "both":
            listing_num         = row.get("Listing #", "")
            salekey             = row.get("SALEKEY", "")
            nopar               = row.get("NOPAR", "")
            additional_parcels  = row.get("ADDITIONAL_PARCELS", "")
            additional_salekeys = row.get("ADDITIONAL_SALEKEYS", "")
            address   = row.get(ADDRESS_COLUMNS["address"], "")
            city      = row.get(ADDRESS_COLUMNS["city"], "")
            state     = row.get(ADDRESS_COLUMNS["state"], "")
            zip_code  = row.get(ADDRESS_COLUMNS["zip"], "")
            yrblt     = row.get("YRBLT", "")
            effyr     = row.get("EFFYR", "")
            grade     = row.get("GRADE", "")
            cdu       = row.get("CDU", "")
            remarks   = row.get("Public Remarks", "")
            taxdist   = row.get("TAXDIST", "")
            zone      = lookup_zone(taxdist)
            fixbath   = row.get("FIXBATH", "")
            fixhalf   = row.get("FIXHALF", "")
            sfla      = row.get("SFLA", "")
            cityname  = str(row.get("CITYNAME", "")).strip().upper()
            if cityname and cityname != "NAN":
                cs = city_stats.setdefault(cityname, {"total": 0, "matched": 0})
                cs["total"]   += 1
                cs["matched"] += 1

            def _base_record(field_mls, field_cama, mls_val, cama_val, diff):
                return {
                    "Parcel_ID":           record_id,
                    "NOPAR":               nopar,
                    "ADDITIONAL_PARCELS":  additional_parcels,
                    "ADDITIONAL_SALEKEYS": additional_salekeys,
                    "Listing_Number":      listing_num,
                    "SALEKEY":             salekey,
                    "Address":            address,
                    "City":               city,
                    "State":              state,
                    "Zip":                zip_code,
                    "Zone":               zone,
                    "YRBLT":              yrblt,
                    "EFFYR":              effyr,
                    "GRADE":              grade,
                    "CDU":                cdu,
                    "FIXBATH":            fixbath,
                    "FIXHALF":            fixhalf,
                    "SFLA":               sfla,
                    "Public_Remarks":     remarks,
                    "Field_MLS":          field_mls,
                    "Field_CAMA":         field_cama,
                    "MLS_Value":          mls_val,
                    "CAMA_Value":         cama_val,
                    "Difference":         diff,
                    "Zillow_URL":         format_zillow_url(address, city, state, zip_code),
                }

            record_mismatches = []
            fields_compared   = []

            # Numeric comparisons
            for m in COLUMNS_TO_COMPARE:
                mc, cc = m["mls_col"], m["cama_col"]
                if mc not in merged.columns or cc not in merged.columns:
                    continue
                mv, cv = row.get(mc), row.get(cc)
                if (pd.isna(mv) or str(mv).strip() == "" or
                        pd.isna(cv) or str(cv).strip() == ""):
                    continue
                fields_compared.append(mc)
                if SKIP_ZERO_VALUES:
                    mn = pd.to_numeric(mv, errors="coerce")
                    cn = pd.to_numeric(cv, errors="coerce")
                    if pd.notna(mn) and mn == 0 and pd.notna(cn) and cn == 0:
                        continue
                if not values_equal(mv, cv):
                    record_mismatches.append(
                        _base_record(mc, cc, mv, cv, calculate_difference(mv, cv)))

            # Sum comparisons (e.g. Below Grade Finished Area)
            for m in COLUMNS_TO_COMPARE_SUM:
                mc, ccs = m["mls_col"], m["cama_cols"]
                if mc not in merged.columns:
                    continue
                if any(c not in merged.columns for c in ccs):
                    continue
                mv = row.get(mc)
                if pd.isna(mv) or str(mv).strip() == "":
                    continue
                cama_sum = sum(
                    pd.to_numeric(row.get(c, 0), errors="coerce") or 0
                    for c in ccs
                )
                fields_compared.append(mc)
                if SKIP_ZERO_VALUES:
                    mn = pd.to_numeric(mv, errors="coerce")
                    if pd.notna(mn) and mn == 0 and cama_sum == 0:
                        continue
                if not values_equal(mv, cama_sum):
                    record_mismatches.append(
                        _base_record(mc, f"SUM({', '.join(ccs)})",
                                     mv, cama_sum,
                                     calculate_difference(mv, cama_sum)))

            # Categorical comparisons (e.g. Cooling / Central Air)
            for m in COLUMNS_TO_COMPARE_CATEGORICAL:
                mc, cc = m["mls_col"], m["cama_col"]
                if mc not in merged.columns or cc not in merged.columns:
                    continue
                mv, cv = row.get(mc), row.get(cc)
                if (pd.isna(mv) or str(mv).strip() == "" or
                        pd.isna(cv) or str(cv).strip() == ""):
                    continue
                fields_compared.append(mc)
                if not categorical_match(mv, cv, m):
                    check = m.get("mls_check_contains", "")
                    exp   = (m["cama_expected_if_true"]
                             if check.lower() in str(mv).lower()
                             else m["cama_expected_if_false"])
                    rec   = _base_record(mc, cc, mv, cv, "")
                    rec["Expected_CAMA_Value"] = exp
                    rec["Match_Rule"] = (
                        f"If '{check}' in {mc}, then {cc} should be "
                        f"{m['cama_expected_if_true']}, "
                        f"else {m['cama_expected_if_false']}"
                    )
                    record_mismatches.append(rec)

            if record_mismatches:
                value_mismatches.extend(record_mismatches)
            elif fields_compared:
                perfect_matches.append({
                    "Parcel_ID":           record_id,
                    "NOPAR":               nopar,
                    "ADDITIONAL_PARCELS":  additional_parcels,
                    "ADDITIONAL_SALEKEYS": additional_salekeys,
                    "Listing_Number":      listing_num,
                    "SALEKEY":             salekey,
                    "Address":            address,
                    "City":               city,
                    "State":              state,
                    "Zip":                zip_code,
                    "Zone":               zone,
                    "YRBLT":              yrblt,
                    "EFFYR":              effyr,
                    "GRADE":              grade,
                    "CDU":                cdu,
                    "FIXBATH":            fixbath,
                    "FIXHALF":            fixhalf,
                    "SFLA":               sfla,
                    "Public_Remarks":     remarks,
                    "Fields_Compared":    len(fields_compared),
                    "Zillow_URL":         format_zillow_url(address, city, state, zip_code),
                })

    # ── Write Excel files ──────────────────────────────────────────────────────
    output_folder.mkdir(parents=True, exist_ok=True)

    def _save(records, filename):
        path = output_folder / filename
        pd.DataFrame(records).to_excel(path, index=False)
        print(f"  ✓ {filename}  ({len(records)} rows)")
        return path

    _save(value_mismatches,  f"value_mismatches_{date_str}.xlsx")
    _save(perfect_matches,   f"perfect_matches_{date_str}.xlsx")
    _save(missing_in_cama,   f"missing_in_cama_{date_str}.xlsx")

    # missing_in_mls gets the batch-formatted Excel
    mls_path = output_folder / f"missing_in_mls_batches_{date_str}.xlsx"
    df_missing = pd.DataFrame(missing_in_mls)
    if not df_missing.empty:
        mls_path.write_bytes(_build_batch_excel(df_missing))
    else:
        df_missing.to_excel(mls_path, index=False)
    print(f"  ✓ missing_in_mls_batches_{date_str}.xlsx  ({len(missing_in_mls)} rows)")

    # ── Fuzzy parcel-typo fallback ────────────────────────────────────────────
    # Cross-checks the missing_in_cama and missing_in_mls residuals for sales
    # that look like the same transaction under a mistyped parcel number.
    likely_matches = _find_likely_parcel_mismatches(merged, output_folder, date_str)

    # ── City match statistics ──────────────────────────────────────────────────
    if city_stats:
        city_rows = sorted(
            [
                {
                    "City":               city,
                    "Total_CAMA_Parcels":  s["total"],
                    "Matched_Parcels":     s["matched"],
                    "Match_Rate":          round(s["matched"] / s["total"] * 100, 2) if s["total"] else 0,
                    "Not_Matched":         s["total"] - s["matched"],
                }
                for city, s in city_stats.items()
            ],
            key=lambda r: -r["Total_CAMA_Parcels"],
        )
        city_path = output_folder / f"city_match_statistics_{date_str}.csv"
        pd.DataFrame(city_rows).to_csv(city_path, index=False)
        print(f"  ✓ city_match_statistics_{date_str}.csv  ({len(city_rows)} cities)")

    # ── SaleTab + MassEntrance ─────────────────────────────────────────────────
    combined = pd.concat(
        [pd.DataFrame(perfect_matches), pd.DataFrame(value_mismatches)],
        ignore_index=True,
    )
    if not combined.empty:
        _generate_mass_update_files(combined, output_folder, date_str)

    return {
        "mismatches":     len(value_mismatches),
        "perfects":       len(perfect_matches),
        "missing_cama":   len(missing_in_cama),
        "missing_mls":    len(missing_in_mls),
        "likely_matches": likely_matches,
    }
