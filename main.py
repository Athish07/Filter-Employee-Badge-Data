import sys
from pathlib import Path
from typing import List, Dict, Tuple
import pandas as pd

DATA_FILE = Path("data") / "EY Badges Tracker.xlsx"

ALLOWED_STATUS = [
    "In Progress", "Completed", "Not Started", "Submitted", "Rejected"
]

ALLOWED_CYCLES = [
    "FY26-Q3(Jan-Mar)",
    "FY26-Q2(Oct-Dec)",
    "FY26-Q1(July-Sep)",
    "FY26-Q4(Apr-Jun)",
]

DEFAULT_COLUMNS = {
    "gpn_id": "GPN",            
    "name": "Name",               
    "status": "Status",            
    "cycle": "Completion-Cycle",
}


def normalize_series(s: pd.Series) -> pd.Series:
    
    return (
        s.astype(str)
         .str.strip()
         .replace({"nan": ""})
         .str.replace(r"\s+", " ", regex=True)
         .str.casefold()
    )

def coalesce_columns(df: pd.DataFrame, primary: str, alternates: List[str]) -> str:
     
    def norm(x: str) -> str:
        return x.strip().lower().replace("-", " ").replace("_", " ")
    
    norm_map = {norm(c): c for c in df.columns}
    
    for cand in [primary] + alternates:
        if norm(cand) in norm_map:
            return norm_map[norm(cand)]
        raise KeyError(f"Column not found. Tried { [primary] + alternates }. Available: {list(df.columns)}")

def compute_exceptional_cases(df: pd.DataFrame, col_gpn: str, col_name: str, col_status: str, col_cycle: str) -> pd.DataFrame:
    
    status_norm = normalize_series(df[col_status].fillna(""))
    cycle_norm = normalize_series(df[col_cycle].fillna(""))

    gpn = df[col_gpn].astype(str).str.strip()
    name = df[col_name].astype(str).str.strip()
    
    status_blank   = status_norm.eq("")
    cycle_blank    = cycle_norm.eq("")
    
    status_allowed = {x.casefold() for x in ALLOWED_STATUS}
    cycle_allowed  = {x.casefold() for x in ALLOWED_CYCLES}
    
    status_invalid = (~status_blank) & (~normalize_series(df[col_status]).isin(status_allowed))
    cycle_invalid  = (~cycle_blank) & (~normalize_series(df[col_cycle]).isin(cycle_allowed))
    gpn_blank      = gpn.eq("") | gpn.eq("nan")
    name_blank     = name.eq("") | name.eq("nan")
    
    reasons = []
    for i in range(len(df)):
        r = []
        if bool(status_blank.iat[i]):   r.append("Blank Status")
        if bool(status_invalid.iat[i]): r.append("Invalid Status")
        if bool(cycle_blank.iat[i]):    r.append("Blank Completion-Cycle")
        if bool(cycle_invalid.iat[i]):  r.append("Invalid Completion-Cycle")
        if bool(gpn_blank.iat[i]):      r.append("Blank GPN ID")
        if bool(name_blank.iat[i]):     r.append("Blank Name")
        reasons.append(", ".join(r))
        
    out = df.copy()
    out["ExceptionReason"] = reasons
    return out[out["ExceptionReason"] != ""].copy()

def prompt_multi_select(title: str, options: List[str]) -> List[str]:
    
    print(f"\n{title}")
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    print("Enter numbers separated by comma (e.g., 1,3) or  'all' for all-options or press Enter to skip.")
    while True:
        s = input("> ").strip()
        if s == "":
            return []
        
        if s.lower() in ("all", "a", "*"):
            return options[:]
        try:
            idxs = [int(x.strip()) for x in s.split(",") if x.strip()]
            picked = [options[i-1] for i in idxs if 1 <= i <= len(options)]
            if picked:
                seen, out = set(), []
                for p in picked:
                    if p not in seen:
                        seen.add(p); out.append(p)
                return out
        except Exception:
            print("Invalid input. Try again (e.g., 1,2 or 'all' or Enter to skip).")

def apply_filters(df: pd.DataFrame, col_status: str, col_cycle: str, sel_status: List[str], sel_cycle: List[str]) -> pd.DataFrame:
    
    s_norm = normalize_series(df[col_status].fillna(""))
    c_norm = normalize_series(df[col_cycle].fillna(""))
    sel_s = set(x.casefold() for x in sel_status) if sel_status else None
    sel_c = set(x.casefold() for x in sel_cycle) if sel_cycle else None
    mask = pd.Series(True, index=df.index)
    
    if sel_s is not None:
        mask &= s_norm.isin(sel_s)
    if sel_c is not None:
        mask &= c_norm.isin(sel_c)
        
    return df[mask].copy()

def main():
    
    if not DATA_FILE.exists():
        print(f"ERROR: Expected file not found: {DATA_FILE.resolve()}")
        print("Place your latest 'EY Badges Tracker.xlsx' under the 'data/' folder and re-run.")
        sys.exit(1)
        
    try:
        df:pd.DataFrame = pd.read_excel(DATA_FILE, dtype=str, engine="openpyxl")
    except Exception as e:
        print(f"ERROR: Failed to read Excel: {e}", file=sys.stderr)
        sys.exit(1)
        
    try:
        col_gpn    = coalesce_columns(df, DEFAULT_COLUMNS["gpn_id"],    ["GPNID", "GPN Id", "GPN_Id"])
        col_name   = coalesce_columns(df, DEFAULT_COLUMNS["name"],      ["Employee Name", "Full Name"])
        col_status = coalesce_columns(df, DEFAULT_COLUMNS["status"],    ["Status"])
        col_cycle  = coalesce_columns(df, DEFAULT_COLUMNS["cycle"],     ["Completion Cycle", "CompletionCycle"])
    except KeyError as e:
        print(f"ERROR: {e}", file=sys.stderr); sys.exit(1)
        
    exceptions = compute_exceptional_cases(df, col_gpn, col_name, col_status, col_cycle)
    
    if exceptions.empty:
        print("All good, there is no issue with data.")
    else:
        print("\n")
        for i in range(len(exceptions)):
            print(exceptions[col_gpn].iat[i] + "\t" + exceptions["ExceptionReason"].iat[i])
            
        print("\n(Consider cleaning these before filtering.)")
        
    sel_status = prompt_multi_select("Select one or more Status options:", ALLOWED_STATUS)
    sel_cycle  = prompt_multi_select("Select one or more Completion-Cycle options:", ALLOWED_CYCLES)

    print("\nSelections:")
    print(f"  Status: {sel_status if sel_status else '(no filter)'}")
    print(f"  Completion-Cycle: {sel_cycle if sel_cycle else '(no filter)'}")
    
    filtered = apply_filters(df, col_status, col_cycle, sel_status, sel_cycle)
    
    output_by_gpn = (
        filtered[[col_gpn, col_name]]
        .assign(
            **{
                col_gpn: filtered[col_gpn].astype(str).str.strip(),
                col_name: filtered[col_name].astype(str).str.strip(),
            }
        )
        .dropna(subset=[col_gpn])
        .sort_values([col_gpn, col_name], kind="stable")
        .drop_duplicates(subset=[col_gpn], keep="first")
        .reset_index(drop=True)
    )
    
    print("\n===== Filtered Output (unique by GPN) =====")
    
    if output_by_gpn.empty:
        print("No rows matched your selections.")
    else:
        print(f"Unique GPNs: {len(output_by_gpn)}\n")
        
        for _, row in output_by_gpn.iterrows():
            print(f"{row[col_gpn]}\t{row[col_name]}")
                
if __name__ == "__main__":
    main()