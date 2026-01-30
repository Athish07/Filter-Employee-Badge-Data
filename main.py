import sys
from pathlib import Path
from typing import List, Dict
import pandas as pd
import json
import pythoncom
from win32com.client import gencache
from constants import ALLOWED_CYCLES, ALLOWED_STATUS, DEFAULT_COLUMNS

BADGES_FILE = Path("data") / "EY Badges Tracker.xlsx"
EMAIL_MASTER_FILE = Path("data") / "Emerging Tech Team - FY 26.xlsx"


def load_email_config(path="email/email_config.json"):
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Email config not found: {config_path.resolve()}")
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)
    
def load_email_template(path):
    template_path = Path(path)
    if not template_path.exists():
        raise FileNotFoundError(f"Email template not found: {template_path.resolve()}")
    with open(template_path, "r", encoding="utf-8") as f:
        return f.read()

def clean_series(s: pd.Series) -> pd.Series:
    
    return (
        s.astype(str)
         .str.replace("\u00A0", " ", regex=False)  
         .str.replace("\u200B", "", regex=False)   
         .str.strip()
         .replace({"nan": ""})
         .str.replace(r"\s+", " ", regex=True)
         .str.casefold()
    )
    

def coalesce_columns(df: pd.DataFrame, primary: str, alternates: List[str]) -> str:
    
    
    def hdr_clean(x: str) -> str:
        return (
            str(x)
            .replace("\u00A0", " ")  
            .replace("\u200B", "")
            .strip()
        )

    def norm(x: str) -> str:
        x = hdr_clean(x)
        return x.strip().lower().replace("-", " ").replace("_", " ")
    
    norm_map = {norm(c): c for c in df.columns}
    
    for cand in [primary] + alternates:
        if norm(cand) in norm_map:
            return norm_map[norm(cand)]
        
    raise KeyError(f"Column not found. Tried { [primary] + alternates }. Available: {list(df.columns)}")

def load_excel(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, dtype=str, engine="openpyxl")

def load_excel_first_sheet(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name= "Emerging Tech Team Members", dtype= str, engine= "openpyxl")
                
def compute_exceptions(df: pd.DataFrame, col_gpn: str, col_name: str, col_status: str, col_cycle: str) -> pd.DataFrame:
    status_norm = clean_series(df[col_status].fillna(""))
    cycle_norm  = clean_series(df[col_cycle].fillna(""))

    gpn  = clean_series(df[col_gpn])
    name = clean_series(df[col_name])

    status_blank = status_norm.eq("")
    cycle_blank  = cycle_norm.eq("")
    status_allowed = {x.casefold() for x in ALLOWED_STATUS}
    cycle_allowed  = {x.casefold() for x in ALLOWED_CYCLES}

    status_invalid = (~status_blank) & (~status_norm.isin(status_allowed))
    cycle_invalid  = (~cycle_blank)  & (~cycle_norm.isin(cycle_allowed))
    gpn_blank      = gpn.eq("")
    name_blank     = name.eq("")

    reasons = []
    for i in range(len(df)):
        r = []
        if bool(status_blank.iat[i]):   r.append("Blank Status")
        if bool(status_invalid.iat[i]): r.append("Invalid Status")
        if bool(cycle_blank.iat[i]):    r.append("Blank Completion-Cycle")
        if bool(cycle_invalid.iat[i]):  r.append("Invalid Completion-Cycle")
        if bool(gpn_blank.iat[i]):      r.append("Blank GPN")
        if bool(name_blank.iat[i]):     r.append("Blank Name")
        reasons.append(", ".join(r))

    out = df.copy()
    out["ExceptionReason"] = reasons
    return out[out["ExceptionReason"] != ""].copy()

def prompt_multi_select(title: str, options: List[str]) -> List[str]:
    print(f"\n{title}")
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    print("Enter numbers separated by comma (e.g., 1,3) — 'all' for all — or press Enter to skip.")
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
            print("No valid selections parsed. Try again.")
        except Exception:
            print("Invalid input. Try again (e.g., 1,2 or 'all' or Enter to skip).")

def apply_filters(df: pd.DataFrame, col_status: str, col_cycle: str, sel_status: List[str], sel_cycle: List[str]) -> pd.DataFrame:
    s_norm = clean_series(df[col_status].fillna(""))
    c_norm = clean_series(df[col_cycle].fillna(""))
    sel_s  = set(x.casefold() for x in sel_status) if sel_status else None
    sel_c  = set(x.casefold() for x in sel_cycle)  if sel_cycle  else None
    mask = pd.Series(True, index=df.index)
    if sel_s is not None:
        mask &= s_norm.isin(sel_s)
    if sel_c is not None:
        mask &= c_norm.isin(sel_c)
    return df[mask].copy()

def build_unique_gpn(df: pd.DataFrame, col_gpn: str, col_name: str) -> pd.DataFrame:
    out = (
        df[[col_gpn, col_name]]
        .assign(
            **{
                col_gpn: clean_series(df[col_gpn]),
                col_name: clean_series(df[col_name]),
            }
        )
        .dropna(subset=[col_gpn])
        .sort_values([col_gpn, col_name], kind="stable")
        .drop_duplicates(subset=[col_gpn], keep="first")
        .reset_index(drop=True)
    )
    return out

def build_gpn_to_email_map(master_path: Path) -> Dict[str, str]:
    
    dfm = load_excel_first_sheet(master_path)
    
    col_gpn   = coalesce_columns(dfm, "GPN", ["GPN ID", "Gpn", "GPN_Id", "GPN Id"])
    col_email = coalesce_columns(dfm, "Email ID", ["EmailID", "Email", "Email Address", "Mail", "E-mail ID"])
    
    dfm["_GPN_norm"]  = clean_series(dfm[col_gpn])
    dfm["_EMAIL_raw"] = (
        dfm[col_email].astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.replace("\u200B", "", regex=False)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    dfm = dfm[dfm["_GPN_norm"] != ""].copy()

    dfm_sorted = dfm.assign(_email_blank=dfm["_EMAIL_raw"].eq("")).sort_values(["_GPN_norm", "_email_blank"])
    dedup = dfm_sorted.drop_duplicates(subset=["_GPN_norm"], keep="first")

    return dict(zip(dedup["_GPN_norm"], dedup["_EMAIL_raw"]))

def perform_action_with_emails(
    emails: List[str],
    subject: str = "Group Email",
    html_body: str = "<p>This email is sent to all recipients in the list.</p>",
    display_before_send: bool = True
) -> None:
        
    clean, seen = [], set()
    for e in emails or []:
        s = (e or "").strip()
        k = s.lower()
        if s and k not in seen:
            seen.add(k)
            clean.append(s)
    if not clean:
        print("No valid emails to send.")
        return

    pythoncom.CoInitialize()
    outlook = gencache.EnsureDispatch("Outlook.Application")
    session = outlook.GetNamespace("MAPI")

    try:
        session.Logon("", "", True, False)   
    except Exception:
        pass 

    mail = outlook.CreateItem(0)  
    mail.To = "; ".join(clean)

    mail.Subject = subject
    mail.HTMLBody = html_body 
    try:
        if session.Accounts.Count > 0:
            mail.SendUsingAccount = session.Accounts.Item(1)
    except Exception:
        pass  

    if display_before_send:
        mail.Display(False)
        print("Displayed email for review. If it looks correct, click Send manually.")
        return

    try:
        mail.Send()
        
        print("Email sent to all recipients!")
    except Exception as e:
        mail.Save()
        raise RuntimeError(f"Outlook aborted send: {repr(e)}, yo can view the email in the draft.")
    
def main() -> None:
    if not BADGES_FILE.exists():
        print(f"ERROR: Expected file not found: {BADGES_FILE.resolve()}")
        print("Place your latest 'EY Badges Tracker.xlsx' under the 'data/' folder and re-run.")
        sys.exit(1)
    if not EMAIL_MASTER_FILE.exists():
        print(f"ERROR: Email master not found: {EMAIL_MASTER_FILE.resolve()}")
        print("Place your 'Emerging Tech Team - FY 26.xlsx' under the 'data/' folder and re-run.")
        sys.exit(1)
        
    df = load_excel(BADGES_FILE)
    
    try:
        col_gpn    = coalesce_columns(df, DEFAULT_COLUMNS["gpn_id"], ["GPNID", "GPN Id", "GPN_Id", "GPN ID"])
        col_name   = coalesce_columns(df, DEFAULT_COLUMNS["name"],   ["Employee Name", "Full Name"])
        col_status = coalesce_columns(df, DEFAULT_COLUMNS["status"], ["Status"])
        col_cycle  = coalesce_columns(df, DEFAULT_COLUMNS["cycle"],  ["Completion Cycle", "CompletionCycle"])
    except KeyError as e:
        print(f"ERROR: {e}", file=sys.stderr); sys.exit(1)
        
    exceptions = compute_exceptions(df, col_gpn, col_name, col_status, col_cycle)
    if exceptions.empty:
        print("All good, there is no issue with data.")
    else:
        print("\n===== Exceptional-Cases =====\n")
        for _, r in exceptions[[col_gpn, "ExceptionReason"]].head(100).iterrows():
            gpn_val = "" if pd.isna(r[col_gpn]) else str(r[col_gpn]).strip()
            reason  = "" if pd.isna(r["ExceptionReason"]) else str(r["ExceptionReason"]).strip()
            print(f"{gpn_val}\t{reason}")
        print("\n(Consider cleaning these before filtering.)")
        
    sel_status = prompt_multi_select("Select one or more Status options:", ALLOWED_STATUS)
    sel_cycle  = prompt_multi_select("Select one or more Completion-Cycle options:", ALLOWED_CYCLES)

    print("\nSelections:")
    print(f"  Status: {sel_status if sel_status else '(no filter)'}")
    print(f"  Completion-Cycle: {sel_cycle if sel_cycle else '(no filter)'}")

    filtered = apply_filters(df, col_status, col_cycle, sel_status, sel_cycle)
    
    output_by_gpn = build_unique_gpn(filtered, col_gpn, col_name)
    
    print("\n===== Filtered Output (unique by GPN) =====")
    if output_by_gpn.empty:
        print("No rows matched your selections.")
        sys.exit(0)
    else:
        print(f"Unique GPNs: {len(output_by_gpn)}\n")
        for _, row in output_by_gpn.iterrows():
            print(f"{str(row[col_gpn]).upper()}\t{row[col_name]}")
            
    gpn_to_email = build_gpn_to_email_map(EMAIL_MASTER_FILE)

    print("\n===== GPN → Email =====")
    missing_gpns: List[str] = []
    valid_emails: List[str] = []
    
    for _, r in output_by_gpn.iterrows():
        gpn_val  = str(r[col_gpn]).strip()
        name_val = str(r[col_name]).strip()
        email    = gpn_to_email.get(gpn_val.casefold(), "")
        
        if email == "":
            missing_gpns.append(gpn_val)
        else:
            print(f"{gpn_val.upper()}\t{name_val}\t{email}")
            valid_emails.append(email)
            
    email_cfg = load_email_config()
    html_template = load_email_template(email_cfg["template_file"])
    
    perform_action_with_emails(
    valid_emails,
    subject=email_cfg["subject"],
    html_body=html_template,
    display_before_send=True
    )
    print("\n")
    
    if missing_gpns:
        print("\n-- Missing emails for these GPNs (not found or blank in master) --")
        for g in missing_gpns:
            print(g)

if __name__ == "__main__":
    main()