"""
SAP Material Master - Change Description (Demo)
-----------------------------------------------
Automates updating material short text and long texts using SAP GUI Scripting.

PUBLIC/SAFE VERSION NOTES:
- No company names, no internal paths, no real data.
- Excel path is provided via --excel or env var SAP_AUTOMATION_EXCEL.
- Keep example files with fake data only.
"""

import os
import time
import argparse
import pandas as pd
import win32com.client


# ----------------------------
# SAFE DEFAULT CONFIG
# ----------------------------
DEFAULT_SHEET_NAME = 0
DEFAULT_CONNECTION_INDEX = 0
DEFAULT_SESSION_INDEX = 0
DEFAULT_DELAY = 0.2

TCODE_MM02 = "/nmm02"
PANE_W, PANE_H = 88, 30

COL_SKU = "SKU"
COL_DESC = "DESCRIPTION"


# ----------------------------
# SAP: CONNECTION
# ----------------------------
def get_session(connection_index=DEFAULT_CONNECTION_INDEX, session_index=DEFAULT_SESSION_INDEX):
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(connection_index)
    session = connection.Children(session_index)
    return session


# ----------------------------
# SAP: LOW-LEVEL HELPERS
# ----------------------------
def exists(session, element_id: str) -> bool:
    try:
        session.findById(element_id)
        return True
    except Exception:
        return False


def press(session, element_id: str):
    session.findById(element_id).press()


def set_text(session, element_id: str, value, focus: bool = True):
    obj = session.findById(element_id)
    txt = "" if value is None else str(value)

    try:
        obj.text = txt
    except Exception:
        obj.Text = txt

    if focus:
        try:
            obj.setFocus()
            obj.caretPosition = len(txt)
        except Exception:
            pass


def send_enter(session, times: int = 1, delay: float = DEFAULT_DELAY):
    for _ in range(times):
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(delay)


def go_tcode(session, tcode: str, delay: float = DEFAULT_DELAY):
    set_text(session, "wnd[0]/tbar[0]/okcd", tcode, focus=False)
    send_enter(session, 1, delay=delay)


def get_status_text(session) -> str:
    try:
        return session.findById("wnd[0]/sbar").Text
    except Exception:
        return ""


def confirm_wnd1_if_exists(session, delay: float = DEFAULT_DELAY) -> bool:
    """Confirms a generic popup (wnd[1]) if present."""
    if exists(session, "wnd[1]"):
        try:
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(delay)
            return True
        except Exception:
            return False
    return False


# ----------------------------
# EXCEL: LOAD + VALIDATE
# ----------------------------
def normalize_col(c) -> str:
    return str(c).strip().upper()


def load_excel(path: str, sheet=DEFAULT_SHEET_NAME) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = [normalize_col(c) for c in df.columns]

    required = {COL_SKU, COL_DESC}
    if not required.issubset(set(df.columns)):
        raise ValueError(f"Excel must contain columns {sorted(required)}. Found: {list(df.columns)}")

    df[COL_SKU] = df[COL_SKU].fillna("").astype(str).str.strip()
    df[COL_DESC] = df[COL_DESC].fillna("").astype(str).str.strip()
    df = df[(df[COL_SKU] != "") & (df[COL_DESC] != "")].reset_index(drop=True)

    return df


# ----------------------------
# BUSINESS: MM02 CHANGE DESCRIPTION
# ----------------------------
def mm02_change_description(session, sku: str, description: str, delay: float = DEFAULT_DELAY):
    session.findById("wnd[0]").resizeWorkingPane(PANE_W, PANE_H, False)

    go_tcode(session, TCODE_MM02, delay=delay)

    set_text(session, "wnd[0]/usr/ctxtRMMG1-MATNR", sku)
    send_enter(session, 1, delay=delay)

    confirm_wnd1_if_exists(session, delay=delay)

    # Choose views (may vary by SAP config). Keeps original IDs as-is.
    press(session, "wnd[0]/tbar[1]/btn[30]")
    time.sleep(delay)

    # Update short text fields (language-dependent table rows)
    set_text(
        session,
        "wnd[0]/usr/tabsTABSPR1/tabpZU01/ssubTABFRA1:SAPLMGMM:2110/"
        "subSUB2:SAPLMGD1:8000/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]",
        description,
    )
    set_text(
        session,
        "wnd[0]/usr/tabsTABSPR1/tabpZU01/ssubTABFRA1:SAPLMGMM:2110/"
        "subSUB2:SAPLMGD1:8000/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,1]",
        description,
    )

    # Optional focus tweak
    try:
        obj_id = (
            "wnd[0]/usr/tabsTABSPR1/tabpZU01/ssubTABFRA1:SAPLMGMM:2110/"
            "subSUB2:SAPLMGD1:8000/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,1]"
        )
        session.findById(obj_id).setFocus()
        session.findById(obj_id).caretPosition = 0
    except Exception:
        pass

    # Back to main tabs
    press(session, "wnd[0]/tbar[1]/btn[27]")
    time.sleep(delay)

    long_text_value = f"{description}\r\n\r\n"

    # Sales text
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09").select()
    time.sleep(delay)
    set_text(
        session,
        "wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2010/"
        "subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell",
        long_text_value,
        focus=False,
    )

    # Purchasing text
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").select()
    time.sleep(delay)
    set_text(
        session,
        "wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2010/"
        "subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell",
        long_text_value,
        focus=False,
    )

    # Save
    press(session, "wnd[0]/tbar[0]/btn[11]")
    time.sleep(delay)

    return "OK", get_status_text(session)


# ----------------------------
# CLI / MAIN
# ----------------------------
def parse_args():
    p = argparse.ArgumentParser(description="Change SAP material description with SAP GUI Scripting (demo).")
    p.add_argument("--excel", default=os.getenv("SAP_AUTOMATION_EXCEL", ""), help="Path to input Excel file.")
    p.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help="Sheet name or index (default: 0).")
    p.add_argument("--connection", type=int, default=DEFAULT_CONNECTION_INDEX, help="SAP connection index.")
    p.add_argument("--session", type=int, default=DEFAULT_SESSION_INDEX, help="SAP session index.")
    p.add_argument("--delay", type=float, default=DEFAULT_DELAY, help="Delay between actions (seconds).")
    return p.parse_args()


def main():
    args = parse_args()
    if not args.excel:
        raise SystemExit("Missing Excel path. Provide --excel <path> or set env var SAP_AUTOMATION_EXCEL.")

    df = load_excel(args.excel, args.sheet)
    session = get_session(args.connection, args.session)

    for i, row in df.iterrows():
        sku = row[COL_SKU]
        description = row[COL_DESC]

        row_number_in_excel = i + 2  # header at row 1
        try:
            status, detail = mm02_change_description(session, sku, description, delay=args.delay)
            print(f"{status} row {row_number_in_excel}: SKU={sku} | {detail}")
        except Exception as e:
            print(f"ERROR row {row_number_in_excel}: SKU={sku} | {e}")


if __name__ == "__main__":
    main()
