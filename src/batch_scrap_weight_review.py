"""
SAP Batch Update - Scrap Weight Review (Demo)
--------------------------------------------
Automates updating batch characteristics in MSC2N using SAP GUI Scripting.

PUBLIC/SAFE VERSION NOTES:
- No internal paths, no company names, no real data.
- Excel path is provided via --excel or env var SAP_AUTOMATION_EXCEL.
- Column names are in English for portfolio consistency.
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

TCODE_MSC2N = "/nmsc2n"

# Public column naming (map from your original Spanish headers)
COL_MATERIAL = "MATERIAL"
COL_FAMILY = "FAMILY"
COL_BATCH = "BATCH"
COL_BASE_UOM = "BASE_UOM"
COL_WEIGHT = "WEIGHT_PER_UNIT"


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


def sap_enter(session, times: int = 1, delay: float = DEFAULT_DELAY):
    for _ in range(times):
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(delay)


def set_text(session, element_id: str, value):
    obj = session.findById(element_id)
    txt = "" if value is None else str(value).strip()

    try:
        obj.text = txt
    except Exception:
        obj.Text = txt

    try:
        obj.setFocus()
        obj.caretPosition = len(txt)
    except Exception:
        pass


def press(session, element_id: str):
    session.findById(element_id).press()


def go_tcode(session, tcode: str, delay: float = DEFAULT_DELAY):
    set_text(session, "wnd[0]/tbar[0]/okcd", tcode)
    sap_enter(session, 1, delay=delay)


def get_status_text(session) -> str:
    try:
        return session.findById("wnd[0]/sbar").Text
    except Exception:
        return ""


# ----------------------------
# INPUT: EXCEL LOAD + VALIDATION
# ----------------------------
def normalize_col(col) -> str:
    return str(col).strip().upper()


def load_excel(path: str, sheet=DEFAULT_SHEET_NAME) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = [normalize_col(c) for c in df.columns]

    required = {
        normalize_col(COL_MATERIAL),
        normalize_col(COL_FAMILY),
        normalize_col(COL_BATCH),
        normalize_col(COL_BASE_UOM),
        normalize_col(COL_WEIGHT),
    }

    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}. Found: {list(df.columns)}")

    for c in required:
        df[c] = df[c].fillna("").astype(str).str.strip()

    df = df[
        (df[normalize_col(COL_MATERIAL)] != "")
        & (df[normalize_col(COL_FAMILY)] != "")
        & (df[normalize_col(COL_BATCH)] != "")
        & (df[normalize_col(COL_WEIGHT)] != "")
    ].reset_index(drop=True)

    return df


def normalize_weight(value: str) -> str:
    # Ensure decimal point format for SAP input
    return str(value).replace(",", ".").strip()


# ----------------------------
# BUSINESS: MSC2N UPDATE (BATCH CHARACTERISTICS)
# ----------------------------
def msc2n_update_batch(session, material: str, batch: str, weight: str, family: str, base_uom: str, delay: float):
    session.findById("wnd[0]").resizeWorkingPane(88, 30, False)

    go_tcode(session, TCODE_MSC2N, delay=delay)

    set_text(
        session,
        "wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/subSUBSCR_HEADER:SAPLCHRG:1501/ctxtDFBATCH-MATNR",
        material,
    )

    set_text(
        session,
        "wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/subSUBSCR_HEADER:SAPLCHRG:1501/ctxtDFBATCH-CHARG",
        batch,
    )

    # Go to Classification tab (IDs vary by SAP customization; keep as-is)
    session.findById(
        "wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/"
        "subSUBSCR_TABSTRIP:SAPLCHRG:2000/tabsTS_BODY/tabpCLAS"
    ).select()
    time.sleep(delay)

    # Weight characteristic
    set_text(
        session,
        "wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/"
        "subSUBSCR_TABSTRIP:SAPLCHRG:2000/tabsTS_BODY/tabpCLAS/"
        "ssubSUBSCR_BODY:SAPLCHRG:2300/ssubSUBSCR_CLASS:SAPLCTMS:5000/"
        "tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/"
        "tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]",
        weight,
    )
    sap_enter(session, 1, delay=delay)

    # If base UoM is MLN, also fill the second weight field (keeps your original business rule)
    if str(base_uom).strip().upper() == "MLN":
        set_text(
            session,
            "wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/"
            "subSUBSCR_TABSTRIP:SAPLCHRG:2000/tabsTS_BODY/tabpCLAS/"
            "ssubSUBSCR_BODY:SAPLCHRG:2300/ssubSUBSCR_CLASS:SAPLCTMS:5000/"
            "tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/"
            "tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,5]",
            weight,
        )
        sap_enter(session, 1, delay=delay)

    # Family characteristic
    set_text(
        session,
        "wnd[0]/usr/subSUBSCR_BATCH_MASTER:SAPLCHRG:1111/"
        "subSUBSCR_TABSTRIP:SAPLCHRG:2000/tabsTS_BODY/tabpCLAS/"
        "ssubSUBSCR_BODY:SAPLCHRG:2300/ssubSUBSCR_CLASS:SAPLCTMS:5000/"
        "tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/"
        "tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,1]",
        family,
    )
    sap_enter(session, 1, delay=delay)

    # Save
    press(session, "wnd[0]/tbar[0]/btn[11]")
    time.sleep(delay)

    return "OK", get_status_text(session)


# ----------------------------
# CLI / MAIN
# ----------------------------
def parse_args():
    p = argparse.ArgumentParser(description="Update batch characteristics in MSC2N (demo).")
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
        material = row[normalize_col(COL_MATERIAL)]
        family = row[normalize_col(COL_FAMILY)]
        batch = row[normalize_col(COL_BATCH)]
        base_uom = row[normalize_col(COL_BASE_UOM)]
        weight = normalize_weight(row[normalize_col(COL_WEIGHT)])

        row_number_in_excel = i + 2
        try:
            status, detail = msc2n_update_batch(
                session,
                material=material,
                batch=batch,
                weight=weight,
                family=family,
                base_uom=base_uom,
                delay=args.delay,
            )
            print(f"{status} row {row_number_in_excel}: MAT={material} BATCH={batch} | {detail}")
        except Exception as e:
            print(f"ERROR row {row_number_in_excel}: MAT={material} BATCH={batch} | {e}")


if __name__ == "__main__":
    main()
