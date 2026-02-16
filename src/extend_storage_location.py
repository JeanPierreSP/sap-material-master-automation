"""
SAP Material Extension - Storage Location (Demo)
-----------------------------------------------
Automates extending an existing material to a storage location using SAP GUI Scripting.

NOTE:
- This is a demo/template version intended for public sharing.
- Do NOT commit real company paths, variants, or production data files to GitHub.
"""

import os
import time
import argparse
import pandas as pd
import win32com.client


# ----------------------------
# CONFIG (safe defaults)
# ----------------------------
DEFAULT_SHEET_NAME = 0
DEFAULT_CONNECTION_INDEX = 0
DEFAULT_SESSION_INDEX = 0
DEFAULT_DELAY = 0.2


# ----------------------------
# SAP: Connection
# ----------------------------
def get_session(connection_index=DEFAULT_CONNECTION_INDEX, session_index=DEFAULT_SESSION_INDEX):
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(connection_index)
    session = connection.Children(session_index)
    return session


# ----------------------------
# SAP: Low-level helpers
# ----------------------------
def exists(session, element_id: str) -> bool:
    try:
        session.findById(element_id)
        return True
    except Exception:
        return False


def sap_enter_wnd0(session, times: int = 1, delay: float = DEFAULT_DELAY):
    for _ in range(times):
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(delay)


def set_field(session, field_id: str, value):
    obj = session.findById(field_id)
    txt = str(value).strip()
    try:
        obj.text = txt
    except Exception:
        obj.Text = txt
    try:
        obj.setFocus()
        obj.caretPosition = len(txt)
    except Exception:
        pass


def press_if_exists(session, element_id: str) -> bool:
    if exists(session, element_id):
        session.findById(element_id).press()
        return True
    return False


# ----------------------------
# SAP: Popups / Status
# ----------------------------
def close_org_levels_popup(session, delay: float = DEFAULT_DELAY):
    """
    Close the 'Organizational levels' popup (wnd[1]).
    Tries common buttons and fallback to F12 (cancel).
    """
    if not exists(session, "wnd[1]"):
        return

    # Common buttons: Cancel/Back/Exit vary by system
    if press_if_exists(session, "wnd[1]/tbar[0]/btn[12]"):
        time.sleep(delay)
        return

    if press_if_exists(session, "wnd[1]/tbar[0]/btn[15]"):
        time.sleep(delay)
        return

    try:
        session.findById("wnd[1]").sendVKey(12)  # F12
        time.sleep(delay)
    except Exception:
        pass


def get_status_text(session) -> str:
    try:
        return session.findById("wnd[0]/sbar").Text
    except Exception:
        return ""


# ----------------------------
# Input: Excel load/validation
# ----------------------------
def load_excel(path: str, sheet=DEFAULT_SHEET_NAME) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = [c.strip().upper() for c in df.columns]

    required = {"SKU", "ALMACEN"}
    if not required.issubset(set(df.columns)):
        raise ValueError(f"Excel must contain columns {sorted(required)}. Found: {list(df.columns)}")

    df["SKU"] = df["SKU"].fillna("").astype(str).str.strip()
    df["ALMACEN"] = df["ALMACEN"].fillna("").astype(str).str.strip()
    df = df[(df["SKU"] != "") & (df["ALMACEN"] != "")].reset_index(drop=True)

    return df


# ----------------------------
# Business process: MM01 (extend storage location)
# ----------------------------
def mm01_extend_storage(session, sku: str, almacen: str, delay: float = DEFAULT_DELAY):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
    sap_enter_wnd0(session, 1, delay=delay)

    set_field(session, "wnd[0]/usr/ctxtRMMG1-MATNR", sku)
    set_field(session, "wnd[0]/usr/ctxtRMMG1_REF-MATNR", sku)

    sap_enter_wnd0(session, 1, delay=delay)

    if not exists(session, "wnd[1]/usr/ctxtRMMG1-LGORT"):
        raise RuntimeError("Organizational levels popup did not appear (LGORT field not found).")

    # Fill storage location
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = almacen
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus()
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = len(almacen)

    session.findById("wnd[1]").sendVKey(0)
    time.sleep(delay)

    # If popup still exists -> likely already extended or blocked
    if exists(session, "wnd[1]"):
        status = get_status_text(session)
        close_org_levels_popup(session, delay=delay)
        return "SKIP", status or "Already extended / did not proceed from organizational levels"

    # Save
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    time.sleep(delay)
    return "OK", get_status_text(session)


# ----------------------------
# Main
# ----------------------------
def parse_args():
    p = argparse.ArgumentParser(description="Extend materials to storage location using SAP GUI Scripting (demo).")
    p.add_argument("--excel", default=os.getenv("SAP_AUTOMATION_EXCEL", ""), help="Path to input Excel file.")
    p.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help="Sheet name or index (default: 0).")
    p.add_argument("--connection", type=int, default=DEFAULT_CONNECTION_INDEX, help="SAP connection index.")
    p.add_argument("--session", type=int, default=DEFAULT_SESSION_INDEX, help="SAP session index.")
    p.add_argument("--delay", type=float, default=DEFAULT_DELAY, help="Delay between actions (seconds).")
    return p.parse_args()


def main():
    args = parse_args()

    if not args.excel:
        raise SystemExit(
            "Missing Excel path. Provide --excel <path> or set env var SAP_AUTOMATION_EXCEL."
        )

    df = load_excel(args.excel, args.sheet)
    session = get_session(args.connection, args.session)

    for i, row in df.iterrows():
        sku = row["SKU"]
        almacen = row["ALMACEN"]

        status, detail = mm01_extend_storage(session, sku, almacen, delay=args.delay)

        row_number_in_excel = i + 2  # assuming header in row 1
        if status == "OK":
            print(f"OK row {row_number_in_excel}: SKU={sku} -> STORAGE={almacen} | {detail}")
        else:
            print(f"SKIP row {row_number_in_excel}: SKU={sku} already in {almacen} | {detail}")


if __name__ == "__main__":
    main()
