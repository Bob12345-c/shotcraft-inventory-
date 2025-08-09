
import streamlit as st
import pandas as pd
import math
from io import BytesIO
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Shotcraft Inventory (Google Sheets Live)", layout="wide")
st.title("📦 Shotcraft — Live Inventory (Google Sheets)")

st.caption(
    "This version works even if SHEET_ID isn't in Secrets. "
    "You can pass it via URL (?sheet_id=...) or paste it below."
)

# -----------------------------
# Secrets & Config resolution
# -----------------------------

def normalize_private_key(sa: dict) -> dict:
    sa = dict(sa) if sa else {}
    pk = sa.get("private_key", "")
    # If the key was pasted with \n escapes, restore real newlines
    if "\\n" in pk:
        sa["private_key"] = pk.replace("\\n", "\n")
    return sa

def read_service_account():
    if "gcp_service_account" not in st.secrets:
        st.error("Missing [gcp_service_account] in Secrets. Add your Google service account JSON under that header.")
        st.stop()
    return normalize_private_key(st.secrets["gcp_service_account"])

def resolve_sheet_id():
    # Priority: query param -> top-level secret -> [app] secret -> manual entry
    qp = st.query_params
    if "sheet_id" in qp and qp["sheet_id"]:
        return qp["sheet_id"]
    sid = st.secrets.get("SHEET_ID")
    if sid: return sid
    appblock = st.secrets.get("app", {})
    sid = appblock.get("SHEET_ID")
    if sid: return sid

    st.info("No SHEET_ID found in Secrets. Paste your Sheet ID or full URL below and click 'Use this Sheet'.")
    with st.expander("Paste your Google Sheet ID or full URL"):
        default_val = st.session_state.get("manual_sheet_input", "")
        user_input = st.text_input("Google Sheet ID **or** full URL", value=default_val, placeholder="1ivuxCDfMu... OR https://docs.google.com/spreadsheets/d/…/edit")
        colA, colB = st.columns([1,1])
        with colA:
            use_btn = st.button("Use this Sheet", type="primary")
        with colB:
            st.caption("Tip: The ID is the long part between /d/ and /edit in the URL.")
        if use_btn and user_input:
            # Extract ID if they pasted full URL
            txt = user_input.strip()
            if "/d/" in txt:
                try:
                    txt = txt.split("/d/")[1].split("/")[0]
                except Exception:
                    pass
            st.session_state["manual_sheet_input"] = txt
            st.rerun()
    # If they previously saved a manual sheet id, use it
    return st.session_state.get("manual_sheet_input")

def resolve_ws_names():
    # Defaults
    form_ws = "FORMULA"
    inv_ws = "INVENTORY"
    # Allow overrides via query params or secrets
    qp = st.query_params
    if "formula_ws" in qp and qp["formula_ws"]:
        form_ws = qp["formula_ws"]
    if "inventory_ws" in qp and qp["inventory_ws"]:
        inv_ws = qp["inventory_ws"]
    # Secrets (top-level or [app])
    if st.secrets.get("FORMULA_WS"):
        form_ws = st.secrets["FORMULA_WS"]
    if st.secrets.get("INVENTORY_WS"):
        inv_ws = st.secrets["INVENTORY_WS"]
    appblock = st.secrets.get("app", {})
    if appblock.get("FORMULA_WS"):
        form_ws = appblock["FORMULA_WS"]
    if appblock.get("INVENTORY_WS"):
        inv_ws = appblock["INVENTORY_WS"]
    return form_ws, inv_ws

SERVICE_ACCOUNT_INFO = read_service_account()
SHEET_ID = resolve_sheet_id()
FORMULA_WS, INVENTORY_WS = resolve_ws_names()

with st.sidebar:
    st.header("🔧 Config")
    st.write("Secrets keys loaded:", list(st.secrets.keys()))
    st.write("Using SHEET_ID ends with:", (SHEET_ID or "")[-6:] if SHEET_ID else "None")
    st.write("FORMULA_WS:", FORMULA_WS, " | INVENTORY_WS:", INVENTORY_WS)

if not SHEET_ID:
    st.stop()

# -----------------------------
# Google Sheets helpers
# -----------------------------

@st.cache_resource(show_spinner=False)
def get_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
    return gspread.authorize(creds)

def read_ws_df(ws):
    vals = ws.get_all_values()
    if not vals:
        return pd.DataFrame()
    df = pd.DataFrame(vals[1:], columns=vals[0])
    # Coerce numeric-ish columns
    for c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="ignore")
    return df

def load_data(gc):
    sh = gc.open_by_key(SHEET_ID)
    try:
        fws = sh.worksheet(FORMULA_WS)
    except Exception as e:
        st.error(f"Could not open FORMULA worksheet '{FORMULA_WS}'. Error: {e}")
        st.stop()
    try:
        iws = sh.worksheet(INVENTORY_WS)
    except Exception as e:
        st.error(f"Could not open INVENTORY worksheet '{INVENTORY_WS}'. Error: {e}")
        st.stop()

    formula = read_ws_df(fws)
    inv = read_ws_df(iws)

    required = {"Component","Per_Case"}
    if not required.issubset(set(formula.columns)):
        st.error(
            "FORMULA sheet must have headers: Component, Per_Case (UOM optional). "
            f"Found: {list(formula.columns)}"
        )
        st.stop()

    comps = formula[["Component","Per_Case"]].copy()
    comps["UOM"] = formula["UOM"] if "UOM" in formula.columns else ""

    if {"Component","On_Hand"}.issubset(set(inv.columns)):
        onhand = inv[["Component","On_Hand"]].copy()
    else:
        # Initialize empty if not present
        onhand = pd.DataFrame({"Component": comps["Component"], "On_Hand": 0.0})

    return sh, comps.reset_index(drop=True), onhand.reset_index(drop=True)

def write_onhand(sh, edited_df):
    ws = sh.worksheet(INVENTORY_WS)
    out = edited_df[["Component","On_Hand"]].copy()
    out["On_Hand"] = pd.to_numeric(out["On_Hand"], errors="coerce").fillna(0).astype(float)
    values = [out.columns.tolist()] + out.astype(object).where(pd.notnull(out), "").values.tolist()
    ws.clear()
    ws.update(values)

def compute(comps, onhand, cases):
    df = comps.merge(onhand, on="Component", how="left")
    if "On_Hand" not in df.columns: df["On_Hand"] = 0.0
    df["Per_Case"]  = pd.to_numeric(df["Per_Case"], errors="coerce").fillna(0.0)
    df["On_Hand"]   = pd.to_numeric(df["On_Hand"], errors="coerce").fillna(0.0)
    df["Required"]  = df["Per_Case"] * float(cases)
    df["Remaining"] = df["On_Hand"] - df["Required"]

    candidates = df[df["Per_Case"] > 0]
    if not candidates.empty:
        max_sellable = int(math.floor((candidates["On_Hand"]/candidates["Per_Case"]).min()))
    else:
        max_sellable = 0

    shortages = df[df["Remaining"] < 0][["Component","UOM","On_Hand","Per_Case","Required","Remaining"]].copy()
    display = df[["Component","UOM","On_Hand","Per_Case","Required","Remaining"]].sort_values("Component")
    return display, max_sellable, shortages

def download_excel(formula_name, display_df):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        display_df[["Component","UOM","Per_Case"]].to_excel(writer, sheet_name=formula_name, index=False)
        display_df.to_excel(writer, sheet_name="INVENTORY", index=False)
    bio.seek(0)
    return bio

# -----------------------------
# Main run
# -----------------------------

try:
    gc = get_client()
    sh, comps, onhand = load_data(gc)
    st.success("Connected to Google Sheet ✓")
except Exception as e:
    st.error(f"Could not connect to Google Sheets: {e}")
    st.stop()

with st.sidebar:
    st.header("Actions")
    if st.button("Reload from Sheet"):
        st.cache_data.clear()
        st.rerun()

st.subheader("Per-case usage (from FORMULA)")
st.dataframe(comps, hide_index=True, use_container_width=True)

st.subheader("Edit On_Hand (writes back to INVENTORY)")
base = comps.merge(onhand, on="Component", how="left")
base["On_Hand"] = pd.to_numeric(base["On_Hand"], errors="coerce").fillna(0.0)

edited = st.data_editor(
    base[["Component","UOM","On_Hand","Per_Case"]],
    hide_index=True,
    column_config={
        "Component": st.column_config.TextColumn(disabled=True),
        "UOM": st.column_config.TextColumn(disabled=True),
        "Per_Case": st.column_config.NumberColumn(format="%.6f", disabled=True),
        "On_Hand": st.column_config.NumberColumn(help="Type your current stock here"),
    },
    use_container_width=True,
    key="edit_table"
)

c1, c2 = st.columns(2)
with c1:
    if st.button("💾 Sync On_Hand to Google Sheets"):
        try:
            write_onhand(sh, edited)
            st.success(f"Synced at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        except Exception as e:
            st.error(f"Sync failed: {e}")
with c2:
    if st.button("↩️ Revert to current sheet values"):
        st.cache_data.clear()
        st.rerun()

st.subheader("Order size")
cases = st.number_input("Cases sold (e.g., LCBO order)", min_value=0.0, step=1.0, value=0.0)

display, max_sell, shortages = compute(comps, edited[["Component","On_Hand"]].copy(), cases)

st.markdown("### Results")
m1, m2 = st.columns(2)
with m1:
    st.metric("Max sellable cases from current stock", max_sell)
with m2:
    st.metric("Order size (cases)", int(cases))

st.dataframe(display, hide_index=True, use_container_width=True)

if not shortages.empty:
    st.warning("Shortages for this order:")
    st.dataframe(shortages, hide_index=True, use_container_width=True)
else:
    st.info("No shortages detected for this order.")

st.markdown("### Download snapshot")
buf = download_excel(FORMULA_WS, display)
st.download_button("Download Excel snapshot", buf, file_name="Shotcraft_Inventory_Snapshot.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
