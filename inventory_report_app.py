# inventory_report_app.py

import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import numpy as np
import re

st.set_page_config(page_title="Inventory Report Generator", layout="wide")

st.title("📦 Inventory Report Generator")
st.markdown("""
Upload your **Master Excel file** (must contain a sheet named `OPEN_ORDERS`)  
and your **Client Excel file**. The client file may be either:
- A **raw workbook** (any sheet names) that contains a weekly **Forecast** sheet and an **Availability** summary, or
- A prepared sheet with columns: `ITEM #` (required), `PO/LOT` (can be blank), `Weekly Sales`, `Total On Hand`, `Total SO`.

The app will **auto-pick the best PO/LOT** per item from the master and then calculate the report.
""")

uploaded_master = st.file_uploader("Upload Master Excel File", type=["xlsx"], key="master")
uploaded_client = st.file_uploader("Upload Client Excel File", type=["xlsx"], key="client")

today = datetime.today().date()
with st.expander("📅 Optional: Override 'Today' Date"):
    custom_today = st.date_input("Choose custom date", today)
    if custom_today:
        today = custom_today

# -------------------- Utilities & Helpers --------------------

def _lower_name(name):
    try:
        return str(name).strip().lower()
    except Exception:
        return ""

def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def clean_key(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    s = str(x).replace("\u00A0", "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"\s+", " ", s)

def _as_item_key(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    s = str(x).replace("\u00A0", "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"\s+", " ", s)
    digits = re.sub(r"\D","",s)
    base = s.replace(" ","").replace("-","")
    if digits and digits == re.sub(r"\D","",base):
        return digits.lstrip("0") or "0"
    return s.lstrip("0")

def to_dt(series: pd.Series) -> pd.Series:
    s = pd.to_datetime(series, errors="coerce")
    need_fix = s.isna() & pd.to_numeric(series, errors="coerce").notna()
    if need_fix.any():
        serial = pd.to_numeric(series[need_fix], errors="coerce")
        s.loc[need_fix] = pd.to_datetime("1899-12-30") + pd.to_timedelta(serial, unit="D")
    return s

def _find_weekly_cols(df: pd.DataFrame) -> list:
    weekly = []
    for c in df.columns:
        c_str = str(c).strip()
        dt = pd.to_datetime(c_str, errors="coerce")
        if pd.notna(dt):
            weekly.append(c)
            continue
        if re.match(r"^\d{1,2}[-/ ]?[A-Za-z]{3}$", c_str):
            weekly.append(c)
    return weekly

# -------------------- NOTES-aware logic --------------------

NOTE_DATE_PATTERNS = [
    r"delivered:\s*(\d{1,2}/\d{1,2})",
    r"da:\s*(\d{1,2}/\d{1,2})",
]

def _parse_mmdd_to_dt(mmdd: str, year_hint: int | None = None):
    """Parse M/D using year_hint if available; else current year."""
    try:
        m, d = [int(x) for x in mmdd.split("/")]
        y = year_hint or datetime.today().year
        dt = datetime(y, m, d)
        return pd.to_datetime(dt.date())
    except Exception:
        return pd.NaT

def analyze_notes(row: pd.Series) -> dict:
    """
    Extracts structured info from NOTES:
      - delivered_date (Timestamp or NaT)
      - da_date (Timestamp or NaT)
      - hot_flag (bool)
      - hold_flag (bool)
    """
    text = ""
    if "NOTES" in row.index and pd.notna(row["NOTES"]):
        text = str(row["NOTES"]).lower()

    hot = "hot container" in text or "hot container" in text.replace("-", " ")
    # 'hold' can appear as its own line; treat presence anywhere as a flag
    hold = "hold" in text

    delivered_date = pd.NaT
    da_date = pd.NaT

    # Use DUE DATE's year as a hint (to pin mm/dd to a year)
    year_hint = None
    if "DUE DATE" in row.index and pd.notna(row["DUE DATE"]):
        try:
            year_hint = pd.to_datetime(row["DUE DATE"]).year
        except Exception:
            year_hint = None

    if text:
        for line in text.splitlines():
            s = line.strip().lower()
            m_del = re.search(r"delivered:\s*(\d{1,2}/\d{1,2})", s)
            if m_del and pd.isna(delivered_date):
                delivered_date = _parse_mmdd_to_dt(m_del.group(1), year_hint)
            m_da = re.search(r"\bda:\s*(\d{1,2}/\d{1,2})", s)
            if m_da and pd.isna(da_date):
                da_date = _parse_mmdd_to_dt(m_da.group(1), year_hint)

    return {
        "delivered_date": delivered_date,
        "da_date": da_date,
        "hot_flag": bool(hot),
        "hold_flag": bool(hold),
    }

# First non-null date in priority order (NOT earliest-of-all)
# (DA from notes inserted after ATD and before DUE DATE)
ETA_PRIORITY = [
    "ETA AT PLACE OF DELIVERY", "ETA", "ETD", "ATD",
    # DA from notes considered next
    "DUE DATE", "RQD. DATE", "LOAD DATE", "ORDER DATE"
]

def best_eta_for_row(row: pd.Series):
    # 1) core ETA fields in priority order
    for col in ["ETA AT PLACE OF DELIVERY", "ETA", "ETD", "ATD"]:
        if col in row.index and pd.notna(row[col]):
            return to_dt(pd.Series([row[col]])).iloc[0]
    # 2) DA from NOTES
    da = analyze_notes(row)["da_date"]
    if pd.notna(da):
        return da
    # 3) fallback planning dates
    for col in ["DUE DATE", "RQD. DATE", "LOAD DATE", "ORDER DATE"]:
        if col in row.index and pd.notna(row[col]):
            return to_dt(pd.Series([row[col]])).iloc[0]
    return pd.NaT

def compute_status_flag(row: pd.Series) -> bool:
    """
    Exclude only if:
      - current status (CLEAN/SHIPMENT/LAST/STATUS) indicates delivered/closed, OR
      - NOTES explicitly shows 'Delivered: mm/dd', OR
      - INVOICED is truthy
    Ignore container-history phrases like 'Received (FCL)' in NOTES.
    """
    fields = []
    for col in ["CLEAN STATUS", "SHIPMENT STATUS", "LAST STATUS", "STATUS"]:
        if col in row.index and pd.notna(row[col]):
            fields.append(str(row[col]).strip().lower())
    joined = " ".join(fields)

    exclude_keywords = ["delivered", "closed"]
    delivered_like = any(k in joined for k in exclude_keywords)

    delivered_in_notes = analyze_notes(row)["delivered_date"]
    delivered_flag = pd.notna(delivered_in_notes)

    invoiced = False
    if "INVOICED" in row.index:
        inv_val = row["INVOICED"]
        invoiced = (pd.notna(inv_val) and str(inv_val).strip().lower() not in ["", "0", "0.0", "nan", "no", "false"])

    return delivered_like or delivered_flag or invoiced

# -------------------- Master & Client builders --------------------

def normalize_master(master_raw: pd.DataFrame) -> pd.DataFrame:
    m = master_raw.copy()
    m.columns = m.columns.str.strip().str.upper()

    if "ITEM #" not in m.columns or "PO/LOT" not in m.columns:
        raise ValueError("Master must contain columns 'ITEM #' and 'PO/LOT' in sheet OPEN_ORDERS.")

    # Optional: map common header aliases (expand if you encounter variants)
    HEADER_ALIASES = {
        "ETA PLACE OF DELIVERY": "ETA AT PLACE OF DELIVERY",
        "ETA PLACE DELIVERY": "ETA AT PLACE OF DELIVERY",
        "ETA POD": "ETA AT PLACE OF DELIVERY",
        "SHIPMENT": "SHIPMENT STATUS",  # unify if needed
    }
    m.columns = [HEADER_ALIASES.get(c, c) for c in m.columns]

    m["ITEM #"] = m["ITEM #"].apply(clean_key).str.lstrip("0")
    m["PO/LOT"] = m["PO/LOT"].apply(clean_key)

    for dcol in list(set(ETA_PRIORITY + ["UPDATE D"])):
        if dcol in m.columns:
            m[dcol] = to_dt(m[dcol])

    if "QTY" in m.columns:
        m["QTY"] = pd.to_numeric(m["QTY"], errors="coerce")
    else:
        m["QTY"] = np.nan

    m["_EXCLUDE"] = m.apply(compute_status_flag, axis=1)
    m["AVAILABLE_QTY"] = m["QTY"].fillna(0)
    return m

def _auto_find_sheet(sheet_names, keywords):
    for sn in sheet_names:
        if any(k in _lower_name(sn) for k in keywords):
            return sn
    return None

def build_client_from_raw_or_prepared(uploaded_client_file) -> pd.DataFrame:
    # Try prepared single-sheet format first
    df0 = pd.read_excel(uploaded_client_file, sheet_name=0, engine="openpyxl")
    df0 = _norm_cols(df0)
    prepared = {"ITEM #","PO/LOT","Weekly Sales","Total On Hand","Total SO"}
    if prepared.issubset(df0.columns):
        df0["ITEM #"] = df0["ITEM #"].apply(_as_item_key)
        for c in ["Weekly Sales","Total On Hand","Total SO"]:
            df0[c] = pd.to_numeric(df0[c], errors="coerce").fillna(0)
        return df0[["ITEM #","PO/LOT","Weekly Sales","Total On Hand","Total SO"]]

    # Else treat as raw workbook
    xl = pd.ExcelFile(uploaded_client_file, engine="openpyxl")
    sheet_names = xl.sheet_names

    forecast_guess = _auto_find_sheet(sheet_names, ["forecast","per item","per-item"])
    avail_guess    = _auto_find_sheet(sheet_names, ["avail","availability","summary"])

    if not forecast_guess or not avail_guess:
        st.warning("Couldn’t auto-detect Forecast/Availability sheets. Please select them below.")
        forecast_guess = st.selectbox("Select the Forecast sheet", sheet_names, index=0)
        avail_default_idx = 1 if len(sheet_names) > 1 else 0
        avail_guess = st.selectbox("Select the Availability sheet", sheet_names, index=avail_default_idx)

    fc = pd.read_excel(uploaded_client_file, sheet_name=forecast_guess, engine="openpyxl")
    av = pd.read_excel(uploaded_client_file, sheet_name=avail_guess,    engine="openpyxl")
    fc = _norm_cols(fc); av = _norm_cols(av)

    fc_item_col = next((c for c in fc.columns if _lower_name(c) in ["item","fin item#","fin item #","item #","item#"]), None) or fc.columns[0]
    av_item_col = next((c for c in av.columns if _lower_name(c) in ["fin item#","fin item #","item","item #","item#"]), None) or av.columns[0]

    fc["_ITEM"] = fc[fc_item_col].apply(_as_item_key)
    av["_ITEM"] = av[av_item_col].apply(_as_item_key)

    weekly_cols = _find_weekly_cols(fc)
    av_wk_col = next((c for c in av.columns if _lower_name(c) in
                     ["wkly sls","weekly sales","wkly sales","avg wkly sls","avg weekly sales"]), None)

    if av_wk_col:
        wkly = av[["_ITEM", av_wk_col]].rename(columns={av_wk_col: "Weekly Sales"})
    else:
        if not weekly_cols:
            weekly_cols = [c for c in fc.columns if pd.to_datetime(str(c), errors="coerce").notna()]
        if not weekly_cols:
            raise ValueError("Couldn't detect weekly forecast columns on the selected Forecast sheet.")
        wkly = (
            fc.set_index("_ITEM")[weekly_cols]
              .apply(pd.to_numeric, errors="coerce")
              .mean(axis=1)
              .rename("Weekly Sales")
              .reset_index()
        )

    fc_oh_col = next((c for c in fc.columns if _lower_name(c) in
                     ["current on hand qty","current on hand","on hand","oh","total on hand"]), None)
    av_oh_col = next((c for c in av.columns if _lower_name(c) in
                     ["total oh","oh","on hand","total on hand"]), None)
    if fc_oh_col:
        oh_df = fc[["_ITEM", fc_oh_col]].rename(columns={fc_oh_col: "Total On Hand"})
    elif av_oh_col:
        oh_df = av[["_ITEM", av_oh_col]].rename(columns={av_oh_col: "Total On Hand"})
    else:
        oh_df = pd.DataFrame({"_ITEM": av["_ITEM"].unique(), "Total On Hand": 0})

    av_so_col = next((c for c in av.columns if _lower_name(c) in
                     ["total so","so","open so","open sales orders","sales orders"]), None)
    if av_so_col:
        so_df = av[["_ITEM", av_so_col]].rename(columns={av_so_col: "Total SO"})
    else:
        so_df = pd.DataFrame({"_ITEM": av["_ITEM"].unique(), "Total SO": 0})

    client = wkly.merge(oh_df, on="_ITEM", how="outer").merge(so_df, on="_ITEM", how="outer")
    client.rename(columns={"_ITEM": "ITEM #"}, inplace=True)
    client["ITEM #"] = client["ITEM #"].apply(_as_item_key)

    for c in ["Weekly Sales","Total On Hand","Total SO"]:
        client[c] = pd.to_numeric(client[c], errors="coerce").fillna(0)

    client["PO/LOT"] = ""
    client = client[["ITEM #","PO/LOT","Weekly Sales","Total On Hand","Total SO"]]
    return client

# -------------------- Picker --------------------

def select_best_po_for_item(master_norm: pd.DataFrame, item: str, maybe_code: str | None = None) -> dict | None:
    sub = master_norm[master_norm["ITEM #"] == item].copy()
    if "C. CODE" in master_norm.columns and maybe_code:
        sub = sub[sub["C. CODE"].astype(str).str.strip() == str(maybe_code).strip()]

    # Filter after computing _EXCLUDE during normalization
    sub = sub[(~sub["_EXCLUDE"]) & (sub["AVAILABLE_QTY"] > 0)]
    if sub.empty:
        return None

    # Per-row date resolver (ETA/ETD/ATD → DA → DUE/RQD/LOAD/ORDER)
    sub["_ETA"] = sub.apply(best_eta_for_row, axis=1)

    od = "ORDER DATE" if "ORDER DATE" in sub.columns else None

    def po_num(x):
        try:
            return int(re.sub(r"\D","",str(x)))
        except Exception:
            return 10**12

    sub["_PO_NUM"] = sub["PO/LOT"].apply(po_num)
    sub["_ETA_SORT"] = sub["_ETA"].fillna(pd.Timestamp.max)

    # Sort: earliest ETA-like date, then available qty (desc), then earliest order date, then smaller PO number
    sort_by = ["_ETA_SORT", "AVAILABLE_QTY"]
    ascending = [True, False]
    if od:
        sort_by.append(od); ascending.append(True)
    sort_by.append("_PO_NUM"); ascending.append(True)

    chosen = sub.sort_values(by=sort_by, ascending=ascending).iloc[0]
    delivery_date = chosen["_ETA"]

    return {
        "PO/LOT": chosen["PO/LOT"],
        "Delivery Date": delivery_date,
        "Delivery Qty": chosen["QTY"],
    }

# -------------------- App --------------------

if uploaded_master and uploaded_client:
    if st.button("🚀 Generate Report"):
        with st.spinner("Processing data..."):
            # Master
            try:
                master_raw = pd.read_excel(uploaded_master, sheet_name="OPEN_ORDERS", engine="openpyxl")
            except ValueError:
                st.error("❌ Could not find a sheet named 'OPEN_ORDERS' in the master file.")
                st.stop()

            # Client
            try:
                client = build_client_from_raw_or_prepared(uploaded_client)
            except Exception as e:
                st.error(f"❌ Error reading client file: {e}")
                st.stop()

            # Normalize master
            try:
                master = normalize_master(master_raw)
            except Exception as e:
                st.error(f"❌ Error normalizing master file: {e}")
                st.stop()

            # Ensure schema
            for col in ["ITEM #","PO/LOT","Weekly Sales","Total On Hand","Total SO"]:
                if col not in client.columns:
                    st.error(f"❌ Client file is missing column: '{col}'")
                    st.stop()

            # Auto-pick PO/LOT
            st.info("🔎 Selecting best PO/LOT for each client item…")
            has_client_code = "C. CODE" in client.columns
            picks = []
            for idx, row in client.iterrows():
                item = row["ITEM #"]
                code = row["C. CODE"] if has_client_code else None
                sel = select_best_po_for_item(master, item, code)
                if sel is None:
                    picks.append({"idx": idx, "ITEM #": item, "PO/LOT": "", "Delivery Date": pd.NaT, "Delivery Qty": np.nan})
                else:
                    picks.append({"idx": idx, "ITEM #": item, **sel})

            picks_df = pd.DataFrame(picks).set_index("idx")

            # Ensure target cols exist
            for c, default in [("Delivery Date", pd.NaT), ("Delivery Qty", np.nan), ("PO/LOT","")]:
                if c not in client.columns:
                    client[c] = default

            # Update client with picks
            client.update(picks_df[["PO/LOT","Delivery Date","Delivery Qty"]])

            # ---------- Forecast calculations ----------
            df = client.copy()
            df["Today"] = today
            df["Today Week"] = pd.to_datetime(today).isocalendar().week
            df["Delivery Week No."] = to_dt(df["Delivery Date"]).dt.isocalendar().week

            total_same = (
                master.groupby("ITEM #", as_index=False)["QTY"]
                      .sum()
                      .rename(columns={"QTY": "Total same item"})
            )
            df = df.merge(total_same, on="ITEM #", how="left")
            df["Total same item"] = df["Total same item"].fillna(0)

            E = pd.to_numeric(df.get("Weekly Sales"), errors="coerce")
            F = pd.to_numeric(df.get("Total On Hand"), errors="coerce")
            G = pd.to_numeric(df.get("Total SO"), errors="coerce")
            deliv_qty = pd.to_numeric(df.get("Delivery Qty", 0), errors="coerce")

            df["Stock + Ordered"]     = (df.get("Total same item", 0).fillna(0) + F - G)
            df["Weeks in Stock"]      = (F - G) / E
            df["Months in Stock"]     = df["Weeks in Stock"] / 4.25
            df["Under/Over"]          = df["Months in Stock"] - 2

            df["Weeks in Supply"]     = df["Stock + Ordered"] / E
            df["Months in Supply"]    = df["Weeks in Supply"] / 4.25
            df["Under/Over2"]         = df["Months in Supply"] - 3.85
            df["Suggested Order Qty"] = -(df["Under/Over2"]) * 4.25 * E

            df["To reach 0 (wks)"]    = ((df["Stock + Ordered"] / E).fillna(0)).astype(int)

            df["Day 0"]            = pd.to_datetime(today) + pd.to_timedelta(df["To reach 0 (wks)"] * 7, unit="D")
            df["Day 0"]            = pd.to_datetime(df["Day 0"], errors="coerce").dt.date

            df["Re order"]         = pd.to_datetime(df["Day 0"]) - timedelta(weeks=11)
            df["Place order 1"]    = pd.to_datetime(df["Re order"]) - timedelta(weeks=12)
            df["Before today"]     = (pd.to_datetime(df["Place order 1"]) - pd.to_datetime(today)).dt.days
            df["Re orders no."]    = (deliv_qty / E) * 7

            df["2"] = pd.to_datetime(df["Place order 1"]) + pd.to_timedelta(df["Re orders no."], unit="D")
            df["3"] = df["2"] + pd.to_timedelta(df["Re orders no."], unit="D")
            df["4"] = df["3"] + pd.to_timedelta(df["Re orders no."], unit="D")

            for col in ["Delivery Date","Day 0","Re order","Place order 1","2","3","4"]:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce").dt.date

            num_cols = [
                "Weeks in Stock","Months in Stock","Under/Over","Stock + Ordered",
                "Weeks in Supply","Months in Supply","Under/Over2","Suggested Order Qty",
                "Before today","Re orders no.","Total same item"
            ]
            for col in num_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").round(2)

            final_cols = [
                "ITEM #","PO/LOT","Delivery Date","Delivery Week No.","Today","Today Week",
                "Weekly Sales","Total On Hand","Total SO","Total same item","Stock + Ordered",
                "Weeks in Stock","Months in Stock","Under/Over","Delivery Qty",
                "Weeks in Supply","Months in Supply","Under/Over2","Suggested Order Qty",
                "To reach 0 (wks)","Day 0","Re order","Place order 1","Before today",
                "Re orders no.","2","3","4"
            ]
            df = df[[c for c in final_cols if c in df.columns]]

            def style_under_over(val):
                if pd.isna(val):
                    return ''
                if val < 0:
                    return 'background-color: red; color: white;'
                if val > 1.5:
                    return 'background-color: yellow;'
                return ''

            def smart_format(x):
                if pd.isna(x):
                    return ""
                if isinstance(x, (int, float, np.integer, np.floating)):
                    if float(x).is_integer():
                        return f"{int(x)}"
                    return f"{float(x):.2f}"
                return str(x)

            styled = df.style.applymap(style_under_over, subset=['Under/Over']).format(smart_format)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Inventory Report")
                audit = client[["ITEM #","PO/LOT","Delivery Date","Delivery Qty"]].copy()
                audit.to_excel(writer, index=False, sheet_name="PO Picks")
            output.seek(0)

        st.success("✅ Report generated!")
        st.dataframe(styled, use_container_width=True)
        st.download_button(
            label="📥 Download Excel Report",
            data=output,
            file_name=f"inventory_report_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Please upload both files to proceed.")
