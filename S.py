
import streamlit as st
import pandas as pd
from script import cdc_receipting_ui

# -------------------------------------------------
# PAGE SETUP
# -------------------------------------------------
st.title("Trades Reconciliation System")

st.caption("ZSE & Sharestock reconciliation with settlement summaries")

# -------------------------------------------------
# RECEIPTING TYPE (ONLY ONE)
# -------------------------------------------------
receipting_type = st.radio(
    "Select Receipting Type",
    ["ZSE Receipting", "CDC Receipting"],
    horizontal=True,
)

st.divider()

# -------------------------------------------------
# CDC RECEIPTING (ONLY SHOWN WHEN SELECTED)
# -------------------------------------------------
if receipting_type == "CDC Receipting":
    cdc_receipting_ui()
    st.stop()   # ⛔ stop here, prevent ZSE from loading

# -------------------------------------------------
# ZSE RECEIPTING (ONLY SHOWN WHEN SELECTED)
# -------------------------------------------------

import camelot
from io import BytesIO
import tempfile
import os

# =====================================================
# PAGE CONFIG
# =====================================================
st.markdown(
"""
<style>

/* ==================================================
   GLOBAL APP BACKGROUND
================================================== */
.stApp {
    background-color: #f5f7fb; /* light dashboard grey */
    color: #0f172a;
    font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

/* ==================================================
   MAIN CONTENT AREA
================================================== */
.main > div {
    background: transparent;
    padding: 1.5rem 2rem;
}

/* ==================================================
   HEADERS
================================================== */
h1 {
    color: #0A2540; /* navy */
    font-weight: 700;
}

h2, h3 {
    color: #102a43;
    font-weight: 600;
}

/* ==================================================
   CARDS (KPI / SECTIONS)
================================================== */
.card,
.main > div > div {
    background: #ffffff;
    border-radius: 14px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
}

/* ==================================================
   BUTTONS (MATCH DASHBOARD BLUE)
================================================== */
div.stButton > button {
    background: #1d4ed8; /* strong dashboard blue */
    color: #ffffff !important;
    border-radius: 10px;
    padding: 0.55rem 1.4rem;
    font-weight: 600;
    border: none;
    box-shadow: 0 6px 16px rgba(29, 78, 216, 0.35);
    transition: all 0.2s ease;
}

div.stButton > button:hover {
    background: #1e40af;
    transform: translateY(-1px);
    box-shadow: 0 10px 26px rgba(29, 78, 216, 0.5);
}

/* ==================================================
   INPUTS & SELECTS
================================================== */
input, textarea, select {
    background-color: #ffffff !important;
    border: 1px solid #d1d5db !important;
    border-radius: 8px !important;
    color: #0f172a !important;
}

input:focus, textarea:focus {
    border-color: #2563eb !important;
    box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.2);
}

/* ==================================================
   RADIO / FILTER PILLS
================================================== */
.stRadio div[role="radiogroup"] label {
    background: #ffffff;
    border: 1px solid #c7d2fe;
    border-radius: 999px;
    padding: 0.3rem 1rem;
    color: #1d4ed8;
    font-weight: 500;
}

/* ==================================================
   DATAFRAMES – MATCH IMAGE STYLE
================================================== */
.stDataFrame {
    background: #ffffff !important;
    border-radius: 12px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 6px 16px rgba(15, 23, 42, 0.08);
}

/* Table headers */
.stDataFrame thead tr th {
    background-color: #dbe4f1 !important; /* light steel blue */
    color: #102a43 !important;
    font-weight: 600;
    border-bottom: 1px solid #cbd5e1 !important;
}

/* Table body */
.stDataFrame tbody tr td {
    background-color: #ffffff !important;
    color: #0f172a !important;
    border-bottom: 1px solid #f1f5f9 !important;
}

/* Row hover */
.stDataFrame tbody tr:hover td {
    background-color: #f1f5fb !important;
}

/* ==================================================
   ALERTS / STATUS
================================================== */
.stAlert {
    background: #ffffff !important;
    border: 1px solid #bfdbfe;
    border-radius: 10px;
    color: #1e3a8a;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08);
}

/* ==================================================
   DIVIDERS
================================================== */
hr {
    border: none;
    height: 1px;
    background: #e5e7eb;
    margin: 1.5rem 0;
}

</style>
""",
unsafe_allow_html=True
)


# =====================================================
# SESSION STATE DEFAULTS (AVOID KEY ERRORS)
# =====================================================
st.session_state.setdefault("sorted", False)
st.session_state.setdefault("zse_df", None)
st.session_state.setdefault("sh_df", None)
st.session_state.setdefault("show_final_summary", False)


# =====================================================
# UTILITY FUNCTIONS (SAFE HELPERS)
# =====================================================
def display_table_with_commas(
    df: pd.DataFrame,
    hide_index=False,
    hide_columns=False,
):
    """
    Display DataFrame with:
    - Comma thousands separator
    - 2 decimal places
    - FIRST COLUMN bold only
    - Other columns normal
    """
    def bold_first_col(col):
        # Bold all non-empty values in first column only
        return ["font-weight: bold" if i == 0 else "" for i in [0] * len(col)]

    styled = df.style.format(precision=2, thousands=",")

    # ✅ Bold first column only
    styled = styled.apply(
        lambda row: ["font-weight: bold"] + [""] * (len(row) - 1),
        axis=1,
    )

    if hide_index:
        styled = styled.hide(axis="index")
    if hide_columns:
        styled = styled.hide(axis="columns")

    st.dataframe(styled, use_container_width=True)

def safe_find_col(df: pd.DataFrame, keywords):
    """
    Safely find the first column whose name contains any keyword.
    Returns None if not found.
    """
    keywords = [k.lower() for k in keywords]
    for c in df.columns:
        cl = str(c).lower()
        if any(k in cl for k in keywords):
            return c
    return None


def ensure_numeric(series: pd.Series):
    """Convert a Series to numeric, coercing errors to NaN."""
    return pd.to_numeric(series, errors="coerce")




# =====================================================
# ZSE RECEIPTING
# =====================================================
if receipting_type == "ZSE Receipting":

    zse_file = st.file_uploader("Upload ZSE Excel File", type=["xlsx", "xls"])

    if zse_file:
        raw_zse = pd.read_excel(zse_file, engine="openpyxl")
        raw_zse.columns = raw_zse.columns.str.lower().str.strip()

        buy_sell_col = safe_find_col(raw_zse, ["buy"])
        if buy_sell_col is None:
            st.error("Could not automatically detect a BUY/SELL column in the ZSE file.")
        else:
            if st.button("Sort ZSE"):
                df = raw_zse.copy()
                df[buy_sell_col] = df[buy_sell_col].astype(str).str.upper()

                # Exclude the 'due' column from totals
                numeric_cols = [
                    c for c in df.select_dtypes(include="number").columns
                    if str(c).lower().strip() != "due"
                ]


                def zse_totals(data, label):
                    total = {
                        c: data[c].sum() if c in numeric_cols else ""
                        for c in data.columns
                    }
                    total[buy_sell_col] = label
                    return pd.concat([data, pd.DataFrame([total])], ignore_index=True)

                st.session_state.zse_df = pd.concat(
                    [
                        zse_totals(df[df[buy_sell_col] == "BUY"], "BUY TOTAL"),
                        zse_totals(df[df[buy_sell_col] == "SELL"], "SELL TOTAL"),
                    ],
                    ignore_index=True,
                )
                st.session_state.sorted = True

        if st.session_state.zse_df is not None:
            st.subheader("ZSE Data (with Totals)")
            st.dataframe(st.session_state.zse_df, use_container_width=True)

    # =====================================================
    # SHARESTOCK SECTION
    # =====================================================
    if st.session_state.sorted:

        st.divider()
        st.subheader("Sharestock Data")

        sh_file = st.file_uploader(
            "Upload Sharestock Excel File", type=["xlsx", "xls"], key="sh"
        )

        if sh_file:
            raw = pd.read_excel(sh_file, header=None, engine="openpyxl")

            # Robust header row detection for "CLIENT"
            header_idx = None
            for i, v in raw.iloc[:, 0].items():
                if isinstance(v, str) and v.strip().upper() == "CLIENT":
                    header_idx = i
                    break

            if header_idx is None:
                st.error("Could not find a 'CLIENT' header row in the Sharestock file.")
            else:
                headers = raw.loc[header_idx]
                clean_headers, seen = [], {}

                for i, h in enumerate(headers):
                    name = str(h).strip() if pd.notna(h) and str(h).strip() else f"UNNAMED_{i}"
                    if name in seen:
                        seen[name] += 1
                        name = f"{name}_{seen[name]}"
                    else:
                        seen[name] = 0
                    clean_headers.append(name)

                sh_df = raw.loc[header_idx + 1 :].copy()
                sh_df.columns = clean_headers
                sh_df = sh_df.loc[:, ~sh_df.columns.str.startswith("UNNAMED")]
                sh_df = sh_df.dropna(how="all").reset_index(drop=True)

                # Convert Qty and all columns after to numeric
                qty_col = safe_find_col(sh_df, ["qty"])
                if qty_col is None:
                    st.error("Could not find a 'Qty' column in the Sharestock data.")
                else:
                    qty_idx = sh_df.columns.get_loc(qty_col)

                    for col in sh_df.columns[qty_idx:]:
                        sh_df[col] = (
                            sh_df[col]
                            .astype(str)
                            .str.replace(",", "", regex=False)
                            .str.strip()
                        )
                        sh_df[col] = ensure_numeric(sh_df[col])

                    st.session_state.sh_df = sh_df
                    st.dataframe(sh_df, use_container_width=True)

        # =================================================
        # MATCH + TOTALS + ZSE + VARIANCE + SUMMARY
        # =================================================
        if (
            st.button("Match Sharestock to ZSE")
            and st.session_state.zse_df is not None
            and st.session_state.sh_df is not None
        ):
            zse = st.session_state.zse_df.copy()
            sh = st.session_state.sh_df.copy()

            zse_type = safe_find_col(zse, ["buy"])
            zse_sym = safe_find_col(zse, ["symbol", "security", "counter"])
            sh_type = safe_find_col(sh, ["type"])
            sh_sym = safe_find_col(sh, ["symbol", "security"])

            if None in (zse_type, zse_sym, sh_type, sh_sym):
                st.error(
                    "Could not reliably detect transaction type / symbol columns in one of the files."
                )
            else:
                # Normalise types
                sh[sh_type] = sh[sh_type].replace(
                    {"BUY": "PURCHASE", "SELL": "SALE"}
                )
                zse["_MATCH_"] = zse[zse_type].map(
                    {"BUY": "PURCHASE", "SELL": "SALE"}
                )

                zse_clean = zse[
                    ~zse[zse_type].isin(["BUY TOTAL", "SELL TOTAL"])
                ].copy()
                keys = set(zip(zse_clean[zse_sym], zse_clean["_MATCH_"]))

                matched = sh[
                    sh.apply(
                        lambda r: (r[sh_sym], r[sh_type]) in keys, axis=1
                    )
                ]

                # Sharestock totals
                def add_total_block(data, label):
                    total = {}
                    for col in data.columns:
                        if "price" in col.lower():
                            total[col] = ""
                        elif pd.api.types.is_numeric_dtype(data[col]):
                            total[col] = data[col].sum()
                        else:
                            total[col] = ""
                    total[sh_type] = label
                    return pd.concat(
                        [data, pd.DataFrame([total])], ignore_index=True
                    )

                purchase_df = add_total_block(
                    matched[matched[sh_type] == "PURCHASE"], "PURCHASE TOTAL"
                )
                sale_df = add_total_block(
                    matched[matched[sh_type] == "SALE"], "SALE TOTAL"
                )

                # Mapping ZSE totals into Sharestock columns
                mapping = {
                    "qty": ["quantity"],
                    "gross proceeds": ["gross consideration"],
                    "brokerage": ["broker", "commission"],
                    "basic": ["basic"],
                    "stamp": ["stamp"],
                    "zse levy": ["zse levy"],
                    "csd levy": ["depository levy"],
                    "cgt": ["capital gains tax"],
                    "sec levy": ["secz levy"],
                    "investor levy": ["inv protection vx"],
                    "vat": ["vat"],
                }

                def append_zse(df, zse_row, label):
                    row = {}
                    for sh_col in df.columns:
                        row[sh_col] = ""
                        if sh_col == sh_type:
                            row[sh_col] = label
                            continue
                        for sh_key, zse_keys in mapping.items():
                            if sh_key in sh_col.lower():
                                for zc in zse_row.index:
                                    if any(k in zc.lower() for k in zse_keys):
                                        row[sh_col] = zse_row[zc]
                    return pd.concat(
                        [df, pd.DataFrame([row])], ignore_index=True
                    )

                # Guard against missing BUY/SELL TOTAL rows
                zse_buy_total = zse[zse[zse_type] == "BUY TOTAL"]
                zse_sell_total = zse[zse[zse_type] == "SELL TOTAL"]

                if not zse_buy_total.empty:
                    purchase_df = append_zse(
                        purchase_df,
                        zse_buy_total.iloc[0],
                        "ZSE BUY TOTAL",
                    )
                if not zse_sell_total.empty:
                    sale_df = append_zse(
                        sale_df,
                        zse_sell_total.iloc[0],
                        "ZSE SELL TOTAL",
                    )

                # Variance
                def append_variance(
                    df, total_label, zse_label, variance_label
                ):
                    total_rows = df[df[sh_type] == total_label]
                    zse_rows = df[df[sh_type] == zse_label]
                    if total_rows.empty or zse_rows.empty:
                        return df

                    t = total_rows.iloc[0]
                    z = zse_rows.iloc[0]

                    row = {}
                    for col in df.columns:
                        if col == sh_type:
                            row[col] = variance_label
                        elif pd.api.types.is_numeric_dtype(df[col]):
                            tv = pd.to_numeric(t[col], errors="coerce")
                            zv = pd.to_numeric(z[col], errors="coerce")
                            tv = 0 if pd.isna(tv) else tv
                            zv = 0 if pd.isna(zv) else zv
                            row[col] = tv - zv
                        else:
                            row[col] = ""
                    return pd.concat(
                        [df, pd.DataFrame([row])], ignore_index=True
                    )

                purchase_df = append_variance(
                    purchase_df,
                    "PURCHASE TOTAL",
                    "ZSE BUY TOTAL",
                    "VARIANCE (PURCHASE - ZSE BUY)",
                )
                sale_df = append_variance(
                    sale_df,
                    "SALE TOTAL",
                    "ZSE SELL TOTAL",
                    "VARIANCE (SALE - ZSE SELL)",
                )

                # Display results
                st.success("Full reconciliation and settlement summary complete")

                st.subheader("Sharestock – Purchases (Reconciled)")
                st.dataframe(purchase_df, use_container_width=True)

                st.subheader("Sharestock – Sales (Reconciled)")
                st.dataframe(sale_df, use_container_width=True)

                # Store these for potential future use (e.g. auto-filling summary)
                st.session_state.purchase_df = purchase_df
                st.session_state.sale_df = sale_df



# =================================================
# HELPER: SUM LEVIES FROM RECONCILED TABLE TOTAL ROW
# =================================================
def sum_levies(df: pd.DataFrame, type_col: str, total_label: str) -> float:
    if df is None or type_col not in df.columns:
        return 0.0

    total_row = df[df[type_col] == total_label]
    if total_row.empty:
        return 0.0

    levy_keywords = ["broker", "stamp", "zse levy", "sec levy", "investor", "vat","cgt"]

    total = 0.0
    for col in df.columns:
        if any(k in col.lower() for k in levy_keywords):
            val = pd.to_numeric(total_row.iloc[0][col], errors="coerce")
            if pd.notna(val):
                total += val

    return round(total, 2)


def extract_specific_levy(df, type_col, total_label, keyword):
    if df is None or type_col not in df.columns:
        return 0.0

    total_row = df[df[type_col] == total_label]
    if total_row.empty:
        return 0.0

    for col in df.columns:
        if keyword in col.lower():
            val = pd.to_numeric(total_row.iloc[0][col], errors="coerce")
            return 0.0 if pd.isna(val) else float(val)

    return 0.0

# =================================================
# FINAL SETTLEMENT SUMMARY
# =================================================
if st.button("Final Settlement Summary"):
    st.session_state.show_final_summary = True


# =================================================
# SHOW SUMMARY + RECALCULATE ON EVERY RERUN
# =================================================
if st.session_state.show_final_summary:

    st.subheader("Final Settlement Summary")

    # -------------------------
    # ✅ BANK INPUT (ALWAYS LIVE)
    # -------------------------
    bank_statement_amount = st.number_input(
        "Bank Statement Amount",
        min_value=0.0,
        value=0.0,
        step=0.01,
        format="%.2f",
    )

    # -------------------------
    # ✅ CALCULATIONS
    # -------------------------
    total_purchases_value = 0.0
    total_sales_value = 0.0
    zse_levy_remitted = 0.0
    ipl_levy_remitted = 0.0
    sec_levy_total = 0.0

    if "purchase_df" in st.session_state:
        p = st.session_state.purchase_df
        p_type = safe_find_col(p, ["type"])
        if p_type:
            total_purchases_value = sum_levies(p, p_type, "PURCHASE TOTAL")
            zse_levy_remitted += extract_specific_levy(p, p_type, "PURCHASE TOTAL", "zse levy")
            ipl_levy_remitted += extract_specific_levy(p, p_type, "PURCHASE TOTAL", "investor")
            sec_levy_total += extract_specific_levy(p, p_type, "PURCHASE TOTAL", "sec levy")

    if "sale_df" in st.session_state:
        s = st.session_state.sale_df
        s_type = safe_find_col(s, ["type"])
        if s_type:
            total_sales_value = sum_levies(s, s_type, "SALE TOTAL")
            zse_levy_remitted += extract_specific_levy(s, s_type, "SALE TOTAL", "zse levy")
            ipl_levy_remitted += extract_specific_levy(s, s_type, "SALE TOTAL", "investor")
            sec_levy_total += extract_specific_levy(s, s_type, "SALE TOTAL", "sec levy")

    total_purchases_and_sales = round(total_purchases_value + total_sales_value, 2)

    total_amount_to_be_received = round(
        total_purchases_and_sales
        - (zse_levy_remitted + ipl_levy_remitted + sec_levy_total),
        2,
    )

    balance_from_bank = round(
        bank_statement_amount - total_amount_to_be_received, 2
    )

    # -------------------------
    # ✅ POST SETTLEMENT TOTAL
    # -------------------------
    post_settlement_total = 0.0
    if st.session_state.zse_df is not None:
        due_col = next(
            (c for c in st.session_state.zse_df.columns if str(c).lower().strip() == "due"),
            None,
        )
        if due_col:
            post_settlement_total = pd.to_numeric(
                st.session_state.zse_df[due_col], errors="coerce"
            ).sum()

    post_settlement_total = round(post_settlement_total, 2)

    difference_bank_vs_post_settlement = round(
        bank_statement_amount - post_settlement_total, 2
    )

    capital_gains_variance = 0.0
    if "sale_df" in st.session_state:
        s = st.session_state.sale_df
        s_type = safe_find_col(s, ["type"])
        if s_type:
            vr = s[s[s_type] == "VARIANCE (SALE - ZSE SELL)"]
            if not vr.empty:
                cgt_col = next(c for c in vr.columns if "cgt" in c.lower())
                val = pd.to_numeric(vr.iloc[0][cgt_col], errors="coerce")
                capital_gains_variance = 0.0 if pd.isna(val) else round(float(val), 4)

    verify_value = round(
        bank_statement_amount + capital_gains_variance - post_settlement_total, 2
    )

    # -------------------------
    # ✅ SUMMARY TABLE
    # -------------------------
    summary_df = pd.DataFrame(
        [
            ["Total Purchases", total_purchases_value, "", "", "", ""],
            ["Total Sales", total_sales_value, "", "", "", ""],
            ["Total Purchases and Sales", total_purchases_and_sales, "", "", "", ""],
            ["", "", "", "", "", ""],
            ["ZSE Levy Remitted Directly", zse_levy_remitted, "", "", "", ""],
            ["IPL Levy Remitted Directly", ipl_levy_remitted, "", "", "", ""],
            ["SEC Levy", sec_levy_total, "", "", "", ""],
            [
                "", "","",
                "Bank Statement Amount",
                "Balance from Amount to be Received and Bank Amount",
                "Capital Gains",
                "Post Settlement Total from Report",
                "Difference Between Bank and Post Settlement",
                "Verify",
            ],
            [
                "Total Amount to be Received",
                total_amount_to_be_received,
                "",
                bank_statement_amount,
                balance_from_bank,
                capital_gains_variance,
                post_settlement_total,
                difference_bank_vs_post_settlement,
                verify_value,
                "",
            ],
            ["", "", "", "", "", ""],
            ["", "", "", "", "", ""],
            ["Prepared by:....................................................................................................", "", "", "", "", ""],
            ["", "", "", "", "", ""],
            ["Reviewed by:....................................................................................................", "", "", "", "", ""],
        ]
    )
    display_table_with_commas(
        summary_df,
        hide_index=True,
        hide_columns=True
    )


    # =================================================
    # ✅ EXCEL DOWNLOAD (AFTER TABLE)
    # =================================================
    from io import BytesIO
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment
    from openpyxl.styles import Border, Side

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # =========================
        # WRITE DATA
        # =========================
        st.session_state.sh_df.to_excel(
            writer, sheet_name="Sharestock Raw", index=False
        )

        st.session_state.purchase_df.to_excel(
            writer, sheet_name="Purchases", index=False
        )

        sales_sheet = "Sales & Summary"  # ✅ FIXED

        st.session_state.sale_df.to_excel(
            writer,
            sheet_name=sales_sheet,
            index=False,
            startrow=0
        )
        summary_df.to_excel(
            writer,
            sheet_name=sales_sheet,
            index=False,
            startrow=len(st.session_state.sale_df) + 3,
            header=False
        )

        # =========================
        # FORMATTING STARTS HERE
        # =========================
        bold = Font(bold=True)

        # ---------- SALES & SUMMARY ----------
        ws_sales = writer.sheets[sales_sheet]

        # 1️⃣ Bold HEADER ROW
        for cell in ws_sales[1]:
            cell.font = bold

        # 2️⃣ Find FIRST EMPTY CELL in COLUMN A
        first_empty_row = None
        for r in range(2, ws_sales.max_row + 1):
            val = ws_sales.cell(row=r, column=1).value
            if val is None or str(val).strip() == "":
                first_empty_row = r
                break

        # 3️⃣ Force-create + bold COLUMN A from empty row down
        if first_empty_row:
            for r in range(first_empty_row, ws_sales.max_row + 1):
                cell = ws_sales.cell(row=r, column=1)
                if cell.value is None:
                    cell.value = ""
                cell.font = bold

        # 4️⃣ Bold LAST 3 FULL ROWS with values in column 12
        rows_with_values = []
        for r in range(ws_sales.max_row, 1, -1):
            val = ws_sales.cell(row=r, column=17).value
            if val not in (None, "", " "):
                rows_with_values.append(r)
            if len(rows_with_values) == 3:
                break

        for r in rows_with_values:
            for c in range(1, ws_sales.max_column + 1):
                ws_sales.cell(row=r, column=c).font = bold

                # =================================================
                # THICK BORDER BELOW "Total Sales" (COLUMNS A & B)
                # =================================================

                thick = Side(style="medium")

        for r in range(1, ws_sales.max_row + 1):
                    val = ws_sales.cell(row=r, column=1).value
                    if val and str(val).strip().lower() == "total sales":
                        ws_sales.cell(row=r, column=1).border = Border(bottom=thick)
                        ws_sales.cell(row=r, column=2).border = Border(bottom=thick)
                        break

        # =================================================
        # THICK BOX FOR BANK / VERIFY SECTION (ROBUST)
        # =================================================

        thick = Side(style="medium")

        target_headers = {
            "bank statement amount": None,
            "balance from amount to be received and bank amount": None,
            "capital gains": None,
            "post settlement total from report": None,
            "difference between bank and post settlement": None,
            "verify": None,
        }

        header_row = None

        # 1️⃣ Find the header row
        for r in range(1, ws_sales.max_row + 1):
            for c in range(1, ws_sales.max_column + 1):
                val = ws_sales.cell(row=r, column=c).value
                if val and str(val).strip().lower() in target_headers:
                    header_row = r
                    break
            if header_row:
                break

        # 2️⃣ Find exact columns for each header (including Verify)
        if header_row:
            for c in range(1, ws_sales.max_column + 1):
                val = ws_sales.cell(row=header_row, column=c).value
                if val:
                    key = str(val).strip().lower()
                    if key in target_headers:
                        target_headers[key] = c

            # Remove any not found (safety)
            found_cols = [c for c in target_headers.values() if c is not None]

            if found_cols:
                start_col = min(found_cols)
                end_col = max(found_cols)  # ✅ Verify INCLUDED
                value_row = header_row

                # 3️⃣ Apply thick borders (inside + outside)
                for r in (header_row, value_row):
                    for c in range(start_col, end_col + 1):
                        ws_sales.cell(row=r, column=c).border = Border(
                            top=thick if r == header_row else None,
                            bottom=thick if r == value_row else None,
                            left=thick if c == start_col else thick,
                            right=thick if c == end_col else thick,
                        )

        # =================================================
        # TOP + DOUBLE BOTTOM BORDER FOR "TOTAL AMOUNT TO BE RECEIVED"
        # (COLUMNS A & B)
        # =================================================

        top_side = Side(style="medium")
        double_bottom = Side(style="double")

        for r in range(1, ws_sales.max_row + 1):
            val = ws_sales.cell(row=r, column=1).value
            if val and str(val).strip().lower() == "total amount to be received":
                ws_sales.cell(row=r, column=1).border = Border(
                    top=top_side,
                    bottom=double_bottom
                )
                ws_sales.cell(row=r, column=2).border = Border(
                    top=top_side,
                    bottom=double_bottom
                )
                break
        # =================================================
        # THIN BORDER FOR VALUE ROW BELOW BANK / VERIFY HEADERS
        # =================================================

        thin = Side(style="thin")

        target_headers = {
            "bank statement amount": None,
            "balance from amount to be received and bank amount": None,
            "capital gains": None,
            "post settlement total from report": None,
            "difference between bank and post settlement": None,
            "verify": None,
        }

        header_row = None

        # 1️⃣ Find header row
        for r in range(1, ws_sales.max_row + 1):
            for c in range(1, ws_sales.max_column + 1):
                val = ws_sales.cell(row=r, column=c).value
                if val and str(val).strip().lower() in target_headers:
                    header_row = r
                    break
            if header_row:
                break

        # 2️⃣ Find columns by header names
        if header_row:
            for c in range(1, ws_sales.max_column):
                val = ws_sales.cell(row=header_row, column=c).value
                if val:
                    key = str(val).strip().lower()
                    if key in target_headers:
                        target_headers[key] = c

            found_cols = [c for c in target_headers.values() if c is not None]

            if found_cols:
                start_col = min(found_cols)
                end_col = max(found_cols)
                value_row = header_row + 1

                # 3️⃣ Apply THIN borders to value row (inside + outside)
                for c in range(start_col, end_col + 1):
                    ws_sales.cell(row=value_row, column=c).border = Border(
                        top=thin,
                        bottom=thin,
                        left=thin if c == start_col else thin,
                        right=thin if c == end_col else thin,
                    )
        # =================================================
        # WRAP + CENTER ROW BELOW "SEC Levy"
        # =================================================

        sec_levy_row = None

        # Find "SEC Levy" in column A
        for r in range(1, ws_sales.max_row + 1):
            val = ws_sales.cell(row=r, column=1).value
            if val and str(val).strip().lower() == "sec levy":
                sec_levy_row = r
                break

        # Apply formatting to the row BELOW it
        if sec_levy_row and sec_levy_row + 1 <= ws_sales.max_row:
            target_row = sec_levy_row + 1

            for c in range(1, ws_sales.max_column + 1):
                cell = ws_sales.cell(row=target_row, column=c)
                cell.alignment = Alignment(
                    wrap_text=True,
                    horizontal="general",
                    vertical="bottom"
                )


        # ---------- PURCHASES ----------
        ws_purchases = writer.sheets["Purchases"]

        # 5️⃣ Bold HEADER ROW
        for cell in ws_purchases[1]:
            cell.font = bold

        # 6️⃣ Bold LAST 3 ROWS except "sharestock"
        last_row = ws_purchases.max_row
        for r in range(max(1, last_row - 2), last_row + 1):
            for c in range(1, ws_purchases.max_column + 1):
                cell = ws_purchases.cell(row=r, column=c)
                val = str(cell.value).lower() if cell.value else ""
                if "sharestock" not in val:
                    cell.font = bold

        # =========================
        # FONT: TIMES NEW ROMAN 12
        # =========================
        for ws in writer.book.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    existing = cell.font or Font()
                    cell.font = Font(
                        name="Times New Roman",
                        size=12,
                        bold=existing.bold,
                        italic=existing.italic,
                        underline=existing.underline,
                        strike=existing.strike,
                        color=existing.color,
                    )

        # =========================
        # AUTO COLUMN WIDTHS
        # =========================

                MAX_WIDTH = 28  # hard cap so columns never explode
                MIN_WIDTH = 8  # prevent overly narrow columns
                PADDING = 2  # small padding (not 3–5)

                for ws in writer.book.worksheets:
                    for col in range(1, ws.max_column + 1):
                        max_len = 0
                        col_letter = get_column_letter(col)

                        for cell in ws[col_letter]:
                            if cell.value is not None:
                                cell_len = len(str(cell.value))
                                max_len = max(max_len, cell_len)

                        adjusted_width = min(MAX_WIDTH, max(MIN_WIDTH, max_len + PADDING))
                        ws.column_dimensions[col_letter].width = adjusted_width

                ws.column_dimensions[col_letter].width = max_len + 3

        writer.book.active = 0  # ✅ ensure visible sheet

    # ✅ writer closed AFTER formatting

    output.seek(0)

    st.download_button(
        label="Download Settlement Excel Workbook",
        data=output,
        file_name=f"ZSE_RECEIPTING_{post_settlement_total:.2f}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

