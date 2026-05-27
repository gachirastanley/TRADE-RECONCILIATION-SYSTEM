import pandas as pd
import streamlit as st
import camelot
import tempfile
import os

import uuid

from io import BytesIO
from datetime import datetime

from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from utils import load_history, save_history, HISTORY_FOLDER


# =====================================================
# HELPERS
# =====================================================
def safe_find_col(df, candidates):
    for c in df.columns:
        if any(c.lower().strip() == cand.lower() for cand in candidates):
            return c
    return None


def get_col_keyword(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k.lower() in col.lower():
                return col
    return None


def ensure_numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


def extract_levy(df, type_col, label, keyword):
    row = df[df[type_col] == label]
    if row.empty:
        return 0.0
    for col in df.columns:
        if keyword.lower() in str(col).lower():
            val = pd.to_numeric(row.iloc[0][col], errors="coerce")
            return 0.0 if pd.isna(val) else float(val)
    return 0.0


def extract_levies(df, type_col, label, keywords):
    if df is None or type_col not in df.columns:
        return 0.0
    row = df[df[type_col] == label]
    if row.empty:
        return 0.0
    total = 0.0
    for col in df.columns:
        if any(k.lower() in str(col).lower() for k in keywords):
            val = pd.to_numeric(row.iloc[0][col], errors="coerce")
            if not pd.isna(val):
                total += float(val)
    return round(total, 2)


# =====================================================
# CDC RECEIPTING UI
# =====================================================
def cdc_receipting_ui():
    st.subheader("Start CDC Reconciliation")

    # ---------- Step 1: Upload & Extract PDF ----------
    pdf = st.file_uploader("Upload Consolidated Trades PDF", type=["pdf"], key="cdc")

    if pdf and st.button("Extract & Sort CDC"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf.read())
            path = tmp.name

        try:
            tables = camelot.read_pdf(path, pages="all", flavor="stream", edge_tol=500, strip_text="\n")
            if tables.n == 0:
                st.error("No tables found in PDF.")
                return

            df = pd.concat([t.df for t in tables], ignore_index=True)

            # Clean
            df = df.iloc[3:, 1:]
            df.columns = df.iloc[0].astype(str).str.strip()
            df = df.iloc[1:]

            if "T+2" in df.columns:
                df.drop(columns=["T+2"], inplace=True)

            if "ZSE Levy" in df.columns:
                df["ZSE Levy"] = df["ZSE Levy"].shift(-1)

            first_col = df.columns[0]
            second_col = df.columns[1]
            df[first_col] = df[first_col].astype(str).str.strip()
            df[second_col] = df[second_col].astype(str).str.strip()

            df = df[
                (df[first_col] != "") &
                (df[second_col] != "") &
                (~df[first_col].str.upper().eq("TRADE DATE"))
            ].reset_index(drop=True)
            # ✅ REMOVE ROWS WHERE INVESTOR NAME HAS NUMBERS
            investor_col = safe_find_col(df, ["investor name", "client", "counter"])

            if investor_col:
                df = df[
                    ~df[investor_col]
                    .astype(str)
                    .str.contains(r"\d", regex=True)  # removes rows with ANY digit
                ].reset_index(drop=True)
            else:
                st.warning("Investor name column not found for numeric filtering.")
            # Numeric columns from col 5 onward
            for col in df.columns[5:]:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(",", "", regex=False).str.strip(),
                    errors="coerce"
                ).fillna(0)

            # Classify BUY / SELL
            stamp_col = safe_find_col(df, ["stamp duty"])
            qty_col = safe_find_col(df, ["quantity"])
            if not stamp_col or not qty_col:
                st.error("Required columns (Stamp Duty / Quantity) missing.")
                return

            qty_idx = df.columns.get_loc(qty_col)
            bs_col = "Buy/Sell"
            df.insert(qty_idx + 1, bs_col, "")
            df.loc[df[stamp_col] > 0, bs_col] = "BUY"
            df.loc[df[stamp_col] == 0, bs_col] = "SELL"

            buy_df = df[df[stamp_col] > 0].copy()
            sell_df = df[df[stamp_col] == 0].copy()

            def make_total_row(data, label):
                row = {c: (data[c].sum() if c != bs_col and df.columns.get_loc(c) >= qty_idx else "") for c in data.columns}
                row[first_col] = label
                return pd.DataFrame([row])

            cdc_df = pd.concat([
                pd.concat([buy_df, make_total_row(buy_df, "BUY TOTAL")], ignore_index=True),
                pd.concat([sell_df, make_total_row(sell_df, "SELL TOTAL")], ignore_index=True),
            ], ignore_index=True)

            # Persist to session state and rerun so display is clean
            st.session_state.cdc_df = cdc_df
            st.session_state.cdc_sorted = True
            st.session_state.cdc_matched = False
            st.session_state.show_final_summary = False
            st.rerun()

        except Exception as e:
            st.error(f"Error extracting tables: {e}")
        finally:
            if os.path.exists(path):
                os.remove(path)

    # ---------- Step 2: Show CDC table ----------
    if not st.session_state.get("cdc_sorted", False):
        return

    st.dataframe(st.session_state.cdc_df, use_container_width=True)
    st.divider()

    # ---------- Step 3: Upload Sharestock ----------
    st.subheader("Upload Sharestock File")
    sh_file = st.file_uploader("Upload Sharestock Excel File", type=["xlsx", "xls"], key="sh")

    if sh_file:
        raw = pd.read_excel(sh_file, header=None, engine="openpyxl")

        header_idx = None
        for i, v in raw.iloc[:, 0].items():
            if isinstance(v, str) and v.strip().upper() == "CLIENT":
                header_idx = i
                break

        if header_idx is None:
            st.error("Could not find a 'CLIENT' header row.")
            return

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

        sh_df = raw.loc[header_idx + 1:].copy()
        sh_df.columns = clean_headers
        sh_df = sh_df.loc[:, ~sh_df.columns.str.startswith("UNNAMED")]
        sh_df = sh_df.dropna(how="all").reset_index(drop=True)

        qty_col = safe_find_col(sh_df, ["qty"])
        if qty_col is None:
            st.error("Qty column not found in Sharestock file.")
            return

        qty_idx = sh_df.columns.get_loc(qty_col)
        for col in sh_df.columns[qty_idx:]:
            sh_df[col] = ensure_numeric(
                sh_df[col].astype(str).str.replace(",", "", regex=False)
            )

        st.session_state.sh_df = sh_df
        st.subheader("Sharestock Data")
        st.dataframe(sh_df, use_container_width=True)

    # ---------- Step 4: Match button ----------
    if st.session_state.get("sh_df") is not None and not st.session_state.get("cdc_matched", False):
        if st.button("Match CDC to Sharestock"):
            _run_cdc_match()   # do all the work, save to session state, rerun

    # ---------- Step 5: Show reconciled tables (persisted) ----------
    if st.session_state.get("cdc_matched", False):
        st.success("CDC ↔ Sharestock matching, totals, and variance complete.")
        st.subheader("Sharestock – Purchases (with CDC + Variance)")
        st.dataframe(st.session_state.purchase_df, use_container_width=True)
        st.subheader("Sharestock – Sales (with CDC + Variance)")
        st.dataframe(st.session_state.sale_df, use_container_width=True)

        if not st.session_state.get("show_final_summary", False):
            if st.button("Final Settlement Summary"):
                st.session_state.show_final_summary = True
                st.rerun()

    # ---------- Step 6: Final settlement summary (persisted) ----------
    if st.session_state.get("show_final_summary", False):
        _show_final_summary()


# =====================================================
# MATCH LOGIC  (called once, results saved to state)
# =====================================================
def take_before_comma(val):
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip()

    if "," in val_str:
        val_str = val_str.split(",")[0]

    return pd.to_numeric(val_str.replace(",", ""), errors="coerce")


def _run_cdc_match():
    cdc = st.session_state.cdc_df.copy()
    sh = st.session_state.sh_df.copy()

    cdc_type = safe_find_col(cdc, ["buy/sell", "type"])
    cdc_qty = safe_find_col(cdc, ["quantity", "qty"])
    sh_type = safe_find_col(sh, ["buy/sell", "type"])
    sh_qty = safe_find_col(sh, ["quantity", "qty"])
    cdc_sec = safe_find_col(cdc, ["counter"])
    sh_sec = safe_find_col(sh, ["security"])

    # ✅ NEW: deal value columns
    cdc_deal = get_col_keyword(cdc, ["deal value"])
    sh_deal = get_col_keyword(sh, ["gross proceeds"])

    if None in (cdc_type, cdc_qty, sh_type, sh_qty, cdc_sec, sh_sec):
        st.error("Could not detect required columns (type / qty / security).")
        return

    # ----------------------------
    # ✅ NORMALISE
    # ----------------------------
    cdc[cdc_qty] = pd.to_numeric(cdc[cdc_qty], errors="coerce").fillna(0)
    sh[sh_qty] = pd.to_numeric(sh[sh_qty], errors="coerce").fillna(0)

    # ✅ CLEAN deal values (FIRST BEFORE COMMA)
    if cdc_deal:
        cdc[cdc_deal] = cdc[cdc_deal].apply(take_before_comma).fillna(0)

    if sh_deal:
        sh[sh_deal] = sh[sh_deal].apply(take_before_comma).fillna(0)

    cdc[cdc_type] = cdc[cdc_type].astype(str).str.upper()
    sh[sh_type] = sh[sh_type].astype(str).str.upper().replace({
        "BUY": "PURCHASE",
        "SELL": "SALE"
    })

    cdc["_MATCH_"] = cdc[cdc_type].map({
        "BUY": "PURCHASE",
        "SELL": "SALE"
    })

    first_col = cdc.columns[0]

    # ----------------------------
    # ✅ SAVE TOTAL ROWS
    # ----------------------------
    cdc_buy_total = cdc[cdc[first_col].astype(str).str.contains("BUY TOTAL", na=False)]
    cdc_sell_total = cdc[cdc[first_col].astype(str).str.contains("SELL TOTAL", na=False)]

    cdc_clean = cdc[~cdc[first_col].astype(str).str.contains("TOTAL", na=False)]

    # ----------------------------
    # ✅ BUILD MATCH KEY (NOW 4 FIELDS)
    # ----------------------------
    keys = set(zip(
        cdc_clean[cdc_sec].astype(str).str.upper().str[:1],
        cdc_clean[cdc_qty],
        cdc_clean["_MATCH_"],
        cdc_clean[cdc_deal] if cdc_deal else [0]*len(cdc_clean)
    ))

    # ----------------------------
    # ✅ MATCH SHARESTOCK
    # ----------------------------
    matched = sh[sh.apply(
        lambda r: (
            str(r[sh_sec]).upper()[:1],
            r[sh_qty],
            r[sh_type],
            r[sh_deal] if sh_deal else 0
        ) in keys,
        axis=1,
    )]

    # ----------------------------
    # ✅ TOTALS
    # ----------------------------
    def add_total(df, label):
        total = {
            col: (df[col].sum() if pd.api.types.is_numeric_dtype(df[col]) else "")
            for col in df.columns
        }
        total[sh_type] = label
        return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

    purchase_df = add_total(matched[matched[sh_type] == "PURCHASE"].copy(), "PURCHASE TOTAL")
    sale_df = add_total(matched[matched[sh_type] == "SALE"].copy(), "SALE TOTAL")

    # ----------------------------
    # ✅ COLUMN MAP
    # ----------------------------
    column_map = {
        "quantity":       ["qty"],
        "deal value":     ["gross proceeds"],
        "commission":     ["brokerage"],
        "vat":            ["vat"],
        "secz levy":      ["sec levy"],
        "csd levy":       ["csd levy"],
        "zse levy":       ["zse levy"],
        "capital gains":  ["cgt"],
        "ipl":            ["investor levy"],
        "stamp duty":     ["stamp duty"],
    }

    def build_cdc_row(source_df, total_row, label):
        row = {col: "" for col in source_df.columns}
        row[sh_type] = label

        for cdc_key, sh_keys in column_map.items():
            cdc_col = get_col_keyword(cdc, [cdc_key])
            sh_col = get_col_keyword(source_df, sh_keys)

            if cdc_col and sh_col:
                val = pd.to_numeric(total_row[cdc_col], errors="coerce")
                row[sh_col] = 0.0 if pd.isna(val).all() else float(val.fillna(0).sum())

        return pd.DataFrame([row])

    def insert_after(df, after_label, new_row):
        idx_list = df.index[df[sh_type] == after_label].tolist()
        if idx_list:
            idx = idx_list[0] + 1
            return pd.concat([df.iloc[:idx], new_row, df.iloc[idx:]], ignore_index=True)
        return pd.concat([df, new_row], ignore_index=True)

    # ----------------------------
    # ✅ INSERT CDC TOTALS
    # ----------------------------
    if not cdc_buy_total.empty:
        purchase_df = insert_after(
            purchase_df,
            "PURCHASE TOTAL",
            build_cdc_row(purchase_df, cdc_buy_total.iloc[0:1], "CDC PURCHASE TOTAL")
        )

    if not cdc_sell_total.empty:
        sale_df = insert_after(
            sale_df,
            "SALE TOTAL",
            build_cdc_row(sale_df, cdc_sell_total.iloc[0:1], "CDC SALE TOTAL")
        )

    # ----------------------------
    # ✅ VARIANCE
    # ----------------------------
    def add_variance(df, total_label, cdc_label):
        t_rows = df[df[sh_type] == total_label]
        c_rows = df[df[sh_type] == cdc_label]

        if t_rows.empty or c_rows.empty:
            return df

        t = t_rows.iloc[0]
        c = c_rows.iloc[0]

        row = {}
        for col in df.columns:
            if col == sh_type:
                row[col] = "VARIANCE"
            elif pd.api.types.is_numeric_dtype(df[col]):
                tv = pd.to_numeric(t[col], errors="coerce")
                cv = pd.to_numeric(c[col], errors="coerce")
                row[col] = (0 if pd.isna(tv) else tv) - (0 if pd.isna(cv) else cv)
            else:
                row[col] = ""

        return insert_after(df, cdc_label, pd.DataFrame([row]))

    purchase_df = add_variance(purchase_df, "PURCHASE TOTAL", "CDC PURCHASE TOTAL")
    sale_df = add_variance(sale_df, "SALE TOTAL", "CDC SALE TOTAL")

    # ----------------------------
    # ✅ SAVE
    # ----------------------------
    st.session_state.purchase_df = purchase_df
    st.session_state.sale_df = sale_df
    st.session_state.cdc_matched = True

    st.rerun()


# =====================================================
# FINAL SETTLEMENT SUMMARY
# =====================================================
def _show_final_summary():
    st.divider()
    st.subheader("Final Settlement Summary")

    csd_reallocation = st.number_input(
        "CSD Reallocation", min_value=0.0, value=0.0, step=0.01, format="%.2f"
    )

    bank_amount = st.number_input(
        "Bank Statement Amount", min_value=0.0, value=0.0, step=0.01, format="%.2f"
    )

    levy_kw = ["broker", "stamp", "zse", "cgt", "sec", "investor", "vat"]

    p = st.session_state.get("purchase_df")
    s = st.session_state.get("sale_df")

    p_type = safe_find_col(p, ["type"]) if p is not None else None
    s_type = safe_find_col(s, ["type"]) if s is not None else None

    total_purchases = extract_levies(p, p_type, "PURCHASE TOTAL", levy_kw) if p_type else 0.0
    total_sales = extract_levies(s, s_type, "SALE TOTAL", levy_kw) if s_type else 0.0

  
    )
    ipl_levy = (
        (extract_levy(p, p_type, "PURCHASE TOTAL", "investor") if p_type else 0.0) +
        (extract_levy(s, s_type, "SALE TOTAL", "investor") if s_type else 0.0)
    )
    sec_levy = (
        (extract_levy(p, p_type, "PURCHASE TOTAL", "sec levy") if p_type else 0.0) +
        (extract_levy(s, s_type, "SALE TOTAL", "sec levy") if s_type else 0.0)
    )

    total_p_and_s = round(total_purchases + total_sales, 2)
    total_to_receive = round(total_p_and_s - ipl_levy - sec_levy, 2)
    balance_from_bank = round( total_to_receive-bank_amount- bank_amount , 2)

    # Post-settlement total from CDC BUY/SELL TOTAL rows
    post_settlement_total = 0.0
    cdc_df = st.session_state.get("cdc_df")
    if cdc_df is not None:
        fc = cdc_df.columns[0]
        total_col = next((c for c in cdc_df.columns if "total" in str(c).lower()), None)
        buy_row = cdc_df[cdc_df[fc].astype(str).str.contains("BUY TOTAL", case=False, na=False)]
        sell_row = cdc_df[cdc_df[fc].astype(str).str.contains("SELL TOTAL", case=False, na=False)]
        if total_col and not buy_row.empty and not sell_row.empty:
            b = pd.to_numeric(buy_row.iloc[0][total_col], errors="coerce")
            sv = pd.to_numeric(sell_row.iloc[0][total_col], errors="coerce")
            post_settlement_total = round(
                (0 if pd.isna(b) else float(b)) +
                (0 if pd.isna(sv) else float(sv)) -
                ipl_levy - sec_levy,
                2,
            )



    # Capital gains variance from VARIANCE row
    cgt_variance = 0.0
    if s is not None and s_type:
        vr = s[s[s_type] == "VARIANCE"]
        if not vr.empty:
            cgt_col = next((c for c in vr.columns if "cgt" in c.lower()), None)
            if cgt_col:
                val = pd.to_numeric(vr.iloc[0][cgt_col], errors="coerce")
                cgt_variance = 0.0 if pd.isna(val) else round(float(val), 2)

    verify_value = round(bank_amount + cgt_variance - post_settlement_total, 2)
    diff_bank_post = round(balance_from_bank - cgt_variance - csd_reallocation, 2)

    # Build display table
    summary_df = pd.DataFrame([
        ["Total Purchases",              total_purchases,    "", "", "", "", "", "", ""],
        ["Total Sales",                  total_sales,        "", "", "", "", "", "", ""],
        ["Total Purchases and Sales",    total_p_and_s,      "", "", "", "", "", "", ""],
        ["",                             "",                 "", "", "", "", "", "", ""],
      
        ["IPL Levy Remitted Directly",   ipl_levy,           "", "", "", "", "", "", ""],
        ["SEC Levy",                     sec_levy,           "", "", "", "", "", "", ""],
        [
            "", "", "",
            "Bank Statement Amount",
            "Balance from Amount to be Received and Bank Amount",
            "Capital Gains",
            "CSD Reallocation ",
            "Balance Bank Charges ",
            "Post Settlement Total from Report",
            "Verify",
            "", "",
        ],
        [
            "Total Amount to be Received",
            total_to_receive,
            "",
            bank_amount,
            balance_from_bank,
            cgt_variance,
            csd_reallocation,
            diff_bank_post,
            post_settlement_total,
            verify_value,
            "",
        ],
        ["", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", ""],

        ["Journals to be Posted in Sybrin", "", "", ""],
        ["LEVIES REMITTED DIRECTLY BY DEPOSITORY", "", "", ""],
        ["Narration", "", "DR", "CR"],
        ["CDC Bank charges", "BANK CHARGES", diff_bank_post, "", "", "", "", "", ""],
        ["", "FBC BANK", "",diff_bank_post],
        ["", "", "", ""],
        ["Total", "",diff_bank_post,diff_bank_post],
        ["", "", "", ""],



        ["Prepared by:....................................................................................................", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", ""],
        ["Reviewed by:....................................................................................................", "", "", "", "", "", "", "", ""],
    ])

    st.dataframe(summary_df, use_container_width=True)

    # Excel export
    _export_excel(summary_df, post_settlement_total)


# =====================================================
# EXCEL EXPORT
# =====================================================
def _export_excel(summary_df, post_settlement_total):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        st.session_state.cdc_df.to_excel(writer, sheet_name="CDC Raw", index=False)
        st.session_state.sh_df.to_excel(writer, sheet_name="Sharestock Raw", index=False)
        st.session_state.purchase_df.to_excel(writer, sheet_name="Purchases", index=False)

        SALES_SHEET = "Sales & Summary"
        st.session_state.sale_df.to_excel(writer, sheet_name=SALES_SHEET, index=False, startrow=0)
        summary_df.to_excel(
            writer, sheet_name=SALES_SHEET,
            index=False, header=False,
            startrow=len(st.session_state.sale_df) + 3,
        )

        bold = Font(bold=True)
        thick = Side(style="medium")
        thin = Side(style="thin")

        # ---- Sales & Summary sheet ----
        ws = writer.sheets[SALES_SHEET]

        for cell in ws[1]:
            cell.font = bold

        # Bold all rows in the summary section (from first blank row in col A downward)
        for r in range(2, ws.max_row + 1):
            val = ws.cell(row=r, column=1).value
            if val is None or str(val).strip() == "":
                for rr in range(r, ws.max_row + 1):
                    cell = ws.cell(row=rr, column=1)
                    cell.value = cell.value or ""
                    cell.font = bold
                break

        # Bold last 3 rows that have values in col 17
        rows_hit = []
        for r in range(ws.max_row, 1, -1):
            if ws.cell(row=r, column=17).value not in (None, "", " "):
                rows_hit.append(r)
            if len(rows_hit) == 3:
                break
        for r in rows_hit:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).font = bold

        # Thick border under "Total Sales"
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value or "").strip().lower() == "total sales":
                ws.cell(row=r, column=1).border = Border(bottom=thick)
                ws.cell(row=r, column=2).border = Border(bottom=thick)
                break

        # Top + double bottom for "Total Amount to be Received"
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value or "").strip().lower() == "total amount to be received":
                ws.cell(row=r, column=1).border = Border(top=thick, bottom=Side(style="double"))
                ws.cell(row=r, column=2).border = Border(top=thick, bottom=Side(style="double"))
                break
        # =================================================
        # ✅ FORMAT JOURNAL BLOCK (FULL BORDERS 4x4)
        # =================================================

        thin = Side(style="thin")
        bold = Font(bold=True)

        # ✅ Locate journal section
        journal_start = None

        for r in range(1, ws.max_row + 1):
            val = str(ws.cell(row=r, column=1).value or "").strip().lower()
            if "journals to be posted in sybrin" in val:
                journal_start = r
                break

        if journal_start:

            header_row = journal_start + 2  # "Narration"
            data_start = header_row + 1
            total_row = data_start + 3  # based on your structure

            # ✅ Apply borders to 4x4 table
            for r in range(header_row, total_row + 1):
                for c in range(1, 5):  # ONLY 4 columns
                    cell = ws.cell(row=r, column=c)
                    cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

            # ✅ Bold header row
            for c in range(1, 5):
                ws.cell(row=header_row, column=c).font = bold

            # ✅ Bold total row
            for c in range(1, 5):
                ws.cell(row=total_row, column=c).font = bold

            # ✅ Align DR / CR right
            for r in range(header_row, total_row + 1):
                ws.cell(row=r, column=3).alignment = Alignment(horizontal="right")
                ws.cell(row=r, column=4).alignment = Alignment(horizontal="right")
        # Thick box around bank/verify header + value rows
        bank_headers = {
            "bank statement amount",
            "balance from amount to be received and bank amount",
            "capital gains",
            "post settlement total from report",
            "Balance Bank Charges",
            "CSD Reallocation",
            "verify",
        }
        header_row, found_cols = None, []
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                if str(ws.cell(row=r, column=c).value or "").strip().lower() in bank_headers:
                    header_row = r
                    break
            if header_row:
                break
        if header_row:
            for c in range(1, ws.max_column + 1):
                if str(ws.cell(row=header_row, column=c).value or "").strip().lower() in bank_headers:
                    found_cols.append(c)
            if found_cols:
                sc, ec = min(found_cols), max(found_cols)
                vr = header_row + 1
                for row_r in (header_row, vr):
                    for c in range(sc, ec + 1):
                        ws.cell(row=row_r, column=c).border = Border(
                            top=thick if row_r == header_row else None,
                            bottom=thick if row_r == vr else None,
                            left=thick, right=thick,
                        )
                for c in range(sc, ec + 1):
                    ws.cell(row=vr, column=c).border = Border(top=thin, bottom=thin, left=thin, right=thin)

        # Wrap row under "SEC Levy"
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value or "").strip().lower() == "sec levy":
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r + 1, column=c).alignment = Alignment(wrap_text=True, vertical="bottom")
                break

        # ---- Purchases sheet ----
        wp = writer.sheets["Purchases"]
        for cell in wp[1]:
            cell.font = bold
        for r in range(max(1, wp.max_row - 2), wp.max_row + 1):
            for c in range(1, wp.max_column + 1):
                cell = wp.cell(row=r, column=c)
                if "sharestock" not in str(cell.value or "").lower():
                    cell.font = bold

        # ---- Global: Times New Roman 12, number format, auto widths ----
        num_fmt = '#,##0.00;-#,##0.00;-'
        for wsh in writer.book.worksheets:
            for row in wsh.iter_rows():
                for cell in row:
                    existing = cell.font or Font()
                    cell.font = Font(
                        name="Times New Roman", size=12,
                        bold=existing.bold, italic=existing.italic,
                    )
                    if isinstance(cell.value, str):
                        try:
                            cell.value = float(cell.value.replace(",", ""))
                        except Exception:
                            pass
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = num_fmt
                        cell.alignment = Alignment(horizontal="right", vertical="center")

            for col_idx in range(1, wsh.max_column + 1):
                col_letter = get_column_letter(col_idx)
                max_len = max(
                    (len(str(cell.value)) for cell in wsh[col_letter] if cell.value is not None),
                    default=0,
                )
                wsh.column_dimensions[col_letter].width = min(28, max(8, max_len + 2))

        writer.book.active = 0

    output.seek(0)
    file_name = f"CDC_RECEIPTING_{post_settlement_total:.2f}.xlsx"
    file_path = os.path.join(HISTORY_FOLDER, file_name)

    if st.download_button(
        label="📥 Save & Download CDC Report",
        data=output,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ):
        with open(file_path, "wb") as f:
            f.write(output.getvalue())

        history = load_history()
        import uuid

        history.append({
            "id": str(uuid.uuid4()),  # ✅ UNIQUE FOREVER
            "file": file_name,
            "user": st.session_state.get("username", ""),
            "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "type": "CDC Receipting",
            "total": float(post_settlement_total),
        })
        save_history(history)
        st.success("CDC report saved to history.")
