import pandas as pd
import streamlit as st
import camelot
import tempfile
import os

# ---------- CUSTOM UI STYLING ----------
st.markdown(
    """
    <style>
    .stApp { background-color: #f5f7fb; color: #0f172a; font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    h1 { color: #0A2540; font-weight: 700; }
    h2, h3 { color: #102a43; font-weight: 600; }
    div.stButton > button { background: #1d4ed8; color: #ffffff !important; border-radius: 10px; padding: 0.55rem 1.4rem; font-weight: 600; border: none; }
    div.stButton > button:hover { background: #1e40af; transform: translateY(-1px); }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------- Helper functions ----------
def safe_find_col(df, candidates):
    for c in df.columns:
        if any(c.lower().strip() == cand.lower() for cand in candidates):
            return c
    return None

def ensure_numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

# ---------- CDC RECEIPTING INTERFACE ----------
def cdc_receipting_ui():
    st.subheader("CDC / Consolidated Trades PDF → Excel")

    # ---------- CDC UPLOAD ----------
    pdf = st.file_uploader(
        "Upload Consolidated Trades PDF",
        type=["pdf"],
        key="cdc"
    )

    if pdf and st.button("Extract & Sort CDC"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf.read())
            path = tmp.name

        try:
            tables = camelot.read_pdf(
                path,
                pages="all",
                flavor="stream",
                edge_tol=500,
                strip_text="\n"
            )

            if tables.n == 0:
                st.error("No tables found in PDF")
                return

            df = pd.concat([t.df for t in tables], ignore_index=True)

            # ---------- CLEAN ----------
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
                (df[first_col].str.upper() != "TRADE DATE")
            ].reset_index(drop=True)

            # ---------- NUMERIC ----------
            for col in df.columns[5:]:
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .str.strip()
                )
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            # ---------- BUY / SELL ----------
            stamp_col = safe_find_col(df, ["stamp duty"])
            qty_col = safe_find_col(df, ["quantity"])

            if not stamp_col or not qty_col:
                st.error("Required columns missing")
                return

            qty_idx = df.columns.get_loc(qty_col)

            bs_col = "Buy/Sell"
            df.insert(qty_idx + 1, bs_col, "")
            df.loc[df[stamp_col] > 0, bs_col] = "BUY"
            df.loc[df[stamp_col] == 0, bs_col] = "SELL"

            buy_df = df[df[stamp_col] > 0]
            sell_df = df[df[stamp_col] == 0]

            def totals(data, label):
                row = {}
                for i, c in enumerate(data.columns):
                    if c == bs_col or i < qty_idx:
                        row[c] = ""
                    else:
                        row[c] = data[c].sum()
                row[first_col] = label
                return pd.DataFrame([row])

            buy_df = pd.concat(
                [buy_df, totals(buy_df, "BUY TOTAL")],
                ignore_index=True
            )
            sell_df = pd.concat(
                [sell_df, totals(sell_df, "SELL TOTAL")],
                ignore_index=True
            )

            st.session_state.cdc_df = pd.concat(
                [buy_df, sell_df],
                ignore_index=True
            )
            st.session_state.cdc_sorted = True

            st.success("CDC data extracted, classified, and sorted successfully")

        except Exception as e:
            st.error(f"Error extracting tables: {e}")

        finally:
            os.remove(path)

    # ---------- CDC PREVIEW ----------
    if st.session_state.get("cdc_sorted", False):
        st.dataframe(st.session_state.cdc_df, use_container_width=True)

    # ---------- SHARESTOCK SECTION ----------
    if st.session_state.get("cdc_sorted", False):
        st.divider()
        st.subheader("Upload Sharestock File")

        sh_file = st.file_uploader(
            "Upload Sharestock Excel File",
            type=["xlsx", "xls"],
            key="sh"
        )

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
                st.error("Qty column not found")
                return

            qty_idx = sh_df.columns.get_loc(qty_col)
            for col in sh_df.columns[qty_idx:]:
                sh_df[col] = ensure_numeric(
                    sh_df[col]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                )

            st.session_state.sh_df = sh_df
            st.session_state.sh_ready = True

            st.subheader("Cleaned Sharestock Data")
            st.dataframe(sh_df, use_container_width=True)

        # =================================================
        # MATCH BUTTON (ONLY SETS STATE)
        # =================================================
    if st.button("Match CDC to Sharestock"):
        st.session_state.run_full_cdc_match = True

        # =================================================
        # FULL MATCH + TOTALS (CDC ↔ SHARESTOCK)
        # =================================================
    if st.session_state.get("run_full_cdc_match", False):

        # ---------- VALIDATION ----------
        if (
                st.session_state.get("cdc_df") is None
                or st.session_state.get("sh_df") is None
        ):
            st.error("CDC data or Sharestock data is missing.")
            st.stop()

        cdc = st.session_state.cdc_df.copy()
        sh = st.session_state.sh_df.copy()

        # ---------- COLUMN DETECTION ----------
        cdc_type = safe_find_col(cdc, ["buy/sell", "type"])
        cdc_qty = safe_find_col(cdc, ["quantity", "qty"])
        sh_type = safe_find_col(sh, ["type"])
        sh_qty = safe_find_col(sh, ["quantity", "qty"])

        if None in (cdc_type, cdc_qty, sh_type, sh_qty):
            st.error("Could not reliably detect required columns.")
            st.stop()

        # ---------- NORMALISE TYPES ----------
        sh[sh_type] = sh[sh_type].replace({"BUY": "PURCHASE", "SELL": "SALE"})
        cdc["_MATCH_"] = cdc[cdc_type].map({"BUY": "PURCHASE", "SELL": "SALE"})

        # ---------- REMOVE CDC TOTAL ROWS ----------
        cdc_clean = cdc[~cdc[cdc_type].isin(["BUY TOTAL", "SELL TOTAL"])].copy()

        # ---------- MATCH USING (QTY + TYPE) ----------
        keys = set(
            zip(
                cdc_clean[cdc_qty].astype(str),
                cdc_clean["_MATCH_"]
            )
        )

        matched = sh[
            sh.apply(
                lambda r: (
                              str(r[sh_qty]),
                              r[sh_type]
                          ) in keys,
                axis=1
            )
        ]

        # =================================================
        # SHARESTOCK TOTALS
        # =================================================
        def add_total_block(data, label):
            total = {}
            for col in data.columns:
                if pd.api.types.is_numeric_dtype(data[col]):
                    total[col] = data[col].sum()
                else:
                    total[col] = ""
            total[sh_type] = label
            return pd.concat([data, pd.DataFrame([total])], ignore_index=True)

        purchase_df = add_total_block(
            matched[matched[sh_type] == "PURCHASE"],
            "PURCHASE TOTAL"
        )

        sale_df = add_total_block(
            matched[matched[sh_type] == "SALE"],
            "SALE TOTAL"
        )

        # =================================================
        # CDC → SHARESTOCK COLUMN MAPPING
        # =================================================
        mapping = {
            "qty": ["quantity"],
            "gross proceeds": ["gross", "gross consideration"],
            "brokerage": ["broker", "commission"],
            "basic": ["basic"],
            "stamp": ["stamp duty", "stamp"],
            "zse levy": ["zse levy"],
            "csd levy": ["depository levy"],
            "cgt": ["capital gains tax", "cgt"],
            "sec levy": ["sec levy", "secz levy"],
            "investor levy": ["investor"],
            "vat": ["vat"],
        }

        def append_cdc(df, cdc_row, label):
            row = {}
            for sh_col in df.columns:
                row[sh_col] = ""
                if sh_col == sh_type:
                    row[sh_col] = label
                    continue

                for sh_key, cdc_keys in mapping.items():
                    if sh_key in sh_col.lower():
                        for cdc_col in cdc_row.index:
                            if any(k in cdc_col.lower() for k in cdc_keys):
                                row[sh_col] = cdc_row[cdc_col]

            return pd.concat([df, pd.DataFrame([row])], ignore_index=True)

        # ---------- CDC BUY / SELL TOTALS ----------
        cdc_buy_total = cdc[cdc[cdc_type] == "BUY TOTAL"]
        cdc_sell_total = cdc[cdc[cdc_type] == "SELL TOTAL"]

        if not cdc_buy_total.empty:
            purchase_df = append_cdc(
                purchase_df,
                cdc_buy_total.iloc[0],
                "CDC BUY TOTAL"
            )

        if not cdc_sell_total.empty:
            sale_df = append_cdc(
                sale_df,
                cdc_sell_total.iloc[0],
                "CDC SELL TOTAL"
            )

        # =================================================
        # STORE + DISPLAY
        # =================================================
        st.session_state.purchase_df = purchase_df
        st.session_state.sale_df = sale_df

        st.success("✅ CDC ↔ Sharestock matching and totals completed")

        st.subheader("Sharestock – Purchases (with CDC Totals)")
        st.dataframe(purchase_df, use_container_width=True)

        st.subheader("Sharestock – Sales (with CDC Totals)")
        st.dataframe(sale_df, use_container_width=True)