"""
app.py  ─  Streamlit UI for the Portfolio Journal Updater
Run locally  : streamlit run app.py
Deploy free  : push to GitHub → connect to share.streamlit.io
"""

import streamlit as st
from Excel_automate import process_streams

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Portfolio Journal Updater",
    page_icon="📈",
    layout="centered",
)

# ─── Header ───────────────────────────────────────────────────────────────────
st.title("📈 Portfolio Journal Updater")
st.markdown(
    """
    Upload your **Zerodha order book CSV** and your **portfolio journal XLSX**,
    and this tool will automatically sync all BUY / SELL trades into the journal.
    """
)
st.divider()

# ─── File uploaders ───────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣  Order Book")
    csv_file = st.file_uploader(
        "Upload Zerodha order book (.csv)",
        type=["csv"],
        help="Download from Zerodha Console → Reports → Tradebook",
    )

with col2:
    st.subheader("2️⃣  Portfolio Journal")
    xlsx_file = st.file_uploader(
        "Upload portfolio journal (.xlsx)",
        type=["xlsx"],
        help="Your Sample_pf.xlsx with the yellow-cell template blocks",
    )

st.divider()

# ─── Process button ───────────────────────────────────────────────────────────
if st.button("🚀  Update Journal", use_container_width=True, type="primary"):
    if not csv_file:
        st.error("Please upload the order book CSV.")
    elif not xlsx_file:
        st.error("Please upload the portfolio journal XLSX.")
    else:
        with st.spinner("Processing orders…"):
            try:
                updated_xlsx_bytes, log_lines = process_streams(
                    csv_file.read(),
                    xlsx_file.read(),
                )
                st.success("✅  Journal updated successfully!")

                # ── Log output ────────────────────────────────────────────
                with st.expander("📋  Processing log", expanded=True):
                    for line in log_lines:
                        # colour-code key lines
                        if line.startswith("  ✏️") or line.startswith("  ➕"):
                            st.markdown(f":green[{line}]")
                        elif line.startswith("  🔴"):
                            st.markdown(f":red[{line}]")
                        elif line.startswith("  ⏭"):
                            st.markdown(f":orange[{line}]")
                        elif line.startswith("  ⚠"):
                            st.markdown(f":orange[{line}]")
                        else:
                            st.text(line)

                # ── Download button ───────────────────────────────────────
                st.download_button(
                    label="⬇️  Download Updated Journal",
                    data=updated_xlsx_bytes,
                    file_name=xlsx_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"❌  Error during processing: {e}")
                st.exception(e)

# ─── How-to sidebar ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("ℹ️  How to use")
    st.markdown(
        """
        **Step 1 — Download order book from Zerodha**
        - Login → Console → Reports → Tradebook
        - Select date range → Download CSV

        **Step 2 — Upload both files**
        - Order book CSV (left panel)
        - Portfolio journal XLSX (right panel)

        **Step 3 — Click "Update Journal"**
        - BUY orders → fill next available tranche
        - SELL orders → mark position Closed + write exit date & price
        - Duplicate orders are automatically skipped

        **Step 4 — Download the updated journal**

        ---
        **Column mapping**

        | Order Book | Journal |
        |---|---|
        | symbol | Symbol (B) |
        | trade_date | Date (F) |
        | quality | Act. Qty (I) |
        | price | Act. Price (J) |

        ---
        **Notes**
        - Only cells with **yellow background** are written to
        - If all 5 tranches are full, a new overflow block is used
        - On SELL: Position → *Closed*, exit date & price saved
        """
    )
    st.divider()
    st.caption("Built with Streamlit · Free to use")
