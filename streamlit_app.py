"""
Streamlit web app for the Private Portfolio Allocation System.

Run locally:
    pip install -r requirements.txt
    streamlit run streamlit_app.py

Deploy for free (see README.md for full steps):
    - Streamlit Community Cloud (streamlit.io/cloud) -- easiest, made for this
    - Hugging Face Spaces (huggingface.co/spaces) -- also free, pick "Streamlit" SDK
"""

import io
import os
from datetime import datetime

import streamlit as st

from portfolio_engine import run_allocation, MONTHS, TICKERS
from run_advisor_session import build_excel_report
from pdf_report import build_pdf_report

st.set_page_config(page_title="Portfolio Allocation System", page_icon="\U0001F4C8", layout="centered")

st.title("Private Portfolio Allocation System")
st.caption("G-sec + equity allocation using CRRA risk aversion, Markowitz optimization, and Merton two-fund separation.")

with st.sidebar:
    st.header("1. Daily market data")

    mode = st.radio(
        "How do you want to provide today's files?",
        ["Upload each time", "Read from a local folder path"],
        help="'Read from a local folder path' only works when you run this app on your own "
             "computer (streamlit run streamlit_app.py) with the Excel files saved on disk -- "
             "it will always pick up your latest saved edits automatically, no re-upload needed. "
             "On the hosted (cloud) version you must use 'Upload each time'.",
    )

    if mode == "Upload each time":
        portfolio_file = st.file_uploader("Portfolio.xlsx (prices, forecasts, T-bill rates)", type=["xlsx"])
        dataset_file = st.file_uploader("Full_Dataset_5_Stocks.xlsx (historical monthly data)", type=["xlsx"])
        if portfolio_file is not None:
            st.session_state["portfolio_bytes"] = portfolio_file.read()
        if dataset_file is not None:
            st.session_state["dataset_bytes"] = dataset_file.read()
        have_files = "portfolio_bytes" in st.session_state and "dataset_bytes" in st.session_state
    else:
        default_folder = st.session_state.get("local_folder", os.getcwd())
        folder = st.text_input("Folder containing both Excel files", value=default_folder)
        st.session_state["local_folder"] = folder
        portfolio_path = os.path.join(folder, "Portfolio.xlsx")
        dataset_path = os.path.join(folder, "Full_Dataset_5_Stocks.xlsx")
        have_files = os.path.isfile(portfolio_path) and os.path.isfile(dataset_path)
        if have_files:
            st.caption(f"Found both files in: {folder}")
        else:
            st.caption("Waiting for Portfolio.xlsx and Full_Dataset_5_Stocks.xlsx in that folder...")

    if have_files:
        st.success("Market data ready.")
    else:
        st.warning("Provide both files to run an allocation.")

    st.divider()
    long_only = st.checkbox("No short-selling (long-only weights)", value=True,
                             help="Keeps every stock weight at 0% or above. Uncheck to allow the "
                                  "unconstrained Markowitz solution, which can assign negative "
                                  "(short) weights to some stocks.")

st.header("2. Investor meeting inputs")
col1, col2 = st.columns(2)
with col1:
    investor_name = st.text_input("Investor name", value="Investor")
    tenor = st.selectbox("Expected tenor (month)", MONTHS, index=5)
with col2:
    amount = st.number_input("Investment amount (LKR)", min_value=0.0, value=1_000_000.0, step=50_000.0, format="%.2f")
    preference = st.multiselect("Preferred stocks (optional)", TICKERS)

run_clicked = st.button("Run allocation", type="primary", disabled=not have_files)

if run_clicked:
    if mode == "Upload each time":
        portfolio_src = io.BytesIO(st.session_state["portfolio_bytes"])
        dataset_src = io.BytesIO(st.session_state["dataset_bytes"])
    else:
        portfolio_src = portfolio_path
        dataset_src = dataset_path

    try:
        res = run_allocation(
            portfolio_xlsx=portfolio_src,
            dataset_xlsx=dataset_src,
            tenor_month=tenor,
            investment_amount=amount,
            preference_list=preference or None,
            long_only=long_only,
        )
    except Exception as e:
        st.error(f"Could not compute the allocation: {e}")
        st.stop()

    st.header("3. Result")
    st.info(res.preference_note)

    m1, m2, m3 = st.columns(3)
    m1.metric("Portfolio expected return", f"{res.portfolio_expected_return_pct:.2f}%")
    m2.metric("Portfolio volatility", f"{res.portfolio_volatility_pct:.2f}%")
    m3.metric("Expected value at horizon", f"Rs. {res.portfolio_expected_value:,.0f}")

    st.subheader("Weights")
    weight_rows = [{"Asset": "G-sec (T-bill)", "Weight": f"{res.gsec_weight*100:.2f}%",
                     "Amount (LKR)": f"{res.gsec_weight*amount:,.2f}"}]
    for t, w in res.equity_weights.items():
        weight_rows.append({"Asset": t, "Weight": f"{w*100:.2f}%", "Amount (LKR)": f"{w*amount:,.2f}"})
    st.table(weight_rows)

    with st.expander("Model diagnostics (gamma, tangency portfolio, gamble amounts)"):
        st.write(f"G-sec (T-bill) rate for this tenor: **{res.rf_pct:.2f}%**")
        st.write(f"Stocks used: **{', '.join(res.used_stocks) if res.used_stocks else 'None'}**")
        if res.used_stocks:
            st.write(f"Tangency portfolio expected return: **{res.mu_p_pct:.2f}%**")
            st.write(f"Tangency portfolio volatility: **{res.sigma_p_pct:.2f}%**")
            st.write(f"Guarantee amount (100% G-sec): **Rs. {res.guarantee_amount:,.2f}**")
            st.write(f"Risky upside amount (+1 std dev): **Rs. {res.upside_amount:,.2f}**")
            st.write(f"Risky downside amount (-1 std dev): **Rs. {res.downside_amount:,.2f}**")
            st.write(f"Implied risk-aversion gamma (CRRA): **{res.gamma:.4f}**")

    out_path = f"/tmp/Allocation_Report_{investor_name}_{tenor}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    build_excel_report(res, amount, investor_name, tenor, out_path)

    pdf_path = out_path.replace(".xlsx", ".pdf")
    build_pdf_report(res, amount, investor_name, tenor, pdf_path)

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        with open(out_path, "rb") as f:
            st.download_button("Download Excel report", data=f.read(),
                                file_name=out_path.split("/")[-1],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with dcol2:
        with open(pdf_path, "rb") as f:
            st.download_button("Download PDF report (with charts)", data=f.read(),
                                file_name=pdf_path.split("/")[-1],
                                mime="application/pdf")
