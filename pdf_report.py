"""
PDF report generator for the Portfolio Allocation System.
Produces a one-page-ish PDF: summary numbers, a weights pie chart, and a
risk/return bar comparison, using matplotlib for charts and reportlab for layout.
"""

import io
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, Image)
from reportlab.lib.enums import TA_CENTER


def _weights_pie_chart(res):
    labels, sizes = [], []
    if res.gsec_weight > 0.001:
        labels.append("G-sec")
        sizes.append(res.gsec_weight)
    for t, w in res.equity_weights.items():
        if w > 0.001:
            labels.append(t)
            sizes.append(w)

    fig, ax = plt.subplots(figsize=(4.2, 4.2), dpi=150)
    colors_list = plt.cm.Set2.colors
    ax.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=90,
           colors=colors_list[:len(sizes)],
           wedgeprops={"edgecolor": "white", "linewidth": 1})
    ax.set_title("Portfolio Weights", fontsize=12, fontweight="bold")
    ax.axis("equal")
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def _risk_return_bar_chart(res):
    categories = ["G-sec\n(risk-free)"]
    returns = [res.rf_pct]
    vols = [0.0]
    if res.used_stocks:
        categories.append("Tangency\nequity mix")
        returns.append(res.mu_p_pct)
        vols.append(res.sigma_p_pct)
    categories.append("Final\nportfolio")
    returns.append(res.portfolio_expected_return_pct)
    vols.append(res.portfolio_volatility_pct)

    x = range(len(categories))
    fig, ax = plt.subplots(figsize=(4.6, 4.2), dpi=150)
    width = 0.35
    ax.bar([i - width/2 for i in x], returns, width, label="Expected return %", color="#4C9A2A")
    ax.bar([i + width/2 for i in x], vols, width, label="Volatility %", color="#D9822B")
    ax.set_xticks(list(x))
    ax.set_xticklabels(categories, fontsize=8)
    ax.set_ylabel("%")
    ax.set_title("Risk / Return Comparison", fontsize=12, fontweight="bold")
    ax.legend(fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf


def build_pdf_report(res, investment_amount, investor_name, tenor, out_path):
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleC", parent=styles["Title"], alignment=TA_CENTER)
    h2 = ParagraphStyle("H2", parent=styles["Heading2"], spaceBefore=10, spaceAfter=6)
    normal = styles["Normal"]

    doc = SimpleDocTemplate(out_path, pagesize=A4,
                             topMargin=1.5*cm, bottomMargin=1.5*cm,
                             leftMargin=1.8*cm, rightMargin=1.8*cm)
    story = []

    story.append(Paragraph("Private Portfolio Allocation Report", title_style))
    story.append(Spacer(1, 4))
    meta = (f"Investor: <b>{investor_name}</b> &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Tenor: <b>{tenor}</b> &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Amount: <b>Rs. {investment_amount:,.2f}</b> &nbsp;&nbsp;|&nbsp;&nbsp; "
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    story.append(Paragraph(meta, ParagraphStyle("Meta", parent=normal, alignment=TA_CENTER, fontSize=9)))
    story.append(Spacer(1, 14))

    # --- Key metrics table -------------------------------------------------
    story.append(Paragraph("Portfolio Summary", h2))
    metrics_data = [
        ["Expected return (horizon)", f"{res.portfolio_expected_return_pct:.2f}%"],
        ["Volatility (horizon)", f"{res.portfolio_volatility_pct:.2f}%"],
        ["Expected value at horizon end", f"Rs. {res.portfolio_expected_value:,.2f}"],
        ["G-sec (T-bill) rate used", f"{res.rf_pct:.2f}%"],
        ["Preference handling", res.preference_note],
    ]
    t = Table(metrics_data, colWidths=[8*cm, 8.5*cm])
    t.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F2F2")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
    ]))
    story.append(t)
    story.append(Spacer(1, 14))

    # --- Weights table -------------------------------------------------
    story.append(Paragraph("Final Allocation", h2))
    weight_data = [["Asset", "Weight", "Amount (LKR)"]]
    weight_data.append(["G-sec (Treasury bill)", f"{res.gsec_weight*100:.2f}%",
                         f"{res.gsec_weight*investment_amount:,.2f}"])
    for tkr, w in res.equity_weights.items():
        weight_data.append([tkr, f"{w*100:.2f}%", f"{w*investment_amount:,.2f}"])
    wt = Table(weight_data, colWidths=[6*cm, 5*cm, 5.5*cm])
    wt.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDEBF7")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))
    story.append(wt)
    story.append(Spacer(1, 14))

    # --- Charts side by side -------------------------------------------------
    story.append(Paragraph("Charts", h2))
    pie_buf = _weights_pie_chart(res)
    bar_buf = _risk_return_bar_chart(res)
    chart_table = Table(
        [[Image(pie_buf, width=7.5*cm, height=7.5*cm), Image(bar_buf, width=7.5*cm, height=7.5*cm)]],
        colWidths=[8*cm, 8*cm]
    )
    story.append(chart_table)
    story.append(Spacer(1, 14))

    # --- Diagnostics -------------------------------------------------
    if res.used_stocks:
        story.append(Paragraph("Model Diagnostics", h2))
        diag_data = [
            ["Stocks used", ", ".join(res.used_stocks)],
            ["Tangency portfolio expected return", f"{res.mu_p_pct:.2f}%"],
            ["Tangency portfolio volatility", f"{res.sigma_p_pct:.2f}%"],
            ["Guarantee amount (100% G-sec)", f"Rs. {res.guarantee_amount:,.2f}"],
            ["Risky upside amount (+1 std dev)", f"Rs. {res.upside_amount:,.2f}"],
            ["Risky downside amount (-1 std dev)", f"Rs. {res.downside_amount:,.2f}"],
            ["Implied risk-aversion (gamma, CRRA)", f"{res.gamma:.4f}"],
        ]
        dt = Table(diag_data, colWidths=[8*cm, 8.5*cm])
        dt.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#EEEEEE")),
        ]))
        story.append(dt)

    story.append(Spacer(1, 16))
    story.append(Paragraph(
        "This report is generated by an internal allocation model combining CRRA risk "
        "profiling, Markowitz (long-only) optimization, and Merton two-fund separation. "
        "For advisor use only.",
        ParagraphStyle("Footer", parent=normal, fontSize=7.5, textColor=colors.grey)
    ))

    doc.build(story)
    return out_path
