# import streamlit as st

# st.title("Assembly Cost Calculator")

# # Direct typing inputs
# assembly = st.text_input("Enter the assembly units:")
# components = st.text_input("Enter the component units:")

# if st.button("Calculate"):

#     try:
#         assembly = int(assembly)
#         components = int(components)

#         if 100 <= assembly < 200:
#             if 0 <= components <= 50:
#                 price = 5.5 * components
#                 st.write("Price:", price)
#             elif components <= 100:
#                 price = 4.5 * components
#                 st.write("Price:", price)
#             else:
#                 st.write("No charge")

            

#         elif 200 <= assembly < 300:
#             if 0 <= components <= 50:
#                 price = 5.5 * components
#                 st.write("Price:", price)
#             elif components <= 100:
#                 price = 4.5 * components
#                 st.write("Price:", price)
#             else:
#                 st.write("No charge")

            

#         elif 300 <= assembly < 500:
#             if 0 <= components <= 50:
#                 price = 5.5 * components
#                 st.write("Price:", price)
#             elif components <= 100:
#                 price = 4.5 * components
#                 st.write("Price:", price)
#             else:
#                 st.write("No charge")

            

#         elif assembly >= 500:
#             if 0 <= components <= 50:
#                 price = 5.5 * components
#                 st.write("Price:", price)
#             elif components <= 100:
#                 price = 4.5 * components
#                 st.write("Price:", price)
#             else:
#                 st.write("No charge")

            

#         else:
#             st.write("No charge")

#     except:
#         st.error("Please enter valid numbers")

import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date
import io

st.set_page_config(
    page_title="PCB Assembly Cost Calculator",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Styling ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

.stApp { background: #0d1117; color: #e6edf3; }

h1 { font-family: 'IBM Plex Mono', monospace !important; color: #58a6ff !important;
     font-size: 1.6rem !important; letter-spacing: -0.5px; margin-bottom: 0 !important; }

h2, h3 { font-family: 'IBM Plex Mono', monospace !important; }

.section-header {
    background: linear-gradient(90deg, #1f2d3d 0%, #161b22 100%);
    border-left: 3px solid #58a6ff;
    padding: 10px 16px;
    border-radius: 4px;
    margin: 24px 0 12px 0;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    font-weight: 600;
    color: #58a6ff;
    letter-spacing: 1px;
    text-transform: uppercase;
}

.badge {
    display: inline-block;
    background: #21262d;
    border: 1px solid #30363d;
    border-radius: 20px;
    padding: 2px 10px;
    font-size: 0.72rem;
    color: #8b949e;
    font-family: 'IBM Plex Mono', monospace;
    margin-bottom: 6px;
}

.stNumberInput > label, .stTextInput > label, .stSelectbox > label {
    font-size: 0.78rem !important;
    color: #8b949e !important;
    font-family: 'IBM Plex Mono', monospace !important;
    letter-spacing: 0.3px;
}

.stNumberInput input, .stTextInput input {
    background: #161b22 !important;
    border: 1px solid #30363d !important;
    color: #e6edf3 !important;
    border-radius: 6px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.85rem !important;
}

.stNumberInput input:focus, .stTextInput input:focus {
    border-color: #58a6ff !important;
    box-shadow: 0 0 0 2px rgba(88,166,255,0.2) !important;
}

.stButton > button {
    background: linear-gradient(135deg, #238636, #2ea043) !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 12px 28px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.5px;
    transition: all 0.2s ease !important;
    width: 100%;
}
.stButton > button:hover { background: linear-gradient(135deg, #2ea043, #3fb950) !important; transform: translateY(-1px); }

.stDownloadButton > button {
    background: linear-gradient(135deg, #1f6feb, #388bfd) !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 12px 28px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
    width: 100%;
}

[data-testid="stExpander"] {
    background: #161b22 !important;
    border: 1px solid #30363d !important;
    border-radius: 8px !important;
}

.divider { border-top: 1px solid #21262d; margin: 20px 0; }

.top-banner {
    background: linear-gradient(90deg, #0d1117 0%, #161b22 50%, #0d1117 100%);
    border: 1px solid #21262d;
    border-radius: 10px;
    padding: 20px 28px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 16px;
}

.metric-box {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 8px;
    padding: 14px 18px;
    text-align: center;
}
.metric-label { font-size: 0.7rem; color: #8b949e; font-family: 'IBM Plex Mono', monospace; text-transform: uppercase; letter-spacing: 1px; }
.metric-value { font-size: 1.4rem; font-weight: 700; color: #58a6ff; font-family: 'IBM Plex Mono', monospace; margin-top: 4px; }
</style>
""", unsafe_allow_html=True)


# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-bottom:6px;">
  <span class="badge">BAUGRUPPE LOHN</span>
</div>
""", unsafe_allow_html=True)

st.title("PCB Assembly Cost Calculator")
st.markdown('<div style="color:#8b949e; font-size:0.88rem; margin-bottom:24px;">Translated from German Excel · Labour · Material · Setup · Pricing</div>', unsafe_allow_html=True)


# ── Helper ───────────────────────────────────────────────────────────────────
def num(label, default, key, fmt="%.4f", min_v=None, step=None, help=None):
    kw = dict(label=label, value=float(default), format=fmt, key=key)
    if min_v is not None: kw["min_value"] = float(min_v)
    if step is not None:  kw["step"] = float(step)
    if help:              kw["help"] = help
    return st.number_input(**kw)

def integer(label, default, key, min_v=0):
    return st.number_input(label, value=int(default), min_value=min_v, step=1, key=key, format="%d")


# ═══════════════════════════════════════════════════════════════════════════
# FORM
# ═══════════════════════════════════════════════════════════════════════════

with st.form("main_form"):

    # ── General ─────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">⚙ General</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        quantity = integer("Total Quantity", 100, "quantity", min_v=1)
    with col2:
        circuits_panel = integer("Circuits per Panel", 4, "circuits_panel", min_v=1)
    with col3:
        created_by = st.text_input("Created by", value="Ashok Kumar", key="created_by")

    # ── SMD Assembly ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">🔵 SMD Assembly</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        smd_time_std  = num("SMD Time/Assembly [Ct] – Standard",  5.5,  "smd_time_std",  "%.2f")
        fix_printer_std  = num("FIX Printer/side [€] – Standard",  30.0, "fix_printer_std",  "%.2f")
        fix_machine_std  = num("FIX Machine/side [€] – Standard",  30.0, "fix_machine_std",  "%.2f")
        feeder_cost_std  = num("Feeder Setup Cost [€] – Standard",  6.5,  "feeder_cost_std",  "%.2f")
    with col2:
        smd_time_cust = num("SMD Time/Assembly [Ct] – Customer", 4.0,  "smd_time_cust",  "%.2f")
        fix_printer_cust = num("FIX Printer/side [€] – Customer", 30.0, "fix_printer_cust", "%.2f")
        fix_machine_cust = num("FIX Machine/side [€] – Customer", 30.0, "fix_machine_cust", "%.2f")
        feeder_cost_cust = num("Feeder Setup Cost [€] – Customer",  6.5,  "feeder_cost_cust", "%.2f")
    with col3:
        smd_qty         = integer("SMD Component Count", 120, "smd_qty")
        soldering_sides = st.selectbox("Soldering Sides", [1, 2], key="soldering_sides")
        feeder_qty      = integer("Number of SMD Feeder Types", 33, "feeder_qty")

    # ── THT Assembly ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">🟠 THT Assembly</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        tht_parts          = integer("THT Component Count", 0, "tht_parts")
        fix_selective_std  = num("FIX Selective/Wave [€] – Standard", 30.0, "fix_selective_std", "%.2f")
        tht_time_std       = num("THT Time [Ct] – Standard", 55.0, "tht_time_std", "%.2f")
    with col2:
        tht_qty            = integer("THT Qty for Time Calc", 0, "tht_qty")
        fix_selective_cust = num("FIX Selective/Wave [€] – Customer", 30.0, "fix_selective_cust", "%.2f")
        tht_time_cust      = num("THT Time [Ct] – Customer", 45.0, "tht_time_cust", "%.2f")

    # ── Manual Soldering ────────────────────────────────────────────────────
    st.markdown('<div class="section-header">🔧 Manual Soldering</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        manual_joints_std  = integer("Manual Solder Joints – Standard", 0, "manual_joints_std")
    with col2:
        manual_joints_cust = integer("Manual Solder Joints – Customer", 0, "manual_joints_cust")
    with col3:
        manual_time        = num("Time per Joint [min]", 0.0, "manual_time", "%.3f")

    # ── QS / Hourly Rate ────────────────────────────────────────────────────
    st.markdown('<div class="section-header">⏱ QS / Hourly Rate</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        qs_time          = num("QS Time [min]", 1.8, "qs_time", "%.2f")
    with col2:
        hourly_rate_std  = num("Hourly Rate [€] – Standard", 60.0, "hourly_rate_std", "%.2f")
    with col3:
        hourly_rate_cust = num("Hourly Rate [€] – Customer", 60.0, "hourly_rate_cust", "%.2f")

    # ── Material ─────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📦 Material</div>', unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        material_cost_std    = num("Material Cost/Unit [€] – Std",  0.0, "material_cost_std",  "%.2f")
    with col2:
        material_cost_cust   = num("Material Cost/Unit [€] – Cust", 0.0, "material_cost_cust", "%.2f")
    with col3:
        material_markup_std  = num("Material Markup % – Standard",  0.0, "material_markup_std",  "%.2f")
    with col4:
        material_markup_cust = num("Material Markup % – Customer",  0.0, "material_markup_cust", "%.2f")

    # ── Pricing ──────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">💰 Pricing</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        skonto        = num("Skonto / Early-Payment Discount (e.g. 0.02 = 2%)", 0.0, "skonto", "%.4f")
    with col2:
        profit_margin = num("Profit Margin (e.g. 0.05 = 5%)", 0.05, "profit_margin", "%.4f")

    # ── Setup / Initial Costs ─────────────────────────────────────────────────
    st.markdown('<div class="section-header">🏗 Setup / Initial Costs</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Standard**")
        setup_smd_std        = num("Setup SMD Assembly/side [€]",   150.0, "setup_smd_std",   "%.2f")
        setup_printer_std    = num("Setup Printer Programme [€]",    45.0, "setup_printer_std","%.2f")
        setup_smd_part_std   = num("Setup SMD per Part [€]",          6.75, "setup_smd_part_std","%.2f")
        setup_tht_part_std   = num("Setup THT per Part [€]",          6.0,  "setup_tht_part_std","%.2f")
        stencil_top_std      = num("Stencil Top [€]",               150.0, "stencil_top_std",  "%.2f")
        stencil_bot_std      = num("Stencil Bottom [€]",            150.0, "stencil_bot_std",  "%.2f")
        order_processing_std = num("Order Processing [€]",           60.0, "order_processing_std","%.2f")
    with col2:
        st.markdown("**Customer**")
        setup_smd_cust       = num("Setup SMD Assembly/side [€]",    90.0, "setup_smd_cust",   "%.2f")
        setup_printer_cust   = num("Setup Printer Programme [€]",    45.0, "setup_printer_cust","%.2f")
        setup_smd_part_cust  = num("Setup SMD per Part [€]",          1.75, "setup_smd_part_cust","%.2f")
        setup_tht_part_cust  = num("Setup THT per Part [€]",          2.5,  "setup_tht_part_cust","%.2f")
        stencil_top_cust     = num("Stencil Top [€]",                 0.0, "stencil_top_cust", "%.2f")
        stencil_bot_cust     = num("Stencil Bottom [€]",              0.0, "stencil_bot_cust", "%.2f")
        order_processing_cust= num("Order Processing [€]",           60.0, "order_processing_cust","%.2f")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        setup_lp_cust        = num("Setup LP Supplier [€] – Customer", 0.0, "setup_lp_cust", "%.2f")
    with col2:
        stencil_top_factor   = num("Stencil Top Surcharge Factor", 0.2, "stencil_top_factor", "%.2f")
    with col3:
        stencil_bot_factor   = num("Stencil Bottom Surcharge Factor", 0.2, "stencil_bot_factor", "%.2f")
    with col4:
        std_days             = integer("Standard Lead Time [days]", 110, "std_days", min_v=1)

    st.markdown("<br>", unsafe_allow_html=True)
    submitted = st.form_submit_button("⚡ Calculate & Generate Excel")


# ═══════════════════════════════════════════════════════════════════════════
# CALCULATION & OUTPUT
# ═══════════════════════════════════════════════════════════════════════════

if submitted:
    p = dict(
        quantity=quantity, circuits_panel=circuits_panel, created_by=created_by,
        smd_time_std=smd_time_std, smd_time_cust=smd_time_cust, smd_qty=smd_qty,
        soldering_sides=soldering_sides,
        fix_printer_std=fix_printer_std, fix_printer_cust=fix_printer_cust,
        fix_machine_std=fix_machine_std, fix_machine_cust=fix_machine_cust,
        feeder_cost_std=feeder_cost_std, feeder_cost_cust=feeder_cost_cust, feeder_qty=feeder_qty,
        tht_parts=tht_parts, fix_selective_std=fix_selective_std, fix_selective_cust=fix_selective_cust,
        tht_time_std=tht_time_std, tht_time_cust=tht_time_cust, tht_qty=tht_qty,
        manual_joints_std=manual_joints_std, manual_joints_cust=manual_joints_cust, manual_time=manual_time,
        qs_time=qs_time, hourly_rate_std=hourly_rate_std, hourly_rate_cust=hourly_rate_cust,
        material_cost_std=material_cost_std, material_cost_cust=material_cost_cust,
        material_markup_std=material_markup_std, material_markup_cust=material_markup_cust,
        skonto=skonto, profit_margin=profit_margin,
        setup_smd_std=setup_smd_std, setup_smd_cust=setup_smd_cust,
        setup_printer_std=setup_printer_std, setup_printer_cust=setup_printer_cust,
        setup_smd_part_std=setup_smd_part_std, setup_smd_part_cust=setup_smd_part_cust,
        setup_tht_part_std=setup_tht_part_std, setup_tht_part_cust=setup_tht_part_cust,
        setup_lp_cust=setup_lp_cust,
        stencil_top_std=stencil_top_std, stencil_top_cust=stencil_top_cust, stencil_top_factor=stencil_top_factor,
        stencil_bot_std=stencil_bot_std, stencil_bot_cust=stencil_bot_cust, stencil_bot_factor=stencil_bot_factor,
        order_processing_std=order_processing_std, order_processing_cust=order_processing_cust,
        std_days=std_days, today=date.today().strftime("%d.%m.%Y"),
    )

    # ── Quick preview calcs (Python) ─────────────────────────────────────────
    hr_s = hourly_rate_std;  hr_c = hourly_rate_cust
    qty  = quantity

    # Labour per unit
    smd_s  = smd_time_std  * smd_qty  / 100
    smd_c  = smd_time_cust * smd_qty  / 100
    sol_s  = (soldering_sides / circuits_panel / 100) if soldering_sides == 2 else 0.0
    sol_c  = sol_s
    fix_s  = ((fix_printer_std  * soldering_sides) + (fix_machine_std  * soldering_sides) + (feeder_cost_std  * feeder_qty)) / qty
    fix_c  = ((fix_printer_cust * soldering_sides) + (fix_machine_cust * soldering_sides) + (feeder_cost_cust * feeder_qty)) / qty
    tht_s  = ((fix_selective_std  * soldering_sides) / qty) + (tht_time_std  * tht_qty / 100)
    tht_c  = ((fix_selective_cust * soldering_sides) / qty) + (tht_time_cust * tht_qty / 100)
    man_s  = manual_joints_std  * hr_s * manual_time / 60
    man_c  = manual_joints_cust * hr_c * manual_time / 60
    qs_s   = qs_time * hr_s / 60
    qs_c   = qs_time * hr_c / 60
    labour_s = (smd_s + sol_s + fix_s + tht_s) * hr_s + man_s + qs_s
    labour_c = (smd_c + sol_c + fix_c + tht_c) * hr_c + man_c + qs_c

    mat_s  = material_cost_std  * (1 + material_markup_std  / 100)
    mat_c  = material_cost_cust * (1 + material_markup_cust / 100)

    unit_s_raw = labour_s + mat_s
    unit_c_raw = labour_c + mat_c
    unit_s = unit_s_raw * (1 + skonto) * (1 + profit_margin)
    unit_c = unit_c_raw * (1 + skonto) * (1 + profit_margin)

    setup_cust = (setup_smd_cust * soldering_sides + setup_printer_cust * soldering_sides
                  + setup_smd_part_cust * feeder_qty + setup_tht_part_cust * tht_qty
                  + setup_lp_cust
                  + stencil_top_cust * (1 + stencil_top_factor)
                  + stencil_bot_cust * (1 + stencil_bot_factor)
                  + order_processing_cust)
    setup_std  = (setup_smd_std * soldering_sides + setup_printer_std * soldering_sides
                  + setup_smd_part_std * feeder_qty + setup_tht_part_std * tht_qty
                  + stencil_top_std * (1 + stencil_top_factor)
                  + stencil_bot_std * (1 + stencil_bot_factor)
                  + order_processing_std)

    total_std  = unit_s * qty + setup_std
    total_cust = unit_c * qty + setup_cust

    # ── Results display ───────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📊 Results Preview</div>', unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""<div class="metric-box">
            <div class="metric-label">Unit Price (Standard)</div>
            <div class="metric-value">€{unit_s:,.2f}</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""<div class="metric-box">
            <div class="metric-label">Unit Price (Customer)</div>
            <div class="metric-value">€{unit_c:,.2f}</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""<div class="metric-box">
            <div class="metric-label">Total Order (Standard)</div>
            <div class="metric-value">€{total_std:,.2f}</div>
        </div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""<div class="metric-box">
            <div class="metric-label">Total Order (Customer)</div>
            <div class="metric-value">€{total_cust:,.2f}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""<div class="metric-box">
            <div class="metric-label">Labour / Unit (Std)</div>
            <div class="metric-value" style="font-size:1.1rem">€{labour_s:,.4f}</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""<div class="metric-box">
            <div class="metric-label">Labour / Unit (Cust)</div>
            <div class="metric-value" style="font-size:1.1rem">€{labour_c:,.4f}</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""<div class="metric-box">
            <div class="metric-label">Setup Costs (Customer)</div>
            <div class="metric-value" style="font-size:1.1rem">€{setup_cust:,.2f}</div>
        </div>""", unsafe_allow_html=True)

    # ── Generate Excel ────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)

    from build_excel import build_excel_bytes
    excel_bytes = build_excel_bytes(p)
    filename = f"Assembly_Cost_{quantity}pcs_{p['today'].replace('.','')}.xlsx"

    col1, col2 = st.columns([1, 1])
    with col1:
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        st.success(f"✅ Calculated for **{quantity} pcs** · {p['today']}")