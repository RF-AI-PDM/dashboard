"""Dashboard Monitoring Pekerjaan - Streamlit Web Application.

A comprehensive monitoring and executive dashboard built with Streamlit, Plotly,
and python-pptx for project tracking, financial performance, and reporting.
Follows PLN / IPS visual identity.
All source code must be ASCII-only.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
import io
import os
import shutil
from typing import Any, Dict, List, Optional
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import data_loader as dl
from data_loader import fmt_pct, fmt_rp, load_all
import excel_writer
import ppt_export
import theme

# Override with env var DASHBOARD_EXCEL to point the app at a different workbook (e.g. a test copy).
EXCEL_PATH = os.environ.get("DASHBOARD_EXCEL", os.path.join("data", "Data.xlsx"))

st.set_page_config(
    layout="wide",
    page_title=theme.APP_TITLE,
    page_icon=":bar_chart:",
    initial_sidebar_state="expanded"
)

# Custom CSS for PLN / IPS professional styling
st.markdown(f"""
<style>
    /* Metric styling: white with navy left border 4px */
    [data-testid="stMetric"] {{
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 4px solid {theme.PRIMARY};
        border-radius: 8px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }}
    [data-testid="stMetricLabel"] {{
        font-size: 0.82rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    [data-testid="stMetricValue"] {{
        font-size: 1.4rem;
        font-weight: 700;
        color: {theme.PRIMARY};
    }}
    /* Tabs styling: active navy with yellow underline */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 6px 6px 0 0;
        padding: 10px 18px;
        font-weight: 600;
        color: #4A5568;
    }}
    .stTabs [aria-selected="true"] {{
        color: {theme.PRIMARY} !important;
        border-bottom: 3px solid {theme.ACCENT} !important;
        background-color: rgba(10, 61, 143, 0.05);
    }}
    /* Section headers */
    .section-title {{
        font-size: 1.2rem;
        font-weight: 700;
        color: {theme.PRIMARY};
        margin-top: 1.2rem;
        margin-bottom: 0.6rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }}
    /* Card box */
    .card-box {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 12px;
    }}
</style>
""", unsafe_allow_html=True)


def get_file_mtime(path: str) -> float:
    """Return modification time of the file, or 0.0 if not found."""
    return os.path.getmtime(path) if os.path.exists(path) else 0.0


@st.cache_data
def get_dashboard_data(file_path: str, mtime: float) -> Dict[str, Any]:
    """Cache data loader based on file modification timestamp."""
    return load_all(file_path)


# ---------------------------------------------------------
# Data Loading & Pre-flight
# ---------------------------------------------------------
if not os.path.exists(EXCEL_PATH):
    st.error(f"File data tidak ditemukan di path: `{EXCEL_PATH}`. Pastikan file Excel tersedia.")
    st.stop()

mtime = get_file_mtime(EXCEL_PATH)
data = get_dashboard_data(EXCEL_PATH, mtime)

df_monitoring = data.get("monitoring", pd.DataFrame())
pencapaian_kpi = data.get("pencapaian", {})
akumulatif_kpi = data.get("akumulatif_kpi", {})
df_monitoring_ak = data.get("monitoring_akumulatif", pd.DataFrame())
df_potensi = data.get("potensi", pd.DataFrame())
penagihan_sections = data.get("penagihan", [])

# ---------------------------------------------------------
# Sidebar: Controls & Filters
# ---------------------------------------------------------
with st.sidebar:
    # Sidebar Logo or Badge
    if os.path.exists(theme.LOGO_PATH):
        st.image(theme.LOGO_PATH, width='stretch')
    else:
        st.markdown(
            f"""<div style="background-color: {theme.PRIMARY}; color: {theme.ACCENT};
            padding: 10px 14px; border-radius: 6px; font-weight: bold; text-align: center;
            font-size: 1.1rem; margin-bottom: 12px;">{theme.ORG_NAME}</div>""",
            unsafe_allow_html=True
        )

    st.title("Monitoring Proyek")
    st.caption("Sistem Pemantauan Progres & Keuangan")

    file_dt = datetime.fromtimestamp(mtime).strftime("%d-%m-%Y %H:%M:%S") if mtime > 0 else "-"
    st.info(f"**File Data:** `{EXCEL_PATH}`\n\n**Terakhir Diperbarui:** {file_dt}")

    if st.button("Muat Ulang Data", width='stretch', type="primary"):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.subheader("Filter Monitoring")

    FILTER_KEYS = ["f_search", "f_user", "f_proses", "f_eksekutor", "f_status", "f_bayar", "f_vendor"]
    if st.session_state.pop("_reset_filters", False):
        for k in FILTER_KEYS:
            st.session_state.pop(k, None)

    # Apply pending filter from chart selection before widgets are created
    if "_pending_filter" in st.session_state:
        p_key, p_val = st.session_state.pop("_pending_filter")
        st.session_state[p_key] = p_val

    # Search filter
    search_query = st.text_input("Cari Judul Pekerjaan", placeholder="Ketik kata kunci...", key="f_search")

    # Options helper
    def _unique_options(col_name: str) -> List[str]:
        if col_name in df_monitoring.columns:
            vals = df_monitoring[col_name].dropna().unique().tolist()
            clean_vals = [str(v).strip() for v in vals if str(v).strip() and str(v).strip().lower() != "nan"]
            return sorted(list(set(clean_vals)))
        return []

    sel_users = st.multiselect("User / Bidang", options=sorted(set(_unique_options("User")) | {"Admin"}), key="f_user")
    sel_proses = st.multiselect("Proses Kontrak", options=_unique_options("Proses Kontrak"), key="f_proses")
    sel_eksekutor = st.multiselect("Eksekutor", options=_unique_options("Eksekutor"), key="f_eksekutor")
    sel_status = st.multiselect("Status Pekerjaan", options=_unique_options("STATUS"), key="f_status")
    sel_pembayaran = st.multiselect("Status Pembayaran", options=_unique_options("STATUS PEMBAYARAN"), key="f_bayar")
    sel_vendors = st.multiselect("Nama Vendor", options=_unique_options("Nama Vendor"), key="f_vendor")

    active_drill = [k for k in ["f_status", "f_bayar", "f_user", "f_proses"] if st.session_state.get(k)]
    if active_drill:
        st.caption("Filter aktif dari klik grafik:")
        if st.button("Hapus Filter Grafik", key="btn_clear_drill", width='stretch'):
            for k in ["f_status", "f_bayar", "f_user", "f_proses"]:
                st.session_state.pop(k, None)
            st.rerun()

    if st.button("Reset Semua Filter", width='stretch'):
        st.session_state["_reset_filters"] = True
        st.rerun()

# ---------------------------------------------------------
# Filtering Logic
# ---------------------------------------------------------
filtered_df = df_monitoring.copy()

if search_query.strip():
    q = search_query.strip().lower()
    filtered_df = filtered_df[filtered_df["Judul Pekerjaan"].str.lower().str.contains(q, na=False)]

if sel_users:
    filtered_df = filtered_df[filtered_df["User"].isin(sel_users)]

if sel_proses:
    filtered_df = filtered_df[filtered_df["Proses Kontrak"].isin(sel_proses)]

if sel_eksekutor:
    filtered_df = filtered_df[filtered_df["Eksekutor"].isin(sel_eksekutor)]

if sel_status:
    filtered_df = filtered_df[filtered_df["STATUS"].isin(sel_status)]

if sel_pembayaran:
    filtered_df = filtered_df[filtered_df["STATUS PEMBAYARAN"].isin(sel_pembayaran)]

if sel_vendors:
    filtered_df = filtered_df[filtered_df["Nama Vendor"].isin(sel_vendors)]

# ---------------------------------------------------------
# App Header Band (PLN Theme)
# ---------------------------------------------------------
st.markdown(f"""
<div style="background-color: {theme.PRIMARY}; border-bottom: 4px solid {theme.ACCENT};
padding: 16px 22px; border-radius: 8px; margin-bottom: 18px;">
    <h1 style="color: #FFFFFF; margin: 0; font-size: 1.75rem; font-weight: 700;">{theme.APP_TITLE}</h1>
    <p style="color: {theme.ACCENT}; margin: 4px 0 0 0; font-size: 0.95rem; font-weight: 600;">{theme.ORG_NAME}</p>
</div>
""", unsafe_allow_html=True)

# One-shot success message that survives st.rerun() after a save
if "_flash" in st.session_state:
    st.success(st.session_state.pop("_flash"))

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "Ringkasan",
    "Monitoring Pekerjaan",
    "Akumulatif",
    "Potensi Pendapatan",
    "Alur Penagihan",
    "Kelola Data",
    "Export PPT",
    "Export Excel"
])

# =========================================================
# TAB 1: RINGKASAN
# =========================================================
with tab1:
    st.markdown('<div class="section-title">Indikator Kinerja Utama (KPI)</div>', unsafe_allow_html=True)

    target_tahun = pencapaian_kpi.get("target_tahun", 0.0)
    realisasi_unfiltered = (
        float(df_monitoring["Nilai Kontrak Setelah PPN (IPS)"].sum())
        if "Nilai Kontrak Setelah PPN (IPS)" in df_monitoring.columns
        else 0.0
    )
    pct_pencapaian = (realisasi_unfiltered / target_tahun) if target_tahun > 0 else 0.0
    deviasi = realisasi_unfiltered - target_tahun

    total_invoice = (
        float(df_monitoring["Nilai Invoice Setelah PPN"].sum())
        if "Nilai Invoice Setelah PPN" in df_monitoring.columns
        else 0.0
    )
    total_potensi = (
        float(df_potensi["POTENSI NILAI"].sum())
        if not df_potensi.empty and "POTENSI NILAI" in df_potensi.columns
        else 0.0
    )

    jml_pekerjaan = len(df_monitoring)
    jml_selesai = (
        int((df_monitoring["STATUS"] == "SELESAI").sum())
        if "STATUS" in df_monitoring.columns
        else 0
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(
            label="Target Tahun 2026",
            value=fmt_rp(target_tahun),
            help="Target anggaran tahun 2026 dari sheet Pencapaian"
        )
    with col2:
        st.metric(
            label="Realisasi Kontrak (IPS)",
            value=fmt_rp(realisasi_unfiltered),
            delta=fmt_pct(pct_pencapaian),
            help="Total nilai kontrak setelah PPN (IPS) dari sheet Monitoring"
        )
    with col3:
        st.metric(
            label="% Pencapaian",
            value=fmt_pct(pct_pencapaian),
            delta=f"Deviasi: {fmt_rp(deviasi)}",
            delta_color="normal" if deviasi >= 0 else "inverse"
        )
    with col4:
        st.metric(
            label="Deviasi Target",
            value=fmt_rp(deviasi),
            delta=fmt_pct(deviasi / target_tahun) if target_tahun > 0 else "-",
            delta_color="normal" if deviasi >= 0 else "inverse"
        )

    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.metric(
            label="Total Nilai Invoice (PPN)",
            value=fmt_rp(total_invoice),
            help="Total nilai invoice yang telah diterbitkan"
        )
    with col6:
        st.metric(
            label="Total Potensi Pendapatan",
            value=fmt_rp(total_potensi),
            help="Total potensi nilai pekerjaan yang berpeluang diperoleh"
        )
    with col7:
        st.metric(
            label="Jumlah Pekerjaan",
            value=f"{jml_pekerjaan} Pekerjaan",
            help="Total pekerjaan terdaftar dalam sistem monitoring"
        )
    with col8:
        st.metric(
            label="Pekerjaan Selesai",
            value=f"{jml_selesai} Selesai",
            delta=f"{fmt_pct(jml_selesai / jml_pekerjaan if jml_pekerjaan > 0 else 0)} tuntas",
            delta_color="normal"
        )

    # -----------------------------------------------------
    # Gauges (Tahun, Semester 1, Semester 2)
    # -----------------------------------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">Pengukur Capaian Target (Gauges)</div>', unsafe_allow_html=True)

    def _make_gauge(title: str, ratio: float, target_val: float, real_val: float) -> go.Figure:
        val_pct = ratio * 100.0
        max_range = max(120.0, val_pct * 1.1)
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=val_pct,
            number={"suffix": "%", "valueformat": ".1f", "font": {"color": theme.PRIMARY, "size": 26}},
            title={
                "text": f"<b>{title}</b><br><span style='font-size:0.75em;color:#64748B'>Target: {fmt_rp(target_val)}<br>Realisasi: {fmt_rp(real_val)}</span>",
                "font": {"size": 13, "color": theme.PRIMARY_DARK}
            },
            gauge={
                "axis": {"range": [0, max_range], "tickwidth": 1, "tickcolor": "#CBD5E1"},
                "bar": {"color": theme.PRIMARY},
                "bgcolor": "#F1F5F9",
                "borderwidth": 1,
                "bordercolor": "#E2E8F0",
                "threshold": {
                    "line": {"color": theme.ACCENT, "width": 4},
                    "thickness": 0.8,
                    "value": 100.0,
                }
            }
        ))
        fig.update_layout(
            height=230,
            margin=dict(l=25, r=25, t=65, b=20),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)"
        )
        return fig

    cg1, cg2, cg3 = st.columns(3)
    with cg1:
        fig_g_tahun = _make_gauge(
            "Capaian Tahun 2026",
            pencapaian_kpi.get("ratio_tahun", 0.0),
            pencapaian_kpi.get("target_tahun", 0.0),
            pencapaian_kpi.get("realisasi_tahun", 0.0)
        )
        st.plotly_chart(fig_g_tahun, width='stretch')

    with cg2:
        fig_g_s1 = _make_gauge(
            "Capaian Semester 1",
            pencapaian_kpi.get("ratio_s1", 0.0),
            pencapaian_kpi.get("target_s1", 0.0),
            pencapaian_kpi.get("capaian_s1", 0.0)
        )
        st.plotly_chart(fig_g_s1, width='stretch')

    with cg3:
        fig_g_s2 = _make_gauge(
            "Capaian Semester 2",
            pencapaian_kpi.get("ratio_s2", 0.0),
            pencapaian_kpi.get("target_s2", 0.0),
            pencapaian_kpi.get("capaian_s2", 0.0)
        )
        st.plotly_chart(fig_g_s2, width='stretch')

    # -----------------------------------------------------
    # Target vs Realisasi Charts
    # -----------------------------------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">Target vs Realisasi Anggaran</div>', unsafe_allow_html=True)

    col_chart_left, col_chart_right = st.columns(2)

    def _make_grouped_bar(kpi: Dict[str, Any], chart_title: str) -> go.Figure:
        categories = ["Tahun 2026", "Semester 1", "Semester 2"]
        targets = [kpi.get("target_tahun", 0.0), kpi.get("target_s1", 0.0), kpi.get("target_s2", 0.0)]
        realisasi = [kpi.get("realisasi_tahun", 0.0), kpi.get("capaian_s1", 0.0), kpi.get("capaian_s2", 0.0)]
        ratios = [kpi.get("ratio_tahun", 0.0), kpi.get("ratio_s1", 0.0), kpi.get("ratio_s2", 0.0)]

        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="Target",
            x=categories,
            y=targets,
            marker_color=theme.ACCENT,
            customdata=[fmt_rp(v) for v in targets],
            hovertemplate="<b>%{x}</b><br>Target: %{customdata}<extra></extra>"
        ))
        fig.add_trace(go.Bar(
            name="Realisasi",
            x=categories,
            y=realisasi,
            marker_color=theme.PRIMARY,
            customdata=[f"{fmt_rp(v)} ({fmt_pct(r)})" for v, r in zip(realisasi, ratios)],
            hovertemplate="<b>%{x}</b><br>Realisasi: %{customdata}<extra></extra>"
        ))

        for idx, (cat, real, rat) in enumerate(zip(categories, realisasi, ratios)):
            fig.add_annotation(
                x=cat,
                y=max(targets[idx], real) * 1.06,
                text=f"<b>{fmt_pct(rat)}</b>",
                showarrow=False,
                font=dict(size=12, color=theme.PRIMARY),
                bgcolor="#FFFFFF",
                bordercolor="#CBD5E1",
                borderwidth=1,
                borderpad=3
            )

        fig.update_layout(
            title=dict(text=chart_title, font=dict(size=14, color=theme.PRIMARY, family="Segoe UI")),
            barmode="group",
            margin=dict(l=20, r=20, t=50, b=40),
            height=370,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#F1F5F9", showgrid=True, title="Nilai (Rupiah)"),
            xaxis=dict(showgrid=False)
        )
        return fig

    with col_chart_left:
        fig_pencapaian = _make_grouped_bar(pencapaian_kpi, "Pencapaian Kinerja 2026 (Sheet: PENCAPAIAN)")
        st.plotly_chart(fig_pencapaian, width='stretch')

    with col_chart_right:
        fig_akumulatif = _make_grouped_bar(akumulatif_kpi, "Kinerja Akumulatif (Sheet: Grafik Akumulatif)")
        st.plotly_chart(fig_akumulatif, width='stretch')

    # -----------------------------------------------------
    # 4 Distribution charts with Drill-Down Filter
    # -----------------------------------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">Distribusi Status & Karakteristik Pekerjaan</div>', unsafe_allow_html=True)
    st.caption("Klik bagian grafik untuk memfilter tabel secara otomatis.")

    c_d1, c_d2, c_d3, c_d4 = st.columns(4)

    with c_d1:
        status_counts = filtered_df["STATUS"].value_counts().reset_index()
        status_counts.columns = ["STATUS", "Jumlah"]
        fig_status = px.pie(
            status_counts,
            names="STATUS",
            values="Jumlah",
            hole=0.55,
            title="Distribusi Status",
            color="STATUS",
            color_discrete_map=theme.STATUS_COLORS
        )
        fig_status.update_traces(textposition="inside", textinfo="percent+value")
        fig_status.update_layout(
            height=300,
            margin=dict(l=10, r=10, t=40, b=10),
            showlegend=False,
            title=dict(font=dict(size=14, color=theme.PRIMARY))
        )
        sel_status_res = st.plotly_chart(
            fig_status,
            width='stretch',
            key="drill_status",
            on_select="rerun",
            selection_mode="points"
        )
        if sel_status_res and sel_status_res.get("selection") and sel_status_res["selection"].get("points"):
            pt = sel_status_res["selection"]["points"][0]
            val = pt.get("label") or (
                status_counts["STATUS"].iloc[pt["point_number"]]
                if "point_number" in pt and pt["point_number"] < len(status_counts)
                else None
            )
            if val and st.session_state.get("f_status") != [val]:
                st.session_state["_pending_filter"] = ("f_status", [val])
                st.rerun()

    with c_d2:
        pay_counts = filtered_df["STATUS PEMBAYARAN"].value_counts().reset_index()
        pay_counts.columns = ["STATUS PEMBAYARAN", "Jumlah"]
        fig_pay = px.pie(
            pay_counts,
            names="STATUS PEMBAYARAN",
            values="Jumlah",
            hole=0.55,
            title="Status Pembayaran",
            color="STATUS PEMBAYARAN",
            color_discrete_map=theme.PAYMENT_COLORS
        )
        fig_pay.update_traces(textposition="inside", textinfo="percent+value")
        fig_pay.update_layout(
            height=300,
            margin=dict(l=10, r=10, t=40, b=10),
            showlegend=False,
            title=dict(font=dict(size=14, color=theme.PRIMARY))
        )
        sel_pay_res = st.plotly_chart(
            fig_pay,
            width='stretch',
            key="drill_pay",
            on_select="rerun",
            selection_mode="points"
        )
        if sel_pay_res and sel_pay_res.get("selection") and sel_pay_res["selection"].get("points"):
            pt = sel_pay_res["selection"]["points"][0]
            val = pt.get("label") or (
                pay_counts["STATUS PEMBAYARAN"].iloc[pt["point_number"]]
                if "point_number" in pt and pt["point_number"] < len(pay_counts)
                else None
            )
            if val and st.session_state.get("f_bayar") != [val]:
                st.session_state["_pending_filter"] = ("f_bayar", [val])
                st.rerun()

    with c_d3:
        user_counts = filtered_df["User"].value_counts().reset_index()
        user_counts.columns = ["User", "Jumlah"]
        fig_user = px.bar(
            user_counts,
            x="User",
            y="Jumlah",
            title="Pekerjaan per User",
            color="User",
            color_discrete_sequence=theme.PALETTE
        )
        fig_user.update_layout(
            height=300,
            margin=dict(l=10, r=10, t=40, b=10),
            showlegend=False,
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#F1F5F9"),
            title=dict(font=dict(size=14, color=theme.PRIMARY))
        )
        sel_user_res = st.plotly_chart(
            fig_user,
            width='stretch',
            key="drill_user",
            on_select="rerun",
            selection_mode="points"
        )
        if sel_user_res and sel_user_res.get("selection") and sel_user_res["selection"].get("points"):
            pt = sel_user_res["selection"]["points"][0]
            val = pt.get("x") or (
                user_counts["User"].iloc[pt["point_number"]]
                if "point_number" in pt and pt["point_number"] < len(user_counts)
                else None
            )
            if val and st.session_state.get("f_user") != [val]:
                st.session_state["_pending_filter"] = ("f_user", [val])
                st.rerun()

    with c_d4:
        proses_counts = filtered_df["Proses Kontrak"].value_counts().reset_index()
        proses_counts.columns = ["Proses Kontrak", "Jumlah"]
        fig_proses = px.bar(
            proses_counts,
            x="Proses Kontrak",
            y="Jumlah",
            title="Proses Kontrak",
            color="Proses Kontrak",
            color_discrete_sequence=[theme.PRIMARY, theme.ACCENT, theme.SECONDARY, "#26A69A"]
        )
        fig_proses.update_layout(
            height=300,
            margin=dict(l=10, r=10, t=40, b=10),
            showlegend=False,
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#F1F5F9"),
            title=dict(font=dict(size=14, color=theme.PRIMARY))
        )
        sel_proses_res = st.plotly_chart(
            fig_proses,
            width='stretch',
            key="drill_proses",
            on_select="rerun",
            selection_mode="points"
        )
        if sel_proses_res and sel_proses_res.get("selection") and sel_proses_res["selection"].get("points"):
            pt = sel_proses_res["selection"]["points"][0]
            val = pt.get("x") or (
                proses_counts["Proses Kontrak"].iloc[pt["point_number"]]
                if "point_number" in pt and pt["point_number"] < len(proses_counts)
                else None
            )
            if val and st.session_state.get("f_proses") != [val]:
                st.session_state["_pending_filter"] = ("f_proses", [val])
                st.rerun()

    # -----------------------------------------------------
    # Gantt Timeline Section: Jadwal Pekerjaan
    # -----------------------------------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">Jadwal Pelaksanaan Pekerjaan (Timeline Gantt)</div>', unsafe_allow_html=True)

    gantt_rows = []
    missing_date_jobs = []

    for _, r in filtered_df.iterrows():
        mulai_val = r.get("Mulai Pekerjaan", "")
        berakhir_val = r.get("Berakhir Pekerjaan", "")
        kontrak_val = r.get("Kontrak IPS", "")
        judul = str(r.get("Judul Pekerjaan", "")).strip()

        start_dt = dl.parse_id_date(mulai_val) or dl.parse_id_date(kontrak_val)
        dur_days = dl.parse_duration_days(mulai_val) or dl.parse_duration_days(berakhir_val)
        end_dt = dl.parse_id_date(berakhir_val)

        if not end_dt and start_dt and dur_days:
            end_dt = start_dt + timedelta(days=dur_days)

        if start_dt and end_dt:
            if end_dt < start_dt:
                end_dt = start_dt + timedelta(days=1)
            short_title = (judul[:45] + "...") if len(judul) > 45 else judul
            gantt_rows.append({
                "Judul": short_title,
                "Judul_Lengkap": judul,
                "Start": start_dt.strftime("%Y-%m-%d"),
                "End": end_dt.strftime("%Y-%m-%d"),
                "STATUS": str(r.get("STATUS", "PROSES")),
                "Vendor": str(r.get("Nama Vendor", "-"))
            })
        else:
            missing_date_jobs.append(f"[{r.get('NO', '-')}] {judul}")

    fig_gantt: Optional[go.Figure] = None

    if gantt_rows:
        df_gantt = pd.DataFrame(gantt_rows)
        fig_gantt = px.timeline(
            df_gantt,
            x_start="Start",
            x_end="End",
            y="Judul",
            color="STATUS",
            color_discrete_map=theme.STATUS_COLORS,
            hover_name="Judul_Lengkap",
            hover_data={"Vendor": True, "Start": True, "End": True, "Judul": False}
        )
        fig_gantt.update_yaxes(autorange="reversed")
        fig_gantt.add_vline(
            x=datetime.now(),
            line_width=2,
            line_dash="dash",
            line_color=theme.DANGER,
            annotation_text="Hari Ini",
            annotation_position="top right"
        )
        fig_gantt.update_layout(
            height=max(320, len(df_gantt) * 36),
            margin=dict(l=20, r=20, t=30, b=30),
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(gridcolor="#F1F5F9", title="Rentang Waktu Pelaksanaan"),
            yaxis=dict(showgrid=False, title="")
        )
        st.plotly_chart(fig_gantt, width='stretch')
    else:
        st.info("Data tanggal pelaksanaan belum terisi lengkap pada pekerjaan terpilih.")

    if missing_date_jobs:
        with st.expander(f"Daftar Pekerjaan Belum Memiliki Jadwal Tanggal ({len(missing_date_jobs)} Pekerjaan)"):
            for job_item in missing_date_jobs:
                st.text(job_item)

    # -----------------------------------------------------
    # Horizontal bar: Nilai Kontrak Setelah PPN (IPS)
    # -----------------------------------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">Nilai Kontrak per Judul Pekerjaan</div>', unsafe_allow_html=True)

    col_ips = "Nilai Kontrak Setelah PPN (IPS)"
    if col_ips in filtered_df.columns:
        df_jobs_val = filtered_df[filtered_df[col_ips] > 0].copy()
        df_jobs_val = df_jobs_val.sort_values(by=col_ips, ascending=True)

        if not df_jobs_val.empty:
            df_jobs_val["Label_Short"] = df_jobs_val["Judul Pekerjaan"].apply(
                lambda s: str(s)[:55] + "..." if len(str(s)) > 55 else str(s)
            )
            df_jobs_val["Nilai_Fmt"] = df_jobs_val[col_ips].apply(fmt_rp)

            fig_bar_jobs = go.Figure(go.Bar(
                x=df_jobs_val[col_ips],
                y=df_jobs_val["Label_Short"],
                orientation="h",
                marker_color=theme.PRIMARY,
                customdata=df_jobs_val["Nilai_Fmt"],
                hovertemplate="<b>%{y}</b><br>Nilai Kontrak: %{customdata}<extra></extra>"
            ))
            fig_bar_jobs.update_layout(
                margin=dict(l=20, r=30, t=30, b=30),
                height=max(350, len(df_jobs_val) * 26),
                xaxis=dict(gridcolor="#F1F5F9", title="Nilai Kontrak Setelah PPN (IPS)"),
                yaxis=dict(showgrid=False),
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_bar_jobs, width='stretch')
        else:
            st.info("Tidak ada data pekerjaan dengan nilai kontrak > 0 pada filter saat ini.")

    # -----------------------------------------------------
    # Total nilai kontrak per Nama Vendor
    # -----------------------------------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">Total Nilai Kontrak per Nama Vendor</div>', unsafe_allow_html=True)
    if "Nama Vendor" in filtered_df.columns and col_ips in filtered_df.columns:
        vendor_df = filtered_df[filtered_df["Nama Vendor"].str.strip() != ""].copy()
        vendor_summary = vendor_df.groupby("Nama Vendor")[col_ips].sum().reset_index()
        vendor_summary = vendor_summary[vendor_summary[col_ips] > 0].sort_values(by=col_ips, ascending=False)

        if not vendor_summary.empty:
            vendor_summary["Nilai_Fmt"] = vendor_summary[col_ips].apply(fmt_rp)
            fig_vendor = px.bar(
                vendor_summary,
                x="Nama Vendor",
                y=col_ips,
                color="Nama Vendor",
                color_discrete_sequence=theme.PALETTE,
                custom_data=["Nilai_Fmt"]
            )
            fig_vendor.update_traces(
                hovertemplate="<b>%{x}</b><br>Total Kontrak: %{customdata[0]}<extra></extra>"
            )
            fig_vendor.update_layout(
                height=380,
                margin=dict(l=20, r=20, t=30, b=40),
                showlegend=False,
                plot_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(gridcolor="#F1F5F9", title="Total Nilai (Rp)"),
                xaxis=dict(showgrid=False, tickangle=-15)
            )
            st.plotly_chart(fig_vendor, width='stretch')
        else:
            st.info("Tidak ada data vendor dengan nilai kontrak > 0 pada filter saat ini.")

# =========================================================
# TAB 2: MONITORING PEKERJAAN
# =========================================================
with tab2:
    st.markdown('<div class="section-title">Daftar Detail Monitoring Pekerjaan</div>', unsafe_allow_html=True)
    st.caption(f"Menampilkan **{len(filtered_df)}** dari **{len(df_monitoring)}** total pekerjaan.")

    all_mon_cols = [c for c in df_monitoring.columns if not c.endswith("(Asli)")]
    default_mon_cols = [
        "NO", "Judul Pekerjaan", "User", "Proses Kontrak", "Eksekutor",
        "Kontrak IPS", "Nilai Kontrak Setelah PPN (IPS)", "Nama Vendor",
        "Kendala", "STATUS", "STATUS PEMBAYARAN"
    ]
    selected_cols = st.multiselect(
        "Pilih Kolom Tabel",
        options=all_mon_cols,
        default=[c for c in default_mon_cols if c in all_mon_cols]
    )

    if not selected_cols:
        selected_cols = default_mon_cols

    display_df = filtered_df[selected_cols].copy()

    col_config: Dict[str, Any] = {}
    rp_columns = [
        "IH Sebelum PPN", "IH Setelah PPN",
        "Nilai Kontrak Sebelum PPN (IPS)", "Nilai Kontrak Setelah PPN (IPS)",
        "Nilai Invoice Sebelum PPN", "Nilai Invoice Setelah PPN",
        "DENDA", "PEMBAYARAN PIPS",
        "Nilai Kontrak Mitra Sebelum PPN", "Nilai Kontrak Mitra Setelah PPN",
        "Nilai BA Vendor Setelah PPN", "DENDA2", "PEMBAYARAN MITRA", "MARGIN"
    ]
    # Pre-format as "Rp 1.234.567" text (Indonesian thousands separator) so the
    # table matches the KPI cards; NumberColumn's format spec has no way to
    # express a period thousands-separator, only comma-based locales.
    for c in rp_columns:
        if c in selected_cols:
            display_df[c] = display_df[c].apply(fmt_rp)
            col_config[c] = st.column_config.Column(c, help=f"Nilai moneter untuk {c}")

    if "NO" in selected_cols:
        col_config["NO"] = st.column_config.Column("NO", width="small")

    st.dataframe(
        display_df,
        width='stretch',
        column_config=col_config,
        hide_index=True,
        height=450
    )

    st.markdown("#### Ringkasan Finansial Kolom Terpilih (Data Terfilter)")
    tot_col1, tot_col2, tot_col3, tot_col4 = st.columns(4)

    tot_kontrak_ips = (
        float(filtered_df["Nilai Kontrak Setelah PPN (IPS)"].sum())
        if "Nilai Kontrak Setelah PPN (IPS)" in filtered_df.columns
        else 0.0
    )
    tot_kontrak_mitra = (
        float(filtered_df["Nilai Kontrak Mitra Setelah PPN"].sum())
        if "Nilai Kontrak Mitra Setelah PPN" in filtered_df.columns
        else 0.0
    )
    tot_invoice = (
        float(filtered_df["Nilai Invoice Setelah PPN"].sum())
        if "Nilai Invoice Setelah PPN" in filtered_df.columns
        else 0.0
    )
    tot_margin = (
        float(filtered_df["MARGIN"].sum())
        if "MARGIN" in filtered_df.columns
        else (tot_kontrak_ips - tot_kontrak_mitra)
    )

    with tot_col1:
        st.metric("Total Kontrak IPS (PPN)", fmt_rp(tot_kontrak_ips))
    with tot_col2:
        st.metric("Total Kontrak Mitra (PPN)", fmt_rp(tot_kontrak_mitra))
    with tot_col3:
        st.metric("Total Invoice (PPN)", fmt_rp(tot_invoice))
    with tot_col4:
        st.metric("Total Estimasi Margin", fmt_rp(tot_margin))

    st.markdown("---")
    with st.expander("Catatan / Kendala Pekerjaan Terdaftar", expanded=False):
        kendala_df = filtered_df[filtered_df["Kendala"].str.strip() != ""]
        if not kendala_df.empty:
            for _, r in kendala_df.iterrows():
                st.markdown(f"**[{r.get('NO', '-')}] {r.get('Judul Pekerjaan', '')}** (User: `{r.get('User', '-')}`) :")
                st.info(r.get("Kendala", ""))
        else:
            st.success("Tidak ada kendala tercatat pada pekerjaan yang difilter saat ini.")

    csv_bytes = filtered_df.to_csv(index=False).encode("utf-8-sig")

    excel_buffer = io.BytesIO()
    filtered_df.to_excel(excel_buffer, index=False, sheet_name="Monitoring", engine="openpyxl")
    excel_bytes = excel_buffer.getvalue()

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            label="Unduh Data Terfilter (.CSV)",
            data=csv_bytes,
            file_name=f"Monitoring_Pekerjaan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            width='stretch'
        )
    with dl_col2:
        st.download_button(
            label="Unduh Data Terfilter (.XLSX)",
            data=excel_bytes,
            file_name=f"Monitoring_Pekerjaan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch'
        )

# =========================================================
# TAB 3: AKUMULATIF
# =========================================================
with tab3:
    st.markdown('<div class="section-title">Kinerja dan Monitoring Akumulatif</div>', unsafe_allow_html=True)
    st.caption("Pencapaian dan daftar pekerjaan tahun jamak / akumulatif berjalan.")

    ak_target = akumulatif_kpi.get("target_tahun", 0.0)
    ak_real = akumulatif_kpi.get("realisasi_tahun", 0.0)
    ak_ratio = akumulatif_kpi.get("ratio_tahun", 0.0)
    ak_dev = akumulatif_kpi.get("deviasi_tahun", 0.0)

    ak1, ak2, ak3, ak4 = st.columns(4)
    with ak1:
        st.metric("Target Akumulatif 2026", fmt_rp(ak_target))
    with ak2:
        st.metric("Realisasi Akumulatif", fmt_rp(ak_real), delta=fmt_pct(ak_ratio))
    with ak3:
        st.metric("% Capaian Akumulatif", fmt_pct(ak_ratio))
    with ak4:
        st.metric("Deviasi Akumulatif", fmt_rp(ak_dev), delta_color="normal" if ak_dev >= 0 else "inverse")

    st.markdown("---")
    st.markdown("#### Tabel Monitoring Akumulatif")

    ak_display_df = df_monitoring_ak.copy()
    for c in ["Nilai Kontrak Setelah PPN", "Nilai Kontrak Mitra Setelah PPN"]:
        if c in ak_display_df.columns:
            ak_display_df[c] = ak_display_df[c].apply(fmt_rp)

    ak_config = {
        "Nilai Kontrak Setelah PPN": st.column_config.Column("Nilai Kontrak Setelah PPN"),
        "Nilai Kontrak Mitra Setelah PPN": st.column_config.Column("Nilai Kontrak Mitra Setelah PPN"),
        "NO": st.column_config.Column("NO", width="small")
    }

    st.dataframe(
        ak_display_df,
        width='stretch',
        column_config=ak_config,
        hide_index=True
    )

    tot_ak_ips = (
        float(df_monitoring_ak["Nilai Kontrak Setelah PPN"].sum())
        if "Nilai Kontrak Setelah PPN" in df_monitoring_ak.columns
        else 0.0
    )
    tot_ak_mitra = (
        float(df_monitoring_ak["Nilai Kontrak Mitra Setelah PPN"].sum())
        if "Nilai Kontrak Mitra Setelah PPN" in df_monitoring_ak.columns
        else 0.0
    )

    m_ak1, m_ak2 = st.columns(2)
    with m_ak1:
        st.metric("Total Nilai Kontrak Setelah PPN", fmt_rp(tot_ak_ips))
    with m_ak2:
        st.metric("Total Nilai Kontrak Mitra Setelah PPN", fmt_rp(tot_ak_mitra))

    st.markdown("---")
    st.markdown("#### Grafik Nilai Kontrak Akumulatif per Pekerjaan")
    if not df_monitoring_ak.empty and "Nilai Kontrak Mitra Setelah PPN" in df_monitoring_ak.columns:
        fig_ak_bar = px.bar(
            df_monitoring_ak,
            x="Judul Pekerjaan",
            y="Nilai Kontrak Mitra Setelah PPN",
            color="User",
            color_discrete_sequence=theme.PALETTE,
            title="Nilai Kontrak Mitra per Pekerjaan Akumulatif",
            custom_data=[df_monitoring_ak["Nilai Kontrak Mitra Setelah PPN"].apply(fmt_rp)]
        )
        fig_ak_bar.update_traces(hovertemplate="<b>%{x}</b><br>Nilai Mitra: %{customdata[0]}<extra></extra>")
        fig_ak_bar.update_layout(
            height=380,
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#F1F5F9", title="Nilai (Rp)"),
            margin=dict(l=20, r=20, t=40, b=40)
        )
        st.plotly_chart(fig_ak_bar, width='stretch')

# =========================================================
# TAB 4: POTENSI PENDAPATAN
# =========================================================
with tab4:
    st.markdown('<div class="section-title">Daftar Potensi Pendapatan 2025 - 2026</div>', unsafe_allow_html=True)
    st.caption("Peluang proyek dan estimasi nilai pendapatan kontrak mendatang.")

    tot_potensi_val = (
        float(df_potensi["POTENSI NILAI"].sum())
        if not df_potensi.empty and "POTENSI NILAI" in df_potensi.columns
        else 0.0
    )
    jml_potensi = len(df_potensi)

    p_c1, p_c2 = st.columns(2)
    with p_c1:
        st.metric("Total Potensi Pendapatan", fmt_rp(tot_potensi_val))
    with p_c2:
        st.metric("Jumlah Item Peluang", f"{jml_potensi} Pekerjaan")

    st.markdown("---")

    pot_display_df = df_potensi.copy()
    if "POTENSI NILAI" in pot_display_df.columns:
        pot_display_df["POTENSI NILAI"] = pot_display_df["POTENSI NILAI"].apply(fmt_rp)

    pot_config = {
        "POTENSI NILAI": st.column_config.Column("Potensi Nilai (Rp)"),
        "JUDUL": st.column_config.Column("Judul Pekerjaan", width="large"),
        "PERIODE": st.column_config.Column("Periode", width="small")
    }

    st.dataframe(
        pot_display_df,
        width='stretch',
        column_config=pot_config,
        hide_index=True
    )

    st.markdown("---")
    st.markdown("#### Grafik Perbandingan Nilai Potensi")
    if not df_potensi.empty and "POTENSI NILAI" in df_potensi.columns:
        df_pot_sorted = df_potensi.sort_values(by="POTENSI NILAI", ascending=True)
        df_pot_sorted["Judul_Short"] = df_pot_sorted["JUDUL"].apply(lambda s: str(s)[:50] + "..." if len(str(s)) > 50 else str(s))
        df_pot_sorted["Nilai_Fmt"] = df_pot_sorted["POTENSI NILAI"].apply(fmt_rp)

        fig_pot = go.Figure(go.Bar(
            x=df_pot_sorted["POTENSI NILAI"],
            y=df_pot_sorted["Judul_Short"],
            orientation="h",
            marker_color=theme.SECONDARY,
            customdata=df_pot_sorted["Nilai_Fmt"],
            hovertemplate="<b>%{y}</b><br>Potensi: %{customdata}<extra></extra>"
        ))
        fig_pot.update_layout(
            height=max(320, len(df_pot_sorted) * 45),
            margin=dict(l=20, r=20, t=20, b=30),
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(gridcolor="#F1F5F9", title="Potensi Nilai (Rp)"),
            yaxis=dict(showgrid=False)
        )
        st.plotly_chart(fig_pot, width='stretch')

# =========================================================
# TAB 5: ALUR PENAGIHAN
# =========================================================
with tab5:
    st.markdown('<div class="section-title">Alur & Persyaratan Dokumen Penagihan</div>', unsafe_allow_html=True)
    st.caption("Checklist standar kelengkapan administrasi dan alur penyerahan berkas penagihan.")

    if penagihan_sections:
        cols_penagihan = st.columns(len(penagihan_sections))
        for idx, sec in enumerate(penagihan_sections):
            with cols_penagihan[idx]:
                st.markdown(f"""
                <div class="card-box">
                    <span style="background-color: {theme.PRIMARY}; color: white; padding: 3px 10px; border-radius: 12px; font-weight: bold; font-size: 0.8rem;">
                        BAGIAN {sec.get('letter', '')}
                    </span>
                    <h4 style="margin-top: 10px; color: {theme.PRIMARY};">{sec.get('title', '')}</h4>
                </div>
                """, unsafe_allow_html=True)

                for item_idx, item in enumerate(sec.get("items", [])):
                    st.checkbox(
                        label=item,
                        key=f"chk_{idx}_{item_idx}",
                        value=False
                    )
    else:
        st.info("Tidak ada data alur penagihan yang ditemukan.")

# =========================================================
# TAB 6: KELOLA DATA
# =========================================================
with tab6:
    st.markdown('<div class="section-title">Kelola & Sunting Data</div>', unsafe_allow_html=True)
    st.info(
        "Perubahan data akan disimpan langsung ke file Excel data/Data.xlsx. "
        "Salinan cadangan (backup) otomatis dibuat sebelum penyimpanan. "
        "Pastikan file Excel ditutup di aplikasi lain saat menyimpan."
    )

    kd_sub1, kd_sub2, kd_sub3, kd_sub4, kd_sub5, kd_sub6 = st.tabs([
        "Monitoring",
        "Tambah Pekerjaan",
        "Akumulatif",
        "Potensi",
        "Target",
        "Cadangan"
    ])

    # -----------------------------------------------------
    # SUB-TAB 1: Monitoring Editor
    # -----------------------------------------------------
    with kd_sub1:
        st.subheader("Edit Data Monitoring Pekerjaan")

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            show_fin_cols = st.checkbox("Tampilkan kolom keuangan", value=False, key="chk_show_fin")
        with col_t2:
            show_doc_cols = st.checkbox("Tampilkan kolom dokumen", value=False, key="chk_show_doc")

        base_cols_ed = [
            "NO", "Judul Pekerjaan", "User", "Proses Kontrak", "Eksekutor",
            "Kontrak IPS", "Nilai Kontrak Sebelum PPN (IPS)", "Nilai Kontrak Setelah PPN (IPS)",
            "Mulai Pekerjaan", "Berakhir Pekerjaan", "Nama Vendor", "Kendala",
            "STATUS", "STATUS PEMBAYARAN"
        ]
        fin_cols_ed = [
            "IH Sebelum PPN", "IH Setelah PPN", "Nilai Invoice Sebelum PPN",
            "Nilai Invoice Setelah PPN", "DENDA", "PEMBAYARAN PIPS",
            "Nilai Kontrak Mitra Sebelum PPN", "Nilai Kontrak Mitra Setelah PPN",
            "Nilai BA Vendor Setelah PPN", "DENDA2", "PEMBAYARAN MITRA", "MARGIN"
        ]
        doc_cols_ed = [
            "Projeck", "DMR/TOR/Spesifikasi", "LOI", "Nomor Informasi Harga Mitra",
            "Undangan PL", "Addendum Kontrak", "Ket Addendum", "BAST/TUG",
            "Invoice", "Nomor Kontrak Vendor", "BA Vendor", "Periode Penagihan"
        ]

        active_ed_cols = [c for c in base_cols_ed if c in df_monitoring.columns]
        if show_fin_cols:
            for c in fin_cols_ed:
                if c in df_monitoring.columns and c not in active_ed_cols:
                    active_ed_cols.append(c)
        if show_doc_cols:
            for c in doc_cols_ed:
                if c in df_monitoring.columns and c not in active_ed_cols:
                    active_ed_cols.append(c)

        editor_mon_df = df_monitoring[active_ed_cols].copy()

        # Build column configs
        ed_col_cfg: Dict[str, Any] = {
            "NO": st.column_config.Column("NO", disabled=True, width="small"),
            "User": st.column_config.SelectboxColumn(
                "User",
                options=sorted(set(_unique_options("User")) | {"Admin"})
            ),
            "Proses Kontrak": st.column_config.SelectboxColumn(
                "Proses Kontrak",
                options=_unique_options("Proses Kontrak") or ["SPK", "PJ", "TERKONTRAK IP", "TIDAK TERKONTRAK IP"]
            ),
            "Eksekutor": st.column_config.SelectboxColumn(
                "Eksekutor",
                options=_unique_options("Eksekutor") or ["RENDAL HAR", "OPERASI", "LOGISTIK", "K3L"]
            ),
            "STATUS": st.column_config.SelectboxColumn(
                "STATUS",
                options=["SELESAI", "PROSES", "CANCEL", "BELUM ADA STATUS"]
            ),
            "STATUS PEMBAYARAN": st.column_config.SelectboxColumn(
                "STATUS PEMBAYARAN",
                options=["PAYMENT", "PROSES", "PROSES PAYMENT", "MENUNGGU", "CANCEL", "BELUM ADA STATUS"]
            ),
            "Kendala": st.column_config.TextColumn("Kendala", width="large")
        }
        for c in rp_columns:
            if c in active_ed_cols:
                # "Rp %d" produces no thousands grouping (printf has none); "localized"
                # gives grouped digits for editing. Label carries the "(Rp)" unit instead.
                ed_col_cfg[c] = st.column_config.NumberColumn(f"{c} (Rp)", format="localized")

        edited_mon_result = st.data_editor(
            editor_mon_df,
            num_rows="dynamic",
            key="ed_monitoring",
            hide_index=True,
            column_config=ed_col_cfg,
            height=450,
            width='stretch'
        )

        col_b1, col_b2 = st.columns([1, 1])
        with col_b1:
            if st.button("Simpan ke Excel", type="primary", key="btn_save_mon"):
                curr_m = get_file_mtime(EXCEL_PATH)
                orig_m = data.get("file_mtime", 0.0)
                if abs(curr_m - orig_m) > 1e-4:
                    st.error("File Excel berubah sejak dimuat, klik Muat Ulang Data dulu")
                else:
                    try:
                        # Overlay edited visible columns onto the full df
                        full_save_cols = [c for c in df_monitoring.columns if not c.endswith("(Asli)")]
                        merged_rows = []
                        # data_editor keeps the original index labels for existing rows
                        # (even after deletions) and assigns new labels to added rows,
                        # so match on label, never on position.
                        for idx, ed_row in edited_mon_result.iterrows():
                            if idx in df_monitoring.index:
                                row_dict = df_monitoring.loc[idx].to_dict()
                            else:
                                row_dict = {c: None for c in full_save_cols}
                            for c in edited_mon_result.columns:
                                row_dict[c] = ed_row[c]
                            merged_rows.append(row_dict)

                        save_df = pd.DataFrame(merged_rows, columns=full_save_cols)
                        backup_path = excel_writer.write_monitoring(EXCEL_PATH, save_df)
                        st.session_state["_flash"] = f"Data monitoring berhasil disimpan! Backup: `{backup_path}`"
                        st.cache_data.clear()
                        st.rerun()
                    except excel_writer.ExcelLocked:
                        st.error("Tutup file Excel terlebih dahulu")

        with col_b2:
            if st.button("Batalkan perubahan", key="btn_cancel_mon"):
                st.session_state.pop("ed_monitoring", None)
                st.rerun()

    # -----------------------------------------------------
    # SUB-TAB 2: Tambah Pekerjaan Form
    # -----------------------------------------------------
    with kd_sub2:
        st.subheader("Formulir Tambah Pekerjaan Baru")

        users_list = sorted(set(_unique_options("User")) | {"Admin"})
        proses_list = _unique_options("Proses Kontrak") or ["SPK", "PJ", "TERKONTRAK IP", "TIDAK TERKONTRAK IP"]
        eksekutor_list = _unique_options("Eksekutor") or ["RENDAL HAR", "OPERASI", "LOGISTIK", "K3L"]
        vendors_list = _unique_options("Nama Vendor") + ["Lainnya"]

        with st.form("form_add_job"):
            form_judul = st.text_input("Judul Pekerjaan *", placeholder="Masukkan judul pekerjaan...")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                form_projeck = st.text_input("Projeck", value="IPS")
                form_user = st.selectbox("User / Bidang", options=users_list)
                form_proses = st.selectbox("Proses Kontrak", options=proses_list)
                form_eksekutor = st.selectbox("Eksekutor", options=eksekutor_list)
                form_kontrak_ips = st.text_input("Nomor Kontrak IPS", placeholder="Contoh: 0001.PJ/...")
            with col_f2:
                form_nilai_sebelum = st.number_input("Nilai Kontrak Sebelum PPN (IPS)", min_value=0.0, step=1000000.0)
                auto_ppn = st.checkbox("Hitung PPN 11% otomatis", value=True)
                if auto_ppn:
                    form_nilai_setelah = form_nilai_sebelum * 1.11
                    st.text(f"Nilai Setelah PPN: {fmt_rp(form_nilai_setelah)}")
                else:
                    form_nilai_setelah = st.number_input("Nilai Kontrak Setelah PPN (IPS)", min_value=0.0, step=1000000.0)

                col_dt1, col_dt2 = st.columns(2)
                with col_dt1:
                    form_mulai_date = st.date_input("Mulai Pekerjaan", value=None)
                with col_dt2:
                    form_berakhir_date = st.date_input("Berakhir Pekerjaan", value=None)

            col_f3, col_f4 = st.columns(2)
            with col_f3:
                sel_vendor_opt = st.selectbox("Nama Vendor", options=vendors_list)
                if sel_vendor_opt == "Lainnya":
                    custom_vendor = st.text_input("Ketik Nama Vendor Baru", "")
                    form_vendor = custom_vendor.strip()
                else:
                    form_vendor = sel_vendor_opt
                form_status = st.selectbox("Status Pekerjaan", options=["PROSES", "SELESAI", "CANCEL", "BELUM ADA STATUS"])
                form_bayar = st.selectbox("Status Pembayaran", options=["PROSES", "PAYMENT", "PROSES PAYMENT", "MENUNGGU", "CANCEL", "BELUM ADA STATUS"])
            with col_f4:
                form_kendala = st.text_area("Catatan / Kendala Pekerjaan", placeholder="Catatan atau kendala...", height=115)

            btn_submit_job = st.form_submit_button("Tambah & Simpan Pekerjaan", type="primary")

        if btn_submit_job:
            if not form_judul.strip():
                st.error("Judul Pekerjaan wajib diisi!")
            else:
                curr_m = get_file_mtime(EXCEL_PATH)
                orig_m = data.get("file_mtime", 0.0)
                if abs(curr_m - orig_m) > 1e-4:
                    st.error("File Excel berubah sejak dimuat, klik Muat Ulang Data dulu")
                else:
                    try:
                        # Format dates to Indonesian string e.g. "01 Januari 2026"
                        def _format_date_entry(d_val: Optional[date]) -> str:
                            if not d_val:
                                return ""
                            mon_id = [
                                "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
                                "Juli", "Agustus", "September", "Oktober", "November", "Desember"
                            ]
                            return f"{d_val.day:02d} {mon_id[d_val.month]} {d_val.year}"

                        mulai_str = _format_date_entry(form_mulai_date)
                        berakhir_str = _format_date_entry(form_berakhir_date)

                        new_row = {
                            "NO": len(df_monitoring) + 1,
                            "Judul Pekerjaan": form_judul.strip(),
                            "Projeck": form_projeck.strip(),
                            "User": form_user,
                            "Proses Kontrak": form_proses,
                            "Eksekutor": form_eksekutor,
                            "Kontrak IPS": form_kontrak_ips.strip(),
                            "Nilai Kontrak Sebelum PPN (IPS)": form_nilai_sebelum,
                            "Nilai Kontrak Setelah PPN (IPS)": form_nilai_setelah,
                            "Mulai Pekerjaan": mulai_str,
                            "Berakhir Pekerjaan": berakhir_str,
                            "Nama Vendor": form_vendor,
                            "Kendala": form_kendala.strip(),
                            "STATUS": form_status,
                            "STATUS PEMBAYARAN": form_bayar,
                        }

                        clean_cols = [c for c in df_monitoring.columns if not c.endswith("(Asli)")]
                        full_updated_df = pd.concat(
                            [df_monitoring[clean_cols], pd.DataFrame([new_row])],
                            ignore_index=True
                        )
                        b_path = excel_writer.write_monitoring(EXCEL_PATH, full_updated_df)
                        st.session_state["_flash"] = f"Pekerjaan baru berhasil ditambahkan dan disimpan! Backup: `{b_path}`"
                        st.cache_data.clear()
                        st.rerun()
                    except excel_writer.ExcelLocked:
                        st.error("Tutup file Excel terlebih dahulu")

    # -----------------------------------------------------
    # SUB-TAB 3: Akumulatif Editor
    # -----------------------------------------------------
    with kd_sub3:
        st.subheader("Edit Data Monitoring Akumulatif")

        ak_ed_cfg = {
            "NO": st.column_config.Column("NO", disabled=True, width="small"),
            "Nilai Kontrak Setelah PPN": st.column_config.NumberColumn("Nilai Kontrak Setelah PPN (Rp)", format="localized"),
            "Nilai Kontrak Mitra Setelah PPN": st.column_config.NumberColumn("Nilai Kontrak Mitra Setelah PPN (Rp)", format="localized"),
            "STATUS": st.column_config.SelectboxColumn("STATUS", options=["SELESAI", "PROSES", "CANCEL", "BELUM ADA STATUS"])
        }

        edited_ak_result = st.data_editor(
            df_monitoring_ak,
            num_rows="dynamic",
            key="ed_akumulatif",
            hide_index=True,
            column_config=ak_ed_cfg,
            width='stretch'
        )

        if st.button("Simpan Akumulatif ke Excel", type="primary", key="btn_save_ak"):
            curr_m = get_file_mtime(EXCEL_PATH)
            orig_m = data.get("file_mtime", 0.0)
            if abs(curr_m - orig_m) > 1e-4:
                st.error("File Excel berubah sejak dimuat, klik Muat Ulang Data dulu")
            else:
                try:
                    b_path = excel_writer.write_monitoring_akumulatif(EXCEL_PATH, edited_ak_result)
                    st.session_state["_flash"] = f"Data akumulatif berhasil disimpan! Backup: `{b_path}`"
                    st.cache_data.clear()
                    st.rerun()
                except excel_writer.ExcelLocked:
                    st.error("Tutup file Excel terlebih dahulu")

    # -----------------------------------------------------
    # SUB-TAB 4: Potensi Editor
    # -----------------------------------------------------
    with kd_sub4:
        st.subheader("Edit Data Potensi Pendapatan")

        pot_ed_cfg = {
            "JUDUL": st.column_config.Column("Judul Pekerjaan", width="large"),
            "POTENSI NILAI": st.column_config.NumberColumn("Potensi Nilai (Rp)", format="localized"),
            "PERIODE": st.column_config.Column("Periode", width="small")
        }

        edited_pot_result = st.data_editor(
            df_potensi,
            num_rows="dynamic",
            key="ed_potensi",
            hide_index=True,
            column_config=pot_ed_cfg,
            width='stretch'
        )

        if st.button("Simpan Potensi ke Excel", type="primary", key="btn_save_pot"):
            curr_m = get_file_mtime(EXCEL_PATH)
            orig_m = data.get("file_mtime", 0.0)
            if abs(curr_m - orig_m) > 1e-4:
                st.error("File Excel berubah sejak dimuat, klik Muat Ulang Data dulu")
            else:
                try:
                    b_path = excel_writer.write_potensi(EXCEL_PATH, edited_pot_result)
                    st.session_state["_flash"] = f"Data potensi berhasil disimpan! Backup: `{b_path}`"
                    st.cache_data.clear()
                    st.rerun()
                except excel_writer.ExcelLocked:
                    st.error("Tutup file Excel terlebih dahulu")

    # -----------------------------------------------------
    # SUB-TAB 5: Target Editor
    # -----------------------------------------------------
    with kd_sub5:
        st.subheader("Edit Anggaran & Target Kinerja")

        col_tp, col_ta = st.columns(2)

        with col_tp:
            st.markdown("#### Target Kinerja 2026 (Sheet: PENCAPAIAN)")
            in_tgt_tahun = st.number_input(
                "Target Tahun 2026",
                value=float(pencapaian_kpi.get("target_tahun", 0.0)),
                step=100000000.0,
                format="%.0f"
            )
            st.caption(f"Preview: **{fmt_rp(in_tgt_tahun)}**")

            in_tgt_s1 = st.number_input(
                "Target Semester 1",
                value=float(pencapaian_kpi.get("target_s1", 0.0)),
                step=100000000.0,
                format="%.0f"
            )
            st.caption(f"Preview: **{fmt_rp(in_tgt_s1)}**")

            in_tgt_s2 = st.number_input(
                "Target Semester 2",
                value=float(pencapaian_kpi.get("target_s2", 0.0)),
                step=100000000.0,
                format="%.0f"
            )
            st.caption(f"Preview: **{fmt_rp(in_tgt_s2)}**")

        with col_ta:
            st.markdown("#### Target Kinerja Akumulatif (Sheet: Grafik Akumulatif)")
            in_ak_tahun = st.number_input(
                "Target Akumulatif Tahun",
                value=float(akumulatif_kpi.get("target_tahun", 0.0)),
                step=100000000.0,
                format="%.0f"
            )
            st.caption(f"Preview: **{fmt_rp(in_ak_tahun)}**")

            in_ak_s1 = st.number_input(
                "Target Akumulatif Semester 1",
                value=float(akumulatif_kpi.get("target_s1", 0.0)),
                step=100000000.0,
                format="%.0f"
            )
            st.caption(f"Preview: **{fmt_rp(in_ak_s1)}**")

            in_ak_cap_s1 = st.number_input(
                "Capaian Akumulatif Semester 1 (Konstanta)",
                value=float(akumulatif_kpi.get("capaian_s1", 0.0)),
                step=100000000.0,
                format="%.0f"
            )
            st.caption(f"Preview: **{fmt_rp(in_ak_cap_s1)}**")

            in_ak_s2 = st.number_input(
                "Target Akumulatif Semester 2",
                value=float(akumulatif_kpi.get("target_s2", 0.0)),
                step=100000000.0,
                format="%.0f"
            )
            st.caption(f"Preview: **{fmt_rp(in_ak_s2)}**")

        if st.button("Simpan Target ke Excel", type="primary", key="btn_save_tgt"):
            curr_m = get_file_mtime(EXCEL_PATH)
            orig_m = data.get("file_mtime", 0.0)
            if abs(curr_m - orig_m) > 1e-4:
                st.error("File Excel berubah sejak dimuat, klik Muat Ulang Data dulu")
            else:
                try:
                    penc_dict = {
                        "target_tahun": in_tgt_tahun,
                        "target_s1": in_tgt_s1,
                        "target_s2": in_tgt_s2
                    }
                    akum_dict = {
                        "target_tahun": in_ak_tahun,
                        "target_s1": in_ak_s1,
                        "capaian_s1": in_ak_cap_s1,
                        "target_s2": in_ak_s2
                    }
                    b_path = excel_writer.write_targets(EXCEL_PATH, penc_dict, akum_dict)
                    st.session_state["_flash"] = f"Target kinerja berhasil disimpan! Backup: `{b_path}`"
                    st.cache_data.clear()
                    st.rerun()
                except excel_writer.ExcelLocked:
                    st.error("Tutup file Excel terlebih dahulu")

    # -----------------------------------------------------
    # SUB-TAB 6: Cadangan (Backup Manager)
    # -----------------------------------------------------
    with kd_sub6:
        st.subheader("Manajemen File Cadangan (Backup)")
        st.write("Daftar salinan cadangan yang dibuat secara otomatis saat penyimpanan:")

        backup_dir = os.path.join("data", "backup")
        if os.path.exists(backup_dir):
            backup_files = [
                f for f in os.listdir(backup_dir)
                if f.endswith(".xlsx") and os.path.isfile(os.path.join(backup_dir, f))
            ]
            backup_files.sort(
                key=lambda f: os.path.getmtime(os.path.join(backup_dir, f)),
                reverse=True
            )
        else:
            backup_files = []

        if not backup_files:
            st.info("Belum ada file cadangan yang tersedia.")
        else:
            confirm_restore = st.checkbox("Saya yakin ingin memulihkan cadangan ini", value=False)

            for b_file in backup_files:
                b_full_path = os.path.join(backup_dir, b_file)
                b_size_kb = os.path.getsize(b_full_path) / 1024.0
                b_mtime_str = datetime.fromtimestamp(os.path.getmtime(b_full_path)).strftime("%d-%m-%Y %H:%M:%S")

                col_bk1, col_bk2 = st.columns([4, 1])
                with col_bk1:
                    st.markdown(f"**`{b_file}`** ({b_size_kb:.1f} KB) - Dibuat: {b_mtime_str}")
                with col_bk2:
                    if st.button("Pulihkan", key=f"btn_res_{b_file}", disabled=not confirm_restore):
                        try:
                            # Make a backup of current file before overwriting
                            excel_writer.backup_excel(EXCEL_PATH)
                            shutil.copy2(b_full_path, EXCEL_PATH)
                            st.cache_data.clear()
                            st.session_state["_flash"] = f"File `{b_file}` berhasil dipulihkan sebagai data aktif!"
                            st.rerun()
                        except PermissionError:
                            st.error("Tutup file Excel terlebih dahulu")

# =========================================================
# TAB 7: EXPORT PPT
# =========================================================
with tab7:
    st.markdown('<div class="section-title">Ekspor Laporan PowerPoint (.PPTX)</div>', unsafe_allow_html=True)
    st.write(
        "Menyusun visualisasi, indikator kinerja utama, dan tabel monitoring ke dalam presentasi "
        "PowerPoint berformat layar lebar 16:9 berstandar visual PLN / IPS."
    )

    st.markdown("#### Slide yang Dihasilkan:")
    slides_info = [
        "Slide 1: Judul Laporan & Identitas Organisasi (PLN IPS)",
        "Slide 2: Ringkasan Indikator Kinerja Utama (KPI Cards)",
        "Slide 3: Target vs Realisasi 2026 (Clustered Bar Chart)",
        "Slide 4: Kinerja Akumulatif (Clustered Bar Chart)",
        "Slide 5: Distribusi Status Pekerjaan & Status Pembayaran (Pie Charts)",
        "Slide 6: Top 10 Nilai Kontrak Pekerjaan (Horizontal Bar Chart)",
        "Slide Jadwal: Timeline Jadwal Pekerjaan (Gantt Chart - Opsional)",
        "Slide Kendala: Daftar Kendala & Hambatan Utama Pekerjaan",
        "Slide Monitoring: Tabel Detail Monitoring Pekerjaan (10 baris per slide)",
        "Slide Akumulatif: Tabel Monitoring Akumulatif beserta Total",
        "Slide Potensi: Tabel Potensi Pendapatan beserta Total",
        "Slide Penagihan: Alur & Checklist Dokumen Penagihan"
    ]
    for s_info in slides_info:
        st.markdown(f"- {s_info}")

    st.markdown("---")

    inc_gantt = st.checkbox(
        "Sertakan slide Jadwal (Gantt)",
        value=False,
        help="Menambahkan visualisasi timeline jadwal pekerjaan jika renderer gambar tersedia."
    )

    if st.button("Generate PowerPoint", type="primary", key="btn_gen_ppt"):
        with st.spinner("Menyusun file presentasi PowerPoint..."):
            gantt_png_bytes: Optional[bytes] = None

            if inc_gantt and fig_gantt is not None:
                try:
                    import kaleido  # Check if kaleido is available
                    gantt_png_bytes = fig_gantt.to_image(format="png")
                except Exception:
                    gantt_png_bytes = None

            pptx_bytes = ppt_export.build_pptx(data, filtered_df, gantt_png=gantt_png_bytes)
            file_name = f"Dashboard_Monitoring_PLN_IPS_{datetime.now().strftime('%Y%m%d')}.pptx"

            st.success("File presentasi berhasil dibuat!")
            st.download_button(
                label="Download .pptx",
                data=pptx_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary"
            )

# =========================================================
# TAB 8: EXPORT EXCEL
# =========================================================
with tab8:
    st.markdown('<div class="section-title">Ekspor Data Excel (.XLSX)</div>', unsafe_allow_html=True)
    st.write(
        "Mengunduh file Excel sumber data (`Data.xlsx`) apa adanya, persis dengan desain aslinya "
        "(format, warna, sheet, dan formula) — sesuai kondisi data terakhir yang tersimpan."
    )

    st.markdown("---")

    if os.path.exists(EXCEL_PATH):
        with open(EXCEL_PATH, "rb") as f:
            xlsx_bytes = f.read()

        file_name = f"Data_PLN_IPS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        st.download_button(
            label="Download Data.xlsx",
            data=xlsx_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.error(f"File sumber data tidak ditemukan: `{EXCEL_PATH}`")
