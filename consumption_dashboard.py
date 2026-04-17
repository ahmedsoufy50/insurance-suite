import re

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from consumption_pdf_export import build_consumption_pptx
from database import (
    CONSUMPTION_MASTER_FIELDS,
    DASHBOARD_FIELDS,
    get_all_consumption_mappings,
    get_consumption_mapping,
)


C_ACC = "#2D3A52"
C_COR = "#E8624A"
C_GREEN = "#2D7A4F"
C_PALETTE = [C_COR, C_ACC, "#A0B0C8", "#F4A990", "#8B9DB8", "#F07A65", "#5A7090"]


FALLBACK_CANDIDATES = {
    "member_name": ["Member Name", "Member", "Patient Name", "Insured Name", "Name"],
    "service_date": ["Service Date", "Date", "Visit Date", "Claim Date"],
    "net_amount": ["Net Payable Amount", "Net Payable", "Net Amount", "Amount", "Paid Amount", "Total Cost"],
    "diagnosis": ["Diagnosis", "Diag", "ICD"],
    "provider": ["Billing Provider", "Provider", "Hospital", "Clinic"],
    "category": ["Category", "Service Type", "Category / Service Type", "Benefit"],
    "card_no": ["Card No", "Card Number", "Member ID"],
    "provider_type": ["Provider Type", "Billing Provider Type", "Type Local"],
    "incident_type": ["Incident Type", "Inpatient / Outpatient", "Incident"],
    "chronic_flag": ["Chronic Flag", "Chronic"],
    "relation": ["Relation", "Relationship", "Dependent"],
    "plan": ["Plan", "Policy Plan"],
    "submission": ["Submission", "Claim Frequency", "Submission Method"],
}


def _norm(value):
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def _chart_layout(fig, height=320):
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family="sans-serif", size=12),
        margin=dict(l=0, r=0, t=32, b=0),
        height=height,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


def _kpi(col, label, value, fmt="num", delta=None, delta_color="normal"):
    with col:
        if fmt == "egp":
            st.metric(label, f"EGP {value:,.0f}", delta=delta, delta_color=delta_color)
        else:
            st.metric(label, f"{value:,.0f}", delta=delta, delta_color=delta_color)


def _resolve_column(df_columns, saved_name="", fallback_candidates=None):
    fallback_candidates = fallback_candidates or []
    normalized = {_norm(col): col for col in df_columns}

    if saved_name and _norm(saved_name) in normalized:
        return normalized[_norm(saved_name)]

    for candidate in fallback_candidates:
        if _norm(candidate) in normalized:
            return normalized[_norm(candidate)]

    for real_col in df_columns:
        real_norm = _norm(real_col)
        for candidate in fallback_candidates:
            cand_norm = _norm(candidate)
            if cand_norm in real_norm or real_norm in cand_norm:
                return real_col

    return None


def _build_mapping(df_columns, db_map):
    resolved = {}
    for field_key, _ in DASHBOARD_FIELDS:
        saved_name = db_map.get(field_key, "") if db_map else ""
        resolved[field_key] = _resolve_column(
            df_columns,
            saved_name=saved_name,
            fallback_candidates=FALLBACK_CANDIDATES.get(field_key, []),
        )
    return resolved


def _render_mapping_status(resolved_columns, selected_insurer):
    rows = []
    for field_key, field_label in CONSUMPTION_MASTER_FIELDS:
        rows.append(
            {
                "Field": field_label,
                "Mapped Header": resolved_columns.get(field_key) or "Missing",
                "Status": "OK" if resolved_columns.get(field_key) else "Missing",
            }
        )
    with st.expander(f"Mapped Headers for {selected_insurer}", expanded=False):
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def _unique_preserve_order(values):
    result = []
    seen = set()
    for value in values:
        if not value or value in seen:
            continue
        seen.add(value)
        result.append(value)
    return result


def render_consumption_dashboard():
    st.markdown("### Consumption Dashboard")
    st.markdown("---")

    mappings = get_all_consumption_mappings()
    insurer_names = [item["insurer_name"] for item in mappings]
    if not insurer_names:
        st.warning("No insurer mappings configured yet. Open the admin settings and add a company first.")
        return

    left, right = st.columns([1.2, 1])
    with left:
        selected_insurer = st.selectbox("Insurance Company", ["Select Company"] + insurer_names, key="dash_insurer")
    with right:
        uploaded_file = st.file_uploader("Upload Consumption File (Excel)", type=["xlsx", "xls"], key="dash_file")

    info_col, ibnr_col, premium_col = st.columns([1.2, 1, 1])
    with info_col:
        st.caption("Headers are resolved using the mapping saved for the selected insurer.")
    with ibnr_col:
        ibnr_pct = st.number_input("IBNR %", min_value=0.0, max_value=50.0, value=10.0, step=1.0)
    with premium_col:
        annual_premium = st.number_input("Annual Premium (EGP)", min_value=0.0, value=245534.0, step=1000.0)

    if selected_insurer == "Select Company":
        st.info("Select an insurance company first.")
        return
    if not uploaded_file:
        st.info("Upload a consumption file to start the analysis.")
        return

    try:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file)
    except Exception as exc:
        st.error(f"Could not read the uploaded file: {exc}")
        return

    df.columns = [str(col).strip() for col in df.columns]
    db_map = get_consumption_mapping(selected_insurer) or {}
    resolved = _build_mapping(df.columns, db_map)
    _render_mapping_status(resolved, selected_insurer)

    missing_required = [label for field_key, label in CONSUMPTION_MASTER_FIELDS if not resolved.get(field_key)]
    if missing_required:
        st.error("Missing required mapped headers: " + ", ".join(missing_required))
        with st.expander("Available Headers", expanded=False):
            st.write(df.columns.tolist())
        return

    col_member = resolved["member_name"]
    col_date = resolved["service_date"]
    col_amount = resolved["net_amount"]
    col_diag = resolved["diagnosis"]
    col_provider = resolved["provider"]
    col_category = resolved["category"]
    col_card = resolved.get("card_no")
    col_ptype = resolved.get("provider_type")
    col_incident = resolved.get("incident_type")
    col_chronic = resolved.get("chronic_flag")
    col_relation = resolved.get("relation")
    col_plan = resolved.get("plan")
    col_submission = resolved.get("submission")

    df["_amt"] = pd.to_numeric(df[col_amount], errors="coerce").fillna(0)
    df["_date"] = pd.to_datetime(df[col_date], errors="coerce")

    with st.expander("Filters", expanded=False):
        filter_cols = st.columns(4)
        with filter_cols[0]:
            provider_options = ["All"] + sorted(df[col_provider].dropna().astype(str).unique().tolist())
            provider_filter = st.selectbox("Billing Provider", provider_options)
        with filter_cols[1]:
            category_options = ["All"] + sorted(df[col_category].dropna().astype(str).unique().tolist())
            category_filter = st.selectbox("Service Category", category_options)
        with filter_cols[2]:
            diag_options = ["All"] + sorted(df[col_diag].dropna().astype(str).unique().tolist())
            diagnosis_filter = st.selectbox("Diagnosis", diag_options)
        with filter_cols[3]:
            valid_dates = df["_date"].dropna()
            if valid_dates.empty:
                date_range = None
                st.caption("No valid service dates found.")
            else:
                min_date = valid_dates.min().date()
                max_date = valid_dates.max().date()
                date_range = st.date_input("Service Date Range", value=(min_date, max_date), min_value=min_date, max_value=max_date)

    dff = df.copy()
    if provider_filter != "All":
        dff = dff[dff[col_provider].astype(str) == provider_filter]
    if category_filter != "All":
        dff = dff[dff[col_category].astype(str) == category_filter]
    if diagnosis_filter != "All":
        dff = dff[dff[col_diag].astype(str) == diagnosis_filter]
    if date_range and len(date_range) == 2:
        dff = dff[(dff["_date"].dt.date >= date_range[0]) & (dff["_date"].dt.date <= date_range[1])]

    total_paid = dff["_amt"].sum()
    total_claims = len(dff)
    active_members = dff[col_member].nunique()
    avg_per_member = total_paid / active_members if active_members else 0
    loss_ratio = (total_paid * (1 + ibnr_pct / 100) / annual_premium * 100) if annual_premium else 0
    if loss_ratio > 100:
        lr_indicator = "RED"
        lr_color = "#E8624A"
        lr_bg = "#FFEBEE"
    elif loss_ratio >= 80:
        lr_indicator = "YELLOW"
        lr_color = "#B58900"
        lr_bg = "#FFF8DB"
    else:
        lr_indicator = "GREEN"
        lr_color = "#2D7A4F"
        lr_bg = "#E8F5E9"

    st.markdown("### KPIs")
    k1, k2, k3, k4, k5 = st.columns(5)
    _kpi(k1, "Total Consumption", total_paid, "egp")
    _kpi(k2, "Claims Count", total_claims)
    _kpi(k3, "Active Members", active_members)
    _kpi(k4, "Avg / Member", avg_per_member, "egp")
    with k5:
        st.markdown(
            f"""
            <div style="
                padding: 0.85rem 0.9rem;
                border-radius: 0.8rem;
                background: {lr_bg};
                border: 1px solid {lr_color}33;
                min-height: 118px;
            ">
                <div style="font-size: 0.85rem; color: #475569; margin-bottom: 0.35rem;">
                    Loss Ratio ({lr_indicator})
                </div>
                <div style="font-size: 1.9rem; font-weight: 700; color: {lr_color}; line-height: 1.1;">
                    {loss_ratio:.1f}%
                </div>
                <div style="font-size: 0.8rem; color: #64748B; margin-top: 0.35rem;">
                    IBNR {ibnr_pct:.0f}%
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    top_members = (
        dff.groupby(col_member)["_amt"].sum().reset_index().sort_values("_amt", ascending=False).head(10)
    )
    top_members.columns = ["Member", "Amount"]

    category_mix = (
        dff.groupby(col_category)["_amt"].sum().reset_index().sort_values("_amt", ascending=False).head(8)
    )
    category_mix.columns = ["Category", "Amount"]

    row1_left, row1_right = st.columns([1.25, 1])
    with row1_left:
        st.markdown("##### Top Members by Amount")
        fig_members = px.bar(top_members, x="Amount", y="Member", orientation="h", color_discrete_sequence=[C_COR], text="Amount")
        fig_members.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig_members.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(_chart_layout(fig_members, 360), use_container_width=True)
    with row1_right:
        st.markdown("##### Category Distribution")
        fig_category = px.pie(category_mix, values="Amount", names="Category", color_discrete_sequence=C_PALETTE, hole=0.45)
        fig_category.update_traces(textposition="outside", textinfo="percent+label")
        st.plotly_chart(_chart_layout(fig_category, 360), use_container_width=True)

    st.markdown("---")

    row2_left, row2_right = st.columns(2)
    top_providers = (
        dff.groupby(col_provider)["_amt"].sum().reset_index().sort_values("_amt", ascending=False).head(10)
    )
    top_providers.columns = ["Provider", "Amount"]

    top_diagnoses = (
        dff.groupby(col_diag)["_amt"].agg(Amount="sum", Claims="count").reset_index().sort_values("Claims", ascending=False).head(10)
    )
    top_diagnoses.columns = ["Diagnosis", "Amount", "Claims"]

    with row2_left:
        st.markdown("##### Top Billing Providers")
        fig_providers = px.bar(top_providers, x="Amount", y="Provider", orientation="h", color_discrete_sequence=[C_ACC], text="Amount")
        fig_providers.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig_providers.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(_chart_layout(fig_providers, 360), use_container_width=True)

    with row2_right:
        st.markdown("##### Top Diagnoses")
        fig_diag = px.bar(top_diagnoses, x="Claims", y="Diagnosis", orientation="h", color_discrete_sequence=[C_COR], text="Claims")
        fig_diag.update_traces(texttemplate="%{text}", textposition="outside")
        fig_diag.update_layout(yaxis=dict(autorange="reversed"))
        st.plotly_chart(_chart_layout(fig_diag, 360), use_container_width=True)

    st.markdown("---")

    row3_left, row3_right = st.columns(2)
    monthly_source = dff.dropna(subset=["_date"]).copy()
    monthly_source["_month"] = monthly_source["_date"].dt.to_period("M").astype(str)
    monthly = monthly_source.groupby("_month")["_amt"].sum().reset_index()
    monthly.columns = ["Month", "Amount"]

    with row3_left:
        st.markdown("##### Monthly Consumption Trend")
        fig_monthly = px.line(monthly, x="Month", y="Amount", markers=True, color_discrete_sequence=[C_COR])
        fig_monthly.update_traces(fill="tozeroy", fillcolor="rgba(232,98,74,0.1)")
        st.plotly_chart(_chart_layout(fig_monthly, 310), use_container_width=True)

    right_title = "Service Category Spend"
    right_df = dff.groupby(col_category)["_amt"].sum().reset_index().sort_values("_amt", ascending=False).head(10)
    right_df.columns = ["Label", "Amount"]
    if col_ptype:
        right_df = dff.groupby(col_ptype)["_amt"].sum().reset_index().sort_values("_amt", ascending=False)
        right_df.columns = ["Label", "Amount"]
        right_title = "Provider Type Mix"

    with row3_right:
        st.markdown(f"##### {right_title}")
        fig_mix = px.bar(right_df, x="Label", y="Amount", color="Label", color_discrete_sequence=C_PALETTE, text="Amount")
        fig_mix.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        st.plotly_chart(_chart_layout(fig_mix, 310), use_container_width=True)

    st.markdown("---")

    row4_left, row4_right = st.columns(2)
    with row4_left:
        if col_incident:
            incident_df = dff.groupby(col_incident)["_amt"].sum().reset_index().sort_values("_amt", ascending=False)
            incident_df.columns = ["Incident Type", "Amount"]
            st.markdown("##### Incident Type Distribution")
            fig_incident = px.pie(incident_df, values="Amount", names="Incident Type", color_discrete_sequence=[C_ACC, C_COR, "#A0B0C8", "#F4A990"], hole=0.4)
            fig_incident.update_traces(textposition="outside", textinfo="percent+label")
            st.plotly_chart(_chart_layout(fig_incident, 300), use_container_width=True)
        elif col_chronic:
            chronic_df = dff.groupby(col_chronic)["_amt"].sum().reset_index().sort_values("_amt", ascending=False)
            chronic_df[col_chronic] = chronic_df[col_chronic].map({"Y": "Chronic", "N": "Non-Chronic"}).fillna(chronic_df[col_chronic])
            chronic_df.columns = ["Chronic Flag", "Amount"]
            st.markdown("##### Chronic vs Non-Chronic")
            fig_chronic = px.bar(chronic_df, x="Chronic Flag", y="Amount", color="Chronic Flag", text="Amount")
            fig_chronic.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            st.plotly_chart(_chart_layout(fig_chronic, 300), use_container_width=True)
        else:
            top_categories = right_df.head(6)
            st.markdown("##### Top Service Categories")
            fig_top_categories = px.bar(top_categories, x="Amount", y="Label", orientation="h", color_discrete_sequence=[C_ACC], text="Amount")
            fig_top_categories.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig_top_categories.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(_chart_layout(fig_top_categories, 300), use_container_width=True)

    with row4_right:
        st.markdown("##### Loss Ratio Gauge")
        fig_gauge = go.Figure(
            go.Indicator(
                mode="gauge+number+delta",
                value=loss_ratio,
                delta={"reference": 100, "increasing": {"color": C_COR}, "decreasing": {"color": C_GREEN}},
                number={"suffix": "%", "font": {"size": 36}},
                gauge={
                    "axis": {"range": [0, 200], "ticksuffix": "%"},
                    "bar": {"color": lr_color},
                    "steps": [
                        {"range": [0, 80], "color": "#E8F5E9"},
                        {"range": [80, 100], "color": "#FFF9C4"},
                        {"range": [100, 200], "color": "#FFEBEE"},
                    ],
                    "threshold": {"line": {"color": "red", "width": 3}, "value": 100},
                },
                title={"text": f"With IBNR {ibnr_pct:.0f}%"},
            )
        )
        fig_gauge.update_layout(paper_bgcolor="rgba(0,0,0,0)", height=300, margin=dict(t=40, b=0, l=20, r=20))
        st.plotly_chart(fig_gauge, use_container_width=True)

    with st.expander("Detailed Data", expanded=False):
        visible_cols = _unique_preserve_order([
            col for col in [col_card, col_member, col_date, col_provider, col_category, col_diag, col_amount, col_ptype, col_incident, col_chronic, col_relation, col_plan, col_submission]
            if col and col in dff.columns
        ])
        st.dataframe(dff[visible_cols].sort_values(col_amount, ascending=False), use_container_width=True, height=420)
        st.caption(f"Rows after filters: {len(dff):,}")

    st.markdown("---")
    st.markdown("### Export Report")

    export_left, export_right = st.columns(2)
    with export_left:
        insurer_input = st.text_input("Insurance Company", value=selected_insurer, key="pptx_insurer")
    with export_right:
        client_input = st.text_input("Client / Policy Holder", placeholder="Example: Alfa Power", key="pptx_client")

    insurer_final = insurer_input.strip() or selected_insurer
    client_final = client_input.strip() or "Client"
    file_name = f"CR_{client_final.replace(' ', '_')}_{insurer_final.replace(' ', '_')}_{pd.Timestamp.now().strftime('%d-%m-%Y')}.pptx"

    if st.button("Generate PPTX Report", type="primary", use_container_width=True):
        try:
            pptx_bytes = build_consumption_pptx(
                df=dff,
                col_member=col_member,
                col_amount=col_amount,
                col_provider=col_provider,
                col_ptype=col_ptype,
                col_diag=col_diag,
                col_date=col_date,
                col_incident=col_incident,
                col_chronic=col_chronic,
                col_card=col_card,
                col_relation=col_relation,
                col_category=col_category,
                annual_premium=annual_premium,
                ibnr_pct=ibnr_pct,
                insurer_name=insurer_final,
                client_name=client_final,
                preparer_name="Ahmed Soufy",
                preparer_title="Corporate Medical Account Manager",
            )
            st.download_button(
                "Download PPTX Report",
                data=pptx_bytes,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Could not generate report: {exc}")
