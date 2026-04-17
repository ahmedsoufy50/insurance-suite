import pandas as pd
import streamlit as st

from database import (
    CONSUMPTION_MASTER_FIELDS,
    OPTIONAL_CONSUMPTION_FIELDS,
    consumption_backend_is_writable,
    consumption_remote_backend_label,
    delete_consumption_mapping,
    get_all_consumption_mappings,
    get_consumption_mapping,
    save_consumption_mapping,
    using_remote_consumption_backend,
)


def _guess_column(field_key, available_columns):
    keywords = {
        "member_name": ["member name", "member", "patient", "insured name", "name"],
        "service_date": ["service date", "service", "date", "visit date"],
        "net_amount": ["net payable amount", "net payable", "net amount", "amount", "paid", "cost", "total cost"],
        "diagnosis": ["diagnosis", "diag", "icd"],
        "provider": ["billing provider", "provider", "hospital", "clinic"],
        "category": ["category", "service type", "benefit", "type of service"],
        "card_no": ["card no", "card", "member id", "policy id"],
        "provider_type": ["provider type", "type local", "provider class"],
        "incident_type": ["incident type", "incident", "inpatient", "outpatient"],
        "chronic_flag": ["chronic flag", "chronic"],
        "relation": ["relation", "dependent", "relationship"],
        "plan": ["plan", "policy plan", "benefit plan"],
        "submission": ["submission", "claim frequency", "claim type", "method"],
    }

    lowered = {str(col).strip().lower(): col for col in available_columns}
    for keyword in keywords.get(field_key, []):
        if keyword in lowered:
            return lowered[keyword]

    for column in available_columns:
        normalized = str(column).strip().lower()
        for keyword in keywords.get(field_key, []):
            if keyword in normalized or normalized in keyword:
                return column
    return ""


def _load_sample_columns(uploaded_file):
    if not uploaded_file:
        return []
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, nrows=0)
    return [str(col).strip() for col in df.columns.tolist()]


def _render_mapping_inputs(existing_map, sample_columns):
    options = [""] + sample_columns
    field_values = {}

    st.markdown("#### Master Consumption Fields")
    st.caption("Write the exact header name from the insurer file, or select it from a sample file.")

    for field_key, field_label in CONSUMPTION_MASTER_FIELDS:
        current_value = existing_map.get(field_key, "") if existing_map else ""
        guessed_value = _guess_column(field_key, sample_columns) if sample_columns else ""
        default_value = current_value or guessed_value

        left, right = st.columns([1.2, 2.3])
        with left:
            st.markdown(f"**{field_label}**")
        with right:
            if sample_columns:
                field_values[field_key] = st.selectbox(
                    field_label,
                    options,
                    index=options.index(default_value) if default_value in options else 0,
                    key=f"req_{field_key}",
                    label_visibility="collapsed",
                )
            else:
                field_values[field_key] = st.text_input(
                    field_label,
                    value=default_value,
                    key=f"req_{field_key}",
                    label_visibility="collapsed",
                    placeholder="Example: Net Payable Amount",
                ).strip()

    with st.expander("Optional Fields", expanded=False):
        st.caption("Optional fields improve charts and export quality if they exist in the file.")
        for field_key, field_label in OPTIONAL_CONSUMPTION_FIELDS:
            current_value = existing_map.get(field_key, "") if existing_map else ""
            guessed_value = _guess_column(field_key, sample_columns) if sample_columns else ""
            default_value = current_value or guessed_value

            if sample_columns:
                field_values[field_key] = st.selectbox(
                    field_label,
                    options,
                    index=options.index(default_value) if default_value in options else 0,
                    key=f"opt_{field_key}",
                )
            else:
                field_values[field_key] = st.text_input(
                    field_label,
                    value=default_value,
                    key=f"opt_{field_key}",
                    placeholder="Optional column name",
                ).strip()

    return field_values


def render_admin():
    server_backend = using_remote_consumption_backend()
    read_only = server_backend and not consumption_backend_is_writable()
    backend_label = consumption_remote_backend_label() or "Remote"

    st.markdown("### Admin Mapping Dashboard")
    st.caption(
        "Manage the data identity of each insurer separately. Mapping is done by header name, not by column order."
    )
    if read_only:
        st.warning(f"{backend_label} backend is active in read-only mode. Add write credentials to enable create, edit, and delete.")
    elif server_backend:
        st.success(f"{backend_label} backend is active. You can add, edit, and delete companies from this screen.")

    mappings = get_all_consumption_mappings()
    insurer_names = [item["insurer_name"] for item in mappings]

    left, right = st.columns([1, 1.8], gap="large")

    with left:
        st.markdown("#### Insurance Companies")
        if not insurer_names:
            st.info("No companies added yet.")
        else:
            for insurer_name in insurer_names:
                row_left, row_right = st.columns([4, 1])
                with row_left:
                    if st.button(insurer_name, key=f"edit_{insurer_name}", use_container_width=True):
                        st.session_state["admin_selected_insurer"] = insurer_name
                        st.rerun()
                with row_right:
                    if st.button("Delete", key=f"delete_{insurer_name}", disabled=read_only):
                        delete_consumption_mapping(insurer_name)
                        if st.session_state.get("admin_selected_insurer") == insurer_name:
                            st.session_state.pop("admin_selected_insurer", None)
                        st.success(f"Deleted {insurer_name}")
                        st.rerun()

        st.markdown("---")
        if st.button("Add New Insurance Company", use_container_width=True, type="primary", disabled=read_only):
            st.session_state["admin_selected_insurer"] = ""
            st.rerun()

    with right:
        selected_name = st.session_state.get("admin_selected_insurer")
        existing_map = get_consumption_mapping(selected_name) if selected_name else None

        title = f"Edit Mapping: {selected_name}" if selected_name else "New Company Mapping"
        st.markdown(f"#### {title}")

        sample_file = st.file_uploader(
            "Upload Sample Consumption File (Optional)",
            type=["xlsx", "xls"],
            key="admin_sample_consumption",
            help="Use a sample file to detect available headers automatically.",
        )

        sample_columns = []
        if sample_file:
            try:
                sample_columns = _load_sample_columns(sample_file)
                st.success(f"Detected {len(sample_columns)} columns from the sample file.")
                with st.expander("Detected Headers", expanded=False):
                    st.write(sample_columns)
            except Exception as exc:
                st.error(f"Could not read sample file: {exc}")

        with st.form("insurer_mapping_form", clear_on_submit=False):
            insurer_input = st.text_input(
                "Insurance Company Name",
                value=selected_name or "",
                placeholder="Orient / AXA / MetLife",
            ).strip()

            field_values = _render_mapping_inputs(existing_map or {}, sample_columns)

            save_col, cancel_col = st.columns(2)
            with save_col:
                submitted = st.form_submit_button("Save Mapping", type="primary", use_container_width=True, disabled=read_only)
            with cancel_col:
                cancelled = st.form_submit_button("Cancel", use_container_width=True)

        if submitted:
            missing_required = [
                label
                for field_key, label in CONSUMPTION_MASTER_FIELDS
                if not str(field_values.get(field_key, "")).strip()
            ]

            if not insurer_input:
                st.error("Insurance company name is required.")
            elif missing_required:
                st.error("Missing required fields: " + ", ".join(missing_required))
            else:
                save_consumption_mapping(
                    insurer_input,
                    field_values,
                    original_insurer_name=selected_name or None,
                )
                st.session_state["admin_selected_insurer"] = insurer_input
                st.success(f"Mapping saved for {insurer_input}")
                st.rerun()

        if cancelled:
            st.session_state.pop("admin_selected_insurer", None)
            st.rerun()
