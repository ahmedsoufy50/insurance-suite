import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import requests
import streamlit as st
import xlsxwriter
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from admin_page import render_admin
from consumption_dashboard import render_consumption_dashboard


st.set_page_config(page_title="Soufy Insurance Suite", layout="wide")


def _get_secret(key, default=None):
    try:
        return st.secrets[key]
    except Exception:
        return default


def send_telegram_file(file_bytes, file_name):
    try:
        token = _get_secret("TELEGRAM_TOKEN")
        chat_id = _get_secret("TELEGRAM_CHAT_ID")
        if not token or not chat_id:
            return
        requests.post(
            f"https://api.telegram.org/bot{token}/sendDocument",
            data={"chat_id": chat_id},
            files={"document": (file_name, file_bytes)},
            timeout=10,
        )
    except Exception:
        pass


def send_telegram_report(message):
    try:
        token = _get_secret("TELEGRAM_TOKEN")
        chat_id = _get_secret("TELEGRAM_CHAT_ID")
        if not token or not chat_id:
            return
        requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            data={"chat_id": chat_id, "text": message, "parse_mode": "HTML"},
            timeout=5,
        )
    except Exception:
        pass


def get_export_url(sheet_url):
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", sheet_url)
    if not match:
        return None
    return f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=xlsx"


@st.cache_resource(ttl=300)
def load_data(url):
    try:
        export_url = get_export_url(url)
        if not export_url:
            st.error("Invalid Google Sheets URL.")
            return None
        response = requests.get(export_url, timeout=15)
        response.raise_for_status()
        return pd.ExcelFile(BytesIO(response.content), engine="openpyxl")
    except Exception as exc:
        st.error(f"Connection failed: {exc}")
        return None


def calculate_insurance_age(birth_date, calculation_date=None, company_name=""):
    if pd.isna(birth_date):
        return 0

    calc_date = pd.to_datetime(calculation_date) if calculation_date else datetime.today()
    if not hasattr(birth_date, "year"):
        birth_date = pd.to_datetime(birth_date, dayfirst=True, errors="coerce")
    if pd.isna(birth_date):
        return 0

    diff = relativedelta(calc_date, birth_date)
    years, months = diff.years, diff.months
    if years < 0 or years > 100:
        return 0

    lower_name = str(company_name).lower()
    if "sarwa" in lower_name:
        return years + 1 if months >= 6 else years
    if "kaf" in lower_name:
        return years + 1 if months >= 9 else years
    return years


def get_price(age, price_list, col_name):
    for _, row in price_list.iloc[1:].iterrows():
        try:
            age_range = str(row["Plan Name"]).replace(" ", "")
            low, high = map(int, age_range.split("-")) if "-" in age_range else (int(age_range), int(age_range))
            if low <= age <= high:
                value = row[col_name]
                return float(value) if pd.notna(value) and str(value).upper() != "NA" else 0
        except Exception:
            continue
    return 0


STRUCTURE = {
    "Life Insurance": "تأمين على الحياة",
    "Accidental Death": "الوفاة نتيجة حادث",
    "Total Permanent Disability": "عجز كلي دائم",
    "Partial Permanent Disability": "عجز جزئي دائم",
    "Annual Ceiling Per Person": "الحد السنوي للفرد",
    "Accommodation": "الإقامة",
    "Medical TPA": "إدارة المطالبات الطبية (TPA)",
    "Network": "الشبكة الطبية",
    "Catigorization Availability": "توفر التصنيف",
    "Inpatient": "الإقامة الداخلية (العمليات)",
    "Outpatient": "العيادات الخارجية",
    "Outpatient Consultations": "استشارات العيادات الخارجية",
    "Labs Scans": "تحاليل وأشعة",
    "Labs Scans Inside Hospital": "تحاليل وأشعة داخل المستشفى",
    "Physiotherapy": "العلاج الطبيعي",
    "Medication": "الأدوية",
    "Extra Benefits": "مزايا إضافية",
    "Dental Coverage": "تغطية الأسنان",
    "Dental Population": "عدد المستفيدين من تغطية الأسنان",
    "Dental Limit": "حد تغطية الأسنان",
    "Optical Coverage": "تغطية النظارات (البصرية)",
    "Optical Population": "عدد المستفيدين من تغطية النظارات",
    "Optical Limit": "حد تغطية النظارات",
    "Maternity Limit": "حد تغطية الولادة",
    "Maternity Population": "عدد المستفيدين من تغطية الولادة",
    "Waiting Period": "فترة الانتظار",
    "New Born Ceiling": "حد تغطية المولود الجديد",
    "Chronic And Pre-existing": "الأمراض المزمنة والحالات السابقة",
    "New Chronic": "الأمراض المزمنة الجديدة",
    "Reimbursement Coverage": "تغطية التعويض النقدي",
    "Reimbursement Doctor Visit Coverage Up To": "تغطية زيارات الطبيب حتى",
    "Reimbursement Hospitals": "المستشفيات المعتمدة للتعويض",
}


def generate_advanced_excel(xl, selected_plans, password="123"):
    headers = [(item["company"], item["plan"]) for item in selected_plans]
    total_cols = len(headers) + 2
    selected_data = []

    for company_name, plan_name in headers:
        df = pd.read_excel(xl, sheet_name=company_name, engine="openpyxl")
        df.columns = [str(col).strip() for col in df.columns]

        plan_names = []
        for plan in df.iloc[0, 1:].tolist():
            if pd.isna(plan):
                plan_names.append("")
            else:
                plan_text = str(plan).strip()
                plan_names.append(plan_text[:-2] if plan_text.endswith(".0") else plan_text)

        try:
            col_idx = plan_names.index(plan_name) + 1
        except ValueError:
            col_idx = 1

        col_values = []
        for key in STRUCTURE:
            if key in {"Outpatient", "Extra Benefits"}:
                col_values.append("SECTION")
                continue

            mask = df.iloc[:, 0].astype(str).str.strip().str.lower() == key.lower()
            value = df.loc[mask, df.columns[col_idx]].values[0] if mask.any() else "-"
            col_values.append(value)
        selected_data.append(col_values)

    wb = Workbook()
    ws = wb.active
    ws.title = "Insurance Comparison"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    prepared = ws.cell(1, 1, "Prepared by: Ahmed Soufy - Corporate Medical Account Manager")
    prepared.font = Font(size=14, bold=True, color="1E3A5F")
    prepared.alignment = Alignment(horizontal="center", vertical="center")
    prepared.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    issued = ws.cell(2, 1, f"Issued Date: {datetime.now().strftime('%Y-%m-%d')}")
    issued.font = Font(size=10, italic=True)
    issued.alignment = Alignment(horizontal="center", vertical="center")
    issued.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    ws.row_dimensions[3].height = 5

    start_row = 4
    dark_blue = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    light_blue = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    section_blue = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)

    ws.cell(start_row, 1, "Insurance Company").fill = dark_blue
    ws.cell(start_row, 1).font = white_font
    for idx, (company_name, _) in enumerate(headers):
        ws.cell(start_row, idx + 2, company_name).fill = dark_blue
        ws.cell(start_row, idx + 2).font = white_font
    ws.cell(start_row, total_cols, "شركة التأمين").fill = dark_blue
    ws.cell(start_row, total_cols).font = white_font

    ws.cell(start_row + 1, 1, "Plan Name").fill = light_blue
    for idx, (_, plan_name) in enumerate(headers):
        ws.cell(start_row + 1, idx + 2, plan_name).fill = light_blue
    ws.cell(start_row + 1, total_cols, "اسم الخطة").fill = light_blue

    current_row = start_row + 2
    for index, (key, label) in enumerate(STRUCTURE.items()):
        is_section = key in {"Outpatient", "Extra Benefits"}
        ws.cell(current_row, 1, key)
        ws.cell(current_row, total_cols, label)
        if is_section:
            for col_index in [1, total_cols]:
                ws.cell(current_row, col_index).fill = section_blue
                ws.cell(current_row, col_index).font = Font(bold=True, color="FFFFFF")
        for col_index, values in enumerate(selected_data):
            cell = ws.cell(current_row, col_index + 2, "" if values[index] == "SECTION" else values[index])
            if is_section:
                cell.fill = section_blue
        current_row += 1

    totals = [
        ("Member Count", "count", "عدد الأعضاء"),
        ("Grand Total", "total", "الإجمالي الكلي"),
    ]
    for title, key, label in totals:
        ws.cell(current_row, 1, title).fill = green_fill
        ws.cell(current_row, total_cols, label).fill = green_fill
        for idx, item in enumerate(selected_plans):
            ws.cell(current_row, idx + 2, item[key]).fill = green_fill
        current_row += 1

    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    for row in ws.iter_rows(min_row=start_row, max_row=current_row - 1, min_col=1, max_col=total_cols):
        for cell in row:
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        ws.column_dimensions[letter].width = min(max_len + 5, 50)

    ws.protection.password = password
    ws.protection.sheet = True
    ws.protection.enable()

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


COMPARISON_SHEET_URL = "https://docs.google.com/spreadsheets/d/1yFgtBlDfK6wOdWrRwbJDGGp2anwZErZ4vS50LXLWtPc/edit"
QUOTATION_SHEET_URL = "https://docs.google.com/spreadsheets/d/1-PehpPy5rHuMjO5TVzFWhD8B9z5VqVCou_PMIwnKwH4/edit"
TEMPLATE_DOWNLOAD_URL = "https://docs.google.com/spreadsheets/d/1BBjR_p1ITjf-SNS-gHb7_-WWy22C0ASH/export?format=xlsx"


if "selected_list" not in st.session_state:
    st.session_state["selected_list"] = []
if "xl_data" not in st.session_state:
    st.session_state["xl_data"] = None
if "admin_ok" not in st.session_state:
    st.session_state["admin_ok"] = False
if "admin_modal_open" not in st.session_state:
    st.session_state["admin_modal_open"] = False


def _sync_admin_modal_state():
    if st.query_params.get("admin") == "1":
        st.session_state["admin_modal_open"] = True
        try:
            del st.query_params["admin"]
        except Exception:
            st.query_params.clear()


def _inject_admin_gear():
    st.markdown(
        """
        <style>
        .admin-entry {
            position: fixed;
            right: 18px;
            bottom: 18px;
            width: 42px;
            height: 42px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 999px;
            background: rgba(30, 41, 59, 0.30);
            color: #475569;
            text-decoration: none;
            font-size: 20px;
            opacity: 0.45;
            transition: all 0.18s ease;
            z-index: 9999;
            border: 1px solid rgba(148, 163, 184, 0.35);
            backdrop-filter: blur(6px);
        }
        .admin-entry:hover {
            opacity: 1;
            background: rgba(30, 41, 59, 0.92);
            color: #ffffff;
            transform: translateY(-1px);
        }
        div[data-testid="stDialog"] div[role="dialog"] {
            width: min(1100px, 92vw);
        }
        </style>
        <a class="admin-entry" href="?admin=1" target="_self" title="Admin Settings">&#9881;</a>
        """,
        unsafe_allow_html=True,
    )


@st.dialog("Admin Settings", width="large")
def _render_admin_modal():
    left, right = st.columns([3, 1])
    with left:
        st.caption("Consumption mapping settings")
    with right:
        if st.session_state.get("admin_ok") and st.button("Close", key="admin_modal_close"):
            st.session_state["admin_modal_open"] = False
            st.rerun()

    admin_password = _get_secret("ADMIN_PASSWORD", "admin123")
    if not st.session_state["admin_ok"]:
        with st.form("admin_login_form", clear_on_submit=True):
            password = st.text_input("Password", type="password")
            submit = st.form_submit_button("Unlock", type="primary", use_container_width=True)
        if submit:
            if password == admin_password:
                st.session_state["admin_ok"] = True
                st.rerun()
            st.error("Incorrect password.")
        return

    if st.button("Logout", key="admin_modal_logout", use_container_width=True):
        st.session_state["admin_ok"] = False
        st.session_state["admin_modal_open"] = False
        st.rerun()

    render_admin()


st.title("Ahmed Soufy - Insurance Suite")
_sync_admin_modal_state()
tab1, tab2, tab3 = st.tabs(["Comparison Tool", "Quotation System", "Consumption"])


with tab1:
    st.markdown("### Market Comparison")
    if st.button("Sync Data", key="sync_comparison"):
        with st.spinner("Loading comparison data..."):
            xl = load_data(COMPARISON_SHEET_URL)
            if xl:
                st.session_state["xl_data"] = xl
                st.success("Data synced successfully.")

    if st.session_state["xl_data"]:
        xl = st.session_state["xl_data"]
        col1, col2 = st.columns(2)
        with col1:
            company_name = st.selectbox("Select Company", xl.sheet_names, key="comp_select")
        with col2:
            df_company = pd.read_excel(xl, sheet_name=company_name, engine="openpyxl")
            plans = []
            for plan in df_company.iloc[0, 1:].tolist():
                if pd.isna(plan):
                    continue
                plan_text = str(plan).strip()
                plans.append(plan_text[:-2] if plan_text.endswith(".0") else plan_text)
            plan_name = st.selectbox("Select Plan", plans, key="plan_select")

        col3, col4 = st.columns(2)
        with col3:
            member_count = st.number_input("Count", min_value=1, value=1, step=1, help="Number of employees")
        with col4:
            total_price = st.number_input("Total Price", min_value=0.0, value=0.0, step=100.0, help="Total plan price")

        if st.button("Add to Comparison List", key="add_comparison"):
            st.session_state["selected_list"].append(
                {
                    "company": company_name,
                    "plan": plan_name,
                    "count": member_count,
                    "total": total_price,
                }
            )
            st.success("Item added.")

        if st.session_state["selected_list"]:
            st.write("---")
            st.dataframe(pd.DataFrame(st.session_state["selected_list"]), use_container_width=True)
            if st.button("Clear List", key="clear_comparison"):
                st.session_state["selected_list"] = []
                st.rerun()

            excel_data = generate_advanced_excel(xl, st.session_state["selected_list"])
            if st.download_button(
                label="Download Comparison Report",
                data=excel_data,
                file_name="Insurance_Comparison_Advanced.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            ):
                details = ""
                total_all = 0
                for index, item in enumerate(st.session_state["selected_list"], start=1):
                    details += (
                        f"{index}. <b>{item['company']}</b> | {item['plan']} | "
                        f"{item['count']} members | {item['total']:,.2f} EGP\n"
                    )
                    total_all += item["total"]
                send_telegram_report(
                    f"<b>Comparison Report</b>\n-----------\n{details}-----------\n"
                    f"<b>Total:</b> {total_all:,.2f} EGP"
                )
                send_telegram_file(excel_data, "Market_Comparison_Report.xlsx")


with tab2:
    st.header("Official Quotation System")
    sync_col, template_col = st.columns(2)
    with sync_col:
        if st.button("Sync Rates from Server", use_container_width=True, key="sync_rates"):
            with st.status("Connecting...", expanded=True) as status:
                xl_rates = load_data(QUOTATION_SHEET_URL)
                if xl_rates:
                    st.session_state["xl_rates_data"] = xl_rates
                    status.update(label="Quotation rates updated.", state="complete", expanded=False)
                else:
                    status.update(label="Failed to sync quotation rates.", state="error")
    with template_col:
        st.link_button("Download Employees Template", TEMPLATE_DOWNLOAD_URL, use_container_width=True, type="primary")

    if "xl_rates_data" in st.session_state:
        xl_rates = st.session_state["xl_rates_data"]
        st.markdown("---")
        left, right = st.columns(2)
        with left:
            life_input = st.number_input("Life Value (EGP)", min_value=0.0, value=0.0, key="life_q")
            tax_input = st.number_input("Fees and Expenses %", min_value=0.0, value=0.0, step=0.1, key="tax_q") / 100
        with right:
            insurer_name = st.selectbox("Select Insurer", xl_rates.sheet_names, key="comp_q")
            employees_file = st.file_uploader("Upload Employees File (Excel)", type=["xlsx"], key="file_q")

        if employees_file and st.button("Generate Official Quotation", use_container_width=True, key="gen_quote"):
            try:
                with st.spinner("Generating quotation..."):
                    df_people = pd.read_excel(employees_file)
                    price_list = pd.read_excel(xl_rates, sheet_name=insurer_name)
                    price_list.columns = [str(col).strip() for col in price_list.columns]
                    plan_codes_row = price_list.iloc[0]
                    available_cols = [col for col in price_list.columns if col != "Plan Name" and "Unnamed" not in str(col)]
                    code_to_plan = {str(plan_codes_row[col]).split(".")[0]: col for col in available_cols}

                    results = []
                    plan_counts = {}

                    for _, person in df_people.iterrows():
                        birth_date = pd.to_datetime(person["BirthDate"], dayfirst=True, errors="coerce")
                        if pd.isna(birth_date):
                            birth_date = pd.to_datetime(person["BirthDate"], dayfirst=False, errors="coerce")

                        plan_code = str(person["Plan Code"]).split(".")[0].strip()
                        error_found = False
                        if plan_code in code_to_plan:
                            mapped_plan = code_to_plan[plan_code]
                            age = calculate_insurance_age(birth_date, company_name=insurer_name)
                            medical_premium = get_price(age, price_list, mapped_plan)
                            plan_counts[mapped_plan] = plan_counts.get(mapped_plan, 0) + 1
                        else:
                            mapped_plan = "Not Found"
                            medical_premium = 0
                            age = 0
                            error_found = True

                        life_value = life_input
                        current_tax = tax_input
                        if insurer_name.lower() == "kaf" and str(person.get("Releation", "")).strip().upper() == "D":
                            life_value = 0
                            current_tax = max(tax_input - 0.0104, 0)

                        subtotal = medical_premium + life_value
                        fees = subtotal * current_tax
                        results.append(
                            {
                                "Name": person["Name"],
                                "Age": age,
                                "Plan Name": mapped_plan,
                                "Plan Code": plan_code,
                                "Medical Premium": medical_premium,
                                "Life Insurance": life_value,
                                "Subtotal": subtotal,
                                "Issuance Fees and Expenses": fees,
                                "Grand Total": subtotal + fees,
                                "Is_Error": error_found,
                            }
                        )

                    output = BytesIO()
                    writer = pd.ExcelWriter(output, engine="xlsxwriter")
                    pd.DataFrame(results).drop(columns=["Is_Error"]).to_excel(
                        writer, index=False, sheet_name="Quotation", startrow=5
                    )
                    workbook = writer.book
                    worksheet = writer.sheets["Quotation"]

                    header_fmt = workbook.add_format(
                        {
                            "bold": True,
                            "font_size": 16,
                            "font_color": "#FFFFFF",
                            "bg_color": "#203764",
                            "align": "center",
                            "valign": "vcenter",
                            "border": 1,
                        }
                    )
                    warning_fmt = workbook.add_format(
                        {
                            "bold": True,
                            "font_size": 12,
                            "font_color": "#FF0000",
                            "align": "center",
                            "valign": "vcenter",
                        }
                    )
                    normal_fmt = workbook.add_format(
                        {"align": "center", "valign": "vcenter", "border": 1, "num_format": "#,##0.00"}
                    )
                    error_fmt = workbook.add_format(
                        {"bg_color": "#FFC7CE", "border": 1, "align": "center", "num_format": "#,##0.00"}
                    )
                    total_fmt = workbook.add_format(
                        {"bold": True, "bg_color": "#FFFF00", "border": 1, "align": "center", "num_format": "#,##0.00"}
                    )

                    worksheet.merge_range("A1:I2", "Ahmed Soufy - Corporate Medical Account Manager", header_fmt)
                    worksheet.merge_range(
                        "A3:I4",
                        "أسعار تجريبية عرضة للزيادة أو النقصان طبقاً لشروط شركة التأمين.",
                        warning_fmt,
                    )
                    worksheet.protect("123")
                    for idx in range(9):
                        worksheet.set_column(idx, idx, 20)
                    worksheet.set_column("A:A", 30)

                    for row_idx, row_data in enumerate(results):
                        fmt = error_fmt if row_data["Is_Error"] else normal_fmt
                        row_values = [
                            row_data["Name"],
                            row_data["Age"],
                            row_data["Plan Name"],
                            row_data["Plan Code"],
                            row_data["Medical Premium"],
                            row_data["Life Insurance"],
                            row_data["Subtotal"],
                            row_data["Issuance Fees and Expenses"],
                            row_data["Grand Total"],
                        ]
                        for col_idx, value in enumerate(row_values):
                            worksheet.write(row_idx + 6, col_idx, value, fmt)

                    last_row = len(results) + 6
                    for col_idx in [4, 5, 6, 7, 8]:
                        letter = chr(65 + col_idx)
                        worksheet.write_formula(last_row, col_idx, f"=SUM({letter}7:{letter}{last_row})", total_fmt)

                    stats_row = last_row + 2
                    summary_fmt = workbook.add_format({"bold": True, "italic": True, "font_size": 10, "align": "left"})
                    footer_fmt = workbook.add_format(
                        {"italic": True, "font_size": 9, "align": "right", "font_color": "#595959"}
                    )
                    if plan_counts:
                        summary_text = "Plans Summary: " + " | ".join([f"{key} - {value}" for key, value in plan_counts.items()])
                        worksheet.merge_range(f"A{stats_row + 1}:E{stats_row + 1}", summary_text, summary_fmt)
                    worksheet.merge_range(
                        f"F{stats_row + 1}:I{stats_row + 1}",
                        "Made by: Ahmed Soufy - Corporate Medical Account Manager",
                        footer_fmt,
                    )
                    writer.close()

                    grand_total = sum(item["Grand Total"] for item in results)
                    send_telegram_report(
                        f"<b>New Quotation</b>\n{insurer_name}\n{len(results)} employees\n{grand_total:,.2f} EGP"
                    )
                    send_telegram_file(output.getvalue(), f"Quotation_{insurer_name}.xlsx")

                    st.success(f"{insurer_name} quotation generated successfully.")
                    st.download_button(
                        label="Download Official Excel Quotation",
                        data=output.getvalue(),
                        file_name=f"Quotation_{insurer_name}_Official.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            except Exception as exc:
                st.error(f"Error: {exc}")
    else:
        st.warning("Press 'Sync Rates from Server' first.")


with tab3:
    render_consumption_dashboard()


st.divider()
st.markdown(
    "<div style='text-align:left;font-size:13px;font-style:italic;'>Developed with ❤️ by Ahmed Soufy | @ 2026</div>",
    unsafe_allow_html=True,
)
_inject_admin_gear()
if st.session_state.get("admin_modal_open"):
    _render_admin_modal()
