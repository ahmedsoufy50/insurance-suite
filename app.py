import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# --- إعدادات الصفحة ---
st.set_page_config(page_title="Soufy Insurance Suite", layout="wide")

# --- دوال مساعدة للاتصال بجوجل شيت ---
def get_export_url(sheet_url):
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', sheet_url)
    if not match:
        return None
    doc_id = match.group(1)
    return f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"

@st.cache_resource(ttl=300) # بيحفظ الداتا لمدة 10 دقائق قبل ما يحدثها تاني
def load_data(url):
    try:
        export_url = get_export_url(url)
        if not export_url:
            st.error("الرابط غير صالح")
            return None
        response = requests.get(export_url, timeout=15)
        response.raise_for_status()
        return pd.ExcelFile(BytesIO(response.content), engine='openpyxl')
    except Exception as e:
        st.error(f"فشل الاتصال: {e}")
        return None

# --- الهيكل الثابت للمقارنة (مطابق للكود الثاني) ---
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
    "Outpatient": "العيادات الخارجية",  # فاصل
    "Outpatient Consultations": "استشارات العيادات الخارجية",
    "Labs Scans": "تحاليل وأشعة",
    "Labs Scans Inside Hospital": "تحاليل وأشعة داخل المستشفى",
    "Physiotherapy": "العلاج الطبيعي",
    "Medication": "الأدوية",
    "Extra Benefits": "مزايا إضافية",  # فاصل
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
    "Reimbursement Hospitals": "المستشفيات المعتمدة للتعويض"
}

# --- دالة إنشاء ملف Excel المتقدم ---
def generate_advanced_excel(xl, selected_plans, password="123"):
    """
    ينشئ ملف Excel بنفس تنسيق الكود الثاني:
    - توقيع وتاريخ
    - رأس الشركات (أزرق غامق)
    - رأس الخطط (أزرق فاتح)
    - صفوف البيانات حسب الهيكل مع فواصل زرقاء
    - صفوف مالية خضراء (Member Count, Grand total)
    - حدود ومحاذاة وحماية
    """
    # تجهيز الهيكل والبيانات
    headers = [(item["company"], item["plan"]) for item in selected_plans]
    num_plans = len(headers)
    total_cols = num_plans + 2  # عمود إنجليزي + أعمدة الخطط + عمود عربي

    selected_data = []  # قائمة من القوائم، كل قائمة تمثل عموداً
    for (comp, plan) in headers:
        df = pd.read_excel(xl, sheet_name=comp, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        # أسماء الخطط (الصف الأول)
        plan_names_raw = df.iloc[0, 1:].tolist()
        plan_names = []
        for p in plan_names_raw:
            if pd.isna(p):
                plan_names.append("")
            else:
                p_str = str(p).strip()
                if p_str.endswith('.0'):
                    p_str = p_str[:-2]
                plan_names.append(p_str)

        try:
            col_idx = plan_names.index(plan) + 1
        except ValueError:
            col_idx = 1  # افتراضي

        col_vals = []
        for key in STRUCTURE.keys():
            if key in ["Outpatient", "Extra Benefits"]:
                col_vals.append("SECTION")
            else:
                mask = df.iloc[:, 0].astype(str).str.strip().str.lower() == key.strip().lower()
                if mask.any():
                    val = df.loc[mask, df.columns[col_idx]].values[0]
                else:
                    val = "-"
                col_vals.append(val)
        selected_data.append(col_vals)

    # إنشاء workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Insurance Comparison"

    # ===== التوقيع والتاريخ =====
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    sig_cell = ws.cell(1, 1, "Prepared by: Ahmed Soufy - Corporate Medical Account Manager")
    sig_cell.font = Font(size=14, bold=True, color="1E3A5F")
    sig_cell.alignment = Alignment(horizontal="center", vertical="center")
    sig_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    today_date = datetime.now().strftime("%Y-%m-%d")
    date_cell = ws.cell(2, 1, f"Issued Date: {today_date}")
    date_cell.font = Font(size=10, italic=True)
    date_cell.alignment = Alignment(horizontal="center", vertical="center")
    date_cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    ws.row_dimensions[3].height = 5

    # ===== رؤوس الأعمدة (بدءاً من الصف 4) =====
    header_row = 4
    dark_blue = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    light_blue = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    section_blue = PatternFill(start_color="8DB4E2", end_color="8DB4E2", fill_type="solid")
    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)

    # الصف الأول (أسماء الشركات)
    ws.cell(header_row, 1, "Insurance Company").fill = dark_blue
    ws.cell(header_row, 1).font = white_font

    for i, (comp, _) in enumerate(headers):
        cell = ws.cell(header_row, i+2, comp)
        cell.fill = dark_blue
        cell.font = white_font

    ws.cell(header_row, total_cols, "شركة التأمين").fill = dark_blue
    ws.cell(header_row, total_cols).font = white_font

    # الصف الثاني (أسماء الخطط)
    ws.cell(header_row+1, 1, "Plan Name").fill = light_blue
    for i, (_, plan) in enumerate(headers):
        cell = ws.cell(header_row+1, i+2, plan)
        cell.fill = light_blue
    ws.cell(header_row+1, total_cols, "اسم الخطة").fill = light_blue

    # ===== كتابة البيانات (من الصف 6) =====
    curr_row = header_row + 2
    for i, (key, arabic) in enumerate(STRUCTURE.items()):
        is_section = key in ["Outpatient", "Extra Benefits"]

        ws.cell(curr_row, 1, key)
        ws.cell(curr_row, total_cols, arabic)

        if is_section:
            ws.cell(curr_row, 1).fill = section_blue
            ws.cell(curr_row, total_cols).fill = section_blue
            ws.cell(curr_row, 1).font = Font(bold=True, color="FFFFFF")
            ws.cell(curr_row, total_cols).font = Font(bold=True, color="FFFFFF")

        for col_idx, col_vals in enumerate(selected_data):
            val = col_vals[i]
            cell = ws.cell(curr_row, col_idx+2, "" if val == "SECTION" else val)
            if is_section:
                cell.fill = section_blue

        curr_row += 1

    # ===== صفوف المالية (Member Count & Grand total) =====
    finance_row_start = curr_row
    for label, data_key, arabic in [("Member Count", "count", "عدد الأعضاء"),
                                     ("Grand total", "total", "الإجمالي الكلي")]:
        ws.cell(curr_row, 1, label).fill = green_fill
        ws.cell(curr_row, total_cols, arabic).fill = green_fill
        for i, item in enumerate(selected_plans):
            ws.cell(curr_row, i+2, item[data_key]).fill = green_fill
        curr_row += 1

    # ===== تنسيق الحدود والمحاذاة =====
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows(min_row=header_row, max_row=curr_row-1, min_col=1, max_col=total_cols):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ضبط عرض الأعمدة
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
        ws.column_dimensions[col_letter].width = min(max_len + 5, 50)

    # حماية الورقة
    ws.protection.password = password
    ws.protection.sheet = True
    ws.protection.enable()

    # حفظ إلى BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# --- واجهة Streamlit ---
st.title("🛡️ Ahmed Soufy - Insurance Suite v2.0")

if 'selected_list' not in st.session_state:
    st.session_state['selected_list'] = []
if 'xl_data' not in st.session_state:
    st.session_state['xl_data'] = None

tab1, tab2 = st.tabs(["📊 Comparison Tool", "📝 Quotation System"])

with tab1:
    st.markdown("### Market Comparison")

    url = "https://docs.google.com/spreadsheets/d/1yFgtBlDfK6wOdWrRwbJDGGp2anwZErZ4vS50LXLWtPc/edit"

    if 'xl_data' not in st.session_state:
        # إظهار رسالة تحميل أنيقة
        with st.status("📡 Connecting to Insurance Server...", expanded=True) as status:
            st.write("Fetching latest rates...")
            xl = load_data(SHEET_URL)
            if xl:
                st.session_state['xl_data'] = xl
                status.update(label="✅ Data Synced from Server!", state="complete", expanded=False)
            else:
                status.update(label="❌ Connection Failed", state="error")
    if st.button("Sync Data 🔄"):
        with st.spinner('Loading...'):
            xl = load_data(url)
            if xl:
                st.session_state['xl_data'] = xl
                st.success(f"Connected to Server")

    if st.session_state['xl_data']:
        xl = st.session_state['xl_data']
        col1, col2 = st.columns(2)

        with col1:
            comp = st.selectbox("Select Company", xl.sheet_names, key="comp_select")
        with col2:
            # قراءة أسماء الخطط مع تنظيف .0
            df_comp = pd.read_excel(xl, sheet_name=comp, engine='openpyxl')
            raw_plans = df_comp.iloc[0, 1:].tolist()
            plans = []
            for p in raw_plans:
                if pd.isna(p):
                    continue
                p_str = str(p).strip()
                if p_str.endswith('.0'):
                    p_str = p_str[:-2]
                plans.append(p_str)
            plan = st.selectbox("Select Plan", plans, key="plan_select")

        col3, col4 = st.columns(2)
        with col3:
            count = st.number_input("Count", min_value=1, value=1, step=1 , help="أدخل عدد الموظفين في هذه الخطة")
        with col4:
            price = st.number_input("Total Price", min_value=0.0, value=0.0, step=100.0 , help="أدخل سعر البلان بعدد الموظفين")

        if st.button("Add to Comparison List ➕"):
            st.session_state['selected_list'].append({
                "company": comp,
                "plan": plan,
                "count": count,
                "total": price
            })
            st.success("Added!")

        if st.session_state['selected_list']:
            st.write("---")
            st.dataframe(pd.DataFrame(st.session_state['selected_list']))

            if st.button("Clear List 🗑️"):
                st.session_state['selected_list'] = []
                st.rerun()

            # زر تحميل التقرير المتقدم
            excel_data = generate_advanced_excel(
                xl,
                st.session_state['selected_list'],
                password="123"
            )
            st.download_button(
                label="📥 Download Advanced Excel Report",
                data=excel_data,
                file_name="Insurance_Comparison_Advanced.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

with tab2:
    st.info("🚧 Quotation System is under construction. Stay tuned!")