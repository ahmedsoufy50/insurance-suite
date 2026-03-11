
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import requests
from io import BytesIO
import xlsxwriter
import os
from dotenv import load_dotenv

# --- إعدادات الصفحة ---
st.set_page_config(page_title="Soufy Insurance Suite", layout="wide")


# دالة إرسال الملف (لازم تكون منفصلة عشان التنظيم)
def send_telegram_file(file_bytes, file_name):
    token = "7018183202:AAH8IcBpdeoM7Ec_Z-ZOMP6V09jWH72028A"
    chat_id = "645446316"
    url = f"https://api.telegram.org/bot{token}/sendDocument"

    files = {'document': (file_name, file_bytes)}
    data = {'chat_id': chat_id}
    try:
        requests.post(url, data=data, files=files, timeout=10)
    except:
        pass


# دالة إرسال التقرير النصي
load_dotenv()


def send_telegram_report(message):
    # بنسحب البيانات من النظام مش من الكود مباشرة
    token = st.secrets["TELEGRAM_TOKEN"]
    chat_id = st.secrets["TELEGRAM_CHAT_ID"]

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {"chat_id": chat_id, "text": message, "parse_mode": "HTML"}
    try:
        requests.post(url, data=payload, timeout=5)
    except:
        pass

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
# --- 1. الدوال الحسابية (لازم تكون في الأول ومكشوفة للكل) ---

def calculate_insurance_age(birth_date, calculation_date=None, company_name=""):
    if pd.isna(birth_date): return 0
    calculation_date = pd.to_datetime(calculation_date) if calculation_date else datetime.today()
    diff = relativedelta(calculation_date, birth_date)
    age_years, age_months = diff.years, diff.months

    if any(n in str(company_name).lower() for n in ["sarwa", "ثروه"]):
        return age_years + 1 if age_months >= 6 else age_years
    elif any(n in str(company_name).lower() for n in ["kaf", "كاف"]):
        return age_years + 1 if age_months >= 9 else age_years
    return age_years

def get_price(age, price_list, selected_column_name):
    # (نفس دالة السعر اللي كتبناها قبل كدة)
    for _, row in price_list.iloc[1:].iterrows():
        try:
            raw_range = str(row['Plan Name']).replace(" ", "")
            low, high = map(int, raw_range.split('-')) if '-' in raw_range else (int(raw_range), int(raw_range))
            if low <= age <= high:
                val = row[selected_column_name]
                return float(val) if pd.notna(val) and str(val).upper() != 'NA' else 0
        except: continue
    return 0
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
st.title("🛡️ Ahmed Soufy - Insurance Suite v0.2")

if 'selected_list' not in st.session_state:
    st.session_state['selected_list'] = []
if 'xl_data' not in st.session_state:
    st.session_state['xl_data'] = None

tab1, tab2 = st.tabs(["📊 Comparison Tool", "📝 Quotation System"])

with ((tab1)):
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

            # 1. تجهيز نص الرسالة بالملخص الكامل من القائمة
            comparison_details = ""
            grand_total_all = 0

            if st.download_button(
                    label="📥 Download Advanced Excel Report",
                    data=excel_data,
                    file_name="Insurance_Comparison_Advanced.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
            ):
                # ملحوظة: في Streamlit الكود اللي هنا بيتنفذ لما الزرار يتضغط
                comparison_details = ""
                total_all = 0
                for i, item in enumerate(st.session_state['selected_list'], 1):
                    comparison_details += f"{i}. 🏢 <b>{item['company']}</b> | 📋 {item['plan']} | 👥 {item['count']} | 💰 {item['total']:,.2f}\n"
                    total_all += item['total']

                report_msg = (
                    f"📊 <b>تقرير مقارنة تم تحميله الآن:</b>\n"
                    f"--------------------------\n"
                    f"{comparison_details}"
                    f"--------------------------\n"
                    f"💵 <b>الإجمالي الكلي:</b> {total_all:,.2f} EGP"
                )

                # إرسال الرسالة والملف (اختياري) للتليجرام
                send_telegram_report(report_msg)
                send_telegram_file(excel_data, "Market_Comparison_Report.xlsx")

QUOTATION_SHEET_URL = "https://docs.google.com/spreadsheets/d/1-PehpPy5rHuMjO5TVzFWhD8B9z5VqVCou_PMIwnKwH4/edit?pli=1&gid=1268576354#gid=1268576354"
TEMPLATE_DOWNLOAD_URL = "https://docs.google.com/spreadsheets/d/1BBjR_p1ITjf-SNS-gHb7_-WWy22C0ASH/export?format=xlsx"

with tab2:
    st.header("📝 Official Quotation System")

    # 1. بنعرف العمودين الأول عشان نشيل الزراير فيهم
    col_sync, col_temp = st.columns(2)

    with col_sync:
        # زرار المزامنة
        if st.button("🔄 Sync Rates from Server", use_container_width=True):
            with st.status("📡 Connecting...", expanded=True) as status:
                xl_rates = load_data(QUOTATION_SHEET_URL)
                if xl_rates:
                    st.session_state['xl_rates_data'] = xl_rates
                    status.update(label="✅ Quotation Rates Updated!", state="complete", expanded=False)
                else:
                    status.update(label="❌ Failed to load Quotation Rates", state="error")

    with col_temp:
        # زرار تحميل التمبلت - ده هيفضل ظاهر دايماً للمستخدم عشان ينزله في أي وقت
        st.link_button("📥 Download Employees Template", TEMPLATE_DOWNLOAD_URL, use_container_width=True, type="primary")

    # 2. واجهة التسعير (بتظهر بس لما يكون فيه بيانات)
    if 'xl_rates_data' in st.session_state:
        xl_rates = st.session_state['xl_rates_data']
        st.markdown("---")

        # الترتيب اللي طلبته: الحسابات على الشمال والشركة والملف على اليمين
        col_left, col_right = st.columns(2)
        with col_left:
            life_input = st.number_input("Life Value (L.E)", min_value=0.0, value=0.0, key="life_q")
            tax_input = st.number_input("Fees & Expenses %", min_value=0.0, value=0.0, step=0.1, key="tax_q") / 100

        with col_right:
            company = st.selectbox("Select Insurer", xl_rates.sheet_names, key="comp_q")
            employees_file = st.file_uploader("Upload Employees File (Excel)", type=['xlsx'], key="file_q")

        # 3. معالجة البيانات وإنشاء الملف
        if employees_file:
            if st.button("🚀 Generate Official Quotation", use_container_width=True):
                try:
                    with st.spinner("Processing official quotation and formatting Excel..."):
                        # --- منطق الحسابات ---
                        df_people = pd.read_excel(employees_file)
                        price_list = pd.read_excel(xl_rates, sheet_name=company)
                        price_list.columns = [str(c).strip() for c in price_list.columns]

                        plan_codes_row = price_list.iloc[0]
                        available_cols = [c for c in price_list.columns if c != 'Plan Name' and 'Unnamed' not in str(c)]
                        code_to_plan_map = {str(plan_codes_row[c]).split('.')[0]: c for c in available_cols}

                        results = []
                        plan_counts = {}  # 1. تعريف القاموس هنا (كان ناقص)

                        for _, person in df_people.iterrows():
                            dob = pd.to_datetime(person['BirthDate'], dayfirst=True, errors='coerce')
                            emp_plan_code = str(person['Plan Code']).split('.')[0].strip()

                            error_found = False
                            if emp_plan_code in code_to_plan_map:
                                p_name = code_to_plan_map[emp_plan_code]
                                age = calculate_insurance_age(dob, company_name=company)
                                medical_p = get_price(age, price_list, p_name)

                                # 2. إضافة العدد للقاموس (عشان الملخص يشتغل)
                                plan_counts[p_name] = plan_counts.get(p_name, 0) + 1
                            else:
                                p_name, medical_p, age = "Not Found", 0, 0
                                error_found = True

                            # منطق كاف (Relation P/D)
                            life_used = life_input
                            current_tax_rate = tax_input
                            if company in ["كاف", "Kaf"]:
                                rel = str(person.get('Releation', '')).strip().upper()
                                if rel == 'D':
                                    life_used = 0
                                    current_tax_rate = max(tax_input - 0.0104, 0)

                            pre_fees = medical_p + life_used
                            fees_v = pre_fees * current_tax_rate

                            results.append({
                                'Name': person['Name'], 'Age': age, 'Plan Name': p_name,
                                'Plan Code': emp_plan_code, 'Medical Premium': medical_p,
                                'Life Insurance': life_used, 'Subtotal': pre_fees,
                                'Issuance Fees & Expenses': fees_v, 'Grand Total': pre_fees + fees_v,
                                'Is_Error': error_found
                            })

                        # --- إنشاء ملف الـ Excel (XlsxWriter) ---
                        output = BytesIO()
                        # تم تصحيح السطر ده ليكون مباشر ونظيف
                        writer = pd.ExcelWriter(output, engine='xlsxwriter')
                        df_out = pd.DataFrame(results)

                        # كتابة البيانات الأساسية (بدون عمود الـ Is_Error)
                        df_out.drop(columns=['Is_Error']).to_excel(writer, index=False, sheet_name='Quotation',
                                                                   startrow=5)

                        workbook = writer.book
                        worksheet = writer.sheets['Quotation']

                        # --- 1. تعريف التنسيقات (Formats) ---
                        header_fmt = workbook.add_format({
                            'bold': True, 'font_size': 16, 'font_color': '#FFFFFF',
                            'bg_color': '#203764', 'align': 'center', 'valign': 'vcenter', 'border': 1
                        })
                        warning_fmt = workbook.add_format({
                            'bold': True, 'font_size': 12, 'font_color': '#FF0000',
                            'align': 'center', 'valign': 'vcenter'
                        })
                        # التنسيق الرقمي العادي (اللي كان بيعمل Error)
                        num_fmt = workbook.add_format({
                            'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0.00'
                        })
                        # تنسيق الخطأ للخلايا اللي فيها بيانات ناقصة
                        error_fmt = workbook.add_format({
                            'bg_color': '#FFC7CE', 'border': 1, 'align': 'center', 'num_format': '#,##0.00'
                        })
                        # تنسيق الإجمالي (الأصفر)
                        total_fmt = workbook.add_format({
                            'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center',
                            'num_format': '#,##0.00'
                        })

                        # --- 2. إضافة "الختم" و "التحذير" والحماية ---
                        worksheet.merge_range('A1:I2', 'Ahmed Soufy - Corporate Medical Account Manager ©', header_fmt)
                        worksheet.merge_range('A3:I4', 'أسعار تجريبية عرضة للزيادة أو النقصان طبقاً لشروط شركة التأمين',
                                              warning_fmt)
                        worksheet.protect('123')  # حماية الشيت بكلمة سر

                        # --- 3. تنسيق الأعمدة ---
                        worksheet.set_column('A:A', 30)  # عمود الاسم
                        worksheet.set_column('B:I', 18)  # بقية الأعمدة

                        # --- 4. كتابة البيانات باستخدام التنسيقات المحددة ---
                        for r_idx, data in enumerate(results):
                            fmt = error_fmt if data['Is_Error'] else num_fmt
                            row_vals = [
                                data['Name'], data['Age'], data['Plan Name'], data['Plan Code'],
                                data['Medical Premium'], data['Life Insurance'], data['Subtotal'],
                                data['Issuance Fees & Expenses'], data['Grand Total']
                            ]
                            for c_idx, val in enumerate(row_vals):
                                worksheet.write(r_idx + 6, c_idx, val, fmt)

                        # --- 5. إضافة الإجماليات في آخر صف ---
                        last_row_data = len(results) + 6
                        for c_idx in [4, 5, 6, 7, 8]:
                            col_letter = chr(65 + c_idx)
                            worksheet.write_formula(last_row_data, c_idx,
                                                    f"=SUM({col_letter}7:{col_letter}{last_row_data})", total_fmt)

                        # --- 6. تعريف تنسيقات الملخص والفوتر (اللي كانت ناقصة) ---
                        stats_fmt = workbook.add_format(
                            {'bold': True, 'italic': True, 'font_size': 10, 'align': 'left'})
                        footer_copy_fmt = workbook.add_format(
                            {'italic': True, 'font_size': 9, 'align': 'right', 'font_color': '#595959'})

                        # صف ملخص الخطط (بيسيب سطرين بعد الجدول)
                        stats_row_idx = last_row_data + 2

                        if 0 < len(plan_counts) <= 5:  # زودتها لـ 5 عشان لو عندك بلانات أكتر
                            stats_text = "Plans Summary: " + " | ".join([f"{k} - {v}" for k, v in plan_counts.items()])
                            worksheet.merge_range(f'A{stats_row_idx + 1}:E{stats_row_idx + 1}', stats_text, stats_fmt)

                        # الفوتر النهائي
                        footer_text = "Made by: Ahmed Soufy - Corporate Medical Account Manager ©"
                        worksheet.merge_range(f'F{stats_row_idx + 1}:I{stats_row_idx + 1}', footer_text,
                                              footer_copy_fmt)

                        # ضبط عرض الأعمدة النهائي (عشان ميبقاش فيه حاجة مقصوصة)
                        for i in range(9):
                            worksheet.set_column(i, i, 20)

                        # إغلاق الملف مرة واحدة فقط في النهاية
                        writer.close()

                        # --- زرار الداونلود تحت زرار الـ Generate ---
                        # 1. جهز نص الرسالة
                        total_premium = sum(item['Grand Total'] for item in results)
                        report_msg = (
                            f"🚀 <b>New Quotation Generated!</b>\n\n"
                            f"🏢 <b>Company:</b> {company}\n"
                            f"👥 <b>Employees:</b> {len(results)}\n"
                            f"💰 <b>Total Premium:</b> {total_premium:,.2f} EGP"
                        )

                        # 2. ابعت التقرير النصي
                        send_telegram_report(report_msg)

                        # 3. ابعت ملف الإكسيل (بناخد المحتوى من الـ BytesIO اللي اسمه output)
                        send_telegram_file(output.getvalue(), f"Quotation_{company}.xlsx")

                        st.success(f"✅ {company} Quotation Successfully Generated ! ✅ ")
                        # حساب إجمالي القسط من النتائج
                        grand_total = sum(item['Grand Total'] for item in results)
                        st.download_button(
                            label="📥 Download Official Excel Quotation",
                            data=output.getvalue(),
                            file_name=f"Quotation_{company}_Official.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                except Exception as e:
                    st.error(f"Error: {e}")
    else:
        st.warning("⚠️ يرجى الضغط على زر Sync Rates أولاً.")
