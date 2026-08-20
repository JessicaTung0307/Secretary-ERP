import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta, date, timezone
import calendar
import io
import json
import zipfile
from weasyprint import HTML
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- 1. Database Connection & Auto-Migration ---
try:
    if "DB_URL" not in st.secrets:
        st.error("❌ `DB_URL` missing in Streamlit Secrets! Please configure it in Settings -> Secrets.")
        st.stop()
    
    DB_URL = st.secrets["DB_URL"]
    engine = create_engine(DB_URL)
    
    with engine.begin() as conn:
        try:
            existing_cols = pd.read_sql("SELECT * FROM companies LIMIT 0", conn).columns
            if 'biz_name' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN biz_name TEXT"))
            if 'branch_code' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN branch_code TEXT"))
            if 'ar_ref_date' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN ar_ref_date DATE"))
            if 'br_ref_date' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN br_ref_date DATE"))
            if 'cessation_date' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN cessation_date DATE"))
            if 'agent' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN agent TEXT"))
            if 'year_end' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN year_end TEXT"))
            if 'billing_mode' not in existing_cols: conn.execute(text("ALTER TABLE companies ADD COLUMN billing_mode TEXT"))
        except Exception:
            pass

except Exception as db_err:
    st.error("### 🛑 Database Connection Critical Failure")
    st.info(f"**Error Details:**\n`{str(db_err)}`")
    st.stop()

# --- 2. Utility Functions & Dynamic Years (HKT Timezone locked) ---
HKT = timezone(timedelta(hours=8))

def to_date(val):
    try:
        if pd.isna(val) or val == "" or str(val).strip() == "" or str(val).lower() in ["none", "nat", "nan"]: return None
        return pd.to_datetime(val).date()
    except: return None

def clean_val(v):
    v = str(v).strip()
    if v.lower() in ["nat", "none", "nan", ""]: return ""
    if v.endswith(" 00:00:00"): return v.replace(" 00:00:00", "")
    return v.strip()

def get_anniv(year, month, day):
    try: return date(year, month, day)
    except ValueError: return date(year, month, day - 1)

def clean_status(val):
    v = str(val)
    for emo in ["🔴 ", "🟡 ", "🟢 ", "✅ ", "⚪ ", "⏳ ", "⚠️ ", "🆕 ", "💰 ", "📄 ", "✨", "✨ "]: 
        v = v.replace(emo, "")
    return v.strip()

def add_months(dt, months):
    if not dt: return None
    m = dt.month + months
    y = dt.year + (m - 1) // 12
    m = (m - 1) % 12 + 1
    d = min(dt.day, calendar.monthrange(y, m)[1])
    return date(y, m, d)

def get_base_date(row_dict):
    place = str(row_dict.get('incorp_place', ''))
    is_hk_reg = str(row_dict.get('is_hk_registered', 'False')).strip().lower() in ['true', 'yes', 'y', '1']
    if place != 'HK' and is_hk_reg:
        return to_date(row_dict.get('hk_incorp_date'))
    return to_date(row_dict.get('incorp_date'))

def calc_ar_deadline(base_dt, ar_ref_date_val, year):
    if ar_ref_date_val:
        base_for_calc = get_anniv(year, ar_ref_date_val.month, ar_ref_date_val.day)
    else:
        base_for_calc = get_anniv(year, base_dt.month, base_dt.day)
    return base_for_calc + timedelta(days=42)

def calc_bvi_fee_deadline(incorp_dt, year):
    if not incorp_dt: return date(year, 5, 31)
    if incorp_dt.month <= 6: return date(year, 5, 31)
    else: return date(year, 11, 30)

def calc_afr_deadline(year_end_str, year):
    if not year_end_str or str(year_end_str).strip() == "": year_end_str = "12/31"
    try: m, d = map(int, str(year_end_str).split('/'))
    except: m, d = 12, 31
    ye_date = get_anniv(year, m, d)
    return add_months(ye_date, 9)

def calc_es_deadline(incorp_dt, year):
    if not incorp_dt: return None
    if incorp_dt.year < 2019:
        return date(year, 12, 31)
    else:
        fp_end = get_anniv(year, incorp_dt.month, incorp_dt.day)
        return add_months(fp_end, 6)

def get_stat(d_val, dl, exempt=False):
    if exempt: return "Exempt"
    if d_val: return "Completed"
    if not dl: return "Pending"
    diff = (dl - datetime.now(HKT).date()).days
    if diff < 0: return "Overdue"
    if diff <= 90: return "Due Soon"
    return "Pending"

current_system_year = datetime.now(HKT).year
active_years = list(range(2024, current_system_year + 5))
report_years = [y for y in active_years if y <= current_system_year]

# --- 3. Navigation ---
st.set_page_config(page_title="Secretary ERP - V130", layout="wide")
choice = st.sidebar.radio("Navigation (V130)", ["📊 Dashboard", "🏢 Company Register", "⚙️ Group Management", "📤 Data Exchange"])

TEMPLATE_COLS = [
    "client_group", "name_en", "name_ch", "biz_name", "incorp_place", "incorp_place_others", 
    "incorp_date", "ci_no", "is_hk_registered", "hk_incorp_date", "hk_ci_no", "br_no", "branch_code",
    "co_type", "reg_addr", "corres_addr", "round_loc", "sign_loc", "seal_loc", 
    "br_ref_date", "ar_ref_date", "cessation_date", "agent", "year_end", "billing_mode",
    "nd2a_eff_date", "nd2a_file_date", "nd2a_download", "nd4_eff_date", "nd4_file_date", "nd4_download", 
    "nn6_eff_date", "nn6_file_date", "nn6_download",
    "dissolution_date", "remark"
]

EXCHANGE_COL_MAPPING = {
    "client_group": "Client Group", "name_en": "Company Name EN", "name_ch": "Company Name CH", 
    "biz_name": "Business Name", "incorp_place": "Incorp Place", "incorp_place_others": "Incorp Place Others", 
    "incorp_date": "Incorp Date", "ci_no": "CI No.", "is_hk_registered": "Non-HK Registered in HK", 
    "hk_incorp_date": "HK Incorp Date", "hk_ci_no": "HK CI No.", "br_no": "BR No.", 
    "branch_code": "Branch Code", "co_type": "Company Type", "reg_addr": "Registered Address", 
    "corres_addr": "Correspondence Address", "round_loc": "Round Stamp", "sign_loc": "Signature Chop", 
    "seal_loc": "Common Seal", "br_ref_date": "BR Ref Date", "ar_ref_date": "AR Ref Date", 
    "cessation_date": "Cessation Date", "agent": "Registered Agent", "year_end": "Financial Year End", "billing_mode": "Billing Mode",
    "nd2a_eff_date": "ND2A Eff Date", "nd2a_file_date": "ND2A File Date", 
    "nd2a_download": "ND2A Download", "nd4_eff_date": "ND4 Eff Date", "nd4_file_date": "ND4 File Date", 
    "nd4_download": "ND4 Download", "nn6_eff_date": "NN6 Eff Date", "nn6_file_date": "NN6 File Date", 
    "nn6_download": "NN6 Download", "dissolution_date": "Dissolution Date", "remark": "Remark"
}
REVERSE_EXCHANGE_MAPPING = {v: k for k, v in EXCHANGE_COL_MAPPING.items()}
REVERSE_EXCHANGE_MAPPING['Business Name (業務名稱)'] = 'biz_name'
REVERSE_EXCHANGE_MAPPING['BR No. (8-digit)'] = 'br_no'

# --- 4. Report Generation ---

@st.cache_data(show_spinner=False)
def generate_custom_pdf(selected_df, hide_client_group=False):
    now = datetime.now(HKT).strftime("%Y/%m/%d %H:%M")
    def fmt_date(val):
        d = to_date(val)
        return d.strftime('%Y/%m/%d') if d else "N/A"
    if selected_df.empty: return b""
    sort_cols = [c for c in ['client_group', 'name_en', 'branch_code', 'incorp_place'] if c in selected_df.columns]
    selected_df = selected_df.sort_values(by=sort_cols, na_position='last')

    html_head = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700&display=swap" rel="stylesheet">
        <style>
            @page {{ size: A4; margin: 15mm 10mm 20mm 10mm; @bottom-left {{ content: "Company Report | Generated on: {now}"; font-size: 8.5pt; color: #7f8c8d; font-family: 'Noto Sans TC', sans-serif; }} @bottom-right {{ content: counter(page) " of " counter(pages) " Page(s)"; font-size: 8.5pt; color: #7f8c8d; font-family: 'Noto Sans TC', sans-serif; }} }}
            body {{ font-family: 'Noto Sans TC', sans-serif; color: #2c3e50; line-height: 1.4; background-color: #ffffff; text-align: justify; }}
            .company-container {{ width: 100%; }}
            .main-table {{ width: 100%; border-collapse: collapse; }}
            .header-content {{ text-align: center; border-bottom: 1px solid #eee; padding-bottom: 15px; margin-bottom: 20px; }}
            .name-en {{ font-size: 20pt; font-weight: bold; color: #2980b9; text-align: center; }}
            .name-ch {{ font-size: 15pt; color: #333333; margin-top: 5px; text-align: center; min-height: 20px; }}
            .section-bar {{ background-color: #f1f4f6; padding: 8px 15px; font-weight: bold; font-size: 11pt; margin: 20px 0 10px 0; border-left: 5px solid #3498db; color: #2c3e50; text-align: left; }}
            .section-group {{ page-break-inside: avoid; }}
            .info-table {{ width: 100%; border-collapse: collapse; }}
            .info-table tr {{ border-bottom: 1px solid #f1f2f6; page-break-inside: avoid; }}
            .info-table th {{ text-align: left; width: 45%; color: #7f8c8d; padding: 8px 0; font-weight: normal; font-size: 10.5pt; }}
            .info-table td {{ text-align: justify; padding: 8px 0; color: #2c3e50; font-size: 10.5pt; font-weight: bold; }}
        </style>
    </head>
    <body>
    """

    cg_row_html = "" if hide_client_group else "<tr><th>Client Group</th><td>__CLIENT_GROUP__</td></tr>"
    docs = []
    companies_with_branches = set(selected_df[selected_df['branch_code'] != '000']['name_en'])
    
    for _, row in selected_df.iterrows():
        ch_name = row.get('name_ch', '')
        if not ch_name or pd.isna(ch_name): ch_name = ''
        biz_name = str(row.get('biz_name', '')).strip()
        biz_html = f'<div class="name-ch" style="font-size: 11.5pt; color: #7f8c8d; font-weight: normal; margin-top: 4px;">Business Name: {biz_name}</div>' if biz_name and biz_name not in ['None', 'nan'] else ''
        place = str(row.get('incorp_place', ''))
        is_hk_reg = str(row.get('is_hk_registered', 'False')).strip().lower() in ['true', 'yes', '1']
        is_bvi = place not in ['HK', ''] and not is_hk_reg
        base_date = get_base_date(row)
        incorp_year = base_date.year if base_date else None
        branch = str(row.get('branch_code', '000')).strip()
        if branch in ['None', 'nan', '', '<NA>']: branch = '000'
        is_branch = branch != '000'
        has_branch = row.get('name_en') in companies_with_branches
        disp_en = str(row.get('name_en', ''))
            
        dynamic_place_rows = ""
        display_place = place
        if place == 'Others': display_place = f"Others ({str(row.get('incorp_place_others', ''))})"
            
        dynamic_place_rows += f"<tr><th>{place} Incorp Date</th><td>{fmt_date(row.get('incorp_date'))}</td></tr>"
        dynamic_place_rows += f"<tr><th>{place} CI No.</th><td>{str(row.get('ci_no', ''))}</td></tr>"

        if is_bvi:
            agent = str(row.get('agent', ''))
            ye = str(row.get('year_end', '12/31'))
            bm = str(row.get('billing_mode', ''))
            if 'All-in' in bm: bm = 'All-in Package'
            elif 'Itemized' in bm: bm = 'Itemized'
            if agent and agent != 'None': dynamic_place_rows += f"<tr><th>Registered Agent</th><td>{agent}</td></tr>"
            if ye and ye != 'None': dynamic_place_rows += f"<tr><th>Financial Year End</th><td>{ye}</td></tr>"
            if bm and bm != 'None': dynamic_place_rows += f"<tr><th>Billing Mode</th><td>{bm}</td></tr>"

        dynamic_hk_rows = ""
        dynamic_annual_rows = ""
        br_ref_raw = to_date(row.get('br_ref_date'))
        ar_ref_raw = to_date(row.get('ar_ref_date'))
        
        if br_ref_raw and (not base_date or (br_ref_raw.month != base_date.month or br_ref_raw.day != base_date.day)):
            dynamic_place_rows += f"<tr><th>BR Ref Date (MM/DD)</th><td>{br_ref_raw.strftime('%m/%d')}</td></tr>"
        if ar_ref_raw and (not base_date or (ar_ref_raw.month != base_date.month or ar_ref_raw.day != base_date.day)):
            dynamic_place_rows += f"<tr><th>AR Ref Date (MM/DD)</th><td>{ar_ref_raw.strftime('%m/%d')}</td></tr>"
            
        br_no_raw = str(row.get('br_no', '')).strip()
        if has_branch and br_no_raw: disp_br = f"{br_no_raw}-{branch}"
        else: disp_br = br_no_raw
            
        if place == 'HK' or is_hk_reg:
            if place == 'HK':
                dynamic_hk_rows += f"<tr><th>HK BR No.</th><td>{disp_br}</td></tr>"
            else:
                dynamic_hk_rows += f"<tr><th>HK Incorp Date</th><td>{fmt_date(row.get('hk_incorp_date'))}</td></tr>"
                dynamic_hk_rows += f"<tr><th>HK CI No.</th><td>{str(row.get('hk_ci_no', ''))}</td></tr>"
                dynamic_hk_rows += f"<tr><th>HK BR No.</th><td>{disp_br}</td></tr>"

        comp_rec_str = str(row.get('compliance_records', '{}'))
        try: rec_dict = json.loads(comp_rec_str)
        except: rec_dict = {}
        if not isinstance(rec_dict, dict): rec_dict = {}
        
        cess_date = to_date(row.get('cessation_date'))
        year_display_data = {}
        prev_br_by = 'Firm'; prev_afr_fee_by = 'Firm'; prev_es_fee_by = 'Firm'
        
        for y in report_years:
            y_str = str(y)
            y_data = rec_dict.get(y_str, {})
            if incorp_year and y < incorp_year:
                year_display_data[y] = {'not_incorp': True}
                br_by = 'N/A'; afr_fee_by = 'N/A'; es_fee_by = 'N/A'
            else:
                raw_br_by = str(y_data.get('fee_by', y_data.get('br_paid_by', ''))).strip()
                br_by = raw_br_by if raw_br_by else (prev_br_by if prev_br_by != 'N/A' else 'Firm')
                
                if is_bvi:
                    raw_afr_fee_by = str(y_data.get('afr_fee_by', '')).strip()
                    afr_fee_by = raw_afr_fee_by if raw_afr_fee_by else (prev_afr_fee_by if prev_afr_fee_by != 'N/A' else 'Firm')
                    raw_es_fee_by = str(y_data.get('es_fee_by', '')).strip()
                    es_fee_by = raw_es_fee_by if raw_es_fee_by else (prev_es_fee_by if prev_es_fee_by != 'N/A' else 'Firm')
                else:
                    afr_fee_by = 'N/A'; es_fee_by = 'N/A'
                
                if is_branch and cess_date and y >= cess_date.year: br_by = "N/A"
                prev_br_by = br_by; prev_afr_fee_by = afr_fee_by; prev_es_fee_by = es_fee_by
                
                br_dt_val = to_date(y_data.get('fee_date', y_data.get('br_date')))
                afr_fee_dt_val = to_date(y_data.get('afr_fee_date'))
                ar_dt_val = to_date(y_data.get('ar_date'))
                es_fee_dt_val = to_date(y_data.get('es_fee_date'))
                es_dt_val = to_date(y_data.get('es_date'))
                
                br_dt = br_dt_val.strftime('%Y/%m/%d') if br_dt_val else 'N/A'
                afr_fee_dt = afr_fee_dt_val.strftime('%Y/%m/%d') if afr_fee_dt_val else 'N/A'
                ar_dt = ar_dt_val.strftime('%Y/%m/%d') if ar_dt_val else 'N/A'
                es_fee_dt = es_fee_dt_val.strftime('%Y/%m/%d') if es_fee_dt_val else 'N/A'
                es_dt = es_dt_val.strftime('%Y/%m/%d') if es_dt_val else 'N/A'
                
                if is_bvi:
                    ar_dl = calc_afr_deadline(row.get('year_end'), y)
                    es_dl = calc_es_deadline(base_date, y)
                    ar_cr_status = get_stat(ar_dt_val, ar_dl, y == incorp_year)
                    es_status = get_stat(es_dt_val, es_dl, y == incorp_year)
                else:
                    ar_dl = calc_ar_deadline(base_date, to_date(row.get('ar_ref_date')), y)
                    ar_cr_status = y_data.get('ar_cr_status', 'Pending')
                    if not ar_cr_status: ar_cr_status = 'Completed' if ar_dt_val else 'Pending'
                    es_status = 'N/A'
                    
                if is_branch:
                    ar_dt_disp = "N/A (Branch)"
                    ar_cr_disp = "N/A"
                    es_dt_disp = "N/A"
                else:
                    if ar_cr_status == 'Exempt (Dormant)': ar_dt_disp = "N/A"; ar_cr_disp = "Exempt (Dormant)"
                    elif ar_cr_status == 'Included in Agent Fee': ar_dt_disp = "N/A"; ar_cr_disp = "Included in Agent Fee"
                    elif incorp_year and y == incorp_year and not ar_dt_val: ar_dt_disp = "Exempt (1st Year)"; ar_cr_disp = "Exempt (1st Year)" if not is_bvi else "N/A"
                    else: ar_dt_disp = ar_dt; ar_cr_disp = ar_cr_status
                    
                    if not is_bvi: es_dt_disp = "N/A"
                    elif incorp_year and y == incorp_year and not es_dt_val: es_dt_disp = "Exempt (1st Year)"
                    else: es_dt_disp = es_dt
                    
                if is_branch and cess_date and y >= cess_date.year: br_dt = "N/A"
                    
                year_display_data[y] = {
                    'not_incorp': False, 'br_by': br_by, 'br_dt': br_dt, 'afr_fee_by': afr_fee_by, 'afr_fee_dt': afr_fee_dt,
                    'ar_dt_disp': ar_dt_disp, 'ar_cr_disp': ar_cr_disp, 'es_fee_by': es_fee_by, 'es_fee_dt': es_fee_dt, 'es_dt_disp': es_dt_disp
                }

        for y in reversed(report_years):
            yd = year_display_data[y]
            if y == current_system_year: header_html = f"<tr><th colspan='2' style='background-color:#ebf5fb; color:#2980b9; text-align:center; border: 1px solid #aed6f1; padding: 6px 0; font-weight: bold;'>=== FY {y} Annual Cycle (Current) ===</th></tr>"
            else: header_html = f"<tr><th colspan='2' style='background-color:#fcfcfc; color:#95a5a6; text-align:center; font-weight:normal; border-top: 1px dashed #eaeded; padding: 6px 0;'>--- FY {y} Annual Cycle ---</th></tr>"
            dynamic_annual_rows += header_html
            
            if yd['not_incorp']: dynamic_annual_rows += f"<tr><td colspan='2' style='text-align:center; color:#bdc3c7; font-weight:normal;'>Not Incorporated Yet</td></tr>"
            else:
                br_dt_disp = yd['br_dt'].replace('-', '/') if yd['br_dt'] != 'N/A' else yd['br_dt']
                afr_fee_dt_disp = yd['afr_fee_dt'].replace('-', '/') if yd['afr_fee_dt'] != 'N/A' else yd['afr_fee_dt']
                ar_dt_disp = yd['ar_dt_disp'].replace('-', '/') if yd['ar_dt_disp'] not in ['N/A', 'Exempt (1st Year)', 'N/A (Branch)'] else yd['ar_dt_disp']
                es_fee_dt_disp = yd['es_fee_dt'].replace('-', '/') if yd['es_fee_dt'] != 'N/A' else yd['es_fee_dt']
                es_dt_disp = yd['es_dt_disp'].replace('-', '/') if yd['es_dt_disp'] not in ['N/A', 'Exempt (1st Year)'] else yd['es_dt_disp']
                    
                text_color = "#2c3e50" if y == current_system_year else "#5d6d7e"
                font_weight = "bold" if y == current_system_year else "normal"
                
                if is_branch and cess_date and y >= cess_date.year:
                    dynamic_annual_rows += f"<tr><th colspan='2' style='text-align:center; color:#e74c3c; font-weight:bold; background-color:#fdedec;'>N/A (Branch Cessed)</th></tr>"
                else:
                    lbl_br = "Annual Fee Paid By" if is_bvi else "BR Paid By"
                    lbl_br_dt = "Annual Fee Paid Date" if is_bvi else "BR Paid Date"
                    lbl_ar_dt = "AFR Filed Date" if is_bvi else "AR Filed Date"
                    lbl_ar_st = "AFR Status" if is_bvi else "AR CR Status"
                    
                    dynamic_annual_rows += f"<tr><th>{lbl_br} (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{yd['br_by']}</td></tr>"
                    dynamic_annual_rows += f"<tr><th>{lbl_br_dt} (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{br_dt_disp}</td></tr>"
                    if is_bvi:
                        dynamic_annual_rows += f"<tr><th>AFR Fee Paid By (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{yd['afr_fee_by']}</td></tr>"
                        dynamic_annual_rows += f"<tr><th>AFR Fee Paid Date (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{afr_fee_dt_disp}</td></tr>"
                    dynamic_annual_rows += f"<tr><th>{lbl_ar_dt} (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{ar_dt_disp}</td></tr>"
                    if not is_bvi:
                        dynamic_annual_rows += f"<tr><th>{lbl_ar_st} (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{yd['ar_cr_disp']}</td></tr>"
                    else:
                        dynamic_annual_rows += f"<tr><th>ES Fee Paid By (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{yd['es_fee_by']}</td></tr>"
                        dynamic_annual_rows += f"<tr><th>ES Fee Paid Date (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{es_fee_dt_disp}</td></tr>"
                        dynamic_annual_rows += f"<tr><th>ES Filed Date (FY {y})</th><td style='color:{text_color}; font-weight:{font_weight};'>{es_dt_disp}</td></tr>"
        
        n2e_val = to_date(row.get('nd2a_eff_date'))
        n2f_val = to_date(row.get('nd2a_file_date'))
        n4e_val = to_date(row.get('nd4_eff_date'))
        n4f_val = to_date(row.get('nd4_file_date'))
        nn6_e_val = to_date(row.get('nn6_eff_date'))
        nn6_f_val = to_date(row.get('nn6_file_date'))

        has_hk_sec = bool(n2e_val or n2f_val or n4e_val or n4f_val)
        has_nonhk_sec = bool(nn6_e_val or nn6_f_val)

        dynamic_sec_rows = ""
        if place == 'HK' and not is_branch:
            if has_hk_sec:
                dynamic_sec_rows += f"""<div class="section-group"><div class="section-bar">Company Secretary Actions</div><table class="info-table"><tr><th>ND2A Eff Date</th><td>{fmt_date(row.get('nd2a_eff_date'))}</td></tr><tr><th>ND4 Eff Date</th><td>{fmt_date(row.get('nd4_eff_date'))}</td></tr></table></div>"""
        elif is_hk_reg and not is_branch:
            if has_nonhk_sec:
                dynamic_sec_rows += f"""<div class="section-group"><div class="section-bar">Non-HK Company Secretary Actions</div><table class="info-table"><tr><th>NN6 Eff Date</th><td>{fmt_date(row.get('nn6_eff_date'))}</td></tr></table></div>"""

        remark_val = str(row.get('remark', ''))
        if remark_val == 'None' or not remark_val: remark_val = 'No remarks.'

        this_cg = "" if hide_client_group else cg_row_html.replace('__CLIENT_GROUP__', str(row.get('client_group', '')))
        
        card_html = f"""<div class="company-container"><table class="main-table"><thead><tr><td><div class="header-content"><div class="name-en">{disp_en}</div><div class="name-ch">{str(ch_name)}</div>{biz_html}</div></td></tr></thead><tbody><tr><td><div class="company-card"><div class="section-group"><div class="section-bar">Registration Details</div><table class="info-table">{this_cg}<tr><th>Incorp Place</th><td>{display_place}</td></tr>{dynamic_place_rows}{dynamic_hk_rows}<tr><th>Company Type</th><td>{str(row.get('co_type', ''))}</td></tr></table></div><div class="section-group"><div class="section-bar">Addresses</div><table class="info-table"><tr><th>Registered Address</th><td>{str(row.get('reg_addr', ''))}</td></tr><tr><th>Correspondence Address</th><td>{str(row.get('corres_addr', ''))}</td></tr></table></div><div class="section-group"><div class="section-bar">Items Storage</div><table class="info-table"><tr><th>Round Stamp</th><td>{str(row.get('round_loc', ''))}</td></tr><tr><th>Signature Chop</th><td>{str(row.get('sign_loc', ''))}</td></tr><tr><th>Common Seal</th><td>{str(row.get('seal_loc', ''))}</td></tr></table></div><div class="section-group"><div class="section-bar">Remarks</div><div style="padding: 8px 15px; font-size: 10.5pt; color: #2c3e50; white-space: pre-wrap;">{remark_val}</div></div>{dynamic_sec_rows}<div class="section-group"><div class="section-bar">Compliance Filings (Yearly)</div><table class="info-table">{dynamic_annual_rows}</table></div></div></td></tr></tbody></table></div>"""
        full_doc = html_head + card_html + "</body></html>"
        docs.append(HTML(string=full_doc).render())

    if docs:
        all_pages = []
        for doc in docs: all_pages.extend(doc.pages)
        buf = io.BytesIO()
        docs[0].copy(all_pages).write_pdf(buf)
        return buf.getvalue()
    return b""

@st.cache_data(show_spinner=False)
def generate_general_excel(selected_df, hide_client_group=False):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Arial", size=10)
    font_red = Font(name="Arial", size=10, color="FF0000", bold=True)
    font_yellow = Font(name="Arial", size=10, color="FF9900", bold=True)
    font_green = Font(name="Arial", size=10, color="00B050", bold=True)
    font_grey = Font(name="Arial", size=10, color="7F8C8D", italic=True)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    groups = selected_df['client_group'].unique() if 'client_group' in selected_df.columns else ['Companies']
    groups = sorted([g for g in groups if pd.notna(g)])
    if not groups: groups = ['Companies']

    for g in groups:
        safe_title = str(g).replace("/", "_").replace("\\", "_").replace("[", "").replace("]", "").replace("*", "").replace(":", "").replace("?", "")[:31] or "Ungrouped"
        ws = wb.create_sheet(title=safe_title)
        ws.views.sheetView[0].showGridLines = True
        group_df = selected_df[selected_df['client_group'] == g] if 'client_group' in selected_df.columns else selected_df
        export_records = group_df.to_dict('records')
        processed_records = []
        
        for row in export_records:
            base_date = get_base_date(row)
            incorp_year = base_date.year if base_date else None
            branch = str(row.get('branch_code', '000')).strip()
            if branch in ['None', 'nan', '', '<NA>']: branch = '000'
            is_branch = branch != '000'
            cess_date = to_date(row.get('cessation_date'))
            is_bvi = str(row.get('incorp_place', '')) not in ['HK', ''] and not str(row.get('is_hk_registered', 'False')).strip().lower() in ['true', 'yes', '1']
            
            row['Branch Code'] = branch 
            row['Business Name'] = row.get('biz_name', '')
            br_no_raw = str(row.get('br_no', '')).strip()
            if br_no_raw:
                companies_with_branches = set(selected_df[selected_df['branch_code'] != '000']['name_en'])
                row['BR No.'] = f"{br_no_raw}-{branch}" if row.get('name_en') in companies_with_branches else br_no_raw
            else: row['BR No.'] = ''
                 
            comp_rec_str = str(row.get('compliance_records', '{}'))
            try: rec_dict = json.loads(comp_rec_str)
            except: rec_dict = {}
            if not isinstance(rec_dict, dict): rec_dict = {}
            
            prev_br_by = 'Firm'; prev_afr_fee_by = 'Firm'; prev_es_fee_by = 'Firm'
            
            for y in report_years:
                if incorp_year and y < incorp_year:
                    row[f'FY{y} Fee Paid By'] = 'N/A'; row[f'FY{y} AR/AFR Status'] = 'N/A'; row[f'FY{y} Fee Paid Date'] = ''
                    row[f'FY{y} AFR Fee Paid By'] = 'N/A'; row[f'FY{y} AFR Fee Paid Date'] = ''; row[f'FY{y} AR/AFR Filed Date'] = ''
                    row[f'FY{y} ES Fee Paid By'] = 'N/A'; row[f'FY{y} ES Fee Paid Date'] = ''; row[f'FY{y} ES Filed Date'] = ''; row[f'FY{y} ES Status'] = 'N/A'
                    prev_br_by, prev_afr_fee_by, prev_es_fee_by = 'N/A', 'N/A', 'N/A'
                    continue
                    
                y_str = str(y)
                y_data = rec_dict.get(y_str, {})
                
                raw_br_by = str(y_data.get('fee_by', y_data.get('br_paid_by', ''))).strip()
                br_by = raw_br_by if raw_br_by else (prev_br_by if prev_br_by != 'N/A' else 'Firm')
                
                if is_bvi:
                    raw_afr_fee_by = str(y_data.get('afr_fee_by', '')).strip()
                    afr_fee_by = raw_afr_fee_by if raw_afr_fee_by else (prev_afr_fee_by if prev_afr_fee_by != 'N/A' else 'Firm')
                    raw_es_fee_by = str(y_data.get('es_fee_by', '')).strip()
                    es_fee_by = raw_es_fee_by if raw_es_fee_by else (prev_es_fee_by if prev_es_fee_by != 'N/A' else 'Firm')
                else:
                    afr_fee_by = 'N/A'; es_fee_by = 'N/A'
                
                if is_branch and cess_date and y >= cess_date.year: br_by = "N/A"
                    
                prev_br_by = br_by; prev_afr_fee_by = afr_fee_by; prev_es_fee_by = es_fee_by
                
                row[f'FY{y} Fee Paid By'] = br_by
                row[f'FY{y} AFR Fee Paid By'] = afr_fee_by
                row[f'FY{y} ES Fee Paid By'] = es_fee_by
                
                ar_dt = str(y_data.get('ar_date', ''))
                if ar_dt in ['None', 'nan', '<NA>']: ar_dt = ''
                ar_dt_val = to_date(ar_dt)
                es_dt_val = to_date(y_data.get('es_date'))
                
                if is_bvi:
                    ar_dl = calc_afr_deadline(row.get('year_end'), y)
                    es_dl = calc_es_deadline(base_date, y)
                    ar_cr_status = get_stat(ar_dt_val, ar_dl, y == incorp_year)
                    es_status = get_stat(es_dt_val, es_dl, y == incorp_year)
                else:
                    ar_cr_status = y_data.get('ar_cr_status', 'Pending')
                    if not ar_cr_status: ar_cr_status = 'Completed' if ar_dt else 'Pending'
                    es_status = 'N/A'
                    
                if is_branch:
                    row[f'FY{y} AR/AFR Status'] = "N/A (Branch)"
                    row[f'FY{y} ES Status'] = "N/A (Branch)"
                else:
                    row[f'FY{y} AR/AFR Status'] = ar_cr_status
                    row[f'FY{y} ES Status'] = es_status
                
                br_d = str(y_data.get('fee_date', y_data.get('br_date', '')))
                if br_d in ['None', 'nan', '<NA>']: br_d = ''
                else: br_d = br_d.replace('-', '/')
                
                afr_fee_dt = str(y_data.get('afr_fee_date', '')) if is_bvi else ''
                if afr_fee_dt in ['None', 'nan', '<NA>']: afr_fee_dt = ''
                else: afr_fee_dt = afr_fee_dt.replace('-', '/')
                
                es_fee_dt = str(y_data.get('es_fee_date', '')) if is_bvi else ''
                if es_fee_dt in ['None', 'nan', '<NA>']: es_fee_dt = ''
                else: es_fee_dt = es_fee_dt.replace('-', '/')
                
                if ar_cr_status in ['Exempt (Dormant)', 'Included in Agent Fee']: ar_dt = ''
                elif ar_dt and not is_branch: ar_dt = ar_dt.replace('-', '/')
                elif is_branch: ar_dt = ''
                
                es_dt = str(y_data.get('es_date', '')) if is_bvi else ''
                if es_dt in ['None', 'nan', '<NA>']: es_dt = ''
                if es_status == 'Exempt' or not is_bvi: es_dt = ''
                elif es_dt and not is_branch: es_dt = es_dt.replace('-', '/')
                elif is_branch: es_dt = ''
                
                if is_branch and cess_date and y >= cess_date.year: row[f'FY{y} Fee Paid Date'] = ""
                else: row[f'FY{y} Fee Paid Date'] = br_d
                    
                row[f'FY{y} AFR Fee Paid Date'] = afr_fee_dt
                row[f'FY{y} AR/AFR Filed Date'] = ar_dt
                row[f'FY{y} ES Fee Paid Date'] = es_fee_dt
                row[f'FY{y} ES Filed Date'] = es_dt

            for col in ["incorp_date", "hk_incorp_date", "br_ref_date", "ar_ref_date", "cessation_date", "nd2a_eff_date", "nd2a_file_date", "nd4_eff_date", "nd4_file_date", "nn6_eff_date", "nn6_file_date", "dissolution_date"]:
                val = to_date(row.get(col))
                row[col] = val.strftime('%Y/%m/%d') if val else ""

            processed_records.append(row)

        df_export = pd.DataFrame(processed_records)
        col_mapping = {
            'client_group': 'Client Group', 'name_en': 'Company Name EN', 'name_ch': 'Company Name CH', 
            'Business Name': 'Business Name', 'incorp_place': 'Incorp Place', 'incorp_place_others': 'Incorp Place Others', 
            'incorp_date': 'Incorp Date', 'ci_no': 'CI No.', 'is_hk_registered': 'Non-HK Registered in HK', 
            'hk_incorp_date': 'HK Incorp Date', 'hk_ci_no': 'HK CI No.', 'BR No.': 'BR No.', 
            'Branch Code': 'Branch Code', 'co_type': 'Company Type', 'reg_addr': 'Registered Address', 
            'corres_addr': 'Correspondence Address', 'round_loc': 'Round Stamp', 'sign_loc': 'Signature Chop', 
            'seal_loc': 'Common Seal', 'br_ref_date': 'BR Ref Date', 'ar_ref_date': 'AR Ref Date', 
            'cessation_date': 'Cessation Date', 'agent': 'Registered Agent', 'year_end': 'Financial Year End', 'billing_mode': 'Billing Mode',
            'nd2a_eff_date': 'ND2A Eff Date', 'nd2a_file_date': 'ND2A File Date', 
            'nd2a_download': 'ND2A Download', 'nd4_eff_date': 'ND4 Eff Date', 'nd4_file_date': 'ND4 File Date', 
            'nd4_download': 'ND4 Download', 'nn6_eff_date': 'NN6 Eff Date', 'nn6_file_date': 'NN6 File Date', 
            'nn6_download': 'NN6 Download', 'dissolution_date': 'Dissolution Date', 'remark': 'Remark'
        }
        df_export.rename(columns=col_mapping, inplace=True, errors='ignore')
        base_cols_ordered = list(col_mapping.values())
        base_cols_ordered.remove('Remark')
        dyn_cols = []
        for y in report_years: dyn_cols.extend([f"FY{y} Fee Paid By", f"FY{y} Fee Paid Date", f"FY{y} AFR Fee Paid By", f"FY{y} AFR Fee Paid Date", f"FY{y} AR/AFR Filed Date", f"FY{y} AR/AFR Status", f"FY{y} ES Fee Paid By", f"FY{y} ES Fee Paid Date", f"FY{y} ES Filed Date", f"FY{y} ES Status"])
        
        headers = [c for c in base_cols_ordered if c in df_export.columns] + [c for c in dyn_cols if c in df_export.columns] + ['Remark']
        if hide_client_group and 'Client Group' in headers: headers.remove('Client Group')
            
        ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
        ws["A1"] = f"Company Report - Group: {g} (Generated on: {datetime.now(HKT).strftime('%Y/%m/%d %H:%M')})"
        ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F497D")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[3].height = 25
        
        current_row = 4
        for _, item in df_export.iterrows():
            for col_idx, h in enumerate(headers, 1):
                val = str(item.get(h, ""))
                if val in ['nan', 'None', '<NA>']: val = ""
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = font_data
                cell.border = thin_border
                if current_row % 2 == 0: cell.fill = fill_zebra
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if "Status" in h:
                    if "Overdue" in val or "Returned" in val: cell.font = font_red
                    elif "Due Soon" in val or "Processing" in val: cell.font = font_yellow
                    elif "Exempt" in val or "Pending" in val or "N/A" in val or "Included" in val: cell.font = font_grey
                    elif "Completed" in val: cell.font = font_green
            ws.row_dimensions[current_row].height = 20
            current_row += 1
            
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row == 1: continue 
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(min(max_len + 4, 50), 12)

    if len(wb.sheetnames) == 0: wb.create_sheet("No Data")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

@st.cache_data(show_spinner=False)
def generate_beautiful_excel(df, hide_client_group=False):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Arial", size=10)
    font_red = Font(name="Arial", size=10, color="FF0000", bold=True)
    font_yellow = Font(name="Arial", size=10, color="FF9900", bold=True)
    font_green = Font(name="Arial", size=10, color="00B050", bold=True)
    font_grey = Font(name="Arial", size=10, color="7F8C8D", italic=True)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    headers = ["Client Group", "Company Name EN", "Company Name CH", "Business Name", "Incorp Place", "Year", "Anniversary (MM/DD)", "BR No.", "Fee Paid By", "Fee Paid Date", "Fee DL", "Fee Status", "AR/AFR Fee By", "AR/AFR Fee Date", "AR/AFR Filed Date", "AR/AFR DL", "AR/AFR Status", "ES Fee By", "ES Fee Date", "ES Filed Date", "ES DL", "ES Status", "Remark"]
    if hide_client_group: headers.remove("Client Group")

    groups = df['Client Group'].unique() if 'Client Group' in df.columns else ['Outstanding']
    groups = sorted([g for g in groups if pd.notna(g)])
    if not groups: groups = ['Outstanding']

    for g in groups:
        safe_title = str(g).replace("/", "_").replace("\\", "_").replace("[", "").replace("]", "").replace("*", "").replace(":", "").replace("?", "")[:31] or "Ungrouped"
        ws = wb.create_sheet(title=safe_title)
        ws.views.sheetView[0].showGridLines = True
        ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
        ws["A1"] = f"Outstanding Report - Group: {g} (Generated on: {datetime.now(HKT).strftime('%Y/%m/%d %H:%M')})"
        ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F497D")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30

        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[3].height = 25

        center_cols = ["Client Group", "Incorp Place", "Year", "Anniversary (MM/DD)", "BR No.", "Fee Paid By", "Fee Paid Date", "Fee DL", "Fee Status", "AR/AFR Fee By", "AR/AFR Fee Date", "AR/AFR Filed Date", "AR/AFR DL", "AR/AFR Status", "ES Fee By", "ES Fee Date", "ES Filed Date", "ES DL", "ES Status"]
        group_df = df[df['Client Group'] == g] if 'Client Group' in df.columns else df
        
        current_row = 4
        for _, item in group_df.iterrows():
            row_dict = {
                "Client Group": item.get("Client Group", ""), "Company Name EN": item.get("Company Name EN", ""), "Company Name CH": item.get("Company Name CH", ""),
                "Business Name": item.get("Business Name", ""), "Incorp Place": item.get("Incorp Place", ""), "Year": str(item.get("Year", "")),
                "Anniversary (MM/DD)": item.get("Anniversary (MM/DD)", ""), "BR No.": item.get("BR No.", ""), "Fee Paid By": item.get("Fee Paid By", ""),
                "Fee Paid Date": item.get("Fee Paid Date", ""), "Fee DL": item.get("Fee Deadline", ""), "Fee Status": item.get("Fee Status", ""),
                "AR/AFR Fee By": item.get("AR/AFR Fee By", ""), "AR/AFR Fee Date": item.get("AR/AFR Fee Date", ""), "AR/AFR Filed Date": item.get("AR/AFR Filed Date", ""),
                "AR/AFR DL": item.get("AR/AFR Deadline", ""), "AR/AFR Status": item.get("AR/AFR Status", ""), "ES Fee By": item.get("ES Fee By", ""),
                "ES Fee Date": item.get("ES Fee Date", ""), "ES Filed Date": item.get("ES Filed Date", ""), "ES DL": item.get("ES Deadline", ""),
                "ES Status": item.get("ES Status", ""), "Remark": item.get("Remark", "")
            }
            for col_idx, h in enumerate(headers, 1):
                val_str = clean_status(row_dict[h]) if row_dict[h] else ""
                cell = ws.cell(row=current_row, column=col_idx, value=val_str)
                cell.font = font_data
                cell.border = thin_border
                if current_row % 2 == 0: cell.fill = fill_zebra
                if h in center_cols: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else: cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if "Status" in h:
                    if "Overdue" in val_str or "Returned" in val_str: cell.font = font_red
                    elif "Due Soon" in val_str or "Processing" in val_str: cell.font = font_yellow
                    elif "Exempt" in val_str or "Not Incorporated" in val_str or "Pending" in val_str or "Branch" in val_str or "Cessed" in val_str or "Included" in val_str: cell.font = font_grey
                    else: cell.font = font_green
            ws.row_dimensions[current_row].height = 20
            current_row += 1

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row == 1: continue 
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(min(max_len + 4, 50), 15)

    if len(wb.sheetnames) == 0: wb.create_sheet("No Data")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

@st.cache_data(show_spinner=False)
def generate_outstanding_pdf(df, hide_client_group=False):
    now_str = datetime.now(HKT).strftime("%Y/%m/%d %H:%M")
    colspan = "17" if hide_client_group else "18"
    cg_th = "" if hide_client_group else '<th style="width:4%">Client Group</th>'
    
    html = f"""<html><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700&display=swap" rel="stylesheet">
        <style>
            @page {{ size: A4 landscape; margin: 10mm; background-color: #ffffff; @bottom-left {{ content: "Outstanding Report | Generated on: {now_str}"; font-size: 8pt; color: #7f8c8d; font-family: 'Noto Sans TC', sans-serif; }} @bottom-right {{ content: counter(page) " of " counter(pages) " Page(s)"; font-size: 8pt; color: #7f8c8d; font-family: 'Noto Sans TC', sans-serif; }} }}
            body {{ font-family: 'Noto Sans TC', sans-serif; font-size: 7pt; color: #2c3e50; margin: 0; padding: 0; }}
            table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
            thead {{ display: table-header-group; }}
            tr {{ page-break-inside: avoid; }}
            th {{ background-color: #1f497d; color: white; padding: 3px 1px; border: 1px solid #d9d9d9; font-size: 6pt; text-align: center; vertical-align: middle; word-wrap: break-word; }}
            td {{ padding: 3px 1px; border: 1px solid #d9d9d9; text-align: center; vertical-align: middle; word-wrap: break-word; }}
            tr:nth-child(even) td {{ background-color: #f8f9fa; }}
            .text-left {{ text-align: left; font-weight: bold; color: #2980b9; padding-left: 3px; }}
            .report-header-cell {{ border: none !important; background-color: white !important; text-align: left !important; padding: 0 0 10px 0 !important; }}
        </style></head><body><table><thead><tr><td colspan="{colspan}" class="report-header-cell"><h2 style="color: #1f497d; margin: 0 0 5px 0;">Outstanding Report</h2></td></tr><tr>{cg_th}
        <th style="width:10%">Company Name EN</th><th style="width:3.5%">Place</th><th style="width:3%">Year</th><th style="width:4.5%">Anniv<br>(MM/DD)</th><th style="width:5%">BR No.</th><th style="width:4%">Fee By</th><th style="width:5.5%">Fee Date</th><th style="width:7%">Fee DL & Status</th><th style="width:4%">AFR Fee By</th><th style="width:5.5%">AFR Fee Dt</th><th style="width:5.5%">AFR Filed</th><th style="width:7%">AFR DL & Status</th><th style="width:4%">ES Fee By</th><th style="width:5.5%">ES Fee Dt</th><th style="width:5.5%">ES Filed</th><th style="width:7%">ES DL & Status</th><th>Remark</th></tr></thead><tbody>"""
    for _, r in df.iterrows():
        br_val = clean_status(r.get('Fee Status', ''))
        ar_val = clean_status(r.get('AR/AFR Status', ''))
        es_val = clean_status(r.get('ES Status', ''))
        def get_color(v):
            if "Overdue" in v or "Returned" in v: return "#ff0000"
            if "Due Soon" in v or "Processing" in v: return "#ff9900"
            if "Exempt" in v or "Not Incorporated" in v or "Pending" in v or "Branch" in v or "Cessed" in v or "Included" in v: return "#7f8c8d"
            return "#00b050"
            
        br_color = get_color(br_val); ar_color = get_color(ar_val); es_color = get_color(es_val)
        cg_td = "" if hide_client_group else f"<td>{r.get('Client Group', '')}</td>"
        name_en = r.get('Company Name EN', '')
        biz = str(r.get('Business Name', '')).strip()
        if biz and biz not in ['None', 'nan']: name_en += f"<br><span style='font-size: 5pt; color: #7f8c8d; font-weight: normal;'>Business Name: {biz}</span>"
        fee_dl_stat = f"{r.get('Fee Deadline', '')}<br><span style='color: {br_color}; font-size: 6pt;'>{br_val}</span>" if r.get('Fee Deadline') not in ["", "N/A"] else r.get('Fee Deadline', '')
        afr_dl_stat = f"{r.get('AR/AFR Deadline', '')}<br><span style='color: {ar_color}; font-size: 6pt;'>{ar_val}</span>" if r.get('AR/AFR Deadline') not in ["", "N/A"] else r.get('AR/AFR Deadline', '')
        es_dl_stat = f"{r.get('ES Deadline', '')}<br><span style='color: {es_color}; font-size: 6pt;'>{es_val}</span>" if r.get('ES Deadline') not in ["", "N/A"] else r.get('ES Deadline', '')
        html += f"<tr>{cg_td}<td class='text-left'>{name_en}</td><td>{r.get('Incorp Place', '')}</td><td style='font-weight: bold; color: #1f497d;'>{r.get('Year', '')}</td><td style='font-weight: bold;'>{r.get('Anniversary (MM/DD)', '')}</td><td style='font-weight: bold;'>{r.get('BR No.', '')}</td><td>{r.get('Fee Paid By', '')}</td><td>{r.get('Fee Paid Date', '')}</td><td>{fee_dl_stat}</td><td>{r.get('AR/AFR Fee By', '')}</td><td>{r.get('AR/AFR Fee Date', '')}</td><td>{r.get('AR/AFR Filed Date', '')}</td><td>{afr_dl_stat}</td><td>{r.get('ES Fee By', '')}</td><td>{r.get('ES Fee Date', '')}</td><td>{r.get('ES Filed Date', '')}</td><td>{es_dl_stat}</td><td style='text-align: left; font-size: 6pt; color: #7f8c8d;'>{r.get('Remark', '')}</td></tr>"
    html += "</tbody></table></body></html>"
    return HTML(string=html).write_pdf()

@st.cache_data(show_spinner=False)
def generate_inv_excel(df, year, month_disp, hide_client_group=False):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Arial", size=10)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    headers = ["Client Group", "Company Name EN", "Company Name CH", "Business Name", "Incorp Place", "Year", "Anniversary (MM/DD)", "BR No.", "Fee Paid By", "Billing Item", "Fee Deadline", "AR/AFR Deadline", "ES Deadline", "Remark"]
    if hide_client_group: headers.remove("Client Group")

    groups = df['Client Group'].unique() if 'Client Group' in df.columns else ['Invoicing']
    groups = sorted([g for g in groups if pd.notna(g)])
    if not groups: groups = ['Invoicing']

    for g in groups:
        safe_title = str(g).replace("/", "_").replace("\\", "_").replace("[", "").replace("]", "").replace("*", "").replace(":", "").replace("?", "")[:31] or "Ungrouped"
        ws = wb.create_sheet(title=safe_title)
        ws.views.sheetView[0].showGridLines = True
        ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
        ws["A1"] = f"Invoicing Schedule Report ({year} - Months: {month_disp}) - Group: {g} (Generated on: {datetime.now(HKT).strftime('%Y/%m/%d %H:%M')})"
        ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1F497D")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 30

        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[3].height = 25

        center_cols = ["Client Group", "Incorp Place", "Year", "Anniversary (MM/DD)", "BR No.", "Fee Paid By", "Billing Item", "Fee Deadline", "AR/AFR Deadline", "ES Deadline"]
        group_df = df[df['Client Group'] == g] if 'Client Group' in df.columns else df

        current_row = 4
        for _, item in group_df.iterrows():
            row_dict = {
                "Client Group": item.get("Client Group", ""), "Company Name EN": item.get("Company Name EN", ""), "Company Name CH": item.get("Company Name CH", ""),
                "Business Name": item.get("Business Name", ""), "Incorp Place": item.get("Incorp Place", ""), "Year": str(item.get("Year", "")),
                "Anniversary (MM/DD)": item.get("Anniversary (MM/DD)", ""), "BR No.": item.get("BR No.", ""), "Fee Paid By": item.get("Fee Paid By", ""),
                "Billing Item": item.get("Billing Item", ""), "Fee Deadline": item.get("Fee Deadline", ""), "AR/AFR Deadline": item.get("AR/AFR Deadline", ""),
                "ES Deadline": item.get("ES Deadline", ""), "Remark": item.get("Remark", "")
            }
            for col_idx, h in enumerate(headers, 1):
                val = str(row_dict[h])
                cell = ws.cell(row=current_row, column=col_idx, value=clean_status(val) if h == "Billing Item" else val)
                cell.font = font_data
                cell.border = thin_border
                if current_row % 2 == 0: cell.fill = fill_zebra
                
                if h in center_cols: cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else: cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                
                if h == "Billing Item":
                    cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                    if "Incorp" in val: cell.fill = PatternFill(start_color="9B59B6", end_color="9B59B6", fill_type="solid")
                    elif "AR + BR" in val or "Package" in val: cell.fill = PatternFill(start_color="E67E22", end_color="E67E22", fill_type="solid")
                    elif "AR Fee Only" in val or "AFR" in val or "ES" in val or "Itemized" in val: cell.fill = PatternFill(start_color="2980B9", end_color="2980B9", fill_type="solid")
                    elif "Branch" in val: cell.fill = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
                    elif "Cessed" in val: cell.fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")
                    else:
                        cell.font = font_data
                        if current_row % 2 == 0: cell.fill = fill_zebra
                        else: cell.fill = PatternFill(fill_type=None)
            ws.row_dimensions[current_row].height = 20
            current_row += 1

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row == 1: continue 
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(min(max_len + 4, 50), 15)

    if len(wb.sheetnames) == 0: wb.create_sheet("No Data")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

@st.cache_data(show_spinner=False)
def generate_inv_pdf(df, year, month_disp, hide_client_group=False):
    now_str = datetime.now(HKT).strftime("%Y/%m/%d %H:%M")
    colspan = "10" if hide_client_group else "11"
    cg_th = "" if hide_client_group else '<th style="width:8%">Client Group</th>'
    
    html = f"""<html><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700&display=swap" rel="stylesheet">
        <style>
            @page {{ size: A4 landscape; margin: 15mm; background-color: #ffffff; @bottom-left {{ content: "Invoicing Schedule Report | Generated on: {now_str}"; font-size: 8pt; color: #7f8c8d; font-family: 'Noto Sans TC', sans-serif; }} @bottom-right {{ content: counter(page) " of " counter(pages) " Page(s)"; font-size: 8pt; color: #7f8c8d; font-family: 'Noto Sans TC', sans-serif; }} }}
            body {{ font-family: 'Noto Sans TC', sans-serif; font-size: 8pt; color: #2c3e50; margin: 0; padding: 0; }}
            table {{ width: 100%; border-collapse: collapse; }}
            thead {{ display: table-header-group; }}
            tr {{ page-break-inside: avoid; }}
            th {{ background-color: #1f497d; color: white; padding: 5px 3px; border: 1px solid #d9d9d9; font-size: 7.5pt; text-align: center; vertical-align: middle; word-wrap: break-word; }}
            td {{ padding: 5px 3px; border: 1px solid #d9d9d9; text-align: center; vertical-align: middle; word-wrap: break-word; }}
            tr:nth-child(even) td {{ background-color: #f8f9fa; }}
            .text-left {{ text-align: left; font-weight: bold; color: #2980b9; }}
            .report-header-cell {{ border: none !important; background-color: white !important; text-align: left !important; padding: 0 0 10px 0 !important; }}
        </style></head><body><table><thead><tr><td colspan="{colspan}" class="report-header-cell"><h2 style="color: #1f497d; margin: 0 0 5px 0;">Invoicing Schedule Report ({year} - Months: {month_disp})</h2></td></tr><tr>{cg_th}
        <th style="width:14%">Company Name EN</th><th style="width:8%">Company Name CH</th><th style="width:5%">Place</th><th style="width:4%">Year</th><th style="width:8%">Anniversary<br>(MM/DD)</th><th style="width:8%">BR No.</th><th style="width:6%">Fee By</th><th style="width:12%">Billing Item</th><th style="width:8%">Fee Deadline</th><th style="width:8%">AFR/ES Deadline</th><th>Remark</th></tr></thead><tbody>"""
    for _, r in df.iterrows():
        cg_td = "" if hide_client_group else f"<td>{r.get('Client Group', '')}</td>"
        bill_raw = str(r.get('Billing Item', ''))
        bill_clean = clean_status(bill_raw)
        if "Incorp" in bill_raw: bill_html = f'<span style="background-color: #9b59b6; color: white; padding: 3px 6px; border-radius: 4px;">{bill_clean}</span>'
        elif "AR + BR" in bill_raw or "Package" in bill_raw: bill_html = f'<span style="background-color: #e67e22; color: white; padding: 3px 6px; border-radius: 4px;">{bill_clean}</span>'
        elif "AR Fee Only" in bill_raw or "AFR" in bill_raw or "ES" in bill_raw or "Itemized" in bill_raw: bill_html = f'<span style="background-color: #2980b9; color: white; padding: 3px 6px; border-radius: 4px;">{bill_clean}</span>'
        elif "Branch" in bill_raw: bill_html = f'<span style="background-color: #27ae60; color: white; padding: 3px 6px; border-radius: 4px;">{bill_clean}</span>'
        elif "Cessed" in bill_raw: bill_html = f'<span style="background-color: #e74c3c; color: white; padding: 3px 6px; border-radius: 4px;">{bill_clean}</span>'
        else: bill_html = bill_clean
            
        name_en = r.get('Company Name EN', '')
        biz = str(r.get('Business Name', '')).strip()
        if biz and biz not in ['None', 'nan']: name_en += f"<br><span style='font-size: 6.5pt; color: #7f8c8d; font-weight: normal;'>Business Name: {biz}</span>"
            
        ar_dl = str(r.get('AR/AFR Deadline', '')).strip()
        es_dl = str(r.get('ES Deadline', '')).strip()
        combo_dl = []
        if ar_dl and ar_dl not in ['N/A', '']: combo_dl.append(f"AFR: {ar_dl}")
        if es_dl and es_dl not in ['N/A', '']: combo_dl.append(f"ES: {es_dl}")
        combo_dl_str = "<br>".join(combo_dl) if combo_dl else "N/A"
        
        html += f"<tr>{cg_td}<td class='text-left'>{name_en}</td><td class='text-left'>{r.get('Company Name CH', '')}</td><td>{r.get('Incorp Place', '')}</td><td style='font-weight: bold; color: #1f497d;'>{r.get('Year', '')}</td><td style='font-weight: bold;'>{r.get('Anniversary (MM/DD)', '')}</td><td style='font-weight: bold;'>{r.get('BR No.', '')}</td><td>{r.get('Fee Paid By', '')}</td><td style='font-weight: bold;'>{bill_html}</td><td>{r.get('Fee Deadline', '')}</td><td>{combo_dl_str}</td><td style='text-align: left; font-size: 7.5pt; color: #7f8c8d;'>{r.get('Remark', '')}</td></tr>"
    html += "</tbody></table></body></html>"
    return HTML(string=html).write_pdf()

@st.cache_data(show_spinner=False)
def create_zip_pdfs(df, report_type="All", year=None, month_disp=None):
    now_d = datetime.now(HKT).strftime('%Y%m%d')
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        col_name = 'Client Group' if 'Client Group' in df.columns else 'client_group'
        groups = df[col_name].unique()
        for g in groups:
            group_df = df[df[col_name] == g]
            safe_g = str(g).replace("/", "_").replace("\\", "_")
            if not safe_g.strip(): safe_g = "Ungrouped"
            
            if report_type == "All": pdf_bytes = generate_custom_pdf(group_df, hide_client_group=True); filename = f"{safe_g}_Company_Report_{now_d}.pdf"
            elif report_type == "Outstanding": pdf_bytes = generate_outstanding_pdf(group_df, hide_client_group=True); filename = f"{safe_g}_Outstanding_Report_{now_d}.pdf"
            elif report_type == "Invoicing": pdf_bytes = generate_inv_pdf(group_df, year, month_disp, hide_client_group=True); filename = f"{safe_g}_Invoicing_Schedule_{year}_{month_disp}_{now_d}.pdf"
            zip_file.writestr(filename, pdf_bytes)
    return zip_buffer.getvalue()

@st.cache_data(show_spinner=False)
def create_zip_excels(df, report_type="All", year=None, month_disp=None):
    now_d = datetime.now(HKT).strftime('%Y%m%d')
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        col_name = 'Client Group' if 'Client Group' in df.columns else 'client_group'
        groups = df[col_name].unique()
        for g in groups:
            group_df = df[df[col_name] == g]
            safe_g = str(g).replace("/", "_").replace("\\", "_")
            if not safe_g.strip(): safe_g = "Ungrouped"
            
            if report_type == "All": excel_bytes = generate_general_excel(group_df, hide_client_group=True); filename = f"{safe_g}_Company_Report_{now_d}.xlsx"
            elif report_type == "Outstanding": excel_bytes = generate_beautiful_excel(group_df, hide_client_group=True); filename = f"{safe_g}_Outstanding_Report_{now_d}.xlsx"
            elif report_type == "Invoicing": excel_bytes = generate_inv_excel(group_df, year, month_disp, hide_client_group=True); filename = f"{safe_g}_Invoicing_Schedule_{year}_{month_disp}_{now_d}.xlsx"
            zip_file.writestr(filename, excel_bytes)
    return zip_buffer.getvalue()

# --- 5. Dashboard ---
if choice == "📊 Dashboard":
    st.header("📊 Compliance Overview")
    df_raw = pd.read_sql("SELECT * FROM companies", engine)
    df_raw['branch_code'] = df_raw['branch_code'].fillna('000').astype(str).replace(['', 'None', 'nan', '<NA>'], '000')
    df_raw['biz_name'] = df_raw['biz_name'].fillna('').astype(str).replace(['None', 'nan', '<NA>'], '')
    groups = pd.read_sql("SELECT group_name FROM client_groups", engine)['group_name'].tolist()
    sorted_groups = sorted([g for g in groups if isinstance(g, str)])
    
    if not df_raw.empty:
        for col in ["incorp_date", "hk_incorp_date", "nd2a_eff_date", "nd4_eff_date", "nd2a_file_date", "nd4_file_date", "nn6_eff_date", "nn6_file_date", "dissolution_date", "ar_ref_date", "br_ref_date", "cessation_date"]:
            if col in df_raw.columns: df_raw[col] = pd.to_datetime(df_raw[col], errors='coerce').dt.date
            
        today = datetime.now(HKT).date()
        outstanding_records = []
        updated_records = []
        raw_dict_list = df_raw.to_dict('records')
        companies_with_branches = set(df_raw[df_raw['branch_code'] != '000']['name_en'])
        
        for row in raw_dict_list:
            comp_rec_str = row.get('compliance_records')
            try: comp_rec = json.loads(comp_rec_str) if isinstance(comp_rec_str, str) else {}
            except: comp_rec = {}
            if not isinstance(comp_rec, dict): comp_rec = {}
            
            place = str(row.get('incorp_place', ''))
            is_hk_reg = str(row.get('is_hk_registered', 'False')).strip().lower() in ['true', 'yes', 'y', '1']
            is_bvi = place not in ['HK', ''] and not is_hk_reg
            base_date = get_base_date(row)
            incorp_year = base_date.year if base_date else None
            cess_date = to_date(row.get('cessation_date'))
            branch_code = str(row.get('branch_code', '000')).strip()
            is_branch = branch_code != '000'
            has_branch = row.get('name_en') in companies_with_branches
            
            prev_br_by = 'Firm'; prev_afr_fee_by = 'Firm'; prev_es_fee_by = 'Firm'
            
            for y in active_years:
                y_str = str(y)
                y_data = comp_rec.get(y_str, {})
                
                if incorp_year and y < incorp_year: br_by = 'N/A'; afr_fee_by = 'N/A'; es_fee_by = 'N/A'
                else:
                    raw_br_by = str(y_data.get('fee_by', y_data.get('br_paid_by', ''))).strip()
                    br_by = raw_br_by if raw_br_by else (prev_br_by if prev_br_by != 'N/A' else 'Firm')
                    if is_bvi:
                        raw_afr_fee_by = str(y_data.get('afr_fee_by', '')).strip()
                        afr_fee_by = raw_afr_fee_by if raw_afr_fee_by else (prev_afr_fee_by if prev_afr_fee_by != 'N/A' else 'Firm')
                        raw_es_fee_by = str(y_data.get('es_fee_by', '')).strip()
                        es_fee_by = raw_es_fee_by if raw_es_fee_by else (prev_es_fee_by if prev_es_fee_by != 'N/A' else 'Firm')
                    else:
                        afr_fee_by = 'N/A'; es_fee_by = 'N/A'
                
                if is_branch and cess_date and y >= cess_date.year: br_by = "N/A"
                    
                prev_br_by = br_by; prev_afr_fee_by = afr_fee_by; prev_es_fee_by = es_fee_by
                
                row[f'{y}_br_paid_by'] = br_by; row[f'{y}_afr_fee_by'] = afr_fee_by; row[f'{y}_es_fee_by'] = es_fee_by
                
                afr_fee_dt = to_date(y_data.get('afr_fee_date')) if is_bvi else None
                ar_dt_val = to_date(y_data.get('ar_date'))
                br_dt_val = to_date(y_data.get('fee_date', y_data.get('br_date')))
                es_dt_val = to_date(y_data.get('es_date')) if is_bvi else None
                es_fee_dt = to_date(y_data.get('es_fee_date')) if is_bvi else None
                
                if is_branch and cess_date and y >= cess_date.year: br_dt_val = None
                    
                row[f'{y}_br_date'] = br_dt_val; row[f'{y}_ar_date'] = ar_dt_val; row[f'{y}_es_date'] = es_dt_val
                row[f'{y}_afr_fee_date'] = afr_fee_dt; row[f'{y}_es_fee_date'] = es_fee_dt
                
                if is_bvi:
                    ar_dl = calc_afr_deadline(row.get('year_end'), y)
                    es_dl = calc_es_deadline(base_date, y)
                    cr_stat = get_stat(ar_dt_val, ar_dl, y == incorp_year)
                    es_stat = get_stat(es_dt_val, es_dl, y == incorp_year)
                else:
                    ar_dl = calc_ar_deadline(base_date, to_date(row.get('ar_ref_date')), y)
                    cr_stat = y_data.get('ar_cr_status', 'Pending')
                    if not cr_stat: cr_stat = 'Completed' if ar_dt_val else 'Pending'
                    es_stat = 'N/A'
                    
                row[f'{y}_ar_cr_status'] = cr_stat
                row[f'{y}_es_status'] = es_stat

            name = str(row.get('name_en', 'Unknown')).strip()
            name_ch = str(row.get('name_ch', ''))
            group = row.get('client_group', '')
            remark_val = str(row.get('remark', ''))
            if remark_val == 'None': remark_val = ""
            
            biz_name = str(row.get('biz_name', '')).strip()
            ar_ref_raw = to_date(row.get('ar_ref_date'))
            br_ref_raw = to_date(row.get('br_ref_date'))
            br_no_raw = str(row.get('br_no', '')).strip()
            if has_branch and br_no_raw: disp_br = f"{br_no_raw}-{branch_code}"
            else: disp_br = br_no_raw
            row['disp_br_no'] = disp_br
            updated_records.append(row)
            
            if not base_date: continue
            
            for y in report_years:
                y_str = str(y)
                if y < incorp_year: continue 
                
                br_by = str(row.get(f'{y}_br_paid_by', 'Firm'))
                last_br = to_date(row.get(f'{y}_br_date'))
                afr_fee_by = str(row.get(f'{y}_afr_fee_by', 'N/A'))
                afr_fee_dt = to_date(row.get(f'{y}_afr_fee_date'))
                last_ar = to_date(row.get(f'{y}_ar_date'))
                ar_cr_status = str(row.get(f'{y}_ar_cr_status', 'Pending'))
                es_fee_by = str(row.get(f'{y}_es_fee_by', 'N/A'))
                es_fee_dt = to_date(row.get(f'{y}_es_fee_date'))
                last_es = to_date(row.get(f'{y}_es_date'))
                es_status_val = str(row.get(f'{y}_es_status', 'Pending'))
                
                if is_bvi:
                    br_dl = calc_bvi_fee_deadline(base_date, y)
                    ar_dl = calc_afr_deadline(row.get('year_end'), y)
                    es_dl = calc_es_deadline(base_date, y)
                else:
                    if br_ref_raw: br_dl = get_anniv(y, br_ref_raw.month, br_ref_raw.day)
                    else: br_dl = get_anniv(y, base_date.month, base_date.day)
                    ar_dl = calc_ar_deadline(base_date, ar_ref_raw, y)
                    es_dl = None
                
                br_dl_str = br_dl.strftime('%Y/%m/%d') if br_dl else "N/A"
                ar_dl_str = ar_dl.strftime('%Y/%m/%d') if ar_dl else "N/A"
                es_dl_str = es_dl.strftime('%Y/%m/%d') if es_dl else "N/A"
                br_dt_str = last_br.strftime('%Y/%m/%d') if last_br else ""
                afr_fee_dt_str = afr_fee_dt.strftime('%Y/%m/%d') if afr_fee_dt else ""
                ar_dt_str = last_ar.strftime('%Y/%m/%d') if last_ar else ""
                es_fee_dt_str = es_fee_dt.strftime('%Y/%m/%d') if es_fee_dt else ""
                es_dt_str = last_es.strftime('%Y/%m/%d') if last_es else ""
                
                if is_branch and cess_date and y >= cess_date.year: br_dl_str = "N/A"
                if ar_cr_status in ['Exempt (Dormant)', 'Included in Agent Fee'] or is_branch or y == incorp_year: ar_dl_str = "N/A"
                if y == incorp_year or not is_bvi: es_dl_str = "N/A"
                
                if br_ref_raw and ar_ref_raw and (br_ref_raw.month != ar_ref_raw.month or br_ref_raw.day != ar_ref_raw.day):
                    anniv_disp = f"BR: {br_ref_raw.strftime('%m/%d')} | AR: {ar_ref_raw.strftime('%m/%d')}"
                elif br_ref_raw and not ar_ref_raw and (br_ref_raw.month != base_date.month or br_ref_raw.day != base_date.day):
                    anniv_disp = f"BR: {br_ref_raw.strftime('%m/%d')} | AR: {base_date.strftime('%m/%d')}"
                elif not br_ref_raw and ar_ref_raw and (base_date.month != ar_ref_raw.month or base_date.day != ar_ref_raw.day):
                    anniv_disp = f"BR: {base_date.strftime('%m/%d')} | AR: {ar_ref_raw.strftime('%m/%d')}"
                else:
                    anniv_disp = base_date.strftime('%m/%d')
                        
                is_alert = False
                br_status = "Normal"; ar_status = "Normal"
                
                if not last_br and br_dl and br_dl_str != "N/A" and br_by != 'N/A' and not (is_branch and cess_date and y >= cess_date.year):
                    days_diff = (br_dl - today).days
                    if days_diff < 0: br_status = "Overdue"
                    elif days_diff <= (90 if is_bvi else 30): br_status = "Due Soon"
                    else: br_status = "Pending"
                elif last_br: br_status = "Completed"
                elif br_by == 'N/A' or (is_branch and cess_date and y >= cess_date.year): br_status = "N/A"
                else: br_status = "Pending"
                    
                if br_status in ["Overdue", "Due Soon"]: is_alert = True
                
                if not is_bvi:
                    if ar_cr_status in ['Completed', 'Exempt (Dormant)', 'Included in Agent Fee']: ar_status = ar_cr_status
                    elif ar_cr_status == 'Processing': ar_status = "Processing"
                    elif ar_cr_status == 'Returned': ar_status = "Returned"
                    elif y == incorp_year: ar_status = "Exempt"
                    elif not last_ar and ar_dl and ar_dl_str != "N/A":
                        ar_days_diff = (ar_dl - today).days
                        if ar_days_diff < 0: ar_status = "Overdue"
                        elif ar_days_diff <= 72: ar_status = "Due Soon"
                        else: ar_status = "Pending"
                    elif last_ar: ar_status = "Completed"
                    else: ar_status = "N/A"
                else:
                    if last_ar: ar_status = "Completed"
                    elif y == incorp_year: ar_status = "Exempt"
                    elif not last_ar and ar_dl and ar_dl_str != "N/A":
                        ar_days_diff = (ar_dl - today).days
                        if ar_days_diff < 0: ar_status = "Overdue"
                        elif ar_days_diff <= 90: ar_status = "Due Soon"
                        else: ar_status = "Pending"
                    else: ar_status = "N/A"
                        
                if ar_status in ["Overdue", "Due Soon", "Processing", "Returned"]: is_alert = True
                
                if is_bvi:
                    if es_status_val == 'Exempt' or y == incorp_year: es_status_disp = "Exempt"
                    elif es_status_val == 'Completed' or last_es: es_status_disp = "Completed"
                    elif not last_es and es_dl and es_dl_str != "N/A":
                        es_days_diff = (es_dl - today).days
                        if es_days_diff < 0: es_status_disp = "Overdue"
                        elif es_days_diff <= 90: es_status_disp = "Due Soon"
                        else: es_status_disp = "Pending"
                    else: es_status_disp = "N/A"
                else:
                    es_status_disp = "N/A"
                    
                if es_status_disp in ["Overdue", "Due Soon"]: is_alert = True
                        
                if is_alert:
                    disp_name = f"{name} (-{branch_code})" if has_branch and is_branch else name
                    outstanding_records.append({
                        "Company Name EN": disp_name, "Company Name CH": name_ch, "Business Name": biz_name, "Client Group": group,
                        "Incorp Place": place, "Year": y_str, "Anniversary (MM/DD)": anniv_disp, "BR No.": disp_br,
                        "Fee Paid By": br_by, "Fee Paid Date": br_dt_str, "Fee Deadline": br_dl_str, "Fee Status": br_status,
                        "AR/AFR Fee By": afr_fee_by if is_bvi else "N/A", "AR/AFR Fee Date": afr_fee_dt_str if is_bvi else "N/A",
                        "AR/AFR Filed Date": ar_dt_str, "AR/AFR Deadline": ar_dl_str, "AR/AFR Status": ar_status,
                        "ES Fee By": es_fee_by if is_bvi else "N/A", "ES Fee Date": es_fee_dt_str if is_bvi else "N/A",
                        "ES Filed Date": es_dt_str if is_bvi else "N/A", "ES Deadline": es_dl_str, "ES Status": es_status_disp,
                        "Remark": remark_val, "branch_code_raw": branch_code
                    })

        df_raw = pd.DataFrame(updated_records)
        tab1, tab2, tab3 = st.tabs(["📊 All Companies", "🚨 Outstanding List", "🧾 Invoicing Schedule"])
        
        with tab1:
            sort_cols = [c for c in ['client_group', 'name_en', 'branch_code', 'incorp_place'] if c in df_raw.columns]
            df_raw = df_raw.sort_values(by=sort_cols, na_position='last')
            
            t1, t2, t3, t4, t5 = st.columns([2, 2, 2, 2, 4])
            filter_g = t1.selectbox("🔍 Filter Group", ["All Groups"] + sorted_groups)
            target_year_disp = t2.selectbox("📅 Display Target Year", active_years, index=active_years.index(current_system_year))
            
            if t3.button("🔄 Refresh"): st.rerun()
            if 'sel_v130' not in st.session_state: st.session_state.sel_v130 = False
            if t4.button("✅ Select All"): st.session_state.sel_v130 = True; st.rerun()
            if t5.button("🧹 Clear All"): st.session_state.sel_v130 = False; st.rerun()
            
            disp_years = [y for y in active_years if target_year_disp - 1 <= y <= target_year_disp + 1]
            df_filtered = df_raw if filter_g == "All Groups" else df_raw[df_raw['client_group'] == filter_g]
            
            display_cols_ordered = [
                "name_ch", "biz_name", "client_group", "incorp_place", "incorp_place_others", 
                "incorp_date", "ci_no", "is_hk_registered", "hk_incorp_date", "hk_ci_no", 
                "disp_br_no", "co_type", "agent", "year_end", "billing_mode", "reg_addr", "corres_addr", "round_loc", "sign_loc", "seal_loc"
            ]
            
            base_cols = [c for c in TEMPLATE_COLS if c in df_filtered.columns]
            for c in base_cols:
                if c not in display_cols_ordered and c not in ["remark", "name_en", "branch_code", "ar_ref_date", "br_ref_date", "cessation_date", "br_no"]: 
                    display_cols_ordered.append(c)
            
            dyn_cols = []
            for y in disp_years:
                dyn_cols.extend([f"{y}_br_paid_by", f"{y}_br_date", f"{y}_afr_fee_by", f"{y}_afr_fee_date", f"{y}_ar_date", f"{y}_ar_cr_status", f"{y}_es_fee_by", f"{y}_es_fee_date", f"{y}_es_date", f"{y}_es_status"])
            
            display_cols_ordered.extend(dyn_cols)
            display_cols_ordered.append("remark")
            display_cols_ordered.append("branch_code")
            
            df_display = df_filtered[["name_en"] + [c for c in display_cols_ordered if c in df_filtered.columns]].copy()
            
            if not df_display.empty:
                df_display['remark'] = df_display['remark'].fillna('')
                def format_name(r):
                    if str(r['branch_code']).strip() != '000' and r['name_en'] in companies_with_branches: return f"{r['name_en']} (-{str(r['branch_code']).strip()})"
                    return r['name_en']
                df_display['name_en'] = df_display.apply(format_name, axis=1)
                df_display.rename(columns=EXCHANGE_COL_MAPPING, inplace=True, errors='ignore')
                df_display.rename(columns={'disp_br_no': 'BR No.'}, inplace=True, errors='ignore')
                
                dyn_rename_dict = {}
                for y in disp_years:
                    dyn_rename_dict[f"{y}_br_paid_by"] = f"FY{y} Fee Paid By"
                    dyn_rename_dict[f"{y}_br_date"] = f"FY{y} Fee Paid Date"
                    dyn_rename_dict[f"{y}_afr_fee_by"] = f"FY{y} AFR Fee Paid By"
                    dyn_rename_dict[f"{y}_afr_fee_date"] = f"FY{y} AFR Fee Paid Date"
                    dyn_rename_dict[f"{y}_ar_date"] = f"FY{y} AR/AFR Filed Date"
                    dyn_rename_dict[f"{y}_ar_cr_status"] = f"FY{y} AR/AFR Status"
                    dyn_rename_dict[f"{y}_es_fee_by"] = f"FY{y} ES Fee Paid By"
                    dyn_rename_dict[f"{y}_es_fee_date"] = f"FY{y} ES Fee Paid Date"
                    dyn_rename_dict[f"{y}_es_date"] = f"FY{y} ES Filed Date"
                    dyn_rename_dict[f"{y}_es_status"] = f"FY{y} ES Status"
                df_display.rename(columns=dyn_rename_dict, inplace=True)
                df_display.insert(0, "Select", st.session_state.sel_v130)
                s = df_display["Company Name EN"].astype(str)
                df_display.index = s + s.groupby(s).cumcount().map(lambda x: '​' * x)
                df_display.index.name = "Company Name EN"
                df_display.drop(columns=["Company Name EN"], inplace=True)
                
                st.markdown(f"📈 Total: **{len(df_filtered)}** records. Only showing columns for years: **{', '.join(map(str, disp_years))}**.")
                
                col_cfg = {"Select": st.column_config.CheckboxColumn("Select", default=False), "Branch Code": None, "Remark": st.column_config.TextColumn("Remark")}
                cr_opts = ["Pending", "Processing", "Returned", "Completed", "Exempt (Dormant)", "Included in Agent Fee"]
                pay_opts = ["Firm", "Client", "N/A"]
                for y in disp_years:
                    col_cfg[f"FY{y} Fee Paid By"] = st.column_config.SelectboxColumn(f"✏️ FY{y} Fee Paid By", options=pay_opts, required=True)
                    col_cfg[f"FY{y} Fee Paid Date"] = st.column_config.DateColumn(f"✏️ FY{y} Fee Paid Date", format="YYYY/MM/DD")
                    col_cfg[f"FY{y} AFR Fee Paid By"] = st.column_config.SelectboxColumn(f"✏️ FY{y} AFR Fee Paid By", options=pay_opts, required=True)
                    col_cfg[f"FY{y} AFR Fee Paid Date"] = st.column_config.DateColumn(f"✏️ FY{y} AFR Fee Paid Date", format="YYYY/MM/DD")
                    col_cfg[f"FY{y} AR/AFR Filed Date"] = st.column_config.DateColumn(f"✏️ FY{y} AR/AFR Filed Date", format="YYYY/MM/DD")
                    col_cfg[f"FY{y} AR/AFR Status"] = st.column_config.SelectboxColumn(f"✏️ FY{y} AR/AFR Status", options=cr_opts, required=True)
                    col_cfg[f"FY{y} ES Fee Paid By"] = st.column_config.SelectboxColumn(f"✏️ FY{y} ES Fee Paid By", options=pay_opts, required=True)
                    col_cfg[f"FY{y} ES Fee Paid Date"] = st.column_config.DateColumn(f"✏️ FY{y} ES Fee Paid Date", format="YYYY/MM/DD")
                    col_cfg[f"FY{y} ES Filed Date"] = st.column_config.DateColumn(f"✏️ FY{y} ES Filed Date", format="YYYY/MM/DD")
                    col_cfg[f"FY{y} ES Status"] = st.column_config.TextColumn(f"FY{y} ES Status (Auto)", disabled=True)
                
                disabled_cols = [c for c in df_display.columns if c not in ["Select", "Remark"] and not any(c.endswith(suffix) for suffix in ["Fee Paid By", "Fee Paid Date", "AFR Fee Paid By", "AFR Fee Paid Date", "AR/AFR Filed Date", "AR/AFR Status", "ES Fee Paid By", "ES Fee Paid Date", "ES Filed Date", "ES Status"])]
                
                edit_df = st.data_editor(df_display, column_config=col_cfg, disabled=disabled_cols, use_container_width=True, key="dash_v130")
                
                if st.button("💾 Save Batch Edits", key="btn_save_grid_v130"):
                    try:
                        with engine.begin() as conn:
                            for c_name_idx, r in edit_df.iterrows():
                                c_n = str(c_name_idx).replace('​', '')
                                b_code = str(r['Branch Code'])
                                suffix = f" (-{b_code})"
                                if c_n.endswith(suffix): c_n = c_n[:-len(suffix)]
                                
                                row_info = df_raw[(df_raw['name_en'] == c_n) & (df_raw['branch_code'] == b_code)].iloc[0]
                                comp_name_safe = str(row_info['name_en']).replace("'", "''")
                                branch_safe = str(row_info['branch_code']).replace("'", "''")
                                base_dt = get_base_date(row_info)
                                inc_yr = base_dt.year if base_dt else None
                                is_bvi = str(row_info['incorp_place']) not in ['HK', ''] and str(row_info['is_hk_registered']).lower() not in ['true', 'yes', '1']
                                b_mode = str(row_info.get('billing_mode', 'Itemized'))
                                
                                try: comp_dict_existing = json.loads(row_info.get('compliance_records', '{}'))
                                except: comp_dict_existing = {}
                                if not isinstance(comp_dict_existing, dict): comp_dict_existing = {}
                                
                                comp_dict = {}; prev_br_by = 'Firm'; prev_afr_fee_by = 'Firm'; prev_es_fee_by = 'Firm'
                                
                                for y in active_years:
                                    y_str = str(y)
                                    if y in disp_years:
                                        raw_br_by = str(r.get(f'FY{y} Fee Paid By', '')).strip()
                                        raw_afr = str(r.get(f'FY{y} AFR Fee Paid By', '')).strip()
                                        raw_es = str(r.get(f'FY{y} ES Fee Paid By', '')).strip()
                                        in_ar_cr = str(r.get(f'FY{y} AR/AFR Status', '')).strip()
                                        
                                        br_date_val = r.get(f'FY{y} Fee Paid Date')
                                        afr_fee_date_val = r.get(f'FY{y} AFR Fee Paid Date')
                                        ar_date_val = r.get(f'FY{y} AR/AFR Filed Date')
                                        es_fee_date_val = r.get(f'FY{y} ES Fee Paid Date')
                                        es_date_val = r.get(f'FY{y} ES Filed Date')
                                    else:
                                        db_y_data = comp_dict_existing.get(y_str, {})
                                        raw_br_by = str(db_y_data.get('fee_by', db_y_data.get('br_paid_by', ''))).strip()
                                        raw_afr = str(db_y_data.get('afr_fee_by', '')).strip()
                                        raw_es = str(db_y_data.get('es_fee_by', '')).strip()
                                        in_ar_cr = str(db_y_data.get('ar_cr_status', 'Pending')).strip()
                                        
                                        br_date_val = db_y_data.get('fee_date', db_y_data.get('br_date'))
                                        afr_fee_date_val = db_y_data.get('afr_fee_date')
                                        ar_date_val = db_y_data.get('ar_date')
                                        es_fee_date_val = db_y_data.get('es_fee_date')
                                        es_date_val = db_y_data.get('es_date')
                                    
                                    br_by = raw_br_by if raw_br_by else (prev_br_by if prev_br_by != 'N/A' else 'Firm')
                                    new_ar_cr = in_ar_cr if in_ar_cr in ["Pending", "Processing", "Returned", "Completed", "Exempt (Dormant)", "Included in Agent Fee"] else "Pending"
                                    
                                    if is_bvi and 'All-in' in b_mode: afr_fee_by = br_by; raw_afr_fee = to_date(br_date_val)
                                    else: afr_fee_by = raw_afr if raw_afr else (prev_afr_fee_by if prev_afr_fee_by != 'N/A' else 'Firm'); raw_afr_fee = to_date(afr_fee_date_val)
                                        
                                    es_fee_by = raw_es if raw_es else (prev_es_fee_by if prev_es_fee_by != 'N/A' else 'Firm')
                                    
                                    if not is_bvi: afr_fee_by = 'N/A'; es_fee_by = 'N/A'; raw_afr_fee = None; raw_es_fee = None; raw_es = None
                                    
                                    prev_br_by = br_by; prev_afr_fee_by = afr_fee_by; prev_es_fee_by = es_fee_by
                                    raw_br = to_date(br_date_val); raw_ar = to_date(ar_date_val)
                                    raw_es_fee = to_date(es_fee_date_val); raw_es = to_date(es_date_val)
                                    
                                    if inc_yr and y < inc_yr:
                                        br_by, afr_fee_by, es_fee_by = 'N/A', 'N/A', 'N/A'
                                        raw_br, raw_afr_fee, raw_ar, raw_es_fee, raw_es = None, None, None, None, None
                                        new_ar_cr = "Pending"
                                    elif inc_yr and y == inc_yr:
                                        raw_ar, raw_es = None, None
                                        new_ar_cr = "Pending"
                                        
                                    if br_by == 'N/A': raw_br = None
                                    if afr_fee_by == 'N/A': raw_afr_fee = None
                                    if es_fee_by == 'N/A': raw_es_fee = None
                                    if new_ar_cr in ['Exempt (Dormant)', 'Included in Agent Fee']: raw_ar = None
                                        
                                    comp_dict[y_str] = {
                                        "br_paid_by": br_by, "br_date": raw_br.strftime('%Y-%m-%d') if raw_br else None,
                                        "afr_fee_by": afr_fee_by, "afr_fee_date": raw_afr_fee.strftime('%Y-%m-%d') if raw_afr_fee else None,
                                        "ar_date": raw_ar.strftime('%Y-%m-%d') if raw_ar else None, "ar_cr_status": new_ar_cr,
                                        "es_fee_by": es_fee_by, "es_fee_date": raw_es_fee.strftime('%Y-%m-%d') if raw_es_fee else None,
                                        "es_date": raw_es.strftime('%Y-%m-%d') if raw_es else None
                                    }
                                json_str = json.dumps(comp_dict).replace("'", "''")
                                rem_str = str(r.get('Remark', '')).replace("'", "''")
                                if rem_str == 'None': rem_str = ""
                                sql = f"UPDATE companies SET compliance_records = '{json_str}', remark = '{rem_str}' WHERE name_en = '{comp_name_safe}' AND branch_code = '{branch_safe}'"
                                conn.execute(text(sql))
                        st.success("✅ Changes saved successfully!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Batch save failed: {e}")
