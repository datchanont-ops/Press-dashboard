import streamlit as st
import pandas as pd
import numpy as np
import os
import plotly.express as px
from io import BytesIO

# ---- Page Config ----
st.set_page_config(page_title="Production Shortage Pro", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# 🎨 CSS Enterprise / Executive Dashboard Styling
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Prompt:wght@300;400;500;600;700&display=swap');

:root {
    --navy:#0f172a; --slate:#475569; --muted:#64748b;
    --line:#e2e8f0; --surface:#ffffff; --bg:#f1f5f9;
    --blue:#2563eb; --red:#dc2626; --amber:#d97706; --green:#059669;
}
* { font-family:'Prompt','Inter',sans-serif !important; }
.stApp { background:#f1f5f9 !important; }
#MainMenu, footer { visibility:hidden; }
header[data-testid="stHeader"] { background:transparent !important; }
.block-container {
    padding-top:1.25rem !important;
    padding-bottom:2rem !important;
    max-width:1500px !important;
}

/* Executive Header */
.dashboard-header {
    background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);
    padding:1.45rem 1.7rem;
    border-radius:16px;
    box-shadow:0 10px 28px rgba(15,23,42,.12);
    border:1px solid rgba(255,255,255,.08);
    margin-bottom:.75rem;
    min-height:105px;
}
.dashboard-header h1 {
    color:#fff !important;
    font-size:27px !important;
    font-weight:800 !important;
    margin:0 !important;
    letter-spacing:-.02em;
}
.dashboard-header p {
    color:#cbd5e1 !important;
    font-size:13px !important;
    margin:5px 0 0 !important;
}

/* Controls */
[data-testid="stDateInput"],
[data-testid="stNumberInput"],
[data-testid="stFileUploader"] {
    background:#fff;
    border-radius:10px;
}
p.input-label {
    font-weight:600;
    color:#334155;
    margin:0 0 5px 2px;
    font-size:13px;
}
.stButton > button, .stDownloadButton > button {
    border-radius:9px !important;
    font-weight:600 !important;
    min-height:38px !important;
    border:1px solid #cbd5e1 !important;
    transition:all .18s ease !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform:translateY(-1px);
    box-shadow:0 5px 14px rgba(15,23,42,.10);
}

/* KPI Cards */
[data-testid="stMetric"] {
    background:#fff !important;
    padding:17px 20px !important;
    border-radius:14px !important;
    box-shadow:0 5px 18px rgba(15,23,42,.06) !important;
    border:1px solid #e2e8f0 !important;
    position:relative !important;
    overflow:hidden !important;
    transition:transform .18s ease,box-shadow .18s ease !important;
    min-height:112px;
}
[data-testid="stMetric"]:hover {
    transform:translateY(-2px) !important;
    box-shadow:0 9px 25px rgba(15,23,42,.10) !important;
}
[data-testid="stMetric"]::before {
    content:'';
    position:absolute;
    left:0; top:0; width:5px; height:100%;
}
div[data-testid="column"]:nth-child(1) [data-testid="stMetric"]::before { background:#dc2626; }
div[data-testid="column"]:nth-child(2) [data-testid="stMetric"]::before { background:#d97706; }
div[data-testid="column"]:nth-child(3) [data-testid="stMetric"]::before { background:#059669; }
[data-testid="stMetricLabel"] > div {
    color:#64748b !important;
    font-size:13px !important;
    font-weight:600 !important;
}
[data-testid="stMetricValue"] > div {
    color:#0f172a !important;
    font-size:29px !important;
    font-weight:800 !important;
    letter-spacing:-.025em !important;
}

/* Sections / Alerts / Tables */
h3,h4 { color:#0f172a !important; font-weight:700 !important; letter-spacing:-.01em; }
.alert-box {
    background:#fff7ed;
    color:#9a3412;
    padding:.85rem 1.1rem;
    border-radius:10px;
    margin-bottom:1rem;
    font-weight:500;
    border:1px solid #fed7aa;
    border-left:5px solid #f97316;
}
div[data-testid="stDataFrame"] {
    background:#fff;
    border-radius:14px;
    padding:5px;
    border:1px solid #e2e8f0;
    box-shadow:0 5px 18px rgba(15,23,42,.05);
}
[data-testid="stTextInput"] > div > div {
    border-radius:10px !important;
    background:#fff !important;
    border:1px solid #cbd5e1 !important;
}
[data-testid="stTextInput"] input { font-size:14px !important; }
[data-testid="stExpander"] {
    background:#fff !important;
    border:1px solid #e2e8f0 !important;
    border-radius:10px !important;
}
[data-testid="stAlert"] { border-radius:10px !important; }
hr {
    border:none !important;
    border-top:1px solid #e2e8f0 !important;
    margin:1.1rem 0 !important;
}
@media (max-width:900px) {
    .dashboard-header { padding:1.1rem 1.2rem; }
    .dashboard-header h1 { font-size:22px !important; }
}
</style>
""", unsafe_allow_html=True)

# ---- Data Processing Logic ----
try:
    current_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    current_dir = os.getcwd()

# กำหนด Path สำหรับเซฟข้อมูลในโฟลเดอร์ปัจจุบัน
SAVED_DB_PATH = os.path.join(current_dir, 'saved_database_upload.xlsx')
TEMPLATE_FILENAME = "@2.daily check aug26 ทุกวัน.XLSX"
TEMPLATE_PATH = os.path.join(current_dir, TEMPLATE_FILENAME)

@st.cache_data
def load_template_data(path):
    if not os.path.exists(path):
        return None, None, None
    df_check = pd.read_excel(path, sheet_name="check dali wipday", header=2)
    df_check = df_check.dropna(subset=['Material'])
    df_ord = pd.read_excel(path, sheet_name="ord", header=0)
    try:
        df_fo = pd.read_excel(path, sheet_name="fo", header=0)
    except:
        df_fo = pd.DataFrame() 
    return df_check, df_ord, df_fo

df_check_bg, df_ord_bg, df_fo_bg = load_template_data(TEMPLATE_PATH)

@st.cache_data
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Shortage Report')
    return output.getvalue()

@st.cache_data
def generate_example_db_template():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame({'Material': ['xxx1'], 'Unrestricted': [0]}).to_excel(writer, index=False, sheet_name='dali wip-fg')
        pd.DataFrame({'SAP Mat.': ['xxx'], 'Outstd.Base Qty': [20], 'Dlv. Date': ['2026-08-07']}).to_excel(writer, index=False, sheet_name='ord')
        pd.DataFrame({'Material': ['xxx1'], 'FO(Pcs)-08.2026': [0], 'ORD(Pcs)-08.2026': [20]}).to_excel(writer, index=False, sheet_name='fo')
        pd.DataFrame({0: ['xxxx'], 2: ['1/INJ_17/A/3/20260801']}).to_excel(writer, index=False, sheet_name='pro')
    return output.getvalue()

# ---- Header Section ----
col_title, col_date, col_workday, col_upload = st.columns([2.2, 0.8, 0.8, 1.2])

with col_title:
    st.markdown("""
        <div class="dashboard-header">
            <h1 style="margin:0; font-size:26px; color:#0f172a; font-weight:800;">📈 Production Shortage Dashboard</h1>
            <p style="margin:0; color:#64748b; font-size:14px; margin-top:6px;">ระบบประเมินความเสี่ยงและติดตามสถานะ B/O Date ระดับองค์กร</p>
        </div>
    """, unsafe_allow_html=True)

with col_date:
    st.markdown("<p class='input-label'>🗓️ ดู Balance ถึงวันที่</p>", unsafe_allow_html=True)
    target_date = st.date_input("", pd.to_datetime('2026-08-31'), label_visibility="collapsed")
    target_date_dt = pd.to_datetime(target_date)

with col_workday:
    st.markdown("<p class='input-label'>⏱️ Working Day (วัน)</p>", unsafe_allow_html=True)
    working_days_input = st.number_input("", min_value=1.0, value=20.0, step=1.0, label_visibility="collapsed")

with col_upload:
    st.markdown("<p class='input-label'>📂 อัปโหลดไฟล์ Database</p>", unsafe_allow_html=True)
    db_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
    
    active_db_file = None
    if db_file is not None:
        active_db_file = db_file
        if st.button("💾 บันทึกไฟล์นี้เป็นค่าเริ่มต้น", use_container_width=True, type="primary"):
            with open(SAVED_DB_PATH, "wb") as f:
                f.write(bytes(db_file.getbuffer()))
            st.success(f"✅ บันทึกข้อมูลเรียบร้อย!")
            
    elif os.path.exists(SAVED_DB_PATH):
        active_db_file = SAVED_DB_PATH
        st.caption(f"📌 แสดงผลจากฐานข้อมูลล่าสุด")
        if st.button("🗑️ ล้างข้อมูลไฟล์ระบบ", use_container_width=True):
            os.remove(SAVED_DB_PATH)
            st.rerun()

    st.download_button(
        label="📥 โหลดไฟล์ Template",
        data=generate_example_db_template(),
        file_name="Template_Database_Upload.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if df_check_bg is None or df_ord_bg is None:
    st.error(f"❌ ไม่พบไฟล์ Template ระบบพื้นหลัง")
    st.stop()

# ---- การคำนวณ ----
if active_db_file:
    with st.spinner("🔄 กำลังประมวลผลข้อมูล (Processing Data)..."):
        df_check = df_check_bg.copy()
        xls_db = pd.ExcelFile(active_db_file)
        sheet_names_lower = [str(s).lower() for s in xls_db.sheet_names]
        
        # 1. อ่านชีทแรก
        df_db = pd.read_excel(xls_db, sheet_name=0) 
        if 'Material' in df_db.columns:
            df_db['Material'] = df_db['Material'].astype(str).str.strip().str.replace(r';A[12]$', '', regex=True)
        stock_agg = df_db.groupby('Material')['Unrestricted'].sum().to_dict()
        
        # 2. อ่านชีท 'ord'
        if 'ord' in sheet_names_lower:
            df_ord = pd.read_excel(xls_db, sheet_name=xls_db.sheet_names[sheet_names_lower.index('ord')], header=0)
        else:
            df_ord = df_ord_bg.copy()
            
        # 3. อ่านชีท 'fo'
        df_fo = pd.DataFrame()
        max_fo_map, max_ord_map = {}, {}
        if 'fo' in sheet_names_lower:
            df_fo = pd.read_excel(xls_db, sheet_name=xls_db.sheet_names[sheet_names_lower.index('fo')], header=0)
            if not df_fo.empty and 'Material' in df_fo.columns:
                fo_cols = [c for c in df_fo.columns if 'FO' in str(c).upper() and '(PCS)' in str(c).upper()]
                ord_cols = [c for c in df_fo.columns if 'ORD' in str(c).upper() and '(PCS)' in str(c).upper()]
                df_fo['Max_FO'] = df_fo[fo_cols].max(axis=1) if fo_cols else 0
                df_fo['Max_ORD'] = df_fo[ord_cols].max(axis=1) if ord_cols else 0
                df_fo['Material'] = df_fo['Material'].astype(str).str.strip()
                max_fo_map = dict(zip(df_fo['Material'], df_fo['Max_FO']))
                max_ord_map = dict(zip(df_fo['Material'], df_fo['Max_ORD']))
        
        # 4. อ่านชีท 'pro'
        machine_mapping = {}
        if 'pro' in sheet_names_lower:
            df_pro = pd.read_excel(xls_db, sheet_name=xls_db.sheet_names[sheet_names_lower.index('pro')], header=None)
            if len(df_pro.columns) >= 3:
                df_pro[0] = df_pro[0].astype(str).str.strip()
                def extract_machine(val):
                    if pd.isna(val): return ""
                    val_str = str(val).strip()
                    parts = val_str.split('/')
                    return parts[1].strip() if len(parts) > 1 else ""
                
                df_pro['extracted_machine'] = df_pro[2].apply(extract_machine)
                machine_mapping = df_pro[df_pro['extracted_machine'] != ""].drop_duplicates(subset=[0, 'extracted_machine']).groupby(0)['extracted_machine'].apply(lambda x: ', '.join(filter(None, x))).to_dict()
        
        # --- คำนวณสต็อก ---
        df_check['fg'] = df_check['Material'].map(stock_agg).fillna(0)
        df_check['wip'] = df_check['Component'].map(stock_agg).fillna(0)
        df_check['Total'] = df_check['fg'] + df_check['wip']
        
        ord_sum = {}
        date_insufficient = {}
        
        for mat in df_check['Material']:
            mat_orders = df_ord[(df_ord['SAP Mat.'] == mat) & (df_ord['Dlv. Date'] <= target_date_dt)]
            ord_sum[mat] = pd.to_numeric(mat_orders['Outstd.Base Qty'], errors='coerce').sum()
            
            mat_orders_all = df_ord[df_ord['SAP Mat.'] == mat].copy()
            mat_orders_all['Outstd.Base Qty'] = pd.to_numeric(mat_orders_all['Outstd.Base Qty'], errors='coerce').fillna(0)
            mat_orders_all = mat_orders_all[mat_orders_all['Outstd.Base Qty'] > 0].sort_values('Dlv. Date')
            
            if not mat_orders_all.empty:
                mat_orders_all['Running_Sum'] = mat_orders_all['Outstd.Base Qty'].cumsum()
                total_stock = df_check.loc[df_check['Material'] == mat, 'Total'].values[0]
                mat_orders_all['Balance'] = total_stock - mat_orders_all['Running_Sum']
                shortage = mat_orders_all[mat_orders_all['Balance'] < 0]
                date_insufficient[mat] = shortage.iloc[0]['Dlv. Date'] if not shortage.empty else pd.NaT
            else:
                date_insufficient[mat] = pd.NaT
                
        df_check['Orders'] = df_check['Material'].map(ord_sum).fillna(0)
        df_check['Balance'] = df_check['Total'] - df_check['Orders']
        df_check['B/O Date'] = df_check['Material'].map(date_insufficient)
        df_check['SCHE'] = df_check['Matl group'].fillna('Unknown')
        
        df_check['Max FO'] = df_check['Material'].astype(str).str.strip().map(max_fo_map).fillna(0)
        df_check['Max ORD'] = df_check['Material'].astype(str).str.strip().map(max_ord_map).fillna(0)
        
        # --- จับคู่สถานะการผลิต ---
        status_list, machine_list = [], []
        for comp, mat in zip(df_check['Component'], df_check['Material']):
            comp_str = str(comp).strip()[:-2] if str(comp).strip().endswith('.0') else str(comp).strip()
            mat_str = str(mat).strip()[:-2] if str(mat).strip().endswith('.0') else str(mat).strip()
            
            if comp_str in machine_mapping:
                status_list.append('ผลิต'), machine_list.append(machine_mapping[comp_str])
            elif mat_str in machine_mapping:
                status_list.append('ผลิต'), machine_list.append(machine_mapping[mat_str])
            else:
                status_list.append('ไม่ได้ผลิต'), machine_list.append('-')
                
        df_check['status การผลิต'], df_check['เครื่องจักร'] = status_list, machine_list

        # --- คำนวณ WIP Days ---
        df_check['order avg/day'] = pd.to_numeric(df_check['total fo+30%'], errors='coerce').fillna(0) / working_days_input
        df_check['wip days value'] = np.where(df_check['order avg/day'] > 0, df_check['Total'] / df_check['order avg/day'], 999)
        
        # 🟢 ฟังก์ชันคำนวณ Status สัญลักษณ์สี
        def determine_status_emoji(wip, bo_dt):
            diff = (bo_dt - target_date_dt).days if pd.notna(bo_dt) else 9999
            if wip < 4 or diff < 4: return '1. 🔴'
            elif (4 <= wip <= 7) or (4 <= diff <= 7): return '2. 🟡'
            else: return '3. 🟠'

        df_check['Status'] = df_check.apply(lambda r: determine_status_emoji(r['wip days value'], r['B/O Date']), axis=1)
        
        df_check = df_check.sort_values(by='B/O Date', na_position='last')
        df_check['B/O Date Format'] = df_check['B/O Date'].dt.strftime('%Y-%m-%d').fillna('-')
        
        # --- เตรียม Dataframe ทั้งหมด (แยก WIP และ FG) ---
        balance_col_name = f'Balance ณ {target_date_dt.strftime("%d %b")}'
        display_df_all = pd.DataFrame({
            'Status': df_check['Status'], 
            'SCHE': df_check['SCHE'],
            'Part No.': df_check['Component'],
            'Material': df_check['Material'],
            'WIP': df_check['wip'].astype(int),
            'FG': df_check['fg'].astype(int),
            'รวม (WIP+FG)': df_check['Total'].astype(int),
            'WIP Day': df_check['wip days value'],
            'Max FO': df_check['Max FO'].astype(int),
            'Max ORD': df_check['Max ORD'].astype(int),
            'Orders': df_check['Orders'].astype(int),
            balance_col_name: df_check['Balance'].astype(int),
            'B/O Date': df_check['B/O Date Format'],
            'Note': df_check['note'].fillna('-') if 'note' in df_check.columns else '-',
            'status การผลิต': df_check['status การผลิต'],
            'เครื่องจักร': df_check['เครื่องจักร']
        })

        condition_short = (display_df_all['WIP Day'] < 7) | (display_df_all[balance_col_name] < 0)
        df_short = display_df_all[condition_short].copy()
        
        total_short_parts = len(df_short)
        total_orders_short = int(df_short['Orders'].sum())

        # --- ตรวจสอบ Part หลุดแผน ---
        if not df_fo.empty and 'Material' in df_fo.columns:
            fo_mats, chk_mats = set(df_fo['Material'].dropna().astype(str).str.strip()), set(df_check['Material'].dropna().astype(str).str.strip())
            missing_mats = [m for m in fo_mats - chk_mats if m not in ['nan', '']]
            if missing_mats:
                st.markdown(f'<div class="alert-box">⚠️ <b>แจ้งเตือนความเสี่ยงหลุดแผน:</b> พบ {len(missing_mats)} Part ใน <b>"fo"</b> ที่ไม่มีใน <b>"check dali wipday"</b></div>', unsafe_allow_html=True)
                with st.expander("👉 ดูรายการที่ตกหล่น"):
                    st.dataframe(df_fo[df_fo['Material'].isin(missing_mats)][['Material'] + (['Description'] if 'Description' in df_fo.columns else [])].drop_duplicates(), use_container_width=True, hide_index=True)

        # ---- สร้างการ์ดแสดงผล (Executive Metrics) ----
        st.markdown("<div style='margin-bottom: -10px;'></div>", unsafe_allow_html=True)
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric("Critical Parts (B/O หรือ WIP<7)", f"{total_short_parts:,} รายการ")
        with col_m2:
            st.metric("Pending Orders (ยอดค้างส่ง)", f"{total_orders_short:,} ชิ้น")
        with col_m3:
            st.metric("System Status (ระบบ)", "✅ Online (Data Synced)" if db_file is not None else "⏳ Cached Data")
        st.markdown("<br>", unsafe_allow_html=True)
        
        # ==========================================
        # 🎨 ระบบไฮไลท์สีคอลัมน์ตารางแบบคลีนๆ (Pastel Palette)
        # ==========================================
        def color_balance(val):
            return 'color: #991b1b; font-weight: bold;' if isinstance(val, (int, float)) and val < 0 else 'color: #065f46; font-weight: 500;'
        
        def color_wip(val):
            if pd.isna(val) or val == 999: return ''
            try:
                v = float(val)
                if v < 4: return 'background-color: #fef2f2; color: #991b1b; font-weight: bold;' # Soft Red
                elif 4 <= v <= 7: return 'background-color: #fefce8; color: #b45309; font-weight: bold;' # Soft Yellow/Amber
                else: return 'background-color: #fff7ed; color: #b45309; font-weight: 500;' # Soft Orange
            except: return ''

        def color_bo_date(val):
            if str(val).strip() in ['-', 'OK', 'nan', 'NaT', '']: return ''
            try:
                diff = (pd.to_datetime(val) - target_date_dt).days
                if diff < 4: return 'background-color: #fef2f2; color: #991b1b; font-weight: bold;' # Soft Red
                elif 4 <= diff <= 7: return 'background-color: #fefce8; color: #b45309; font-weight: bold;' # Soft Yellow/Amber
                else: return 'background-color: #fff7ed; color: #b45309; font-weight: 500;' # Soft Orange
            except: return ''
                
        format_dict = {'WIP Day': '{:.2f}'}

        # --- ส่วนค้นหาหลาย Part พร้อมกัน ---
        st.markdown("<h3 style='font-size:18px;'>🔍 ค้นหาข้อมูลเชิงลึก (Multi-Part Search)</h3>", unsafe_allow_html=True)
        search_query = st.text_input("พิมพ์รหัส Part No. หรือ Material (คั่นด้วยลูกน้ำ ',' เพื่อดูเทียบกันหลายเบอร์)", "")
        
        if search_query:
            queries = [q.strip() for q in search_query.split(',') if q.strip()]
            if queries:
                search_mask = pd.Series(False, index=display_df_all.index)
                for q in queries:
                    search_mask |= display_df_all['Part No.'].astype(str).str.contains(q, case=False, na=False) | display_df_all['Material'].astype(str).str.contains(q, case=False, na=False)
                
                searched_df = display_df_all[search_mask]
                if not searched_df.empty:
                    styled_search = searched_df.style.map(color_balance, subset=[balance_col_name]).map(color_wip, subset=['WIP Day']).map(color_bo_date, subset=['B/O Date']).format(format_dict)
                    st.dataframe(styled_search, use_container_width=True, hide_index=True)
                else:
                    st.warning(f"❌ ไม่พบข้อมูล Part ที่ตรงกับ '{search_query}' ในฐานข้อมูล")

        st.markdown("<hr style='margin-top:20px; margin-bottom:20px;'>", unsafe_allow_html=True)

        # --- ส่วนแสดงกราฟและตาราง Part ที่ต้องระวัง ---
        col_chart, col_table = st.columns([1.1, 2.5]) 
        
        with col_chart:
            st.markdown("<h3 style='font-size:17px;'>📊 สัดส่วนตามแผนก (SCHE)</h3>", unsafe_allow_html=True)
            if not df_short.empty:
                sche_counts = df_short.groupby('SCHE').size().reset_index(name='count')
                # โทนสีสวยหรู (Corporate Palette)
                corp_colors = ['#ef4444', '#f59e0b', '#10b981', '#3b82f6', '#8b5cf6', '#64748b', '#ec4899', '#0ea5e9']
                fig_sche = px.pie(sche_counts, values='count', names='SCHE', hole=0.6, color_discrete_sequence=corp_colors)
                fig_sche.update_layout(showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5), margin=dict(t=0, b=0, l=0, r=0), height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                fig_sche.update_traces(textposition='none', hoverinfo='label+percent')
                st.plotly_chart(fig_sche, use_container_width=True)
                
                st.markdown("<br><h3 style='font-size:17px;'>⚙️ สถานะการผลิต</h3>", unsafe_allow_html=True)
                chart_status = df_short.groupby('status การผลิต').size().reset_index(name='count')
                fig_status = px.pie(chart_status, values='count', names='status การผลิต', hole=0.6, color='status การผลิต', color_discrete_map={'ผลิต': '#10b981', 'ไม่ได้ผลิต': '#ef4444'})
                fig_status.update_layout(showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5), margin=dict(t=0, b=0, l=0, r=0), height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                fig_status.update_traces(textposition='none', hoverinfo='label+percent')
                st.plotly_chart(fig_status, use_container_width=True)
            else:
                st.success("🎉 ไม่มีรายการที่ B/O หรือ WIP ต่ำกว่ากำหนด")

        with col_table:
            c_header, c_btn = st.columns([2.5, 1])
            with c_header:
                st.markdown("<h3 style='font-size:18px;'>⚠️ รายการ Part ที่น่าเป็นห่วง (B/O Date หรือ WIP < 7 วัน)</h3>", unsafe_allow_html=True)
            
            with c_btn:
                st.download_button(label="📥 โหลดไฟล์ Excel", data=convert_df_to_excel(df_short), file_name=f"Shortage_{target_date_dt.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
            table_height = max(int((len(df_short) + 1) * 36) + 40, 300)
            
            # จัดการสีแบบ Clean & Corporate
            styled_df = df_short.style.map(color_balance, subset=[balance_col_name])\
                                      .map(color_wip, subset=['WIP Day'])\
                                      .map(color_bo_date, subset=['B/O Date'])\
                                      .map(lambda x: 'color: #3b82f6; font-weight: 600;', subset=['Part No.', 'Material'])\
                                      .map(lambda x: 'color: #ef4444; font-weight: 600;' if x > 0 else '', subset=['Orders'])\
                                      .format(format_dict)
                
            st.dataframe(
                styled_df, 
                use_container_width=True, 
                height=table_height, 
                hide_index=True,
                column_config={
                    "Status": st.column_config.TextColumn("Status", width="small"),
                    "WIP": st.column_config.NumberColumn("WIP", format="%d"),
                    "FG": st.column_config.NumberColumn("FG", format="%d"),
                    "รวม (WIP+FG)": st.column_config.NumberColumn("รวม", format="%d")
                }
            )

else:
    st.info("👋 ยินดีต้อนรับสู่ระบบ Executive Dashboard! กรุณาอัปโหลดไฟล์ Database รายวันเพื่อเริ่มต้นวิเคราะห์ข้อมูลแบบ Real-time ครับ")