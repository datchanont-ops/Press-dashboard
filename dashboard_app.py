import streamlit as st
import pandas as pd
import numpy as np
import os
import plotly.express as px
from io import BytesIO
import base64
import requests

# ---- Page Config ----
st.set_page_config(page_title="Production Shortage Dashboard", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# 🔗 GitHub Persistence Layer
# ==========================================
try:
    GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]
    GITHUB_REPO = st.secrets["GITHUB_REPO"]
    GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")
    GITHUB_DATA_DIR = st.secrets.get("GITHUB_DATA_DIR", "data")
    GITHUB_ENABLED = True
except Exception:
    GITHUB_ENABLED = False

GITHUB_API = "https://api.github.com"

def gh_headers():
    return {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }

def gh_get_file(remote_path):
    if not GITHUB_ENABLED:
        return None, None
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}?ref={GITHUB_BRANCH}"
    try:
        r = requests.get(url, headers=gh_headers(), timeout=15)
        if r.status_code == 200:
            data = r.json()
            content = base64.b64decode(data["content"])
            return content, data["sha"]
    except Exception:
        pass
    return None, None

def gh_put_file(remote_path, content_bytes, message):
    if not GITHUB_ENABLED:
        return False
    _, sha = gh_get_file(remote_path)
    url = f"{GITHUB_API}/repos/{GITHUB_REPO}/contents/{remote_path}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    try:
        r = requests.put(url, headers=gh_headers(), json=payload, timeout=15)
        return r.status_code in (200, 201)
    except Exception:
        return False

# ---- Custom CSS ----
st.markdown("""
    <style>
        .main-header {
            background-color: white;
            padding: 1.5rem;
            border-radius: 0.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            margin-bottom: 1rem;
        }
        .metric-card {
            background-color: white;
            padding: 1.5rem;
            border-radius: 0.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            border-top: 4px solid;
            text-align: left;
        }
        .metric-card.red { border-top-color: #ef4444; }
        .metric-card.orange { border-top-color: #f97316; }
        .metric-card.green { border-top-color: #10b981; }
        .metric-value {
            font-size: 2.5rem;
            font-weight: bold;
            color: #1f2937;
        }
        .metric-label {
            color: #6b7280;
            font-size: 1rem;
            font-weight: 600;
            margin-bottom: 0.5rem;
        }
        .alert-box {
            background-color: #fee2e2;
            color: #b91c1c;
            padding: 1rem;
            border-radius: 0.5rem;
            margin-bottom: 1rem;
            font-weight: 500;
            border-left: 5px solid #b91c1c;
        }
    </style>
""", unsafe_allow_html=True)

# ---- Data Processing Logic ----
try:
    current_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    current_dir = os.getcwd()

TEMPLATE_FILENAME = "@2.daily check aug26 ทุกวัน.XLSX"
TEMPLATE_PATH = os.path.join(current_dir, TEMPLATE_FILENAME)
SAVED_DB_PATH = os.path.join(current_dir, 'saved_database_upload.xlsx')

# 🔄 ระบบซิงค์ข้อมูลลง Local
if GITHUB_ENABLED and not st.session_state.get("github_synced"):
    with st.spinner("🔄 กำลังซิงค์ข้อมูลล่าสุดจาก GitHub..."):
        content, _ = gh_get_file(f"{GITHUB_DATA_DIR}/saved_database_upload.xlsx")
        if content:
            with open(SAVED_DB_PATH, "wb") as f:
                f.write(content)
    st.session_state["github_synced"] = True
elif not GITHUB_ENABLED:
    st.warning("⚠️ ยังไม่ได้ตั้งค่า GitHub Secrets ข้อมูลจะไม่ถูกบันทึกข้ามการเปิดเว็บใหม่ (รันแบบ Local เท่านั้น)", icon="⚠️")

# Load background template data
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
    processed_data = output.getvalue()
    return processed_data

@st.cache_data
def generate_example_db_template():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_wip = pd.DataFrame({
            'Material': ['xxx1', 'xxx2;A1', 'xxx2;A2'], 
            'Plant': [1200, 1200, 1200],
            'Storage Location': [1201, 1201, 1201],
            'DF stor. loc. level': ['', '', ''],
            'Base Unit of Measure': ['PC', 'PC', 'PC'],
            'Unrestricted': [0, 1000, 388]
        })
        df_wip.to_excel(writer, index=False, sheet_name='dali wip-fg')
        
        df_ord = pd.DataFrame({
            'Customer': ['0000100013', '0000100013'],
            'Name': ['ASIAN HONDA MOTOR CO.,LTD', 'ASIAN HONDA MOTOR CO.,LTD'],
            'SAP Mat.': ['xxx', 'xxx2'],
            'Cust.Mat.': ['abc', 'cae'],
            'Base Qty': [20, 12],
            'Outstd.Base Qty': [20, 5],
            'Dlv. Date': ['2026-08-07', '2026-08-14'],
            'On Hand': [17, 0],
            'Wait Ins.(Unr)': [0, 0],
            'Group': ['F-PUNCH', 'F-RP'],
            'Create Date': ['2026-07-16', '2026-07-16'],
            'Order': ['0002240801', '0002240813']
        })
        df_ord.to_excel(writer, index=False, sheet_name='ord')
        
        df_fo = pd.DataFrame({
            'Cust.Code': ['123', '1234'],
            'Name': ['x', 'x'],
            'Material': ['xxx1', 'xxx2'],
            'Description': ['aaa', 'bbb'],
            'Mat.Group': ['F-PUNCH', 'F-RP'],
            'Mat.Group4': ['PRESS', 'PRESS'],
            'Cust.Group Name': ['Spare parts', 'Spare parts'],
            'FO(Pcs)-08.2026': [0, 19],
            'ORD(Pcs)-08.2026': [20, 29],
            'FO(Pcs)-09.2026': [10, 25],
            'ORD(Pcs)-09.2026': [15, 35]
        })
        df_fo.to_excel(writer, index=False, sheet_name='fo')
        
        df_pro = pd.DataFrame({
            'Material': ['xxxx', 'xxxx1'],
            'Material Description': ['abc', 'aaa'],
            'Document Header Text': ['1/INJ_17/A/3/20260801', '1/INJ_88/A/2/20260805'],
            'Batch': ['', ''],
            'Storage Location': ['PP01', 'PP01'],
            'Movement Type': [131, 131],
            'Qty in Un. of Entry': [126, 72],
            'Unit of Entry': ['PC', 'PC'],
            'Amount in LC': ['', ''],
            'Material Document': ['4957809142', '4957800488'],
            'Posting Date': ['2026-08-11', '2026-08-11']
        })
        df_pro.to_excel(writer, index=False, sheet_name='pro')
        
    return output.getvalue()

# ---- Header Section ----
col_title, col_date, col_workday, col_upload = st.columns([2, 1, 1, 1.2])

with col_title:
    st.markdown("""
        <div class="main-header">
            <h1 style="margin:0; font-size:1.8rem; color:#111827;">📈 Production Shortage Dashboard</h1>
            <p style="margin:0; color:#6b7280; font-size:0.9rem; margin-top:0.5rem;">
                แสดงผลข้อมูลและสถานะการ Short / WIP น้อยกว่ากำหนด เฉพาะช่วงเวลาที่เลือก
            </p>
        </div>
    """, unsafe_allow_html=True)

with col_date:
    st.markdown("🗓️ **ดู Balance ถึงวันที่**")
    target_date = st.date_input("", pd.to_datetime('2026-08-31'), label_visibility="collapsed")
    target_date_dt = pd.to_datetime(target_date) # ดึงค่าวันที่ที่เลือกมาใช้คำนวณ

with col_workday:
    st.markdown("⏱️ **Working Day (วัน)**")
    working_days_input = st.number_input("", min_value=1.0, value=20.0, step=1.0, label_visibility="collapsed")

with col_upload:
    st.markdown("📁 **อัปโหลดไฟล์ Database**")
    db_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
    
    # --- ระบบอัปโหลดและเซฟไฟล์ ---
    active_db_file = None
    if db_file is not None:
        active_db_file = db_file
        if st.button("💾 บันทึกไฟล์นี้ไว้ใช้รอบหน้า", use_container_width=True):
            file_bytes = bytes(db_file.getbuffer())
            with open(SAVED_DB_PATH, "wb") as f:
                f.write(file_bytes)
            
            if GITHUB_ENABLED:
                with st.spinner("☁️ กำลังบันทึกไฟล์ขึ้น GitHub..."):
                    ok = gh_put_file(f"{GITHUB_DATA_DIR}/saved_database_upload.xlsx", file_bytes, f"Auto-save db upload: {db_file.name}")
                if ok:
                    st.success("✅ บันทึกไฟล์ขึ้น GitHub เรียบร้อย!")
                else:
                    st.warning("⚠️ บันทึกลง local ได้ แต่ push ไป GitHub ไม่สำเร็จ")
            else:
                st.success("✅ บันทึกไฟล์ลงในระบบชั่วคราวเรียบร้อย!")
    elif os.path.exists(SAVED_DB_PATH):
        active_db_file = SAVED_DB_PATH
        st.caption("📌 กำลังแสดงผลจาก: **ไฟล์ที่บันทึกไว้ล่าสุด**")
        if st.button("🗑️ ล้างข้อมูลไฟล์ที่บันทึกไว้", use_container_width=True):
            os.remove(SAVED_DB_PATH)
            st.rerun()

    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Template อัปโหลด",
        data=generate_example_db_template(),
        file_name="Template_Database_Upload.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

if df_check_bg is None or df_ord_bg is None:
    st.error(f"❌ ไม่พบไฟล์ Template ระบบพื้นหลัง")
    st.stop()

# ---- การคำนวณ (ถ้ามีไฟล์ครบ) ----
if active_db_file:
    with st.spinner("กำลังคำนวณข้อมูลเบื้องหลัง..."):
        df_check = df_check_bg.copy()
        
        xls_db = pd.ExcelFile(active_db_file)
        sheet_names_lower = [str(s).lower() for s in xls_db.sheet_names]
        
        # 1. อ่านชีทแรก
        df_db = pd.read_excel(xls_db, sheet_name=0) 
        if 'Material' in df_db.columns:
            df_db['Material'] = df_db['Material'].astype(str).str.strip()
            df_db['Material'] = df_db['Material'].str.replace(r';A[12]$', '', regex=True)
        stock_agg = df_db.groupby('Material')['Unrestricted'].sum().to_dict()
        
        # 2. อ่านชีท 'ord'
        if 'ord' in sheet_names_lower:
            ord_sheet_name = xls_db.sheet_names[sheet_names_lower.index('ord')]
            df_ord = pd.read_excel(xls_db, sheet_name=ord_sheet_name, header=0)
        else:
            df_ord = df_ord_bg.copy()
            
        # 3. อ่านชีท 'fo'
        df_fo = pd.DataFrame()
        max_fo_map = {}
        max_ord_map = {}
        if 'fo' in sheet_names_lower:
            fo_sheet_name = xls_db.sheet_names[sheet_names_lower.index('fo')]
            df_fo = pd.read_excel(xls_db, sheet_name=fo_sheet_name, header=0)
            
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
            pro_sheet_name = xls_db.sheet_names[sheet_names_lower.index('pro')]
            df_pro = pd.read_excel(xls_db, sheet_name=pro_sheet_name, header=None)
            if len(df_pro.columns) >= 3:
                df_pro[0] = df_pro[0].astype(str).str.strip()
                def extract_machine(val):
                    if pd.isna(val): return ""
                    val_str = str(val).strip()
                    if val_str == '' or val_str.lower() == 'nan': return ""
                    parts = val_str.split('/')
                    if len(parts) > 1: return parts[1].strip()
                    return ""
                
                df_pro['extracted_machine'] = df_pro[2].apply(extract_machine)
                df_pro_valid = df_pro[df_pro['extracted_machine'] != ""]
                df_pro_unique = df_pro_valid.drop_duplicates(subset=[0, 'extracted_machine'])
                machine_mapping = df_pro_unique.groupby(0)['extracted_machine'].apply(lambda x: ', '.join(filter(None, x))).to_dict()
        
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
            mat_orders_all = mat_orders_all[mat_orders_all['Outstd.Base Qty'] > 0]
            
            if not mat_orders_all.empty:
                mat_orders_all = mat_orders_all.sort_values('Dlv. Date')
                mat_orders_all['Running_Sum'] = mat_orders_all['Outstd.Base Qty'].cumsum()
                total_stock = df_check.loc[df_check['Material'] == mat, 'Total'].values[0]
                mat_orders_all['Balance'] = total_stock - mat_orders_all['Running_Sum']
                
                shortage = mat_orders_all[mat_orders_all['Balance'] < 0]
                if not shortage.empty:
                    date_insufficient[mat] = shortage.iloc[0]['Dlv. Date']
                else:
                    date_insufficient[mat] = pd.NaT
            else:
                date_insufficient[mat] = pd.NaT
                
        df_check['Orders'] = df_check['Material'].map(ord_sum).fillna(0)
        df_check['Balance'] = df_check['Total'] - df_check['Orders']
        df_check['Short Date'] = df_check['Material'].map(date_insufficient)
        df_check['SCHE'] = df_check['Matl group'].fillna('Unknown')
        
        df_check['Max FO'] = df_check['Material'].astype(str).str.strip().map(max_fo_map).fillna(0)
        df_check['Max ORD'] = df_check['Material'].astype(str).str.strip().map(max_ord_map).fillna(0)
        
        # --- จับคู่สถานะการผลิต ---
        status_list = []
        machine_list = []
        for comp, mat in zip(df_check['Component'], df_check['Material']):
            comp_str = str(comp).strip() if pd.notna(comp) else ""
            mat_str = str(mat).strip() if pd.notna(mat) else ""
            if comp_str.endswith('.0'): comp_str = comp_str[:-2]
            if mat_str.endswith('.0'): mat_str = mat_str[:-2]
            
            if comp_str != "" and comp_str in machine_mapping:
                status_list.append('ผลิต')
                machine_list.append(machine_mapping[comp_str])
            elif mat_str != "" and mat_str in machine_mapping:
                status_list.append('ผลิต')
                machine_list.append(machine_mapping[mat_str])
            else:
                status_list.append('ไม่ได้ผลิต')
                machine_list.append('-')
                
        df_check['status การผลิต'] = status_list
        df_check['เครื่องจักร'] = machine_list

        # --- คำนวณ WIP Days ---
        df_check['order avg/day'] = pd.to_numeric(df_check['total fo+30%'], errors='coerce').fillna(0) / working_days_input
        df_check['wip days value'] = np.where(df_check['order avg/day'] > 0, df_check['Total'] / df_check['order avg/day'], 999)
        
        df_check = df_check.sort_values(by='Short Date', na_position='last')
        df_check['Short Date Format'] = df_check['Short Date'].dt.strftime('%Y-%m-%d').fillna('-')
        
        # --- เตรียม Dataframe ทั้งหมด ---
        balance_col_name = f'Balance ณ {target_date_dt.strftime("%d %b")}'
        display_df_all = pd.DataFrame({
            'SCHE': df_check['SCHE'],
            'Part No.': df_check['Component'],
            'Material': df_check['Material'],
            'WIP+FG': df_check['Total'].astype(int),
            'WIP Day': df_check['wip days value'],
            'Max FO': df_check['Max FO'].astype(int),
            'Max ORD': df_check['Max ORD'].astype(int),
            'Orders': df_check['Orders'].astype(int),
            balance_col_name: df_check['Balance'].astype(int),
            'Short Date': df_check['Short Date Format'],
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
            fo_materials = df_fo['Material'].dropna().astype(str).str.strip().unique()
            check_materials = df_check['Material'].dropna().astype(str).str.strip().unique()
            missing_materials = [m for m in fo_materials if m not in check_materials and m != 'nan' and m != '']
            if len(missing_materials) > 0:
                st.markdown(f'''
                    <div class="alert-box">⚠️ <b>แจ้งเตือนความเสี่ยงหลุดแผน:</b> พบ {len(missing_materials)} Part ที่มีใน Sheet <b>"fo"</b> แต่ไม่ได้นำมาคำนวณใน <b>"check dali wipday"</b></div>
                ''', unsafe_allow_html=True)
                with st.expander("👉 คลิกเพื่อดูรายการ Part ที่ตกหล่น"):
                    cols_to_show = ['Material']
                    if 'Description' in df_fo.columns: cols_to_show.append('Description')
                    missing_df = df_fo[df_fo['Material'].astype(str).str.strip().isin(missing_materials)][cols_to_show].drop_duplicates(subset=['Material'])
                    missing_df.columns = ['Part No. ที่ตกหล่น', 'รายละเอียด (Description)'][:len(cols_to_show)]
                    st.dataframe(missing_df, use_container_width=True, hide_index=True)

        # ---- สร้างการ์ดแสดงผล ----
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.markdown(f'<div class="metric-card red"><div class="metric-label">Part ที่ต้องระวัง (Short หรือ WIP<7)</div><div class="metric-value">{total_short_parts:,}</div></div>', unsafe_allow_html=True)
        with col_m2:
            st.markdown(f'<div class="metric-card orange"><div class="metric-label">จำนวน Order ค้างส่ง (ชิ้น)</div><div class="metric-value">{total_orders_short:,}</div></div>', unsafe_allow_html=True)
        with col_m3:
            st.markdown('<div class="metric-card green"><div class="metric-label">สถานะระบบ</div><div class="metric-value" style="font-size:1.5rem; margin-top:0.8rem;">✨ ข้อมูลอัปเดตล่าสุด<br>(Live File)</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        
        # ==========================================
        # 🎨 ระบบไฮไลท์สีคอลัมน์ (Styling Functions)
        # ==========================================
        def color_balance(val):
            color = 'red' if isinstance(val, (int, float)) and val < 0 else 'black'
            return f'color: {color}'
        
        # สีสำหรับ WIP Day
        def color_wip(val):
            if pd.isna(val) or val == 999: return ''
            try:
                v = float(val)
                if v < 4:
                    return 'background-color: #fee2e2; color: #b91c1c; font-weight: bold;' # สีแดง
                elif 4 <= v <= 7:
                    return 'background-color: #ffedd5; color: #c2410c; font-weight: bold;' # สีส้ม
                else:
                    return 'background-color: #ffedd5; color: #c2410c; font-weight: bold;' # สีส้ม (มากกว่า 7 วัน)
            except:
                return ''

        # สีสำหรับ Short Date (คำนวณระยะห่างจากวันที่ระบุใน Dashboard)
        def color_short_date(val):
            if str(val).strip() in ['-', 'OK', 'nan', 'NaT', '']: return ''
            try:
                s_dt = pd.to_datetime(val)
                diff = (s_dt - target_date_dt).days
                if diff < 4:
                    return 'background-color: #fee2e2; color: #b91c1c; font-weight: bold;' # สีแดง
                elif 4 <= diff <= 7:
                    return 'background-color: #fef08a; color: #854d0e; font-weight: bold;' # สีเหลือง
                else:
                    return 'background-color: #ffedd5; color: #c2410c; font-weight: bold;' # สีส้ม
            except:
                return ''
                
        format_dict = {'WIP Day': '{:.2f}'}

        # --- ส่วนค้นหาหลาย Part พร้อมกัน ---
        st.markdown("### 🔍 ค้นหาสถานะ Part ข้อมูลทั้งหมด (เทียบหลายรายการได้)")
        search_query = st.text_input("พิมพ์รหัส Part No. หรือ Material (คั่นด้วยลูกน้ำ ',' หากต้องการเทียบหลายตัว เช่น 1184469, BZ130)", "")
        
        if search_query:
            queries = [q.strip() for q in search_query.split(',') if q.strip()]
            if queries:
                search_mask = pd.Series(False, index=display_df_all.index)
                for q in queries:
                    mask = display_df_all['Part No.'].astype(str).str.contains(q, case=False, na=False) | \
                           display_df_all['Material'].astype(str).str.contains(q, case=False, na=False)
                    search_mask = search_mask | mask
                
                searched_df = display_df_all[search_mask]
                if not searched_df.empty:
                    try:
                        styled_search = searched_df.style.map(color_balance, subset=[balance_col_name])\
                                                    .map(color_wip, subset=['WIP Day'])\
                                                    .map(color_short_date, subset=['Short Date'])\
                                                    .format(format_dict)
                    except AttributeError:
                        styled_search = searched_df.style.applymap(color_balance, subset=[balance_col_name])\
                                                    .applymap(color_wip, subset=['WIP Day'])\
                                                    .applymap(color_short_date, subset=['Short Date'])\
                                                    .format(format_dict)
                    st.dataframe(styled_search, use_container_width=True, hide_index=True)
                else:
                    st.warning(f"❌ ไม่พบข้อมูล Part ที่ตรงกับ '{search_query}' ในฐานข้อมูล")

        st.markdown("<hr>", unsafe_allow_html=True)

        # --- ส่วนแสดงกราฟและตาราง Part ที่ต้องระวัง ---
        col_chart, col_table = st.columns([1, 2.5]) 
        
        with col_chart:
            st.markdown("**แยกตามแผนก (SCHE)**")
            if not df_short.empty:
                chart_data_sche = df_short.groupby('SCHE').size().reset_index(name='count')
                fig_sche = px.pie(chart_data_sche, values='count', names='SCHE', hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_sche.update_layout(showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5), margin=dict(t=0, b=0, l=0, r=0), height=300)
                fig_sche.update_traces(textposition='none')
                st.plotly_chart(fig_sche, use_container_width=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                st.markdown("**สถานะการผลิต**")
                chart_data_status = df_short.groupby('status การผลิต').size().reset_index(name='count')
                color_map = {'ผลิต': '#10b981', 'ไม่ได้ผลิต': '#ef4444'}
                fig_status = px.pie(chart_data_status, values='count', names='status การผลิต', hole=0.5, color='status การผลิต', color_discrete_map=color_map)
                fig_status.update_layout(showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5), margin=dict(t=0, b=0, l=0, r=0), height=300)
                fig_status.update_traces(textposition='none')
                st.plotly_chart(fig_status, use_container_width=True)
            else:
                st.info("ไม่มีรายการที่ Short หรือ WIP ต่ำกว่ากำหนด")

        with col_table:
            c_header, c_btn = st.columns([2, 1])
            with c_header:
                st.markdown("**รายการ Part ที่ติดลบ (Short Date) หรือ WIP < 7 วัน**")
            
            with c_btn:
                excel_data = convert_df_to_excel(df_short)
                st.download_button(label="📥 ดาวน์โหลดไฟล์ Excel (ที่ Short)", data=excel_data, file_name=f"Production_Shortage_{target_date_dt.strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
            try:
                styled_df = df_short.style.map(color_balance, subset=[balance_col_name])\
                                          .map(color_wip, subset=['WIP Day'])\
                                          .map(color_short_date, subset=['Short Date'])\
                                          .format(format_dict)
            except AttributeError:
                styled_df = df_short.style.applymap(color_balance, subset=[balance_col_name])\
                                          .applymap(color_wip, subset=['WIP Day'])\
                                          .applymap(color_short_date, subset=['Short Date'])\
                                          .format(format_dict)
            
            table_height = (len(df_short) + 1) * 35 + 10
            if table_height < 150: table_height = 150
                
            st.dataframe(styled_df, use_container_width=True, height=table_height, hide_index=True)

else:
    st.markdown('<div class="metric-card" style="text-align:center; color:#6b7280; margin-top:2rem;">กรุณาอัปโหลดไฟล์ Database รายวัน หรือใช้ข้อมูลที่บันทึกไว้ เพื่อเริ่มการคำนวณ</div>', unsafe_allow_html=True)