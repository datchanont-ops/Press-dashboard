import streamlit as st
import pandas as pd
import numpy as np
import os
import plotly.express as px
from io import BytesIO

# ---- Page Config ----
st.set_page_config(page_title="Production Shortage Dashboard", layout="wide", initial_sidebar_state="collapsed")

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

# Load background template data
@st.cache_data
def load_template_data(path):
    if not os.path.exists(path):
        return None, None
    df_check = pd.read_excel(path, sheet_name="check dali wipday", header=2)
    df_check = df_check.dropna(subset=['Material'])
    df_ord = pd.read_excel(path, sheet_name="ord", header=0)
    return df_check, df_ord

df_check_bg, df_ord_bg = load_template_data(TEMPLATE_PATH)

# ฟังก์ชันสำหรับแปลง DataFrame เป็น Excel เพื่อให้ดาวน์โหลด
@st.cache_data
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Shortage Report')
    processed_data = output.getvalue()
    return processed_data

# ---- Header Section ----
col_title, col_date, col_upload = st.columns([2, 1, 1])

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
    st.markdown("🗓️ **ดู Balance ถึงวันที่ (คลิก)**")
    target_date = st.date_input("", pd.to_datetime('2026-08-31'), label_visibility="collapsed")

with col_upload:
    st.markdown("📁 **อัปโหลดไฟล์ Database**")
    db_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")

# ---- ระบบตรวจจับและแจ้งเตือนข้อผิดพลาดเรื่องไฟล์ Template ----
if df_check_bg is None or df_ord_bg is None:
    st.error(f"❌ ไม่พบไฟล์ Template ระบบพื้นหลัง")
    st.markdown("### 🔍 ระบบช่วยวิเคราะห์ปัญหา (Debug Info):")
    st.write(f"1. **โปรแกรมกำลังพยายามหาไฟล์ที่ตำแหน่งนี้:** `{TEMPLATE_PATH}`")
    try:
        available_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.xlsx')]
        if available_files:
            st.write(f"2. **เจอไฟล์ Excel อื่นๆ ในโฟลเดอร์นี้ ได้แก่:**")
            for f in available_files:
                st.write(f"- `{f}`")
        else:
            st.write("2. **ไม่พบไฟล์ Excel ใดๆ ในโฟลเดอร์นี้เลย**")
    except Exception as e:
        st.write(f"ไม่สามารถตรวจสอบโฟลเดอร์ได้: {e}")
    st.stop()

# ---- การคำนวณ (ถ้ามีไฟล์ครบ) ----
if db_file:
    with st.spinner("กำลังคำนวณข้อมูลเบื้องหลัง..."):
        df_check = df_check_bg.copy()
        df_ord = df_ord_bg.copy()
        
        # โหลดไฟล์ Database
        xls_db = pd.ExcelFile(db_file)
        
        # 1. อ่านชีทแรก (สต็อกปกติ)
        df_db = pd.read_excel(xls_db, sheet_name=0) 
        stock_agg = df_db.groupby('Material')['Unrestricted'].sum().to_dict()
        
        # 2. อ่านชีท 'pro' เพื่อดึงสถานะการผลิต
        machine_mapping = {}
        sheet_names_lower = [str(s).lower() for s in xls_db.sheet_names]
        
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
                    if len(parts) > 1:
                        return parts[1].strip()
                    return ""
                
                df_pro['extracted_machine'] = df_pro[2].apply(extract_machine)
                df_pro_valid = df_pro[df_pro['extracted_machine'] != ""]
                df_pro_unique = df_pro_valid.drop_duplicates(subset=[0, 'extracted_machine'])
                machine_mapping = df_pro_unique.groupby(0)['extracted_machine'].apply(
                    lambda x: ', '.join(filter(None, x))
                ).to_dict()
        
        # --- คำนวณสต็อก ---
        df_check['fg'] = df_check['Material'].map(stock_agg).fillna(0)
        df_check['wip'] = df_check['Component'].map(stock_agg).fillna(0)
        df_check['Total'] = df_check['fg'] + df_check['wip']
        
        target_date_dt = pd.to_datetime(target_date)
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
        
        # --- คำนวณ WIP Days ---
        L3_val = 20
        df_check['order avg/day'] = pd.to_numeric(df_check['total fo+30%'], errors='coerce').fillna(0) / L3_val
        df_check['wip days value'] = np.where(df_check['order avg/day'] > 0, df_check['Total'] / df_check['order avg/day'], 999)
        
        # --- กรองข้อมูล: WIP Day < 7 หรือ Balance < 0 ---
        condition = (df_check['wip days value'] < 7) | (df_check['Balance'] < 0)
        df_short = df_check[condition].copy()
        
        df_short = df_short.sort_values(by='Short Date', na_position='last')
        df_short['Short Date Format'] = df_short['Short Date'].dt.strftime('%Y-%m-%d').fillna('-')
        
        total_short_parts = len(df_short)
        total_orders_short = int(df_short['Orders'].sum())

        # --- จับคู่สถานะการผลิต (ย้ายมาคำนวณก่อนเพื่อให้กราฟดึงไปใช้ได้) ---
        status_list = []
        machine_list = []
        for comp, mat in zip(df_short['Component'], df_short['Material']):
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
                
        df_short['status การผลิต'] = status_list
        df_short['เครื่องจักร'] = machine_list

        # ---- สร้างการ์ดแสดงผล ----
        col_m1, col_m2, col_m3 = st.columns(3)
        target_str = target_date_dt.strftime('%d/%m/%Y')
        
        with col_m1:
            st.markdown(f"""
                <div class="metric-card red">
                    <div class="metric-label">Part ที่ต้องระวัง (Short หรือ WIP<7)</div>
                    <div class="metric-value">{total_short_parts:,}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with col_m2:
            st.markdown(f"""
                <div class="metric-card orange">
                    <div class="metric-label">จำนวน Order ค้างส่ง (ชิ้น)</div>
                    <div class="metric-value">{total_orders_short:,}</div>
                </div>
            """, unsafe_allow_html=True)
            
        with col_m3:
            st.markdown("""
                <div class="metric-card green">
                    <div class="metric-label">สถานะระบบ</div>
                    <div class="metric-value" style="font-size:1.5rem; margin-top:0.8rem;">✨ ข้อมูลอัปเดตล่าสุด<br>(Live File)</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        col_chart, col_table = st.columns([1, 2.5]) 
        
        with col_chart:
            # 1. กราฟแยกตามแผนก
            st.markdown("**แยกตามแผนก (SCHE)**")
            if not df_short.empty:
                chart_data_sche = df_short.groupby('SCHE').size().reset_index(name='count')
                fig_sche = px.pie(chart_data_sche, values='count', names='SCHE', hole=0.5,
                             color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_sche.update_layout(showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                                  margin=dict(t=0, b=0, l=0, r=0), height=300)
                fig_sche.update_traces(textposition='none')
                st.plotly_chart(fig_sche, use_container_width=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # 2. กราฟเปรียบเทียบสถานะการผลิต
                st.markdown("**สถานะการผลิต**")
                chart_data_status = df_short.groupby('status การผลิต').size().reset_index(name='count')
                
                # กำหนดสี: เขียว=ผลิต, แดง=ไม่ได้ผลิต
                color_map = {'ผลิต': '#10b981', 'ไม่ได้ผลิต': '#ef4444'}
                
                fig_status = px.pie(chart_data_status, values='count', names='status การผลิต', hole=0.5,
                                    color='status การผลิต', color_discrete_map=color_map)
                fig_status.update_layout(showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
                                  margin=dict(t=0, b=0, l=0, r=0), height=300)
                fig_status.update_traces(textposition='none')
                st.plotly_chart(fig_status, use_container_width=True)
                
            else:
                st.info("ไม่มีรายการที่ Short หรือ WIP ต่ำกว่ากำหนด")

        with col_table:
            c_header, c_btn = st.columns([2, 1])
            with c_header:
                st.markdown("**รายการ Part ที่ติดลบ (Short Date) หรือ WIP < 7 วัน**")
            
            # เตรียมข้อมูลสำหรับแสดงผล (เพิ่มคอลัมน์ Material)
            display_df = pd.DataFrame({
                'SCHE': df_short['SCHE'],
                'Part No.': df_short['Component'],
                'Material': df_short['Material'],
                'WIP+FG': df_short['Total'].astype(int),
                'WIP Day': df_short['wip days value'],
                'Orders': df_short['Orders'].astype(int),
                f'Balance ณ {target_date_dt.strftime("%d %b")}': df_short['Balance'].astype(int),
                'Short Date': df_short['Short Date Format'],
                'Note': df_short['note'].fillna('-') if 'note' in df_short.columns else '-',
                'status การผลิต': df_short['status การผลิต'],
                'เครื่องจักร': df_short['เครื่องจักร']
            })
            
            with c_btn:
                # ปุ่มดาวน์โหลด Excel
                excel_data = convert_df_to_excel(display_df)
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=excel_data,
                    file_name=f"Production_Shortage_{target_date_dt.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # ฟังก์ชันปรับสีตัวเลข
            def color_balance(val):
                color = 'red' if isinstance(val, (int, float)) and val < 0 else 'black'
                return f'color: {color}'
            
            def color_wip(val):
                color = 'red' if isinstance(val, (int, float)) and val < 7 else 'black'
                return f'color: {color}'
            
            # จัด Format ทศนิยม 2 ตำแหน่งให้ WIP Day
            format_dict = {'WIP Day': '{:.2f}'}
            
            try:
                styled_df = display_df.style.map(color_balance, subset=[f'Balance ณ {target_date_dt.strftime("%d %b")}'])\
                                            .map(color_wip, subset=['WIP Day'])\
                                            .format(format_dict)
            except AttributeError:
                styled_df = display_df.style.applymap(color_balance, subset=[f'Balance ณ {target_date_dt.strftime("%d %b")}'])\
                                            .applymap(color_wip, subset=['WIP Day'])\
                                            .format(format_dict)
            
            # คำนวณความสูงให้แสดงผลได้เต็ม
            table_height = (len(display_df) + 1) * 35 + 10
            if table_height < 150: 
                table_height = 150
                
            st.dataframe(styled_df, use_container_width=True, height=table_height, hide_index=True)

else:
    st.markdown('<div class="metric-card" style="text-align:center; color:#6b7280; margin-top:2rem;">กรุณาอัปโหลดไฟล์ Database รายวันเพื่อเริ่มการคำนวณ</div>', unsafe_allow_html=True)