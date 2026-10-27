import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import seaborn as sns
import json
import os
import numpy as np
import io
import PIL.Image

# Desactivar límite de píxeles por si alguna librería intenta procesar el artefacto
PIL.Image.MAX_IMAGE_PIXELS = None

# ¡Esta línea debe ser la primera!
st.set_page_config(page_title="Salary Planning Dashboard", layout="wide")

st.title("Salary Planning Dashboard")
st.write("Upload your Excel files to view and filter metrics.")

# --- FILTER MEMORY ---
ARCHIVO_FILTROS = 'mis_filtros_guardados.json'

def cargar_memoria():
    if os.path.exists(ARCHIVO_FILTROS):
        try:
            with open(ARCHIVO_FILTROS, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

if 'memoria_filtros' not in st.session_state:
    st.session_state['memoria_filtros'] = cargar_memoria()

# --- FUNCIONES ROBUSTAS PARA LEER DATOS EXCEL ---
def get_num(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).replace('$', '').replace(',', '').replace('%', '').strip()
    if val_str.lower() in ['', '-', 'nan', 'none', 'null']: return 0.0
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def get_date(val):
    if pd.isna(val): return pd.NaT
    val_str = str(val).strip().lower()
    if val_str in ['', 'nan', 'none', 'nat', 'null']: return pd.NaT
    try:
        # Añadido dayfirst=True según la sugerencia del log del servidor
        return pd.to_datetime(val, dayfirst=True, errors='coerce')
    except Exception:
        return pd.NaT

# --- FILE UPLOADER ---
col_up1, col_up2 = st.columns(2)
with col_up1:
    file1 = st.file_uploader("1. Upload the Main file (e.g., colgate_...xlsx)", type=["xlsx", "xls"], key="up_main")
with col_up2:
    file2 = st.file_uploader("2. Upload the SuperFlex file (SuperFlex-DefaultView...)", type=["xlsx", "xls"], key="up_flex")

if file1 is not None and file2 is not None:
    try:
        # LECTURA ESTRICTA PARA IGNORAR IMÁGENES GIGANTES Y ARTEFACTOS
        df1 = pd.read_excel(file1, sheet_name="Salary", engine="openpyxl", engine_kwargs={"read_only": True, "data_only": True})
        df2 = pd.read_excel(file2, header=2, engine="openpyxl", engine_kwargs={"read_only": True, "data_only": True})
        
        df1['Global ID'] = pd.to_numeric(df1['Global ID'], errors='coerce')
        df2['Global ID'] = pd.to_numeric(df2['Global ID'], errors='coerce')
        
        if len(df2.columns) > 28:
            df2['Employee Subgroup'] = df2.iloc[:, 28].astype(str)
        else:
            df2['Employee Subgroup'] = "N/A"
            
        if len(df2.columns) > 47:
            df2['Gender'] = df2.iloc[:, 47].astype(str)
        else:
            df2['Gender'] = "Unknown"
        
        df = pd.merge(df1, df2[['Global ID', 'Chief Name', 'Employee Subgroup', 'Gender']], on='Global ID', how='left')
        
        df['Chief Name'] = df['Chief Name'].fillna("No Manager Assigned")
        df['Employee Subgroup'] = df['Employee Subgroup'].replace('nan', 'Unknown').fillna('Unknown')
        df['Gender'] = df['Gender'].replace('nan', 'Unknown').fillna('Unknown')
        
        if 'Potential' in df.columns:
            df['Potential'] = df['Potential'].fillna('Not Assigned').astype(str)
            df['Potential'] = df['Potential'].replace({'nan': 'Not Assigned', 'None Selected': 'Not Assigned'})
        else:
            df['Potential'] = 'Not Assigned'
        
        # --- SIDEBAR: LOAD FILTERS ---
        st.sidebar.header("💾 My Saved Filters")
        nombres_disponibles = ["-- None --"] + list(st.session_state['memoria_filtros'].keys())
        filtro_elegido = st.sidebar.selectbox("Load a configuration:", nombres_disponibles, key="load_cfg")
        
        def_gerentes, def_orgs, def_funcs, def_comps, def_subgroups, def_potentials = [], [], [], [], [], []
        
        gerente_options = sorted(df['Chief Name'].astype(str).unique().tolist())
        org_options = sorted(df['Reporting Organization'].dropna().unique().tolist())
        func_options = sorted(df['Function'].dropna().unique().tolist())
        comp_options = sorted(df['Compensation Area'].dropna().unique().tolist())
        subgroup_options = sorted(df['Employee Subgroup'].unique().tolist())
        potential_options = sorted(df['Potential'].unique().tolist())

        if filtro_elegido != "-- None --":
            config = st.session_state['memoria_filtros'][filtro_elegido]
            def_gerentes = config.get("gerentes", [])
            def_orgs = config.get("orgs", [])
            def_funcs = config.get("funcs", [])
            def_comps = config.get("comps", [])
            def_subgroups = config.get("subgroups", [])
            def_potentials = config.get("potentials", [])

        st.sidebar.markdown("---")
        
        # --- SIDEBAR: DATA SELECTION ---
        st.sidebar.header("🔍 Data Filters")
        st.sidebar.info("Leave blank to include all.")
        
        def_gerentes = [x for x in def_gerentes if x in gerente_options]
        def_orgs = [x for x in def_orgs if x in org_options]
        def_funcs = [x for x in def_funcs if x in func_options]
        def_comps = [x for x in def_comps if x in comp_options]
        def_subgroups = [x for x in def_subgroups if x in subgroup_options]
        def_potentials = [x for x in def_potentials if x in potential_options]

        selected_gerentes = st.sidebar.multiselect("Manager(s):", gerente_options, default=def_gerentes, key="f_mgr")
        selected_orgs = st.sidebar.multiselect("Reporting Organization:", org_options, default=def_orgs, key="f_org")
        selected_funcs = st.sidebar.multiselect("Function:", func_options, default=def_funcs, key="f_func")
        selected_comps = st.sidebar.multiselect("Compensation Area:", comp_options, default=def_comps, key="f_comp")
        selected_subgroups = st.sidebar.multiselect("Employee Subgroup:", subgroup_options, default=def_subgroups, key="f_sg")
        selected_potentials = st.sidebar.multiselect("Potential (Col Y):", potential_options, default=def_potentials, key="f_pot")
        
        st.sidebar.markdown("---")
        st.sidebar.subheader("✏️ Save current combination")
        nuevo_nombre = st.sidebar.text_input("Give this filter a title:", key="save_title")
        
        if st.sidebar.button("Save Filter", key="save_btn"):
            if nuevo_nombre:
                st.session_state['memoria_filtros'][nuevo_nombre] = {
                    "gerentes": selected_gerentes,
                    "orgs": selected_orgs,
                    "funcs": selected_funcs,
                    "comps": selected_comps,
                    "subgroups": selected_subgroups,
                    "potentials": selected_potentials
                }
                with open(ARCHIVO_FILTROS, 'w', encoding='utf-8') as f:
                    json.dump(st.session_state['memoria_filtros'], f)
                st.sidebar.success(f"Filter '{nuevo_nombre}' saved successfully!")
            else:
                st.sidebar.warning("Please enter a title before saving.")

        # --- APPLY FILTERS ---
        filtros_finales_gerentes = selected_gerentes if selected_gerentes else gerente_options
        filtros_finales_orgs = selected_orgs if selected_orgs else org_options
        filtros_finales_funcs = selected_funcs if selected_funcs else func_options
        filtros_finales_comps = selected_comps if selected_comps else comp_options
        filtros_finales_subgroups = selected_subgroups if selected_subgroups else subgroup_options
        filtros_finales_potentials = selected_potentials if selected_potentials else potential_options

        df_filtered = df[
            (df['Chief Name'].isin(filtros_finales_gerentes)) &
            (df['Reporting Organization'].isin(filtros_finales_orgs)) &
            (df['Function'].isin(filtros_finales_funcs)) &
            (df['Compensation Area'].isin(filtros_finales_comps)) &
            (df['Employee Subgroup'].isin(filtros_finales_subgroups)) &
            (df['Potential'].isin(filtros_finales_potentials))
        ].copy()
        
        st.success(f"Showing {len(df_filtered)} records matching your sidebar search.")
        
        FECHA_ACTUAL = pd.to_datetime('2026-07-23')

        # --- CÁLCULOS GLOBALES ---
        col_date_sg = None
        col_curr_qrtle = None
        col_new_qrtle = None
        
        for c in df_filtered.columns:
            c_lower = str(c).lower().strip()
            if 'date in sg' in c_lower or 'date in salary group' in c_lower: col_date_sg = c
            elif 'new qrtle' in c_lower or 'new quartile' in c_lower or ('quartile' in c_lower and 'new' in c_lower): col_new_qrtle = c
            elif ('qrtle' in c_lower or 'quartile' in c_lower) and 'new' not in c_lower: col_curr_qrtle = c
                    
        if not col_date_sg and len(df_filtered.columns) > 25: col_date_sg = df_filtered.columns[25]
        if not col_curr_qrtle and len(df_filtered.columns) > 22: col_curr_qrtle = df_filtered.columns[22]
        if not col_new_qrtle and len(df_filtered.columns) > 49: col_new_qrtle = df_filtered.columns[49]

        if col_date_sg and not df_filtered.empty:
            df_filtered['Date_in_SG_clean'] = df_filtered[col_date_sg].apply(get_date)
            df_filtered['Years_Raw'] = ((FECHA_ACTUAL - df_filtered['Date_in_SG_clean']).dt.days / 365.25)
            def format_years(x):
                if pd.isna(x) or x < 0: return "Unknown"
                val = int(x)
                if val == 0: return "< 1 Year"
                elif val == 1: return "1 Year"
                else: return f"{val} Years"
            df_filtered['Years in SG'] = df_filtered['Years_Raw'].apply(format_years)
        else:
            df_filtered['Years in SG'] = "Unknown"

        mapeo_q = {"Below Minimum": "Below Min", "Below Min": "Below Min", "1Q": "1Q", "2Q": "2Q", "3Q": "3Q", "4Q": "4Q", "AboveMax": "Above max", "Above max": "Above max"}

        if col_curr_qrtle and not df_filtered.empty:
            df_filtered['Current Quartile'] = df_filtered[col_curr_qrtle].astype(str).str.strip().replace(mapeo_q)
        else:
            df_filtered['Current Quartile'] = "Unknown"

        if col_new_qrtle and not df_filtered.empty:
            df_filtered['New Quartile'] = df_filtered[col_new_qrtle].astype(str).str.strip().replace(mapeo_q)
        else:
            df_filtered['New Quartile'] = "Unknown"

        if df_filtered.empty:
            st.warning("No data matches these filters.")
        else:
            # ==========================================
            # SECCIÓN DE PDF (SIDEBAR)
            # ==========================================
            st.sidebar.markdown("---")
            st.sidebar.subheader("📄 Export Report")
            pdf_mode = st.sidebar.checkbox("🖨️ Enable PDF Print Mode (All-in-one view)", key="pdf_mode_check")
            
            if pdf_mode:
                st.markdown("""
                <style>
                @media print {
                    [data-testid="stSidebar"] { display: none !important; }
                    header { display: none !important; }
                    footer { display: none !important; }
                    .main .block-container { max-width: 100% !important; width: 100% !important; padding: 1cm !important; }
                    @page { size: landscape; margin: 1cm; }
                }
                </style>
                """, unsafe_allow_html=True)
                st.sidebar.success("✅ PDF mode active. Press **Ctrl + P** (or Cmd + P) in your browser and select 'Save as PDF'. Ensure 'Background graphics' is enabled.")

            # ==========================================
            #           FUNCIÓN: SALARY PLANNING
            # ==========================================
            def render_salary():
                df_breakdown_org, df_breakdown_func, df_detalle = None, None, None
                st.subheader("💰 Cost Summary")
                
                col_b1, col_b2 = st.columns([1, 3])
                with col_b1:
                    budget_pct = st.number_input("🎯 Set Target Budget %:", min_value=0.0, max_value=100.0, value=3.0, step=0.1, key="sal_bgt")
                
                adj_pct = pd.to_numeric(df_filtered.get('%Adjustment', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
                promo_pct = pd.to_numeric(df_filtered.get('%Growth Promotion', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
                col_t_annual_usd = pd.to_numeric(df_filtered.get('$ Annual Salary(in USD)', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
                col_au_new_annual_usd = pd.to_numeric(df_filtered.get('$ New Annual Salary(in USD)', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)

                cost_adj = ((adj_pct / 100) * col_t_annual_usd)[adj_pct > 0].sum()
                cost_promo = ((promo_pct / 100) * col_t_annual_usd)[promo_pct > 0].sum()
                total_cost = cost_adj + cost_promo
                sum_t = col_t_annual_usd.sum()
                sum_au = col_au_new_annual_usd.sum()
                
                pct_adj_vs_total = (cost_adj / sum_t) * 100 if sum_t > 0 else 0
                pct_promo_vs_total = (cost_promo / sum_t) * 100 if sum_t > 0 else 0
                pct_total_cost_vs_total = (total_cost / sum_t) * 100 if sum_t > 0 else 0
                pct_incremento = ((sum_au / sum_t) - 1) * 100 if sum_t > 0 else 0

                budget_amount = sum_t * (budget_pct / 100)
                variance = total_cost - budget_amount

                if variance <= 0:
                    status_html = f"<span style='color: green; font-weight: bold;'>${variance:,.2f} (On/Under Budget)</span>"
                else:
                    status_html = f"<span style='color: red; font-weight: bold;'>+${variance:,.2f} (Over Budget)</span>"

                cost_df = pd.DataFrame({
                    "Concept": ["Adjustment Cost", "Growth Promotion Cost", "Total Cost", "Total Budget Amount", "Total % Increment"],
                    "Value": [f"${cost_adj:,.2f}", f"${cost_promo:,.2f}", f"${total_cost:,.2f}", f"${budget_amount:,.2f}", f"{pct_incremento:,.2f}%"],
                    "% of Total Salary": [f"{pct_adj_vs_total:,.2f}%", f"{pct_promo_vs_total:,.2f}%", f"{pct_total_cost_vs_total:,.2f}%", f"{budget_pct:,.2f}%", "-"]
                })
                
                summary_html = f"""
                <table style="width:100%; text-align:left; border-collapse: collapse; margin-bottom: 25px; font-size: 15px;">
                  <tr style="background-color: #f0f2f6; border-bottom: 2px solid #ccc;">
                    <th style="padding: 10px; border: 1px solid #e0e0e0; color: #333;">Adj Cost</th>
                    <th style="padding: 10px; border: 1px solid #e0e0e0; color: #333;">Promo Cost</th>
