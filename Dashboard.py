import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json
import os
import numpy as np
import io

# Page Configuration
st.set_page_config(page_title="Salary Planning Dashboard", layout="wide")
st.title("Salary Planning Dashboard")
st.write("Upload your Excel files to view and filter metrics.")

# --- FILTER MEMORY ---
ARCHIVO_FILTROS = 'mis_filtros_guardados.json'

def cargar_memoria():
    if os.path.exists(ARCHIVO_FILTROS):
        with open(ARCHIVO_FILTROS, 'r', encoding='utf-8') as f:
            return json.load(f)
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
        return pd.to_datetime(val)
    except Exception:
        return pd.NaT

# --- FILE UPLOADER ---
col_up1, col_up2 = st.columns(2)
with col_up1:
    file1 = st.file_uploader("1. Upload the Main file (e.g., colgate_...xlsx)", type=["xlsx", "xls"])
with col_up2:
    file2 = st.file_uploader("2. Upload the SuperFlex file (SuperFlex-DefaultView...)", type=["xlsx", "xls"])

if file1 is not None and file2 is not None:
    try:
        df1 = pd.read_excel(file1, sheet_name="Salary")
        df2 = pd.read_excel(file2, header=2)
        
        # Format Global ID to merge correctly
        df1['Global ID'] = pd.to_numeric(df1['Global ID'], errors='coerce')
        df2['Global ID'] = pd.to_numeric(df2['Global ID'], errors='coerce')
        
        # Extraer columna AC (índice 28) como Employee Subgroup
        if len(df2.columns) > 28:
            df2['Employee Subgroup'] = df2.iloc[:, 28].astype(str)
        else:
            df2['Employee Subgroup'] = "N/A"
            
        # Extraer columna AV (índice 47) como Gender
        if len(df2.columns) > 47:
            df2['Gender'] = df2.iloc[:, 47].astype(str)
        else:
            df2['Gender'] = "Unknown"
        
        # Merge Main File with SuperFlex data
        df = pd.merge(df1, df2[['Global ID', 'Chief Name', 'Employee Subgroup', 'Gender']], on='Global ID', how='left')
        
        # Limpieza inicial de variables críticas
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
        filtro_elegido = st.sidebar.selectbox("Load a configuration:", nombres_disponibles)
        
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

        selected_gerentes = st.sidebar.multiselect("Manager(s):", gerente_options, default=def_gerentes)
        selected_orgs = st.sidebar.multiselect("Reporting Organization:", org_options, default=def_orgs)
        selected_funcs = st.sidebar.multiselect("Function:", func_options, default=def_funcs)
        selected_comps = st.sidebar.multiselect("Compensation Area:", comp_options, default=def_comps)
        selected_subgroups = st.sidebar.multiselect("Employee Subgroup:", subgroup_options, default=def_subgroups)
        selected_potentials = st.sidebar.multiselect("Potential (Col Y):", potential_options, default=def_potentials)
        
        # --- SIDEBAR: SAVE NEW FILTER ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("✏️ Save current combination")
        nuevo_nombre = st.sidebar.text_input("Give this filter a title:")
        
        if st.sidebar.button("Save Filter"):
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

        # --- CÁLCULOS GLOBALES (YEARS EN SG Y QUARTILES) ---
        col_date_sg = None
        col_curr_qrtle = None
        col_new_qrtle = None
        
        for c in df_filtered.columns:
            c_lower = str(c).lower().strip()
            if 'date in sg' in c_lower or 'date in salary group' in c_lower:
                col_date_sg = c
            elif 'new qrtle' in c_lower or 'new quartile' in c_lower or ('quartile' in c_lower and 'new' in c_lower):
                col_new_qrtle = c
            elif ('qrtle' in c_lower or 'quartile' in c_lower) and 'new' not in c_lower:
                col_curr_qrtle = c
                    
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

        mapeo_q = {
            "Below Minimum": "Below Min", "Below Min": "Below Min",
            "1Q": "1Q", "2Q": "2Q", "3Q": "3Q", "4Q": "4Q",
            "AboveMax": "Above max", "Above max": "Above max"
        }

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
            # SECCIÓN DE EXPORTACIÓN Y PDF (SIDEBAR)
            # ==========================================
            st.sidebar.markdown("---")
            st.sidebar.subheader("📄 Export Report")
            pdf_mode = st.sidebar.checkbox("🖨️ Enable PDF Print Mode (All-in-one view)")
            
            if pdf_mode:
                components.html(
                    """
                    <script>
                    function printDashboard() {
                        window.parent.print();
                    }
                    </script>
                    <button onclick="printDashboard()" style="width: 100%; padding: 10px; background-color: #4CAF50; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold; font-family: sans-serif;">
                        Save Dashboard as PDF
                    </button>
                    """,
                    height=50
                )
                st.sidebar.info("Tip: In the print dialog, select 'Save as PDF', enable 'Background Graphics', and hide Headers/Footers.")

            # ==========================================
            #           FUNCIÓN: SALARY PLANNING
            # ==========================================
            def render_salary():
                df_breakdown_org, df_breakdown_func, df_detalle = None, None, None
                
                st.subheader("💰 Cost Summary")
                
                col_b1, col_b2 = st.columns([1, 3])
                with col_b1:
                    budget_pct = st.number_input("🎯 Set Target Budget %:", min_value=0.0, max_value=100.0, value=3.0, step=0.1)
                
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
                
                st.table(cost_df)
                st.markdown(f"<div style='margin-bottom: 20px; font-size: 18px;'><strong>Variance vs BGT:</strong> {status_html}</div>", unsafe_allow_html=True)
                
                # --- DESGLOSE DE COSTOS (BREAKDOWN) ---
                st.markdown("### 🏢 Cost Breakdown (By Org & Function)")
                tab_org, tab_func = st.tabs(["By Reporting Organization", "By Function"])
                
                def build_cost_breakdown(df_base, col_name, b_pct):
                    a_pct = pd.to_numeric(df_base.get('%Adjustment', pd.Series(0, index=df_base.index)), errors='coerce').fillna(0)
                    p_pct = pd.to_numeric(df_base.get('%Growth Promotion', pd.Series(0, index=df_base.index)), errors='coerce').fillna(0)
                    s_t = pd.to_numeric(df_base.get('$ Annual Salary(in USD)', pd.Series(0, index=df_base.index)), errors='coerce').fillna(0)
                    s_au = pd.to_numeric(df_base.get('$ New Annual Salary(in USD)', pd.Series(0, index=df_base.index)), errors='coerce').fillna(0)
                    
                    df_temp = pd.DataFrame({
                        'Group': df_base[col_name].fillna('Unknown'),
                        'Adj_Cost': np.where(a_pct > 0, (a_pct / 100) * s_t, 0),
                        'Promo_Cost': np.where(p_pct > 0, (p_pct / 100) * s_t, 0),
                        'Total_Salary': s_t,
                        'New_Salary': s_au
                    })
                    
                    grp = df_temp.groupby('Group').sum().reset_index()
                    grp['Total Cost'] = grp['Adj_Cost'] + grp['Promo_Cost']
                    grp['Budget Amount'] = grp['Total_Salary'] * (b_pct / 100)
                    grp['Variance vs BGT'] = grp['Total Cost'] - grp['Budget Amount']
                    grp['Status'] = np.where(grp['Variance vs BGT'] > 0, 'Over Budget', 'On/Under Budget')
                    grp['Total % Increment'] = np.where(grp['Total_Salary'] > 0, ((grp['New_Salary'] / grp['Total_Salary']) - 1) * 100, 0)
                    
                    grp = grp[['Group', 'Adj_Cost', 'Promo_Cost', 'Total Cost', 'Budget Amount', 'Variance vs BGT', 'Status', 'Total % Increment']]
                    grp.rename(columns={'Group': col_name, 'Adj_Cost': 'Adjustment Cost', 'Promo_Cost': 'Growth Promo Cost'}, inplace=True)
                    return grp
                
                def color_salary_variance(row):
                    if row['Variance vs BGT'] > 0:
                        return ['color: red; font-weight: bold' if col in ['Variance vs BGT', 'Status'] else '' for col in row.index]
                    else:
                        return ['color: green; font-weight: bold' if col in ['Variance vs BGT', 'Status'] else '' for col in row.index]

                format_dict = {
                    'Adjustment Cost': '${:,.2f}',
                    'Growth Promo Cost': '${:,.2f}',
                    'Total Cost': '${:,.2f}',
                    'Budget Amount': '${:,.2f}',
                    'Variance vs BGT': '${:,.2f}',
                    'Total % Increment': '{:.2f}%'
                }
                
                with tab_org:
                    if 'Reporting Organization' in df_filtered.columns:
                        df_breakdown_org = build_cost_breakdown(df_filtered, 'Reporting Organization', budget_pct)
                        st.dataframe(df_breakdown_org.style.format(format_dict).apply(color_salary_variance, axis=1), use_container_width=True, hide_index=True)
                    else:
                        st.info("Reporting Organization column not found.")
                        
                with tab_func:
                    if 'Function' in df_filtered.columns:
                        df_breakdown_func = build_cost_breakdown(df_filtered, 'Function', budget_pct)
                        st.dataframe(df_breakdown_func.style.format(format_dict).apply(color_salary_variance, axis=1), use_container_width=True, hide_index=True)
                    else:
                        st.info("Function column not found.")

                st.markdown("---")

                # --- PREPARE DATA FOR CHARTS ---
                cond_adj = adj_pct > 0
                cond_promo = promo_pct > 0
                tiene_movimiento = cond_adj | cond_promo
                solo_adj = (cond_adj & ~cond_promo)
                solo_promo = (~cond_adj & cond_promo)
                ambos = (cond_adj & cond_promo)
                sin_mov = (~cond_adj & ~cond_promo)
                
                num_solo_adj = solo_adj.sum()
                num_solo_promo = solo_promo.sum()
                num_ambos = ambos.sum()
                num_sin_mov = sin_mov.sum()
                total_personas = len(df_filtered)
                num_movimientos = num_solo_adj + num_solo_promo + num_ambos

                def make_autopct(values):
                    def my_autopct(pct):
                        total = sum(values)
                        val = int(round(pct * total / 100.0))
                        return f'{val}\n({pct:.1f}%)' if pct >= 5 else ''
                    return my_autopct

                st.subheader("📊 Graphical Summary")
                col1, col2 = st.columns(2)
                
                with col1:
                    fig1, ax1 = plt.subplots(figsize=(7, 6))
                    if total_personas > 0:
                        if num_movimientos == 0:
                            ax1.pie([100], colors=['#d3d3d3'], startangle=90)
                            ax1.legend(["No Movement (100%)"], loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        else:
                            sizes = [num_movimientos, total_personas - num_movimientos]
                            labels = ['With Movement', 'No Movement']
                            colors = ['#ff9999', '#d3d3d3']
                            
                            wedges, texts, autotexts = ax1.pie(
                                sizes, startangle=90, colors=colors, 
                                autopct=make_autopct(sizes), pctdistance=0.65,
                                wedgeprops=dict(edgecolor='white', linewidth=1.5),
                                textprops=dict(fontsize=10, fontweight='bold', color='#333333')
                            )
                            leyenda1 = [f"{l} - {s} ({s/total_personas*100:.1f}%)" for l, s in zip(labels, sizes)]
                            ax1.legend(wedges, leyenda1, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        
                        ax1.axis('equal') 
                        ax1.set_title('Overall Movement', fontweight='bold', pad=15)
                    else:
                        ax1.text(0.5, 0.5, "No data", ha='center', va='center')
                    st.pyplot(fig1)

                with col2:
                    fig2, ax2 = plt.subplots(figsize=(7, 6))
                    raw_sizes2 = [num_solo_adj, num_solo_promo, num_ambos, num_sin_mov]
                    raw_labels2 = ['Adjustment Only', 'Promotion Only', 'Both', 'No Movement']
                    raw_colors2 = ['#ffb3e6', '#c2c2f0', '#ff6666', '#c2f0c2']
                    
                    sizes2 = [s for s in raw_sizes2 if s > 0]
                    labels2 = [l for s, l in zip(raw_sizes2, raw_labels2) if s > 0]
                    colors2 = [c for s, c in zip(raw_sizes2, raw_colors2) if s > 0]
                    total_chart2 = sum(sizes2)
                    
                    if total_chart2 > 0:
                        wedges2, texts2, autotexts2 = ax2.pie(
                            sizes2, startangle=90, colors=colors2,
                            autopct=make_autopct(sizes2), pctdistance=0.65,
                            wedgeprops=dict(edgecolor='white', linewidth=1.5),
                            textprops=dict(fontsize=10, fontweight='bold', color='#333333')
                        )
                        leyenda2 = [f"{l} - {s} ({s/total_chart2*100:.1f}%)" for l, s in zip(labels2, sizes2)]
                        ax2.legend(wedges2, leyenda2, title="Breakdown", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        ax2.axis('equal')
                        ax2.set_title('Movement Details', fontweight='bold', pad=15)
                    else:
                        ax2.text(0.5, 0.5, "No data", ha='center', va='center')
                    st.pyplot(fig2)

                st.markdown("<br>", unsafe_allow_html=True)
                col3, col4 = st.columns(2)

                with col3:
                    fig3, ax3 = plt.subplots(figsize=(7, 6))
                    df_adj = df_filtered[cond_adj].copy()
                    
                    if not df_adj.empty and 'Adjustment Reason' in df_adj.columns:
                        df_adj['Adjustment Reason'] = df_adj['Adjustment Reason'].fillna('No Reason Assigned').replace({'None Selected': 'No Reason Assigned'})
                        reason_counts = df_adj['Adjustment Reason'].value_counts()
                        total_reasons = reason_counts.sum()
                        
                        if total_reasons > 0:
                            colores_motivos = sns.color_palette("pastel", len(reason_counts))
                            
                            wedges3, texts3, autotexts3 = ax3.pie(
                                reason_counts.values, startangle=90, colors=colores_motivos,
                                autopct=make_autopct(reason_counts.values), pctdistance=0.65,
                                wedgeprops=dict(edgecolor='white', linewidth=1.5),
                                textprops=dict(fontsize=10, fontweight='bold', color='#333333')
                            )
                            leyenda3 = [f"{i} - {v} ({v/total_reasons*100:.1f}%)" for i, v in reason_counts.items()]
                            ax3.legend(wedges3, leyenda3, title="Reasons", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                            ax3.axis('equal')
                            ax3.set_title('Adjustment Split (Reason)', fontweight='bold', pad=15)
                        else:
                            ax3.text(0.5, 0.5, "No valid data", ha='center', va='center')
                    else:
                        ax3.text(0.5, 0.5, "No adjustments to analyze", ha='center', va='center')
                    st.pyplot(fig3)

                with col4:
                    fig4, ax4 = plt.subplots(figsize=(7, 6))
                    df_pot_mov = df_filtered[tiene_movimiento].copy()
                    
                    if not df_pot_mov.empty and 'Potential' in df_pot_mov.columns:
                        pot_counts = df_pot_mov['Potential'].value_counts()
                        total_pot = pot_counts.sum()
                        
                        if total_pot > 0:
                            colores_pot = sns.color_palette("Set3", len(pot_counts))
                            
                            wedges4, texts4, autotexts4 = ax4.pie(
                                pot_counts.values, startangle=90, colors=colores_pot,
                                autopct=make_autopct(pot_counts.values), pctdistance=0.65,
                                wedgeprops=dict(edgecolor='white', linewidth=1.5),
                                textprops=dict(fontsize=10, fontweight='bold', color='#333333')
                            )
                            leyenda4 = [f"{i} - {v} ({v/total_pot*100:.1f}%)" for i, v in pot_counts.items()]
                            ax4.legend(wedges4, leyenda4, title="Potential Rating", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                            ax4.axis('equal')
                            ax4.set_title('Movements by Potential Split', fontweight='bold', pad=15)
                        else:
                            ax4.text(0.5, 0.5, "No valid data", ha='center', va='center')
                    else:
                        ax4.text(0.5, 0.5, "No movements", ha='center', va='center')
                    st.pyplot(fig4)

                st.markdown("<br>", unsafe_allow_html=True)
                col5, col6 = st.columns(2)

                with col5:
                    fig5, ax5 = plt.subplots(figsize=(7, 6))
                    if not df_filtered.empty and 'Current Quartile' in df_filtered.columns and 'New Quartile' in df_filtered.columns:
                        counts_curr = df_filtered['Current Quartile'].value_counts()
                        counts_new = df_filtered['New Quartile'].value_counts()
                        
                        final_order = ["Below Min", "1Q", "2Q", "3Q", "4Q", "Above max"]
                        val_curr = [counts_curr.get(c, 0) for c in final_order]
                        val_new = [counts_new.get(c, 0) for c in final_order]
                        
                        x_pos = np.arange(len(final_order))
                        ancho_barra = 0.35
                        
                        label_curr = str(col_curr_qrtle)[:20] if col_curr_qrtle else 'Current Qrtle'
                        label_new = str(col_new_qrtle)[:20] if col_new_qrtle else 'New Qrtle'
                        
                        bars_curr = ax5.bar(x_pos - ancho_barra/2, val_curr, ancho_barra, label=label_curr, color='#ffb347')
                        bars_new = ax5.bar(x_pos + ancho_barra/2, val_new, ancho_barra, label=label_new, color='#87cefa')
                        
                        max_y = max(max(val_curr, default=0), max(val_new, default=0))
                        if max_y == 0: max_y = 1
                        
                        ax5.bar_label(bars_curr, padding=3, fontsize=9, color='#333333')
                        ax5.bar_label(bars_new, padding=3, fontsize=9, color='#333333')
                        
                        ax5.set_xticks(x_pos)
                        ax5.set_xticklabels(final_order, rotation=45, ha='right', fontsize=9)
                        ax5.set_title('Overall Quartile Distribution', fontweight='bold', pad=15)
                        ax5.set_ylabel('Employees')
                        ax5.set_ylim(0, max_y * 1.15)
                        ax5.spines['top'].set_visible(False)
                        ax5.spines['right'].set_visible(False)
                        ax5.legend(loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2, fontsize=10)
                        fig5.tight_layout()
                    else:
                        ax5.text(0.5, 0.5, "Quartile Columns missing", ha='center', va='center')
                    st.pyplot(fig5)

                with col6:
                    fig6, ax6 = plt.subplots(figsize=(7, 6))
                    df_gender_mov = df_filtered[tiene_movimiento].copy()
                    
                    if not df_gender_mov.empty and 'Gender' in df_gender_mov.columns:
                        gender_counts = df_gender_mov['Gender'].value_counts()
                        total_gender = gender_counts.sum()
                        
                        if total_gender > 0:
                            colores_gender = sns.color_palette("pastel", len(gender_counts))
                            bars = ax6.bar(gender_counts.index, gender_counts.values, color=colores_gender)
                            for bar in bars:
                                yval = bar.get_height()
                                pct = (yval / total_gender) * 100
                                ax6.text(bar.get_x() + bar.get_width()/2, yval + (total_gender * 0.01), 
                                         f'{int(yval)}\n({pct:.1f}%)', ha='center', va='bottom', fontsize=10, fontweight='bold')
                                
                            ax6.set_title('Movements by Gender Split', fontweight='bold', pad=15)
                            ax6.set_ylabel('Employees')
                            ax6.spines['top'].set_visible(False)
                            ax6.spines['right'].set_visible(False)
                            fig6.tight_layout()
                        else:
                            ax6.text(0.5, 0.5, "No data", ha='center', va='center')
                    else:
                        ax6.text(0.5, 0.5, "No Gender data", ha='center', va='center')
                    st.pyplot(fig6)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("##### 📈 Years in SG vs New Quartile (High Potential & Strategic Few)")
                
                mask_pot = df_filtered['Potential'].astype(str).str.strip().str.lower().isin(['high potential', 'strategic few'])
                df_pot_filtered = df_filtered[mask_pot].copy()
                
                if not df_pot_filtered.empty and 'Years in SG' in df_pot_filtered.columns and 'New Quartile' in df_pot_filtered.columns:
                    orden_qrtle = ["Below Min", "1Q", "2Q", "3Q", "4Q", "Above max"]
                    pivot_df = pd.crosstab(df_pot_filtered['Years in SG'], df_pot_filtered['New Quartile'])
                    
                    cols_presentes = [c for c in orden_qrtle if c in pivot_df.columns]
                    otras_cols = [c for c in pivot_df.columns if c not in orden_qrtle]
                    pivot_df = pivot_df[cols_presentes + otras_cols]
                    
                    def sort_years(idx):
                        if "Unknown" in idx: return 999
                        if "< 1" in idx: return 0
                        try: return int(idx.split()[0])
                        except: return 999
                    
                    pivot_df = pivot_df.loc[sorted(pivot_df.index, key=sort_years)]
                    
                    fig_new, ax_new = plt.subplots(figsize=(10, 5))
                    bars = pivot_df.plot(kind='bar', ax=ax_new, stacked=False, colormap='viridis', width=0.8)
                    
                    for container in ax_new.containers:
                        labels = [f'{int(v)}' if v > 0 else '' for v in container.datavalues]
                        ax_new.bar_label(container, labels=labels, padding=3, fontsize=9, color='#333333')
                    
                    ax_new.set_title("Distribution of Years in SG by New Quartile", fontweight='bold', pad=15)
                    ax_new.set_ylabel("Number of Employees")
                    ax_new.set_xlabel("Years in Subgroup")
                    
                    ax_new.spines['top'].set_visible(False)
                    ax_new.spines['right'].set_visible(False)
                    
                    max_y = pivot_df.values.max() if not pivot_df.empty else 1
                    ax_new.set_ylim(0, max_y * 1.15)
                    
                    plt.xticks(rotation=0)
                    ax_new.legend(title="New Quartile", bbox_to_anchor=(1.05, 1), loc='upper left')
                    fig_new.tight_layout()
                    
                    st.pyplot(fig_new)
                else:
                    st.info("No data found for 'High Potential' or 'Strategic Few' in the current filter.")
                
                st.markdown("---")

                # --- DYNAMIC STAFF TABLE ---
                st.subheader("👥 Employee Detailed List & Alerts")
                
                col_filt1, col_filt2, col_filt3 = st.columns(3)
                
                with col_filt1:
                    opcion_detalle = st.radio(
                        "1. Filter by Movement Type:",
                        ["All Employees", "Adjustment Only", "Promotion Only", "Both", "No Movement"],
                        horizontal=True,
                        key='rad_mov'
                    )
                    
                with col_filt2:
                    opcion_alerta = st.selectbox(
                        "2. Filter by Alerts (Colors):",
                        [
                            "Show All", 
                            "⚠️ Show Only with Alerts (Any Color)", 
                            "🔴 Red Alerts Only (Critical)", 
                            "🟠 Orange Alerts Only (Warning)", 
                            "🟡 Yellow Alerts Only (Notice)"
                        ],
                        key='sel_alert'
                    )
                    
                with col_filt3:
                    search_emp = st.text_input(
                        "3. Search Employee (Exact ID or Name):", 
                        "", 
                        placeholder="Type Name or Exact Global ID...", 
                        key='search_salary'
                    )

                col_filt4, col_filt5 = st.columns(2)
                
                def sort_years_opts(x):
                    if x == "Unknown": return 999
                    if "< 1" in x: return 0
                    try: return int(x.split()[0])
                    except: return 999
                    
                with col_filt4:
                    opts_years = sorted(df_filtered['Years in SG'].unique().tolist(), key=sort_years_opts)
                    selected_years = st.multiselect("4. Filter by Years in SG:", options=opts_years, default=[])
                    
                with col_filt5:
                    base_qrtles = ["Below Min", "1Q", "2Q", "3Q", "4Q", "Above max", "Unknown"]
                    opts_qrtle = [q for q in base_qrtles if q in df_filtered['New Quartile'].unique().tolist()]
                    opts_qrtle += [q for q in df_filtered['New Quartile'].unique().tolist() if q not in opts_qrtle]
                    selected_qrtle = st.multiselect("5. Filter by New Quartile:", options=opts_qrtle, default=[])

                if opcion_detalle == "Adjustment Only": mask = solo_adj
                elif opcion_detalle == "Promotion Only": mask = solo_promo
                elif opcion_detalle == "Both": mask = ambos
                elif opcion_detalle == "No Movement": mask = sin_mov
                else: mask = pd.Series(True, index=df_filtered.index)
                    
                df_detalle = df_filtered[mask].copy()
                df_detalle = df_detalle.rename(columns={'Chief Name': 'Manager'})
                
                if selected_years:
                    df_detalle = df_detalle[df_detalle['Years in SG'].isin(selected_years)]
                if selected_qrtle:
                    df_detalle = df_detalle[df_detalle['New Quartile'].isin(selected_qrtle)]

                def evaluar_alertas(row):
                    comentarios = []
                    color = ''
                    if len(row) > 37:
                        val_j = get_num(row.iloc[9])
                        val_z = get_date(row.iloc[25])
                        val_ab = get_num(row.iloc[27])
                        val_ac = get_num(row.iloc[28])
                        val_af = get_num(row.iloc[31])
                        val_ah = str(row.iloc[33]).strip().lower()
                        val_ai = get_num(row.iloc[34])
                        val_ak = get_num(row.iloc[36])
                        val_al_raw = str(row.iloc[37]).strip().lower()
                        al_vacio = val_al_raw in ['', 'nan', 'nat', 'none', 'null']
                        
                        flag_rojo = False
                        flag_naranja = False
                        flag_amarillo = False
                        
                        if pd.notna(val_z):
                            delta_dias = abs((FECHA_ACTUAL - val_z).days)
                            if (0.01 < val_af < 1) and (val_ab > 0 or val_ac > 0) and delta_dias <= 182:
                                flag_naranja = True
                                comentarios.append("Revisar Adjustment vs Fecha reciente (<=6 meses)")
                        if val_af > 0 and val_ah in ["none selected", "nan", "", "none"]:
                            flag_amarillo = True
                            comentarios.append("Adjustment con 'None Selected'")
                        if (0 < val_ak < 6) or (val_ak > 15):
                            flag_amarillo = True
                            comentarios.append("Valor AK fuera de rango recomendado")
                        if val_ak > 0 and al_vacio:
                            flag_amarillo = True
                            comentarios.append("AK > 0 pero Columna AL está vacía")
                        if val_ai > val_j and val_ak > 0:
                            flag_rojo = True
                            comentarios.append("AI supera el valor de J y AK > 0")
                        
                        if flag_rojo: color = '#ffcccc'
                        elif flag_naranja: color = '#ffe4b5'
                        elif flag_amarillo: color = '#ffffcc'
                        
                    return [" | ".join(comentarios), color]

                if not df_detalle.empty:
                    alertas_lista = df_detalle.apply(evaluar_alertas, axis=1).tolist()
                    df_detalle['ALERT_COMMENTS'] = [x[0] for x in alertas_lista]
                    df_detalle['RowColor'] = [x[1] for x in alertas_lista]
                    
                    if opcion_alerta == "⚠️ Show Only with Alerts (Any Color)":
                        df_detalle = df_detalle[df_detalle['RowColor'] != '']
                    elif opcion_alerta == "🔴 Red Alerts Only (Critical)":
                        df_detalle = df_detalle[df_detalle['RowColor'] == '#ffcccc']
                    elif opcion_alerta == "🟠 Orange Alerts Only (Warning)":
                        df_detalle = df_detalle[df_detalle['RowColor'] == '#ffe4b5']
                    elif opcion_alerta == "🟡 Yellow Alerts Only (Notice)":
                        df_detalle = df_detalle[df_detalle['RowColor'] == '#ffffcc']
                    
                    if search_emp.strip() and not df_detalle.empty:
                        term = search_emp.strip()
                        if term.isdigit() and 'Global ID' in df_detalle.columns:
                            mask_exact = df_detalle['Global ID'].astype(str).str.replace(r'\.0$', '', regex=True) == term
                            df_detalle = df_detalle[mask_exact]
                        else:
                            name_cols = [c for c in df_detalle.columns if c.strip().lower() in ['name', 'employee name', 'full name', 'first name', 'last name']]
                            if not name_cols:
                                name_cols = [c for c in df_detalle.columns if 'name' in c.lower() and 'manager' not in c.lower()]
                            if name_cols:
                                mask_search = np.column_stack([df_detalle[col].astype(str).str.contains(term, case=False, na=False) for col in name_cols]).any(axis=1)
                                df_detalle = df_detalle[mask_search]
                            else:
                                df_detalle = df_detalle.iloc[0:0]

                    if not df_detalle.empty:
                        st.write(f"Showing **{len(df_detalle)}** matching employees.")
                        columnas_todas = list(df_detalle.columns)
                        if 'ALERT_COMMENTS' in columnas_todas:
                            columnas_todas.insert(4, columnas_todas.pop(columnas_todas.index('ALERT_COMMENTS')))
                            df_detalle = df_detalle[columnas_todas]
                        
                        cols_to_drop = ['RowColor', 'Date_in_SG_clean', 'Years_Raw', 'Current Quartile', 'New Quartile']
                        df_visual = df_detalle.drop(columns=[c for c in cols_to_drop if c in df_detalle.columns])
                        colores_array = df_detalle['RowColor'].values
                        
                        cols_fijas = list(df_visual.columns[:4])
                        df_visual = df_visual.set_index(cols_fijas)
                        
                        def aplicar_colores(df_vista):
                            estilos = [f"background-color: {c}" if c else "" for c in colores_array]
                            style_dict = {col: estilos for col in df_vista.columns}
                            return pd.DataFrame(style_dict, index=df_vista.index)
                        
                        df_estilizado = df_visual.style.apply(aplicar_colores, axis=None)
                        st.dataframe(df_estilizado, use_container_width=True)
                    else:
                        st.info("No employees match your search term or filters.")
                else:
                    st.info("No employees match the selected Movement Type or Years/Quartile filters.")
                    
                return cost_df, df_breakdown_org, df_breakdown_func, df_detalle


            # ==========================================
            #           FUNCIÓN: EQUITY PLANNING
            # ==========================================
            def render_equity():
                eq_summary_df, df_brk_org, df_brk_func, df_equity_filtered = None, None, None, None
                st.subheader("📈 Equity Planning")
                st.info("This section displays data exclusively for employees eligible for stock (Column BJ = 'Yes').")
                
                col_bj_idx = 61  
                col_bl_idx = 63  
                col_bn_idx = 65  
                col_bp_idx = 67  
                
                if len(df_filtered.columns) > max(col_bn_idx, col_bp_idx):
                    mask_stock = df_filtered.iloc[:, col_bj_idx].astype(str).str.strip().str.lower() == 'yes'
                    df_equity_full = df_filtered[mask_stock].copy()
                    
                    if not df_equity_full.empty:
                        df_equity_full = df_equity_full.rename(columns={'Chief Name': 'Manager'})
                        
                        val_bn_full = pd.to_numeric(df_equity_full.iloc[:, col_bn_idx], errors='coerce').fillna(0)
                        val_bl_full = pd.to_numeric(df_equity_full.iloc[:, col_bl_idx], errors='coerce').fillna(0)
                        
                        df_equity_full['Proposed Recommendation Value'] = val_bl_full * (val_bn_full / 100)
                        
                        c_unplanned = (val_bn_full == 0).sum()
                        c_below = ((val_bn_full > 0) & (val_bn_full < 100)).sum()
                        c_midpoint = (val_bn_full == 100).sum()
                        c_above = (val_bn_full > 100).sum()
                        
                        st.markdown("### Equity Overview")
                        col_pie, col_summary = st.columns([1, 1.5])
                        
                        def make_autopct(values):
                            def my_autopct(pct):
                                total = sum(values)
                                val = int(round(pct * total / 100.0))
                                return f'{val}\n({pct:.1f}%)' if pct >= 5 else ''
                            return my_autopct
                        
                        with col_pie:
                            pie_labels_raw = ["Unplanned equity", "Below Midpoint", "Midpoint", "Above midpoint"]
                            pie_sizes_raw = [c_unplanned, c_below, c_midpoint, c_above]
                            pie_colors_raw = ['#d3d3d3', '#ffb347', '#87cefa', '#ff9999'] 
                            
                            pie_labels = [l for l, s in zip(pie_labels_raw, pie_sizes_raw) if s > 0]
                            pie_sizes = [s for s in pie_sizes_raw if s > 0]
                            pie_colors = [c for c, s in zip(pie_colors_raw, pie_sizes_raw) if s > 0]
                            
                            fig_eq, ax_eq = plt.subplots(figsize=(5, 4))
                            if sum(pie_sizes) > 0:
                                wedges_eq, texts_eq, autotexts_eq = ax_eq.pie(
                                    pie_sizes, startangle=90, colors=pie_colors,
                                    autopct=make_autopct(pie_sizes), pctdistance=0.65,
                                    wedgeprops=dict(edgecolor='white', linewidth=1.5),
                                    textprops=dict(fontsize=10, fontweight='bold', color='#333333')
                                )
                                leyenda_eq = [f"{l} - {s}" for l, s in zip(pie_labels, pie_sizes)]
                                ax_eq.legend(wedges_eq, leyenda_eq, title="Distribution", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                                ax_eq.axis('equal')
                            else:
                                ax_eq.text(0.5, 0.5, "No data available", ha='center', va='center')
                            st.pyplot(fig_eq)
                        
                        with col_summary:
                            equity_budget = val_bl_full.sum()
                            total_proposed = df_equity_full['Proposed Recommendation Value'].sum()
                            variance = total_proposed - equity_budget
                            
                            eq_summary_df = pd.DataFrame({
                                "Equity Budget (Total BL)": [f"${equity_budget:,.2f}"],
                                "Proposed Recommendation Value": [f"${total_proposed:,.2f}"],
                                "Variance vs BGT": [f"${variance:,.2f}"]
                            })
                            
                            if variance <= 0:
                                status_html = f"<span style='color: green; font-weight: bold;'>${variance:,.2f} (On/Under Budget)</span>"
                            else:
                                status_html = f"<span style='color: red; font-weight: bold;'>+${variance:,.2f} (Over Budget)</span>"
                                
                            st.markdown("##### Total Equity Budget Summary")
                            summary_html = f"""
                            <table style="width:100%; text-align:left; border-collapse: collapse; margin-bottom: 25px;">
                              <tr style="background-color: #f0f2f6; border-bottom: 2px solid #ccc;">
                                <th style="padding: 12px; border: 1px solid #e0e0e0; color: #333;">Equity Budget (Total BL)</th>
                                <th style="padding: 12px; border: 1px solid #e0e0e0; color: #333;">Proposed Recommendation Value</th>
                                <th style="padding: 12px; border: 1px solid #e0e0e0; color: #333;">Status (Variance)</th>
                              </tr>
                              <tr style="background-color: white;">
                                <td style="padding: 12px; border: 1px solid #e0e0e0; font-size: 16px;">${equity_budget:,.2f}</td>
                                <td style="padding: 12px; border: 1px solid #e0e0e0; font-size: 16px;">${total_proposed:,.2f}</td>
                                <td style="padding: 12px; border: 1px solid #e0e0e0; font-size: 16px;">{status_html}</td>
                              </tr>
                            </table>
                            """
                            st.markdown(summary_html, unsafe_allow_html=True)
                            
                        # --- DESGLOSE DE EQUITY (BREAKDOWN) ---
                        st.markdown("### 🏢 Equity Breakdown (By Org & Function)")
                        tab_eq_org, tab_eq_func = st.tabs(["By Reporting Organization", "By Function"])
                        
                        def build_equity_breakdown(df_base, col_name):
                            df_temp = pd.DataFrame({
                                'Group': df_base[col_name].fillna('Unknown'),
                                'Equity Budget': pd.to_numeric(df_base.iloc[:, col_bl_idx], errors='coerce').fillna(0),
                                'Proposed Value': df_base['Proposed Recommendation Value']
                            })
                            grp = df_temp.groupby('Group').sum().reset_index()
                            grp['Variance vs BGT'] = grp['Proposed Value'] - grp['Equity Budget']
                            grp['Status'] = np.where(grp['Variance vs BGT'] > 0, 'Over Budget', 'On/Under Budget')
                            grp.rename(columns={'Group': col_name}, inplace=True)
                            return grp

                        def color_variance_row(row):
                            if row['Variance vs BGT'] > 0:
                                return ['color: red; font-weight: bold' if col in ['Variance vs BGT', 'Status'] else '' for col in row.index]
                            else:
                                return ['color: green; font-weight: bold' if col in ['Variance vs BGT', 'Status'] else '' for col in row.index]
                                
                        format_dict_eq = {
                            'Equity Budget': '${:,.2f}',
                            'Proposed Value': '${:,.2f}',
                            'Variance vs BGT': '${:,.2f}'
                        }

                        with tab_eq_org:
                            if 'Reporting Organization' in df_filtered.columns:
                                df_brk_org = build_equity_breakdown(df_equity_full, 'Reporting Organization')
                                st.dataframe(df_brk_org.style.format(format_dict_eq).apply(color_variance_row, axis=1), use_container_width=True, hide_index=True)
                            else:
                                st.info("Reporting Organization column not found.")
                                
                        with tab_eq_func:
                            if 'Function' in df_filtered.columns:
                                df_brk_func = build_equity_breakdown(df_equity_full, 'Function')
                                st.dataframe(df_brk_func.style.format(format_dict_eq).apply(color_variance_row, axis=1), use_container_width=True, hide_index=True)
                            else:
                                st.info("Function column not found.")
                                
                        st.markdown("---")
                        st.markdown("### 👥 Eligible Employees List")
                        
                        col_eq_filt1, col_eq_filt2 = st.columns(2)
                        opciones_conceptos = ["0%", "Below 100%", "On Midpoint 100%", "Above Midpoint", "Above Midpoint w/o Comments"]
                        
                        with col_eq_filt1:
                            seleccion_conceptos = st.multiselect("1. Select Categories to display:", options=opciones_conceptos, default=[])
                        with col_eq_filt2:
                            search_eq = st.text_input("2. Search Employee (Exact ID or Name):", "", placeholder="Type Name or Exact Global ID...", key='search_equity')

                        def clasificar_bn(row):
                            val_bn = row.iloc[col_bn_idx]
                            val_bp = str(row.iloc[col_bp_idx]).strip()
                            if pd.isna(val_bn) or str(val_bn).strip() == '' or str(val_bn).lower() == 'nan': return "0%"
                            try: bn_num = float(val_bn)
                            except ValueError: return "0%"
                                
                            bp_vacio = (val_bp == '' or val_bp.lower() in ['nan', 'none', 'null'])
                            if bn_num == 0: return "0%"
                            elif bn_num < 100: return "Below 100%"
                            elif bn_num == 100: return "On Midpoint 100%"
                            elif bn_num > 100:
                                if bp_vacio: return "Above Midpoint w/o Comments" 
                                else: return "Above Midpoint" 
                            return "0%"

                        df_equity_full['ALERT_CATEGORY'] = df_equity_full.apply(clasificar_bn, axis=1)
                        
                        if seleccion_conceptos:
                            df_equity_filtered = df_equity_full[df_equity_full['ALERT_CATEGORY'].isin(seleccion_conceptos)]
                        else:
                            df_equity_filtered = df_equity_full.iloc[0:0] 
                            
                        if search_eq.strip() and not df_equity_filtered.empty:
                            term_eq = search_eq.strip()
                            if term_eq.isdigit() and 'Global ID' in df_equity_filtered.columns:
                                mask_exact_eq = df_equity_filtered['Global ID'].astype(str).str.replace(r'\.0$', '', regex=True) == term_eq
                                df_equity_filtered = df_equity_filtered[mask_exact_eq]
                            else:
                                name_cols_eq = [c for c in df_equity_filtered.columns if c.strip().lower() in ['name', 'employee name', 'full name', 'first name', 'last name']]
                                if not name_cols_eq: name_cols_eq = [c for c in df_equity_filtered.columns if 'name' in c.lower() and 'manager' not in c.lower()]
                                
                                if name_cols_eq:
                                    mask_search_eq = np.column_stack([df_equity_filtered[col].astype(str).str.contains(term_eq, case=False, na=False) for col in name_cols_eq]).any(axis=1)
                                    df_equity_filtered = df_equity_filtered[mask_search_eq]
                                else:
                                    df_equity_filtered = df_equity_filtered.iloc[0:0]
                            
                        if not df_equity_filtered.empty:
                            st.write(f"Showing **{len(df_equity_filtered)}** employees based on your criteria.")
                            cols_eq = list(df_equity_filtered.columns)
                            if 'ALERT_CATEGORY' in cols_eq:
                                cols_eq.insert(4, cols_eq.pop(cols_eq.index('ALERT_CATEGORY')))
                                df_equity_filtered = df_equity_filtered[cols_eq]

                            categorias_array = df_equity_filtered['ALERT_CATEGORY'].values
                            
                            cols_to_drop_eq = ['Date_in_SG_clean', 'Years_Raw', 'Current Quartile', 'New Quartile']
                            df_visual_eq = df_equity_filtered.drop(columns=[c for c in cols_to_drop_eq if c in df_equity_filtered.columns])
                            
                            cols_fijas_eq = list(df_visual_eq.columns[:4])
                            df_visual_eq = df_visual_eq.set_index(cols_fijas_eq)
                            
                            def color_equity_totales(df_vista):
                                estilos = []
                                for cat in categorias_array:
                                    if cat == "Above Midpoint w/o Comments": estilos.append('background-color: #ffcccc')
                                    elif cat == "Above Midpoint": estilos.append('background-color: #ffffcc')
                                    else: estilos.append('')
                                style_dict = {col: estilos for col in df_vista.columns}
                                return pd.DataFrame(style_dict, index=df_vista.index)
                            
                            df_eq_styled = df_visual_eq.style.apply(color_equity_totales, axis=None)
                            st.dataframe(df_eq_styled, use_container_width=True)
                        else:
                            if seleccion_conceptos: st.info("No employees match your search term in the selected categories.")
                            else: st.info("Select one or more categories above to view employee details in this table.")
                    else:
                        st.warning("No employees in the current filtered selection have 'Yes' in Stock Eligibility (Column BJ).")
                else:
                    st.error("The uploaded file does not contain enough columns to process Equity Planning metrics (requires up to Column BP).")
                    
                return eq_summary_df, df_brk_org, df_brk_func, df_equity_filtered

            # ==========================================
            #   LÓGICA PRINCIPAL DE RENDERIZADO
            # ==========================================
            if pdf_mode:
                sal_data = render_salary()
                st.markdown("<div style='page-break-after: always; margin-top: 50px;'></div>", unsafe_allow_html=True)
                eq_data = render_equity()
            else:
                tab_salary, tab_equity = st.tabs(["💰 Salary Planning", "📈 Equity Planning"])
                with tab_salary:
                    sal_data = render_salary()
                with tab_equity:
                    eq_data = render_equity()

            # ==========================================
            #   MOTOR DE EXPORTACIÓN A EXCEL EN SIDEBAR
            # ==========================================
            st.sidebar.markdown("---")
            st.sidebar.subheader("📥 Download Data (Excel)")
            st.sidebar.info("Download a complete Excel report with all summaries, breakdowns, and employee details matching your active filters.")
            
            output = io.BytesIO()
            
            # Intentamos usar xlsxwriter (ideal), si no existe, openpyxl, y si no, por defecto.
            try:
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
            except ImportError:
                try:
                    writer = pd.ExcelWriter(output, engine='openpyxl')
                except ImportError:
                    writer = pd.ExcelWriter(output)
                    
            with writer:
                # Escribimos hojas de Salary
                if sal_data:
                    cost_df, df_breakdown_org, df_breakdown_func, df_detalle = sal_data
                    if cost_df is not None: cost_df.to_excel(writer, sheet_name='Salary Summary', index=False)
                    if df_breakdown_org is not None: df_breakdown_org.to_excel(writer, sheet_name='Salary by Org', index=False)
                    if df_breakdown_func is not None: df_breakdown_func.to_excel(writer, sheet_name='Salary by Func', index=False)
                    if df_detalle is not None and not df_detalle.empty: 
                        df_det_excel = df_detalle.copy()
                        cols_d = ['RowColor', 'Date_in_SG_clean', 'Years_Raw', 'Current Quartile', 'New Quartile']
                        df_det_excel = df_det_excel.drop(columns=[c for c in cols_d if c in df_det_excel.columns])
                        df_det_excel.to_excel(writer, sheet_name='Salary Details', index=False)
                
                # Escribimos hojas de Equity
                if eq_data:
                    eq_summary_df, eq_brk_org, eq_brk_func, df_eq_filtered = eq_data
                    if eq_summary_df is not None: eq_summary_df.to_excel(writer, sheet_name='Equity Summary', index=False)
                    if eq_brk_org is not None: eq_brk_org.to_excel(writer, sheet_name='Equity by Org', index=False)
                    if eq_brk_func is not None: eq_brk_func.to_excel(writer, sheet_name='Equity by Func', index=False)
                    if df_eq_filtered is not None and not df_eq_filtered.empty:
                        df_eq_excel = df_eq_filtered.copy()
                        cols_e = ['Date_in_SG_clean', 'Years_Raw', 'Current Quartile', 'New Quartile']
                        df_eq_excel = df_eq_excel.drop(columns=[c for c in cols_e if c in df_eq_excel.columns])
                        df_eq_excel.to_excel(writer, sheet_name='Equity Details', index=False)
                        
            # ¡CRUCIAL!: Extraer el valor DESPUÉS de que se haya cerrado el bloque 'with writer:'
            excel_data = output.getvalue()
            
            st.sidebar.download_button(
                label="📊 Download Full Excel Report",
                data=excel_data,
                file_name="Dashboard_Final_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
else:
    st.info("Please upload both files (Main and SuperFlex) to begin.")
