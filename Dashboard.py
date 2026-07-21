import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json
import os

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
        
        df1['Global ID'] = pd.to_numeric(df1['Global ID'], errors='coerce')
        df2['Global ID'] = pd.to_numeric(df2['Global ID'], errors='coerce')
        
        df = pd.merge(df1, df2[['Global ID', 'Chief Name']], on='Global ID', how='left')
        df['Chief Name'] = df['Chief Name'].fillna("No Manager Assigned")
        
        # --- SIDEBAR: LOAD FILTERS ---
        st.sidebar.header("💾 My Saved Filters")
        nombres_disponibles = ["-- None --"] + list(st.session_state['memoria_filtros'].keys())
        filtro_elegido = st.sidebar.selectbox("Load a configuration:", nombres_disponibles)
        
        def_gerentes, def_orgs, def_funcs, def_comps = [], [], [], []
        if filtro_elegido != "-- None --":
            config = st.session_state['memoria_filtros'][filtro_elegido]
            def_gerentes = config.get("gerentes", [])
            def_orgs = config.get("orgs", [])
            def_funcs = config.get("funcs", [])
            def_comps = config.get("comps", [])

        st.sidebar.markdown("---")
        
        # --- SIDEBAR: DATA SELECTION ---
        st.sidebar.header("🔍 Data Filters")
        st.sidebar.info("Leave blank to include all.")
        
        gerente_options = sorted(df['Chief Name'].astype(str).unique().tolist())
        org_options = sorted(df['Reporting Organization'].dropna().unique().tolist())
        func_options = sorted(df['Function'].dropna().unique().tolist())
        comp_options = sorted(df['Compensation Area'].dropna().unique().tolist())

        def_gerentes = [x for x in def_gerentes if x in gerente_options]
        def_orgs = [x for x in def_orgs if x in org_options]
        def_funcs = [x for x in def_funcs if x in func_options]
        def_comps = [x for x in def_comps if x in comp_options]

        selected_gerentes = st.sidebar.multiselect("Manager(s):", gerente_options, default=def_gerentes)
        selected_orgs = st.sidebar.multiselect("Reporting Organization:", org_options, default=def_orgs)
        selected_funcs = st.sidebar.multiselect("Function:", func_options, default=def_funcs)
        selected_comps = st.sidebar.multiselect("Compensation Area:", comp_options, default=def_comps)
        
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
                    "comps": selected_comps
                }
                with open(ARCHIVO_FILTROS, 'w', encoding='utf-8') as f:
                    json.dump(st.session_state['memoria_filtros'], f)
                st.sidebar.success(f"Filter '{nuevo_nombre}' saved successfully!")
                
                if hasattr(st, "rerun"):
                    st.rerun()
                elif hasattr(st, "experimental_rerun"):
                    st.experimental_rerun()
            else:
                st.sidebar.warning("Please enter a title before saving.")

        # --- APPLY FILTERS ---
        filtros_finales_gerentes = selected_gerentes if selected_gerentes else gerente_options
        filtros_finales_orgs = selected_orgs if selected_orgs else org_options
        filtros_finales_funcs = selected_funcs if selected_funcs else func_options
        filtros_finales_comps = selected_comps if selected_comps else comp_options

        df_filtered = df[
            (df['Chief Name'].isin(filtros_finales_gerentes)) &
            (df['Reporting Organization'].isin(filtros_finales_orgs)) &
            (df['Function'].isin(filtros_finales_funcs)) &
            (df['Compensation Area'].isin(filtros_finales_comps))
        ]
        
        st.success(f"Showing {len(df_filtered)} records matching your search.")
        
        if df_filtered.empty:
            st.warning("No data matches these filters.")
        else:
            # --- STAFF TABLE ---
            st.subheader("👥 Employee List")
            tabla_personal = df_filtered[['Name', 'Global ID', 'Chief Name', 'Position', 'Reporting Organization', 'Function', 'Compensation Area']]
            tabla_personal = tabla_personal.rename(columns={'Chief Name': 'Manager'})
            st.dataframe(tabla_personal, width=700) 
            
            st.markdown("---")

            # --- PREPARE DATA FOR PIE CHARTS ---
            adj_col = pd.to_numeric(df_filtered.get('%Adjustment', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
            growth_col = pd.to_numeric(df_filtered.get('%Growth Promotion', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
            
            cond_adj = adj_col > 0
            cond_promo = growth_col > 0
            
            solo_adj = (cond_adj & ~cond_promo).sum()
            solo_promo = (~cond_adj & cond_promo).sum()
            ambos = (cond_adj & cond_promo).sum()
            sin_mov = (~cond_adj & ~cond_promo).sum()
            
            num_movimientos = solo_adj + solo_promo + ambos
            total_personas = len(df_filtered)

            # Allow all percentages to display inside the pie chart
            def my_autopct(pct):
                return f'{pct:.1f}%' if pct > 0 else ''

            # --- CHARTS ---
            st.subheader("📊 Graphical Summary")
            col1, col2, col3 = st.columns(3)
            
            # --- CHART 1: Overall Percentage ---
            with col1:
                pct_mov = (num_movimientos / total_personas) * 100 if total_personas > 0 else 0
                fig1, ax1 = plt.subplots(figsize=(4, 4))
                if total_personas > 0:
                    if num_movimientos == 0:
                        ax1.pie([100], labels=['No Movement'], colors=['#d3d3d3'], startangle=90)
                    else:
                        sizes = [num_movimientos, total_personas - num_movimientos]
                        labels = ['With Movement', 'No Movement']
                        colors = ['#ff9999', '#d3d3d3']
                        ax1.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors)
                    
                    ax1.axis('equal') 
                    ax1.set_title('Overall %', fontweight='bold', pad=15)
                else:
                    ax1.text(0.5, 0.5, "No data", ha='center', va='center')
                    
                st.pyplot(fig1)
                if pct_mov > 10:
                    st.warning(f"⚠️ Alert: **{pct_mov:.1f}%** of the selected staff have movements (> 10%).")
                elif pct_mov > 0:
                    st.info(f"✅ {pct_mov:.1f}% of the selected staff have movements.")

            # --- CHART 2: Breakdown %Adjustment vs %GPromotion ---
            with col2:
                fig2, ax2 = plt.subplots(figsize=(5, 4))
                
                raw_sizes2 = [solo_adj, solo_promo, ambos, sin_mov]
                raw_labels2 = ['Adjustment Only', 'Promotion Only', 'Both', 'No Movement']
                raw_colors2 = ['#ffb3e6', '#c2c2f0', '#ff6666', '#c2f0c2']
                
                sizes2 = [s for s in raw_sizes2 if s > 0]
                labels2 = [l for s, l in zip(raw_sizes2, raw_labels2) if s > 0]
                colors2 = [c for s, c in zip(raw_sizes2, raw_colors2) if s > 0]
                
                if len(sizes2) > 0:
                    wedges, texts, autotexts = ax2.pie(sizes2, autopct=my_autopct, startangle=90, colors=colors2)
                    
                    leyenda2 = [f"{l} ({s})" for l, s in zip(labels2, sizes2)]
                    ax2.legend(wedges, leyenda2, title="Breakdown", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                    
                    ax2.axis('equal')
                    ax2.set_title('Movement Details', fontweight='bold', pad=15)
                else:
                    ax2.text(0.5, 0.5, "No data", ha='center', va='center')
                
                st.pyplot(fig2)

            # --- CHART 3: Adjustment Reason ---
            with col3:
                fig3, ax3 = plt.subplots(figsize=(5, 4))
                
                df_adj = df_filtered[cond_adj].copy()
                
                if not df_adj.empty and 'Adjustment Reason' in df_adj.columns:
                    df_adj['Adjustment Reason'] = df_adj['Adjustment Reason'].fillna('No Reason Assigned')
                    df_adj['Adjustment Reason'] = df_adj['Adjustment Reason'].replace({'None Selected': 'No Reason Assigned'})
                    
                    reason_counts = df_adj['Adjustment Reason'].value_counts()
                    
                    if not reason_counts.empty:
                        colores_motivos = sns.color_palette("pastel", len(reason_counts))
                        
                        wedges3, texts3, autotexts3 = ax3.pie(reason_counts, autopct=my_autopct, 
                                                              startangle=90, colors=colores_motivos)
                        
                        leyenda3 = [f"{i} ({v})" for i, v in reason_counts.items()]
                        ax3.legend(wedges3, leyenda3, title="Reasons", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        
                        ax3.axis('equal')
                        ax3.set_title('Adjustment Split (Reason)', fontweight='bold', pad=15)
                    else:
                        ax3.text(0.5, 0.5, "No valid data", ha='center', va='center')
                else:
                    ax3.text(0.5, 0.5, "No adjustments to analyze", ha='center', va='center')
                
                st.pyplot(fig3)
                
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
else:
    st.info("Please upload both files (Main and SuperFlex) to begin.")
