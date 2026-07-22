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
        
        # Format Global ID to merge correctly
        df1['Global ID'] = pd.to_numeric(df1['Global ID'], errors='coerce')
        df2['Global ID'] = pd.to_numeric(df2['Global ID'], errors='coerce')
        
        # Extraer columna AC (índice 28) del SuperFlex como Employee Subgroup
        if len(df2.columns) > 28:
            df2['Employee Subgroup'] = df2.iloc[:, 28].astype(str)
        else:
            df2['Employee Subgroup'] = "N/A"
        
        # Merge Main File with SuperFlex data
        df = pd.merge(df1, df2[['Global ID', 'Chief Name', 'Employee Subgroup']], on='Global ID', how='left')
        df['Chief Name'] = df['Chief Name'].fillna("No Manager Assigned")
        df['Employee Subgroup'] = df['Employee Subgroup'].replace('nan', 'Unknown').fillna('Unknown')
        
        # --- SIDEBAR: LOAD FILTERS ---
        st.sidebar.header("💾 My Saved Filters")
        nombres_disponibles = ["-- None --"] + list(st.session_state['memoria_filtros'].keys())
        filtro_elegido = st.sidebar.selectbox("Load a configuration:", nombres_disponibles)
        
        def_gerentes, def_orgs, def_funcs, def_comps, def_subgroups = [], [], [], [], []
        if filtro_elegido != "-- None --":
            config = st.session_state['memoria_filtros'][filtro_elegido]
            def_gerentes = config.get("gerentes", [])
            def_orgs = config.get("orgs", [])
            def_funcs = config.get("funcs", [])
            def_comps = config.get("comps", [])
            def_subgroups = config.get("subgroups", [])

        st.sidebar.markdown("---")
        
        # --- SIDEBAR: DATA SELECTION ---
        st.sidebar.header("🔍 Data Filters")
        st.sidebar.info("Leave blank to include all.")
        
        gerente_options = sorted(df['Chief Name'].astype(str).unique().tolist())
        org_options = sorted(df['Reporting Organization'].dropna().unique().tolist())
        func_options = sorted(df['Function'].dropna().unique().tolist())
        comp_options = sorted(df['Compensation Area'].dropna().unique().tolist())
        subgroup_options = sorted(df['Employee Subgroup'].unique().tolist())

        def_gerentes = [x for x in def_gerentes if x in gerente_options]
        def_orgs = [x for x in def_orgs if x in org_options]
        def_funcs = [x for x in def_funcs if x in func_options]
        def_comps = [x for x in def_comps if x in comp_options]
        def_subgroups = [x for x in def_subgroups if x in subgroup_options]

        selected_gerentes = st.sidebar.multiselect("Manager(s):", gerente_options, default=def_gerentes)
        selected_orgs = st.sidebar.multiselect("Reporting Organization:", org_options, default=def_orgs)
        selected_funcs = st.sidebar.multiselect("Function:", func_options, default=def_funcs)
        selected_comps = st.sidebar.multiselect("Compensation Area:", comp_options, default=def_comps)
        selected_subgroups = st.sidebar.multiselect("Employee Subgroup:", subgroup_options, default=def_subgroups)
        
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
                    "subgroups": selected_subgroups
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
        filtros_finales_subgroups = selected_subgroups if selected_subgroups else subgroup_options

        df_filtered = df[
            (df['Chief Name'].isin(filtros_finales_gerentes)) &
            (df['Reporting Organization'].isin(filtros_finales_orgs)) &
            (df['Function'].isin(filtros_finales_funcs)) &
            (df['Compensation Area'].isin(filtros_finales_comps)) &
            (df['Employee Subgroup'].isin(filtros_finales_subgroups))
        ]
        
        st.success(f"Showing {len(df_filtered)} records matching your sidebar search.")
        
        if df_filtered.empty:
            st.warning("No data matches these filters.")
        else:
            # --- COST SUMMARY TABLE ---
            st.subheader("💰 Cost Summary")
            
            # Obtenemos las variables base
            adj_pct = pd.to_numeric(df_filtered.get('%Adjustment', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
            promo_pct = pd.to_numeric(df_filtered.get('%Growth Promotion', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
            
            col_t_annual_usd = pd.to_numeric(df_filtered.get('$ Annual Salary(in USD)', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
            col_au_new_annual_usd = pd.to_numeric(df_filtered.get('$ New Annual Salary(in USD)', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)

            # Cálculos de Costos
            cost_adj = ((adj_pct / 100) * col_t_annual_usd)[adj_pct > 0].sum()
            cost_promo = ((promo_pct / 100) * col_t_annual_usd)[promo_pct > 0].sum()
            total_cost = cost_adj + cost_promo

            sum_t = col_t_annual_usd.sum()
            sum_au = col_au_new_annual_usd.sum()
            
            pct_adj_vs_total = (cost_adj / sum_t) * 100 if sum_t > 0 else 0
            pct_promo_vs_total = (cost_promo / sum_t) * 100 if sum_t > 0 else 0
            pct_total_cost_vs_total = (total_cost / sum_t) * 100 if sum_t > 0 else 0
            
            pct_incremento = ((sum_au / sum_t) - 1) * 100 if sum_t > 0 else 0

            cost_df = pd.DataFrame({
                "Concept": ["Adjustment Cost", "Growth Promotion Cost", "Total Cost", "Total % Increment"],
                "Value": [
                    f"${cost_adj:,.2f}",
                    f"${cost_promo:,.2f}",
                    f"${total_cost:,.2f}",
                    f"{pct_incremento:,.2f}%"
                ],
                "% of Total Salary": [
                    f"{pct_adj_vs_total:,.2f}%",
                    f"{pct_promo_vs_total:,.2f}%",
                    f"{pct_total_cost_vs_total:,.2f}%",
                    "-"
                ]
            })
            
            st.table(cost_df)
            st.markdown("---")

            # --- PREPARE DATA FOR PIE CHARTS ---
            cond_adj = adj_pct > 0
            cond_promo = promo_pct > 0
            
            solo_adj = (cond_adj & ~cond_promo)
            solo_promo = (~cond_adj & cond_promo)
            ambos = (cond_adj & cond_promo)
            sin_mov = (~cond_adj & ~cond_promo)
            
            num_solo_adj = solo_adj.sum()
            num_solo_promo = solo_promo.sum()
            num_ambos = ambos.sum()
            num_sin_mov = sin_mov.sum()
            
            num_movimientos = num_solo_adj + num_solo_promo + num_ambos
            total_personas = len(df_filtered)

            # --- CHARTS (2x2 GRID FOR LARGER DISPLAY) ---
            st.subheader("📊 Graphical Summary")
            
            # FILA 1 DE GRÁFICAS
            col1, col2 = st.columns(2)
            
            # --- CHART 1: Overall Percentage ---
            with col1:
                pct_mov = (num_movimientos / total_personas) * 100 if total_personas > 0 else 0
                fig1, ax1 = plt.subplots(figsize=(7, 6))
                if total_personas > 0:
                    if num_movimientos == 0:
                        ax1.pie([100], colors=['#d3d3d3'], startangle=90)
                        ax1.legend(["No Movement (100%)"], loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                    else:
                        sizes = [num_movimientos, total_personas - num_movimientos]
                        labels = ['With Movement', 'No Movement']
                        colors = ['#ff9999', '#d3d3d3']
                        
                        wedges, _ = ax1.pie(sizes, startangle=90, colors=colors)
                        leyenda1 = [f"{l} - {s} ({s/total_personas*100:.1f}%)" for l, s in zip(labels, sizes)]
                        ax1.legend(wedges, leyenda1, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                    
                    ax1.axis('equal') 
                    ax1.set_title('Overall Movement', fontweight='bold', pad=15)
                else:
                    ax1.text(0.5, 0.5, "No data", ha='center', va='center')
                    
                st.pyplot(fig1)

            # --- CHART 2: Breakdown %Adjustment vs %GPromotion ---
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
                    wedges2, _ = ax2.pie(sizes2, startangle=90, colors=colors2)
                    leyenda2 = [f"{l} - {s} ({s/total_chart2*100:.1f}%)" for l, s in zip(labels2, sizes2)]
                    ax2.legend(wedges2, leyenda2, title="Breakdown", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                    
                    ax2.axis('equal')
                    ax2.set_title('Movement Details', fontweight='bold', pad=15)
                else:
                    ax2.text(0.5, 0.5, "No data", ha='center', va='center')
                
                st.pyplot(fig2)

            # FILA 2 DE GRÁFICAS
            st.markdown("<br>", unsafe_allow_html=True)
            col3, col4 = st.columns(2)

            # --- CHART 3: Adjustment Reason ---
            with col3:
                fig3, ax3 = plt.subplots(figsize=(7, 6))
                
                df_adj = df_filtered[cond_adj].copy()
                
                if not df_adj.empty and 'Adjustment Reason' in df_adj.columns:
                    df_adj['Adjustment Reason'] = df_adj['Adjustment Reason'].fillna('No Reason Assigned')
                    df_adj['Adjustment Reason'] = df_adj['Adjustment Reason'].replace({'None Selected': 'No Reason Assigned'})
                    
                    reason_counts = df_adj['Adjustment Reason'].value_counts()
                    total_reasons = reason_counts.sum()
                    
                    if total_reasons > 0:
                        colores_motivos = sns.color_palette("pastel", len(reason_counts))
                        
                        wedges3, _ = ax3.pie(reason_counts, startangle=90, colors=colores_motivos)
                        leyenda3 = [f"{i} - {v} ({v/total_reasons*100:.1f}%)" for i, v in reason_counts.items()]
                        ax3.legend(wedges3, leyenda3, title="Reasons", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        
                        ax3.axis('equal')
                        ax3.set_title('Adjustment Split (Reason)', fontweight='bold', pad=15)
                    else:
                        ax3.text(0.5, 0.5, "No valid data", ha='center', va='center')
                else:
                    ax3.text(0.5, 0.5, "No adjustments to analyze", ha='center', va='center')
                
                st.pyplot(fig3)

            # --- CHART 4: Potential (Columna Y) SOLAMENTE para los que tienen movimiento ---
            with col4:
                fig4, ax4 = plt.subplots(figsize=(7, 6))
                
                # Filtramos para analizar solo a aquellos que tienen Adjustment o Promo (>0)
                tiene_movimiento = cond_adj | cond_promo
                df_pot_mov = df_filtered[tiene_movimiento].copy()
                
                if not df_pot_mov.empty and 'Potential' in df_pot_mov.columns:
                    # Limpiamos los vacíos
                    df_pot = df_pot_mov['Potential'].fillna('Not Assigned').astype(str)
                    df_pot = df_pot.replace({'nan': 'Not Assigned', 'None Selected': 'Not Assigned'})
                    
                    pot_counts = df_pot.value_counts()
                    total_pot = pot_counts.sum()
                    
                    if total_pot > 0:
                        colores_pot = sns.color_palette("Set3", len(pot_counts))
                        
                        wedges4, _ = ax4.pie(pot_counts, startangle=90, colors=colores_pot)
                        leyenda4 = [f"{i} - {v} ({v/total_pot*100:.1f}%)" for i, v in pot_counts.items()]
                        ax4.legend(wedges4, leyenda4, title="Potential Rating", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        
                        ax4.axis('equal')
                        # Título aclaratorio para que el usuario entienda qué está viendo
                        ax4.set_title('Potential Split (Only Staff w/ Adjustments or Promotions)', fontweight='bold', pad=15)
                    else:
                        ax4.text(0.5, 0.5, "No valid data", ha='center', va='center')
                else:
                    ax4.text(0.5, 0.5, "No movements to analyze for Potential", ha='center', va='center')
                
                st.pyplot(fig4)

            st.markdown("---")

            # --- DYNAMIC STAFF TABLE (BELOW CHARTS) ---
            st.subheader("👥 Employee Detailed List")
            st.info("💡 Tip: Select a movement category below to see the full information of those specific employees.")
            
            # Selector for the detail table
            opcion_detalle = st.radio(
                "Filter list by movement type (Chart 2 categories):",
                ["All Employees", "Adjustment Only", "Promotion Only", "Both", "No Movement"],
                horizontal=True
            )
            
            # Apply corresponding mask
            if opcion_detalle == "Adjustment Only":
                mask = solo_adj
            elif opcion_detalle == "Promotion Only":
                mask = solo_promo
            elif opcion_detalle == "Both":
                mask = ambos
            elif opcion_detalle == "No Movement":
                mask = sin_mov
            else:
                mask = pd.Series(True, index=df_filtered.index)
                
            df_detalle = df_filtered[mask]
            
            # Renombrar para mayor limpieza
            df_detalle = df_detalle.rename(columns={'Chief Name': 'Manager'})
            
            st.write(f"Showing **{len(df_detalle)}** employees for: **{opcion_detalle}**")
            st.dataframe(df_detalle, use_container_width=True)
                
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
else:
    st.info("Please upload both files (Main and SuperFlex) to begin.")
