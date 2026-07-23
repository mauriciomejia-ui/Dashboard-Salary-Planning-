import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json
import os
import numpy as np

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
        
        # Opciones disponibles
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
        else:
            # Filtro por defecto para Potential si no se carga ninguna configuración
            def_potentials = [x for x in potential_options if x.strip().lower() in ['strategic few', 'high potential']]

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

        # Filtros en el orden solicitado (Potential hasta abajo)
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
        filtros_finales_potentials = selected_potentials if selected_potentials else potential_options

        df_filtered = df[
            (df['Chief Name'].isin(filtros_finales_gerentes)) &
            (df['Reporting Organization'].isin(filtros_finales_orgs)) &
            (df['Function'].isin(filtros_finales_funcs)) &
            (df['Compensation Area'].isin(filtros_finales_comps)) &
            (df['Employee Subgroup'].isin(filtros_finales_subgroups)) &
            (df['Potential'].isin(filtros_finales_potentials))
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

            # --- CHARTS (2x3 GRID ARCHITECTURE) ---
            st.subheader("📊 Graphical Summary")
            
            # FILA 1 DE GRÁFICAS
            col1, col2 = st.columns(2)
            
            # --- CHART 1: Overall Percentage ---
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

            # --- CHART 4: Potential (SOLAMENTE para los que tienen movimiento) ---
            with col4:
                fig4, ax4 = plt.subplots(figsize=(7, 6))
                
                df_pot_mov = df_filtered[tiene_movimiento].copy()
                
                if not df_pot_mov.empty and 'Potential' in df_pot_mov.columns:
                    pot_counts = df_pot_mov['Potential'].value_counts()
                    total_pot = pot_counts.sum()
                    
                    if total_pot > 0:
                        colores_pot = sns.color_palette("Set3", len(pot_counts))
                        
                        wedges4, _ = ax4.pie(pot_counts, startangle=90, colors=colores_pot)
                        leyenda4 = [f"{i} - {v} ({v/total_pot*100:.1f}%)" for i, v in pot_counts.items()]
                        ax4.legend(wedges4, leyenda4, title="Potential Rating", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        
                        ax4.axis('equal')
                        ax4.set_title('Potential Split (Only Staff w/ Adjustments or Promotions)', fontweight='bold', pad=15)
                    else:
                        ax4.text(0.5, 0.5, "No valid data", ha='center', va='center')
                else:
                    ax4.text(0.5, 0.5, "No movements to analyze for Potential", ha='center', va='center')
                
                st.pyplot(fig4)

            # FILA 3 DE GRÁFICAS
            st.markdown("<br>", unsafe_allow_html=True)
            col5, col6 = st.columns(2)

            # --- CHART 5: Distribution vs Columns W and AW ---
            with col5:
                fig5, ax5 = plt.subplots(figsize=(7, 6))
                
                # Índices de columnas
                col_w_idx = 22
                col_aw_idx = 48
                
                if total_personas > 0 and len(df_filtered.columns) > max(col_w_idx, col_aw_idx):
                    col_w_name = df_filtered.columns[col_w_idx]
                    col_aw_name = df_filtered.columns[col_aw_idx]
                    
                    data_w = df_filtered.iloc[:, col_w_idx].astype(str).str.strip()
                    data_aw = df_filtered.iloc[:, col_aw_idx].astype(str).str.strip()
                    
                    counts_w = data_w.value_counts()
                    counts_aw = data_aw.value_counts()
                    
                    orden_deseado = ["1Q", "2Q", "3Q", "4Q", "AboveMax"]
                    categorias_encontradas = set(counts_w.index).union(set(counts_aw.index))
                    categorias_extra = [
                        c for c in categorias_encontradas 
                        if c not in orden_deseado 
                        and c.lower() not in ['nan', 'none', 'null', '', 'below minimum']
                    ]
                    
                    final_order = orden_deseado + categorias_extra
                    
                    val_w = [counts_w.get(c, 0) for c in final_order]
                    val_aw = [counts_aw.get(c, 0) for c in final_order]
                    
                    if sum(val_w) > 0 or sum(val_aw) > 0:
                        x_pos = np.arange(len(final_order))
                        ancho_barra = 0.35
                        
                        bars_w = ax5.bar(x_pos - ancho_barra/2, val_w, ancho_barra, label=str(col_w_name)[:20], color='#ffb347')
                        bars_aw = ax5.bar(x_pos + ancho_barra/2, val_aw, ancho_barra, label=str(col_aw_name)[:20], color='#87cefa')
                        
                        max_y = max(max(val_w), max(val_aw))
                        
                        ax5.set_xticks(x_pos)
                        ax5.set_xticklabels(final_order, rotation=45, ha='right')
                        ax5.legend(loc="upper right", fontsize=9)
                        
                        ax5.set_title('Hipos distribution', fontweight='bold', pad=15)
                        ax5.set_ylabel('Number of Employees')
                        
                        ax5.set_ylim(0, max_y * 1.20)
                        ax5.spines['top'].set_visible(False)
                        ax5.spines['right'].set_visible(False)
                    else:
                        ax5.text(0.5, 0.5, "No data matches these categories", ha='center', va='center')
                else:
                    ax5.text(0.5, 0.5, "Columns W or AW not found or missing data", ha='center', va='center')
                    
                st.pyplot(fig5)

            # --- CHART 6: Gender Split Bar Chart ---
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
                            
                        ax6.set_title('Gender Split (Only Staff w/ Adjustments or Promotions)', fontweight='bold', pad=15)
                        ax6.set_ylabel('Number of Employees')
                        
                        ax6.spines['top'].set_visible(False)
                        ax6.spines['right'].set_visible(False)
                    else:
                        ax6.text(0.5, 0.5, "No valid data", ha='center', va='center')
                else:
                    ax6.text(0.5, 0.5, "No Gender data to analyze", ha='center', va='center')
                
                st.pyplot(fig6)
            
            st.markdown("---")

            # --- DYNAMIC STAFF TABLE (BELOW CHARTS) ---
            st.subheader("👥 Employee Detailed List & Alerts")
            
            # FILTROS DE LA TABLA
            col_filt1, col_filt2 = st.columns(2)
            
            with col_filt1:
                opcion_detalle = st.radio(
                    "1. Filter by Movement Type:",
                    ["All Employees", "Adjustment Only", "Promotion Only", "Both", "No Movement"],
                    horizontal=True
                )
                
            with col_filt2:
                # FILTRO DE ALERTAS (COLORES)
                opcion_alerta = st.selectbox(
                    "2. Filter by Alerts (Colors):",
                    [
                        "Show All", 
                        "⚠️ Show Only with Alerts (Any Color)", 
                        "🔴 Red Alerts Only (Critical)", 
                        "🟠 Orange Alerts Only (Warning)", 
                        "🟡 Yellow Alerts Only (Notice)"
                    ]
                )
            
            # 1. Apply Movement Type Mask
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
                
            df_detalle = df_filtered[mask].copy()
            df_detalle = df_detalle.rename(columns={'Chief Name': 'Manager'})
            
            # --- EVALUACIÓN AUTOMÁTICA DE ALERTAS EN COLUMNAS (J, Z, AB, AC, AF, AH, AI, AK, AL) ---
            FECHA_ACTUAL = pd.to_datetime('2026-07-23')

            def evaluar_alertas(row):
                comentarios = []
                color = ''
                
                try:
                    val_j = pd.to_numeric(row.iloc[9], errors='coerce') if len(row) > 9 else 0
                    val_z = pd.to_datetime(row.iloc[25], errors='coerce') if len(row) > 25 else pd.NaT
                    val_ab = pd.to_numeric(row.iloc[27], errors='coerce') if len(row) > 27 else 0
                    val_ac = pd.to_numeric(row.iloc[28], errors='coerce') if len(row) > 28 else 0
                    val_af = pd.to_numeric(row.iloc[31], errors='coerce') if len(row) > 31 else 0
                    val_ah = str(row.iloc[33]).strip() if len(row) > 33 else ""
                    val_ai = pd.to_numeric(row.iloc[34], errors='coerce') if len(row) > 34 else 0
                    val_ak = pd.to_numeric(row.iloc[36], errors='coerce') if len(row) > 36 else 0
                    
                    val_al_raw = str(row.iloc[37]).strip() if len(row) > 37 else ""
                    al_vacio = val_al_raw == "" or val_al_raw.lower() in ['nan', 'nat', 'none']
                    
                    flag_rojo = False
                    flag_naranja = False
                    flag_amarillo = False
                    
                    # 1. ORANGE: AF < 1 AND (AB > 0 OR AC > 0) AND Z <= 6 meses
                    if pd.notna(val_z):
                        delta_dias = abs((FECHA_ACTUAL - val_z).days)
                        if val_af < 1 and (val_ab > 0 or val_ac > 0) and delta_dias <= 182:
                            flag_naranja = True
                            comentarios.append("Revisar Adjustment vs Fecha reciente (<=6 meses)")
                            
                    # 2. YELLOW: AF > 0 AND AH == "None Selected"
                    if val_af > 0 and val_ah == "None Selected":
                        flag_amarillo = True
                        comentarios.append("Adjustment con 'None Selected'")
                        
                    # 3. YELLOW: (AK > 0 y AK < 6) OR AK > 15
                    if (0 < val_ak < 6) or (val_ak > 15):
                        flag_amarillo = True
                        comentarios.append("Valor AK fuera de rango recomendado")
                        
                    # 4. YELLOW: AK > 0 y Columna AL está vacía
                    if val_ak > 0 and al_vacio:
                        flag_amarillo = True
                        comentarios.append("AK > 0 pero Columna AL está vacía")
                        
                    # 5. RED: AI > J y AK > 0
                    if val_ai > val_j and val_ak > 0:
                        flag_rojo = True
                        comentarios.append("AI supera el valor de J y AK > 0")
                    
                    # Asignar colores por jerarquía
                    if flag_rojo:
                        color = '#ffcccc' # Rojo pastel
                    elif flag_naranja:
                        color = '#ffe4b5' # Naranja pastel
                    elif flag_amarillo:
                        color = '#ffffcc' # Amarillo pastel
                        
                except Exception:
                    pass
                
                return pd.Series([", ".join(comentarios), color])

            # Aplicar motor de alertas
            res_alertas = df_detalle.apply(evaluar_alertas, axis=1)
            df_detalle['Comments'] = res_alertas[0]
            df_detalle['RowColor'] = res_alertas[1]
            
            # 2. Apply Alerts (Color) Filter
            if opcion_alerta == "⚠️ Show Only with Alerts (Any Color)":
                df_detalle = df_detalle[df_detalle['RowColor'] != '']
            elif opcion_alerta == "🔴 Red Alerts Only (Critical)":
                df_detalle = df_detalle[df_detalle['RowColor'] == '#ffcccc']
            elif opcion_alerta == "🟠 Orange Alerts Only (Warning)":
                df_detalle = df_detalle[df_detalle['RowColor'] == '#ffe4b5']
            elif opcion_alerta == "🟡 Yellow Alerts Only (Notice)":
                df_detalle = df_detalle[df_detalle['RowColor'] == '#ffffcc']
            
            # --- RENDERIZAR TABLA COLOREADA ---
            st.write(f"Showing **{len(df_detalle)}** matching employees.")
            
            def aplicar_color_fila(row, colores_serie):
                color_hex = colores_serie.loc[row.name]
                css = f"background-color: {color_hex}" if color_hex else ""
                return [css] * len(row)

            if not df_detalle.empty:
                serie_colores = df_detalle['RowColor']
                df_visual = df_detalle.drop(columns=['RowColor'])
                
                df_estilizado = df_visual.style.apply(aplicar_color_fila, colores_serie=serie_colores, axis=1)
                st.dataframe(df_estilizado, use_container_width=True)
            else:
                st.info("No employees match the selected table filters.")
                
    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")
else:
    st.info("Please upload both files (Main and SuperFlex) to begin.")
