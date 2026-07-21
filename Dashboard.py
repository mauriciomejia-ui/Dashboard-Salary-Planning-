import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json
import os

# Configuración de la página
st.set_page_config(page_title="Colgate HR Data Dashboard", layout="wide")
st.title("Dashboard de Recursos Humanos")
st.write("Sube tus archivos de Excel para ver y filtrar las métricas.")

# --- MEMORIA DE FILTROS ---
ARCHIVO_FILTROS = 'mis_filtros_guardados.json'

def cargar_memoria():
    if os.path.exists(ARCHIVO_FILTROS):
        with open(ARCHIVO_FILTROS, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

if 'memoria_filtros' not in st.session_state:
    st.session_state['memoria_filtros'] = cargar_memoria()

# --- CARGADOR DE ARCHIVOS ---
col_up1, col_up2 = st.columns(2)
with col_up1:
    file1 = st.file_uploader("1. Sube el archivo Principal (Ej. colgate_...xlsx)", type=["xlsx", "xls"])
with col_up2:
    file2 = st.file_uploader("2. Sube el archivo SuperFlex (SuperFlex-DefaultView...)", type=["xlsx", "xls"])

if file1 is not None and file2 is not None:
    try:
        df1 = pd.read_excel(file1, sheet_name="Salary")
        df2 = pd.read_excel(file2, header=2)
        
        df1['Global ID'] = pd.to_numeric(df1['Global ID'], errors='coerce')
        df2['Global ID'] = pd.to_numeric(df2['Global ID'], errors='coerce')
        
        df = pd.merge(df1, df2[['Global ID', 'Chief Name']], on='Global ID', how='left')
        df['Chief Name'] = df['Chief Name'].fillna("Sin Gerente Asignado")
        
        # --- MENÚ LATERAL: CARGAR FILTROS ---
        st.sidebar.header("💾 Mis Filtros Guardados")
        nombres_disponibles = ["-- Ninguno --"] + list(st.session_state['memoria_filtros'].keys())
        filtro_elegido = st.sidebar.selectbox("Cargar una configuración:", nombres_disponibles)
        
        def_gerentes, def_orgs, def_funcs, def_comps = [], [], [], []
        if filtro_elegido != "-- Ninguno --":
            config = st.session_state['memoria_filtros'][filtro_elegido]
            def_gerentes = config.get("gerentes", [])
            def_orgs = config.get("orgs", [])
            def_funcs = config.get("funcs", [])
            def_comps = config.get("comps", [])

        st.sidebar.markdown("---")
        
        # --- MENÚ LATERAL: SELECCIÓN DE DATOS ---
        st.sidebar.header("🔍 Filtros de Datos")
        st.sidebar.info("Deja el espacio en blanco para incluir todo.")
        
        gerente_options = sorted(df['Chief Name'].astype(str).unique().tolist())
        org_options = sorted(df['Reporting Organization'].dropna().unique().tolist())
        func_options = sorted(df['Function'].dropna().unique().tolist())
        comp_options = sorted(df['Compensation Area'].dropna().unique().tolist())

        def_gerentes = [x for x in def_gerentes if x in gerente_options]
        def_orgs = [x for x in def_orgs if x in org_options]
        def_funcs = [x for x in def_funcs if x in func_options]
        def_comps = [x for x in def_comps if x in comp_options]

        selected_gerentes = st.sidebar.multiselect("Gerente(s):", gerente_options, default=def_gerentes)
        selected_orgs = st.sidebar.multiselect("Reporting Organization:", org_options, default=def_orgs)
        selected_funcs = st.sidebar.multiselect("Function:", func_options, default=def_funcs)
        selected_comps = st.sidebar.multiselect("Compensation Area:", comp_options, default=def_comps)
        
        # --- MENÚ LATERAL: GUARDAR NUEVO FILTRO ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("✏️ Guardar combinación actual")
        nuevo_nombre = st.sidebar.text_input("Dale un título a este filtro:")
        
        if st.sidebar.button("Guardar Filtro"):
            if nuevo_nombre:
                st.session_state['memoria_filtros'][nuevo_nombre] = {
                    "gerentes": selected_gerentes,
                    "orgs": selected_orgs,
                    "funcs": selected_funcs,
                    "comps": selected_comps
                }
                with open(ARCHIVO_FILTROS, 'w', encoding='utf-8') as f:
                    json.dump(st.session_state['memoria_filtros'], f)
                st.sidebar.success(f"¡Filtro '{nuevo_nombre}' guardado!")
                
                if hasattr(st, "rerun"):
                    st.rerun()
                elif hasattr(st, "experimental_rerun"):
                    st.experimental_rerun()
            else:
                st.sidebar.warning("Escribe un título antes de guardar.")

        # --- APLICAR FILTROS ---
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
        
        st.success(f"Mostrando {len(df_filtered)} registros correspondientes a tu búsqueda.")
        
        if df_filtered.empty:
            st.warning("No hay datos que coincidan con estos filtros.")
        else:
            # --- TABLA DE PERSONAL ---
            st.subheader("👥 Lista de Personal")
            tabla_personal = df_filtered[['Name', 'Global ID', 'Chief Name', 'Position', 'Reporting Organization', 'Function', 'Compensation Area']]
            tabla_personal = tabla_personal.rename(columns={'Chief Name': 'Gerente'})
            st.dataframe(tabla_personal, width=700) 
            
            st.markdown("---")

            # --- PREPARAR DATOS PARA GRÁFICAS DE PASTEL ---
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

            # Función para evitar que porcentajes pequeños se encimen
            def my_autopct(pct):
                return f'{pct:.1f}%' if pct > 3 else ''

            # --- GRÁFICAS ---
            st.subheader("📊 Resumen Gráfico")
            col1, col2, col3 = st.columns(3)
            
            # --- GRÁFICA 1: Porcentaje General ---
            with col1:
                pct_mov = (num_movimientos / total_personas) * 100 if total_personas > 0 else 0
                fig1, ax1 = plt.subplots(figsize=(4, 4))
                if total_personas > 0:
                    if num_movimientos == 0:
                        ax1.pie([100], labels=['Sin Movimiento'], colors=['#d3d3d3'], startangle=90)
                    else:
                        sizes = [num_movimientos, total_personas - num_movimientos]
                        labels = ['Con Movimiento', 'Sin Movimiento']
                        colors = ['#ff9999', '#d3d3d3']
                        ax1.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors)
                    
                    ax1.axis('equal') 
                    ax1.set_title('% General', fontweight='bold', pad=15)
                else:
                    ax1.text(0.5, 0.5, "Sin datos", ha='center', va='center')
                    
                st.pyplot(fig1)
                if pct_mov > 10:
                    st.warning(f"⚠️ Alerta: **{pct_mov:.1f}%** del personal supera el 10%.")
                elif pct_mov > 0:
                    st.info(f"✅ {pct_mov:.1f}% tiene movimientos.")

            # --- GRÁFICA 2: Desglose %Adjustment vs %GPromotion ---
            with col2:
                fig2, ax2 = plt.subplots(figsize=(5, 4))
                
                raw_sizes2 = [solo_adj, solo_promo, ambos, sin_mov]
                raw_labels2 = ['Solo Ajuste', 'Solo Promoción', 'Ambos', 'Sin Movimiento']
                raw_colors2 = ['#ffb3e6', '#c2c2f0', '#ff6666', '#c2f0c2']
                
                # Filtramos para que no dibuje ceros
                sizes2 = [s for s in raw_sizes2 if s > 0]
                labels2 = [l for s, l in zip(raw_sizes2, raw_labels2) if s > 0]
                colors2 = [c for s, c in zip(raw_sizes2, raw_colors2) if s > 0]
                
                if len(sizes2) > 0:
                    wedges, texts, autotexts = ax2.pie(sizes2, autopct=my_autopct, startangle=90, colors=colors2)
                    
                    # Agregamos la leyenda externa para evitar texto encimado
                    leyenda2 = [f"{l} ({s})" for l, s in zip(labels2, sizes2)]
                    ax2.legend(wedges, leyenda2, title="Desglose", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                    
                    ax2.axis('equal')
                    ax2.set_title('Detalle de Movimientos', fontweight='bold', pad=15)
                else:
                    ax2.text(0.5, 0.5, "Sin datos", ha='center', va='center')
                
                st.pyplot(fig2)

            # --- GRÁFICA 3: Motivos de Ajuste (Adjustment Reason) ---
            with col3:
                fig3, ax3 = plt.subplots(figsize=(5, 4))
                
                # Tomamos SOLO a las personas que realmente tuvieron un ajuste (>0)
                df_adj = df_filtered[cond_adj].copy()
                
                if not df_adj.empty and 'Adjustment Reason' in df_adj.columns:
                    # Limpiamos los vacíos y "None Selected" para que no se pierdan en la suma
                    df_adj['Adjustment Reason'] = df_adj['Adjustment Reason'].fillna('Sin Razón Asignada')
                    df_adj['Adjustment Reason'] = df_adj['Adjustment Reason'].replace({'None Selected': 'Sin Razón Asignada'})
                    
                    reason_counts = df_adj['Adjustment Reason'].value_counts()
                    
                    if not reason_counts.empty:
                        colores_motivos = sns.color_palette("pastel", len(reason_counts))
                        
                        wedges3, texts3, autotexts3 = ax3.pie(reason_counts, autopct=my_autopct, 
                                                              startangle=90, colors=colores_motivos)
                        
                        # Creamos la leyenda externa con el motivo y la cantidad exacta de casos
                        leyenda3 = [f"{i} ({v})" for i, v in reason_counts.items()]
                        ax3.legend(wedges3, leyenda3, title="Motivos", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                        
                        ax3.axis('equal')
                        ax3.set_title('Split de Ajustes (Reason)', fontweight='bold', pad=15)
                    else:
                        ax3.text(0.5, 0.5, "Sin datos válidos", ha='center', va='center')
                else:
                    ax3.text(0.5, 0.5, "Sin ajustes para analizar", ha='center', va='center')
                
                st.pyplot(fig3)
                
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos (el Principal y el SuperFlex) para comenzar.")
