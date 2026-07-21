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
                
                # Recargar a prueba de errores de versión
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
            st.dataframe(tabla_personal, use_container_width=True)
            
            st.markdown("---")

            # --- GRÁFICAS ---
            st.subheader("📊 Resumen Gráfico")
            
            col1, col2, col3 = st.columns(3)
            
            # --- GRÁFICA 1: Porcentaje de Movimientos ---
            with col1:
                # OJO: Cambié '% adjustment' a '%Adjustment' (sin espacio) que es como viene en tu Excel
                adj_col = pd.to_numeric(df_filtered.get('%Adjustment', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
                growth_col = pd.to_numeric(df_filtered.get('%Growth Promotion', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
                
                # Condición: Ajuste > 0 O Crecimiento > 0
                movimientos_mask = (adj_col > 0) | (growth_col > 0)
                num_movimientos = movimientos_mask.sum()
                total_personas = len(df_filtered)
                
                pct_mov = (num_movimientos / total_personas) * 100 if total_personas > 0 else 0
                
                fig1, ax1 = plt.subplots(figsize=(5, 4))
                if total_personas > 0:
                    labels = ['Con Movimiento', 'Sin Movimiento']
                    sizes = [num_movimientos, total_personas - num_movimientos]
                    colors = ['#ff9999', '#66b3ff']
                    
                    # Evitar graficar pastel si no hay movimientos para que no marque error visual
                    if num_movimientos == 0:
                        ax1.pie([100], labels=['Sin Movimiento'], colors=['#66b3ff'], startangle=90)
                    else:
                        ax1.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors)
                    
                    ax1.axis('equal') 
                    ax1.set_title('% de Empleados con Ajuste/Promo', fontweight='bold')
                else:
                    ax1.text(0.5, 0.5, "Sin datos", ha='center', va='center')
                    
                st.pyplot(fig1)
                
                if pct_mov > 10:
                    st.warning(f"⚠️ Alerta: El **{pct_mov:.1f}%** del personal seleccionado tiene movimientos (supera el límite del 10%).")
                elif pct_mov > 0:
                    st.info(f"✅ El {pct_mov:.1f}% del personal tiene movimientos (dentro de los límites).")

            # --- GRÁFICA 2: Top 10 Reporting Org ---
            with col2:
                fig2, ax2 = plt.subplots(figsize=(5, 4))
                sns.countplot(y='Reporting Organization', data=df_filtered, order=df_filtered['Reporting Organization'].value_counts().index[:10], palette='viridis', ax=ax2)
                ax2.set_xlabel('Cantidad')
                ax2.set_ylabel('')
                ax2.set_title('Top 10 Reporting Org')
                st.pyplot(fig2)

            # --- GRÁFICA 3: Top 10 Functions ---
            with col3:
                fig3, ax3 = plt.subplots(figsize=(5, 4))
                sns.countplot(y='Function', data=df_filtered, order=df_filtered['Function'].value_counts().index[:10], palette='viridis', ax=ax3)
                ax3.set_xlabel('Cantidad')
                ax3.set_ylabel('')
                ax3.set_title('Top 10 Functions')
                st.pyplot(fig3)
                
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos (el Principal y el SuperFlex) para comenzar.")
