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
            # Se usa width='stretch' para evitar advertencias de Streamlit
            st.dataframe(tabla_personal, width=700) 
            
            st.markdown("---")

            # --- PREPARAR DATOS PARA GRÁFICAS DE PASTEL ---
            adj_col = pd.to_numeric(df_filtered.get('%Adjustment', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
            growth_col = pd.to_numeric(df_filtered.get('%Growth Promotion', pd.Series(0, index=df_filtered.index)), errors='coerce').fillna(0)
            
            cond_adj = adj_col > 0
            cond_promo = growth_col > 0
            
            # Cantidades exactas para el desglose
            solo_adj = (cond_adj & ~cond_promo).sum()
            solo_promo = (~cond_adj & cond_promo).sum()
            ambos = (cond_adj & cond_promo).sum()
            sin_mov = (~cond_adj & ~cond_promo).sum()
            
            num_movimientos = solo_adj + solo_promo + ambos
            total_personas = len(df_filtered)

            # --- GRÁFICAS ---
            st.subheader("📊 Resumen Gráfico")
            
            col1, col2, col3 = st.columns(3)
            
            # --- GRÁFICA 1: Porcentaje General de Movimientos ---
            with col1:
                pct_mov = (num_movimientos / total_personas) * 100 if total_personas > 0 else 0
                
                fig1, ax1 = plt.subplots(figsize=(5, 4))
                if total_personas > 0:
                    if num_movimientos == 0:
                        ax1.pie([100], labels=['Sin Movimiento'], colors=['#66b3ff'], startangle=90)
                    else:
                        sizes = [num_movimientos, total_personas - num_movimientos]
                        labels = ['Con Movimiento', 'Sin Movimiento']
                        colors = ['#ff9999', '#66b3ff']
                        ax1.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors)
                    
                    ax1.axis('equal') 
                    ax1.set_title('% General de Movimientos', fontweight='bold')
                else:
                    ax1.text(0.5, 0.5, "Sin datos", ha='center', va='center')
                    
                st.pyplot(fig1)
                
                # Alerta condicional
                if pct_mov > 10:
                    st.warning(f"⚠️ Alerta: El **{pct_mov:.1f}%** del personal seleccionado tiene movimientos (supera el 10%).")
                elif pct_mov > 0:
                    st.info(f"✅ El {pct_mov:.1f}% del personal tiene movimientos (dentro del límite).")

            # --- GRÁFICA 2: Desglose del Tipo de Movimiento ---
            with col2:
                fig2, ax2 = plt.subplots(figsize=(5, 4))
                
                sizes2 = []
                labels2 = []
                colors2 = []
                
                if solo_adj > 0:
                    sizes2.append(solo_adj)
                    labels2.append('Solo Ajuste')
                    colors2.append('#ffb3e6')
                if solo_promo > 0:
                    sizes2.append(solo_promo)
                    labels2.append('Solo Promoción')
                    colors2.append('#c2c2f0')
                if ambos > 0:
                    sizes2.append(ambos)
                    labels2.append('Ambos')
                    colors2.append('#ff6666')
                if sin_mov > 0:
                    sizes2.append(sin_mov)
                    labels2.append(f'Ninguno\n({sin_mov} pers.)') # Aquí se añade el conteo exacto de personas
                    colors2.append('#c2f0c2')
                
                if sum(sizes2) > 0:
                    ax2.pie(sizes2, labels=labels2, autopct='%1.1f%%', startangle=90, colors=colors2)
                    ax2.axis('equal')
                    ax2.set_title('Desglose: Ajuste vs Promo', fontweight='bold')
                else:
                    ax2.text(0.5, 0.5, "Sin datos", ha='center', va='center')
                
                st.pyplot(fig2)

            # --- GRÁFICA 3: Top 10 Reporting Org ---
            with col3:
                fig3, ax3 = plt.subplots(figsize=(5, 4))
                # Ajustado para evitar los "Warnings" de PowerShell
                sns.countplot(y='Reporting Organization', data=df_filtered, 
                              order=df_filtered['Reporting Organization'].value_counts().index[:10], 
                              hue='Reporting Organization', legend=False, palette='viridis', ax=ax3)
                ax3.set_xlabel('Cantidad')
                ax3.set_ylabel('')
                ax3.set_title('Top 10 Reporting Org')
                st.pyplot(fig3)
                
    except Exception as e:
        st.error(f"Ocurrió un error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos (el Principal y el SuperFlex) para comenzar.")
