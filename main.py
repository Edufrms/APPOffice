import streamlit as st
import pandas as pd
import os
from datetime import datetime

EXCEL_FILE = 'empresas_objetivo.xlsx'
SHEET_NAME = 'Empresas'

CAMPOS = [
    'Nombre de la empresa',
    'País',
    'Sector',
    'Nivel de interés',
    'Fecha de contacto'
]

NIVELES_INTERES = ['Alto', 'Medio', 'Bajo']

@st.cache_data(show_spinner=False)
def cargar_datos():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        except Exception:
            df = pd.DataFrame(columns=CAMPOS)
    else:
        df = pd.DataFrame(columns=CAMPOS)
    return df

def guardar_datos(df):
    df.to_excel(EXCEL_FILE, index=False, sheet_name=SHEET_NAME)

st.title('Gestión de Empresas Objetivo - Red Exterior Castilla y León')
st.markdown('---')

if 'df_empresas' not in st.session_state:
    st.session_state.df_empresas = cargar_datos()

st.header('Añadir nueva empresa')
with st.form('form_nueva_empresa', clear_on_submit=True):
    nombre = st.text_input('Nombre de la empresa')
    pais = st.text_input('País')
    sector = st.text_input('Sector')
    nivel = st.selectbox('Nivel de interés', NIVELES_INTERES)
    fecha = st.date_input('Fecha de contacto', value=datetime.today())
    submitted = st.form_submit_button('Añadir empresa')

    if submitted:
        if not nombre or not pais or not sector:
            st.warning('Por favor, completa todos los campos.')
        else:
            nueva_fila = {
                'Nombre de la empresa': nombre,
                'País': pais,
                'Sector': sector,
                'Nivel de interés': nivel,
                'Fecha de contacto': fecha.strftime('%Y-%m-%d')
            }
            st.session_state.df_empresas = pd.concat([
                st.session_state.df_empresas,
                pd.DataFrame([nueva_fila])
            ], ignore_index=True)
            guardar_datos(st.session_state.df_empresas)
            st.success('¡Empresa añadida correctamente!')

st.markdown('---')

def mostrar_tabla():
    st.header('Empresas registradas')
    df = st.session_state.df_empresas.copy()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        paises = ['Todos'] + sorted(df['País'].dropna().unique().tolist())
        filtro_pais = st.selectbox('Filtrar por país', paises)
    with col2:
        sectores = ['Todos'] + sorted(df['Sector'].dropna().unique().tolist())
        filtro_sector = st.selectbox('Filtrar por sector', sectores)
    with col3:
        niveles = ['Todos'] + NIVELES_INTERES
        filtro_nivel = st.selectbox('Filtrar por nivel de interés', niveles)

    if filtro_pais != 'Todos':
        df = df[df['País'] == filtro_pais]
    if filtro_sector != 'Todos':
        df = df[df['Sector'] == filtro_sector]
    if filtro_nivel != 'Todos':
        df = df[df['Nivel de interés'] == filtro_nivel]

    st.dataframe(df, use_container_width=True)

    st.markdown('---')
    st.subheader('Exportar tabla filtrada')
    exportar = st.button('Exportar a Excel')
    if exportar:
        nombre_export = f'empresas_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        df.to_excel(nombre_export, index=False)
        with open(nombre_export, 'rb') as f:
            st.download_button(
                label='Descargar archivo Excel',
                data=f,
                file_name=nombre_export,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

mostrar_tabla()