import streamlit as st
import pandas as pd
from datetime import datetime
import requests
from github import Github
import io

# --- CONFIGURACIÓN ---
URL_APP_PRINCIPAL = "https://vazquezpariente.streamlit.app"
REPO_NAME = "carrerajudicial2008/gestion" 
EXCEL_SEGURIDAD = "SEGURIDAD.xlsx"
EXCEL_HISTORICO = "HISTORICO.xlsx"  # Tu base de datos real

def obtener_ubicacion():
    try:
        res = requests.get('https://ipapi.co/json/').json()
        return res.get('ip', '0.0.0.0'), res.get('city', 'Desconocida')
    except:
        return "Error", "Error"

def guardar_en_github(nueva_fila):
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(REPO_NAME)
        contents = repo.get_contents(EXCEL_SEGURIDAD)
        df = pd.read_excel(io.BytesIO(contents.decoded_content))
        df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        repo.update_file(contents.path, f"Registro: {nueva_fila['Cliente']}", output.getvalue(), contents.sha)
        return True
    except Exception as e:
        st.error(f"Error al registrar: {e}")
        return False

# --- LÓGICA DE VERIFICACIÓN ---
params = st.query_params
if "id" in params:
    id_l = params["id"]
    cli = params.get("u", "Usuario")

    # 1. Verificar si el ID existe en HISTORICO.xlsx
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(REPO_NAME)
        contents_h = repo.get_contents(EXCEL_HISTORICO)
        df_h = pd.read_excel(io.BytesIO(contents_h.decoded_content))
        
        # Comprobamos si el ID está en la columna correspondiente (ajusta el nombre 'ID_LIBRO' si es distinto)
        if id_l in df_h.values:
            ip, ciudad = obtener_ubicacion()
            datos_registro = {
                "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ID_Libro": id_l,
                "Cliente": cli,
                "IP": ip,
                "Ciudad": ciudad
            }
            
            if guardar_en_github(datos_registro):
                st.success(f"✅ Acceso verificado para {cli}")
                st.markdown(f'<a href="{URL_APP_PRINCIPAL}/?id={id_l}" target="_self" style="background-color: #28a745; color: white; padding: 15px; border-radius: 10px; text-decoration: none;">🚀 ENTRAR AL TEMARIO</a>', unsafe_allow_html=True)
        else:
            st.error("❌ CÓDIGO NO VÁLIDO. Este ejemplar no figura en el HISTORICO.")
    except Exception as e:
        st.error(f"Error al conectar con la base de datos: {e}")
else:
    st.warning("Esperando código QR...")