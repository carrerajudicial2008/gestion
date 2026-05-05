import streamlit as st
import pandas as pd
from datetime import datetime
import requests
from github import Github
import io

# --- CONFIGURACIÓN ---
URL_APP_PRINCIPAL = "https://vazquezpariente.streamlit.app"
REPO_NAME = "carrerajudicial2008/gestion" 
EXCEL_FILE = "SEGURIDAD.xlsx"

def obtener_ubicacion():
    try:
        res = requests.get('https://ipapi.co/json/').json()
        return res.get('ip', '0.0.0.0'), res.get('city', 'Desconocida')
    except:
        return "Error", "Error"

def guardar_en_github(nueva_fila):
    try:
        if "GITHUB_TOKEN" not in st.secrets:
            st.error("Falta GITHUB_TOKEN en Secrets")
            return False
            
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(REPO_NAME)
        
        # Intentar obtener el archivo
        contents = repo.get_contents(EXCEL_FILE)
        df = pd.read_excel(io.BytesIO(contents.decoded_content))
        
        # Añadir datos
        df_nueva = pd.DataFrame([nueva_fila])
        df = pd.concat([df, df_nueva], ignore_index=True)
        
        # Convertir a Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        # Subir
        repo.update_file(
            contents.path, 
            f"Registro: {nueva_fila['Cliente']}", 
            output.getvalue(), 
            contents.sha
        )
        return True
    except Exception as e:
        st.error(f"Error detallado: {e}")
        return False

# --- LÓGICA PRINCIPAL ---
st.title("Sistema de Verificación")

# Usamos st.query_params directamente
params = st.query_params

if "id" in params:
    id_l = params["id"]
    cli = params.get("u", "Usuario_Anonimo")
    
    # REGISTRO OBLIGATORIO (Sin st.session_state para probar que funcione siempre)
    ip, ciudad = obtener_ubicacion()
    datos_registro = {
        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ID_Libro": id_l,
        "Cliente": cli,
        "IP": ip,
        "Dispositivo": "Web_Access",
        "Ciudad": ciudad,
        "Alerta": "TEST"
    }
    
    st.info(f"Registrando acceso para {cli}...")
    
    # LLAMADA A LA FUNCIÓN
    if guardar_en_github(datos_registro):
        st.success(f"✅ ¡Hola {cli}! Acceso registrado en la nube.")
        
        enlace_final = f"{URL_APP_PRINCIPAL}/?id={id_l}"
        st.markdown(f"""
            <div style="text-align: center; margin-top: 30px;">
                <a href="{enlace_final}" target="_blank" 
                   style="text-decoration: none; background-color: #28a745; color: white; padding: 20px 40px; border-radius: 15px; font-weight: bold; font-size: 22px;">
                   🚀 PULSA AQUÍ PARA ENTRAR
                </a>
            </div>
        """, unsafe_allow_html=True)
    else:
        st.error("No se pudo validar el registro. Inténtalo de nuevo.")
else:
    st.warning("⚠️ Esperando código QR...")