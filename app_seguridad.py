import streamlit as st
import pandas as pd
from datetime import datetime
import requests

# --- CONFIGURACIÓN ---
URL_APP_PRINCIPAL = "https://vazquezpariente.streamlit.app" 

def obtener_ubicacion():
    try:
        res = requests.get('https://ipapi.co/json/').json()
        return res.get('ip', '0.0.0.0'), res.get('city', 'Desconocida')
    except:
        return "Error", "Error"

# --- LÓGICA DE REGISTRO ---
params = st.query_params

if "id" in params:
    id_l = params["id"]
    cli = params.get("u", "Usuario")
    ip, ciudad = obtener_ubicacion()
    dispositivo = st.context.headers.get("User-Agent", "Desconocido")

    # Guardar en el Excel
    try:
        nueva_fila = {
            "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ID_Libro": id_l,
            "Cliente": cli,
            "IP": ip,
            "Dispositivo": dispositivo[:50],
            "Ciudad": ciudad,
            "Alerta": "NORMAL"
        }
        df = pd.read_excel("SEGURIDAD.xlsx")
        df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
        df.to_excel("SEGURIDAD.xlsx", index=False)
    except Exception:
        pass 

    # --- PANTALLA DE ACCESO ---
    st.success(f"✅ ¡Hola {cli}! Acceso verificado correctamente.")
    
    enlace_final = f"{URL_APP_PRINCIPAL}/?id={id_l}"

    # Creamos un botón grande que abra en pestaña nueva (target="_blank")
    # Esto es lo que rompe el error de "too many redirects"
    st.markdown(f"""
        <div style="text-align: center; margin-top: 30px;">
            <a href="{enlace_final}" target="_blank" 
               style="text-decoration: none; background-color: #28a745; color: white; padding: 20px 40px; border-radius: 15px; font-weight: bold; font-size: 22px; border: 2px solid #1e7e34;">
               🚀 PULSA AQUÍ PARA ENTRAR
            </a>
            <p style="margin-top: 20px; color: #666;">
                El temario se abrirá en una ventana nueva para asegurar tu conexión.
            </p>
        </div>
    """, unsafe_allow_html=True)

else:
    st.warning("⚠️ Por favor, escanea un código QR válido para acceder.")