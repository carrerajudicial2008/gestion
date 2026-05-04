import streamlit as st
import pandas as pd
from datetime import datetime
import requests

# --- CONFIGURACIÓN ---
# Aquí pondremos más adelante la URL real de tu aplicación principal
URL_APP_PRINCIPAL = "https://vazquezpariente.streamlit.app" 

def obtener_ubicacion():
    try:
        # Servicio gratuito para obtener IP y Ciudad
        res = requests.get('https://ipapi.co/json/').json()
        return res.get('ip', '0.0.0.0'), res.get('city', 'Desconocida')
    except:
        return "Error", "Error"

# --- LÓGICA DE REGISTRO ---
# Capturamos los datos que vienen del QR (id del libro y usuario)
params = st.query_params

if "id" in params:
    id_l = params["id"]
    cli = params.get("u", "Anonimo")
    ip, ciudad = obtener_ubicacion()
    dispositivo = st.context.headers.get("User-Agent", "Desconocido")

    # Guardar en el Excel localmente (mientras probamos en el PC)
    nueva_fila = {
        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ID_Libro": id_l,
        "Cliente": cli,
        "IP": ip,
        "Dispositivo": dispositivo[:50],
        "Ciudad": ciudad,
        "Alerta": "NORMAL"
    }
    
    # Añadimos la fila al Excel
    df = pd.read_excel("SEGURIDAD.xlsx")
    df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
    df.to_excel("SEGURIDAD.xlsx", index=False)

    # REDIRECCIÓN: Envía al usuario a la app principal
    enlace_destino = f"{URL_APP_PRINCIPAL}/?id={id_l}"
    
    # CORRECCIÓN AQUÍ: Cambiado unsafe_content por unsafe_allow_html
    st.markdown(f'<meta http-equiv="refresh" content="0;URL={enlace_destino}">', unsafe_allow_html=True)
    st.write("Verificando acceso...")

else:
    st.warning("Por favor, escanea un código QR válido.")