import streamlit as st
import pandas as pd
from datetime import datetime
import requests

# --- CONFIGURACIÓN ---
# Tu URL de destino real
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

    # --- PANTALLA DE REDIRECCIÓN ---
    st.success(f"¡Acceso verificado para {cli}!")
    st.write("Redirigiendo al temario...")

    enlace_final = f"{URL_APP_PRINCIPAL}/?id={id_l}"

    # BOTÓN MANUAL (Es lo más seguro para evitar el bucle de redirecciones)
    st.link_button("HAGA CLIC AQUÍ PARA ENTRAR", enlace_final)
    
    # Redirección automática un poco más lenta para no marear al navegador
    st.markdown(f'<meta http-equiv="refresh" content="4;URL={enlace_final}">', unsafe_allow_html=True)

else:
    st.warning("Por favor, escanea un código QR válido.")