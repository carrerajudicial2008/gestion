import streamlit as st
import pandas as pd
from datetime import datetime
import requests

# --- CONFIGURACIÓN ---
# Esta es tu URL de destino real
URL_APP_PRINCIPAL = "https://vazquezpariente.streamlit.app" 

def obtener_ubicacion():
    try:
        # Servicio gratuito para obtener IP y Ciudad
        res = requests.get('https://ipapi.co/json/').json()
        return res.get('ip', '0.0.0.0'), res.get('city', 'Desconocida')
    except:
        return "Error", "Error"

# --- LÓGICA DE REGISTRO ---
params = st.query_params

if "id" in params:
    id_l = params["id"]
    cli = params.get("u", "Anonimo")
    ip, ciudad = obtener_ubicacion()
    dispositivo = st.context.headers.get("User-Agent", "Desconocido")

    # Guardar en el Excel (Importante: en la nube esto puede dar error hasta que configuremos GitHub)
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
    except Exception as e:
        st.error(f"Error al registrar datos: {e}")

    # REDIRECCIÓN MEJORADA
    enlace_destino = f"{URL_APP_PRINCIPAL}/?id={id_l}"
    
    st.success(f"¡Bienvenido {cli}! Verificando credenciales...")
    
    # Opción A: Redirección por Meta-Refresh (con 2 segundos de cortesía)
    st.markdown(f'<meta http-equiv="refresh" content="2;URL={enlace_destino}">', unsafe_allow_html=True)
    
    # Opción B: Botón manual por si el navegador bloquea el auto-salto
    st.markdown(f"""
        <div style="text-align: center; margin-top: 20px;">
            <p>Si no eres redirigido automáticamente en 2 segundos...</p>
            <a href="{enlace_destino}" target="_self" 
               style="text-decoration: none; background-color: #ff4b4b; color: white; padding: 10px 20px; border-radius: 10px;">
               HAGA CLIC AQUÍ PARA ACCEDER
            </a>
        </div>
    """, unsafe_allow_html=True)

else:
    st.warning("Por favor, escanea un código QR válido para acceder al temario.")