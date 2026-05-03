import streamlit as st
import pandas as pd

# CONFIGURACIÓN DE LA WEB (Tu diseño original)
st.set_page_config(page_title="Verificación de Temarios", page_icon="⚖️")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .status-box { padding: 20px; border-radius: 15px; text-align: center; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("⚖️ Verificación de Autenticidad")
st.subheader("Temarios Vázquez Pariente")

# Cargar el histórico que sube el otro script
@st.cache_data(ttl=60)
def cargar_datos():
    try:
        return pd.read_excel("HISTORICO.xlsx", dtype=str)
    except:
        return None

df = cargar_datos()
params = st.query_params
token = params.get("id")

if token and df is not None:
    res = df[df['TOKEN'] == token]
    if not res.empty:
        st.balloons()
        st.success(f"### ✅ DOCUMENTO VÁLIDO")
        st.write(f"**Autorizado para:** {res.iloc[0]['NOMBRE'].upper()}")
        st.write(f"**Libro:** {params.get('libro')}")
    else:
        st.error("❌ CÓDIGO NO ENCONTRADO")
else:
    st.info("Escanee el QR para verificar.")