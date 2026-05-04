import streamlit as st
import pandas as pd

# 1. Configuración visual
st.set_page_config(page_title="Verificador Oficial", page_icon="🛡️")

# 2. Leer el histórico que genera tu script de Python
try:
    # Intentamos leer el Excel que estará en el mismo repo
    df = pd.read_excel("HISTORICO.xlsx", engine='openpyxl')
    df.columns = [str(c).strip().upper() for c in df.columns]
except Exception as e:
    st.error("Base de datos no encontrada o error al leer.")
    st.stop()

# 3. Lógica de escaneo (Captura datos del QR)
# El QR envía: ?id=TOKEN123&libro=CO
params = st.query_params

if "id" in params:
    token_cliente = params["id"]
    libro_cod = params.get("libro", "Desconocido")
    
    # Buscar el token en la columna 'TOKEN'
    coincidencia = df[df["TOKEN"] == token_cliente]
    
    if not coincidencia.empty:
        st.success("### ✅ EJEMPLAR AUTÉNTICO")
        nombre = coincidencia.iloc[0].get("NOMBRE", "Usuario").title()
        st.write(f"Este libro pertenece a: **{nombre}**")
        st.write(f"Código de materia: **{libro_cod}**")
        st.balloons() 
    else:
        st.error("### ❌ CÓDIGO NO VÁLIDO")
        st.write("Este ejemplar no figura en nuestros registros oficiales.")
else:
    st.info("Por favor, escanea el código QR de tu libro para verificar la autenticidad.")