import os
import datetime

# =========================================================================
# 1. CONFIGURACIÓN DE RUTA PARA TU PC (CORREGIDA)
# =========================================================================
# Hemos cambiado 'bin' por 'cmd' que es donde tu PC tiene guardado el programa
os.environ["GIT_PYTHON_GIT_EXECUTABLE"] = r"C:\Program Files\Git\cmd\git.exe"

import fitz  # PyMuPDF
import segno
import pandas as pd
import io
import secrets
import string
import openpyxl 

try:
    import git
    from git import Repo
except ImportError:
    print("⚠️ No se pudo cargar la librería GitPython. Instálala con: pip install GitPython")

# =========================================================================
# 2. CONFIGURACIÓN DE RUTAS
# =========================================================================
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_PATH, "PEDIDOS.xlsx")
HISTORICO_FILE = os.path.join(BASE_PATH, "HISTORICO.xlsx")
BASE_FOLDER = BASE_PATH
OUTPUT_FOLDER = os.path.abspath(os.path.join(BASE_PATH, "..", "LIBROS_FINALIZADOS"))

MIS_CODIGOS = ["CO", "CI1", "CI2", "PE1", "PE2", "PC1", "PC2", "PP", "ME", "AD"]
URL_WEB_VERIFICACION = "https://seguridad-libros-kazw34k4ngbtvzgcr8iqnt.streamlit.app"
URL_PARA_PIE = "www.vazquezpariente.com/verificacion"

NOMBRES_LIBROS = {
    "CO": "Derecho Constitucional y de la Unión Europea",
    "CI1": "Derecho Civil I", "CI2": "Derecho Civil II",
    "PE1": "Derecho Penal I", "PE2": "Derecho Penal II",
    "PC1": "Derecho Procesal Civil I", "PC2": "Derecho Procesal Civil II",
    "PP": "Derecho Procesal Penal", "ME": "Derecho Mercantil", "AD": "Derecho Administrativo y Laboral"
}

# =========================================================================
# 3. FUNCIONES
# =========================================================================
def generar_token_seguro():
    caracteres = string.ascii_uppercase + string.digits
    return ''.join(secrets.choice(caracteres) for _ in range(12))

def anexar_a_historico(df_procesados):
    """Guarda en Excel local y sube a GitHub"""
    try:
        df_procesados["FECHA_REGISTRO"] = datetime.datetime.now().strftime("%d/%m/%Y")
        cols_pedidos = [c for c in df_procesados.columns if c not in ["FECHA_REGISTRO", "TOKEN"]]
        nuevo_orden = ["FECHA_REGISTRO", "TOKEN"] + cols_pedidos
        df_reordenado = df_procesados[nuevo_orden]
        
        if not os.path.exists(HISTORICO_FILE):
            df_reordenado.to_excel(HISTORICO_FILE, index=False, engine='openpyxl')
        else:
            wb = openpyxl.load_workbook(HISTORICO_FILE)
            ws = wb.active
            for r in df_reordenado.values.tolist():
                ws.append(r)
            wb.save(HISTORICO_FILE)
        
        print(f"✅ Local: Datos añadidos a HISTORICO.xlsx")

        # --- SINCRONIZACIÓN AUTOMÁTICA ---
        try:
            print("🚀 Sincronizando con GitHub...")
            repo = Repo(BASE_PATH)
            repo.index.add([HISTORICO_FILE])
            ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            repo.index.commit(f"Update automático: {ahora}")
            origin = repo.remote(name='origin')
            origin.push()
            print("✨ ¡GitHub actualizado! Web lista.")
        except Exception as git_err:
            print(f"⚠️ Error al subir a GitHub: {git_err}")

    except Exception as e:
        print(f"❌ Error en histórico: {e}")

# =========================================================================
# 4. EJECUCIÓN (LÓGICA DE PROCESAMIENTO)
# =========================================================================
def procesar():
    if not os.path.exists(EXCEL_FILE):
        print("❌ No se encuentra PEDIDOS.xlsx"); return
    
    df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
    df.columns = [str(c).strip().upper() for c in df.columns]
    if "TOKEN" not in df.columns: df["TOKEN"] = ""
    
    if not os.path.exists(OUTPUT_FOLDER): os.makedirs(OUTPUT_FOLDER)
    
    col_nombre = next((c for c in df.columns if "NOMBRE" in c), "NOMBRE")
    col_tel = next((c for c in df.columns if "TEL" in c), "TELEFONO")

    procesados = []

    for index, row in df.iterrows():
        nombre = row.get(col_nombre, "").strip()
        if not nombre: continue
        
        # Filtramos los temas marcados (con 1, SI, si...)
        temas = [c for c in MIS_CODIGOS if str(row.get(c, "")).strip() in ["1", "SI", "si", "1.0", "1"]]
        
        if temas:
            token = generar_token_seguro()
            df.at[index, "TOKEN"] = token
            print(f"🔨 Generando libros para: {nombre}")
            
            folder_c = os.path.join(OUTPUT_FOLDER, nombre.title().replace(" ", "_"))
            if not os.path.exists(folder_c): os.makedirs(folder_c)

            for cod in temas:
                pdf_in = os.path.join(BASE_PATH, f"{cod}.pdf")
                if os.path.exists(pdf_in):
                    doc = fitz.open(pdf_in)
                    qr_buf = io.BytesIO()
                    segno.make(f"{URL_WEB_VERIFICACION}/?id={token}&libro={cod}", error='h').save(qr_buf, kind='png', scale=5)
                    qr_img = qr_buf.getvalue()
                    
                    for pag in doc:
                        pag.insert_image(fitz.Rect(50, 25, 86, 61), stream=qr_img)
                        pag.insert_textbox(fitz.Rect(pag.rect.width-25, 220, pag.rect.width-5, 620), f"{nombre.upper()} - {token}", fontsize=8, rotate=90)
                    
                    doc.save(os.path.join(folder_c, f"{cod}_{nombre.replace(' ','_')}.pdf"))
                    doc.close()
            
            procesados.append(index)

    if procesados:
        anexar_a_historico(df.loc[procesados].copy())
        print("\n🏆 ¡Todo listo! PDFs creados y GitHub actualizado.")

if __name__ == "__main__":
    procesar()