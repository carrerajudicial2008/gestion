import os
import fitz  # PyMuPDF
import segno
import pandas as pd
import io
import secrets
import string
import openpyxl 
import datetime
import git

# === CONFIGURACIÓN DE RUTAS ===
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_PATH, "PEDIDOS.xlsx")
HISTORICO_FILE = os.path.join(BASE_PATH, "HISTORICO.xlsx")
BASE_FOLDER = BASE_PATH
OUTPUT_FOLDER = os.path.abspath(os.path.join(BASE_PATH, "..", "LIBROS_FINALIZADOS"))

MIS_CODIGOS = ["CO", "CI1", "CI2", "PE1", "PE2", "PC1", "PC2", "PP", "ME", "AD"]
URL_WEB_VERIFICACION = "https://vazquezpariente.streamlit.app"
URL_PARA_PIE = "www.vazquezpariente.com/verificacion"

NOMBRES_LIBROS = {
    "CO": "Derecho Constitucional y de la Unión Europea",
    "CI1": "Derecho Civil I", "CI2": "Derecho Civil II",
    "PE1": "Derecho Penal I", "PE2": "Derecho Penal II",
    "PC1": "Derecho Procesal Civil I", "PC2": "Derecho Procesal Civil II",
    "PP": "Derecho Procesal Penal", "ME": "Derecho Mercantil", "AD": "Derecho Administrativo y Laboral"
}

def generar_token_seguro():
    caracteres = string.ascii_uppercase + string.digits
    return ''.join(secrets.choice(caracteres) for _ in range(12))

def anexar_a_historico(df_procesados):
    """Anexa registros al histórico con orden: FECHA, TOKEN, DATOS..."""
    try:
        df_procesados["FECHA_REGISTRO"] = datetime.datetime.now().strftime("%d/%m/%Y")
        
        # Reordenar: FECHA y TOKEN al principio, luego el resto
        cols_pedidos = [c for c in df_procesados.columns if c not in ["FECHA_REGISTRO", "TOKEN"]]
        nuevo_orden = ["FECHA_REGISTRO", "TOKEN"] + cols_pedidos
        df_reordenado = df_procesados[nuevo_orden]
        
        if not os.path.exists(HISTORICO_FILE):
            df_reordenado.to_excel(HISTORICO_FILE, index=False, engine='openpyxl')
            print(f"✨ HISTORICO.xlsx creado con éxito.")
            return

        wb = openpyxl.load_workbook(HISTORICO_FILE)
        ws = wb.active
        for r in df_reordenado.values.tolist():
            ws.append(r)
        wb.save(HISTORICO_FILE)
        print(f"✅ {len(df_reordenado)} registros añadidos al HISTORICO.xlsx.")
    except Exception as e:
        print(f"❌ Error al anexar al histórico: {e}")

def generar_pdf_copisteria(datos_agrupados, ruta_salida):
    print(f"\n--- GENERANDO HOJA DE ENCARGO PARA COPISTERÍA ---")
    try:
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        page.insert_text((30, 45), "LISTA DE PEDIDOS - HOJA DE ENCARGO", fontsize=14, fontname="helv")
        header_y = 70
        x = [30, 170, 420, 510] 
        headers = ["NOMBRE", "DIRECCIÓN / DETALLES", "TELÉFONO", "TEMAS"]
        for i, h in enumerate(headers):
            page.insert_text((x[i], header_y), h, fontsize=9, fontname="helv")
        page.draw_line((30, header_y + 5), (565, header_y + 5), width=1)
        y_actual = header_y + 15
        
        for item in datos_agrupados:
            nom = str(item.get('OPOSITOR', '')).upper()
            tel = str(item.get('TELÉFONO', '')).upper()
            calle = str(item.get('DIRECCIÓN', '')).strip().upper()
            detalle = str(item.get('DETALLE', '')).strip().upper()
            cp = str(item.get('CP', '')).strip().upper()
            ciudad = str(item.get('CIUDAD', '')).strip().upper()
            prov = str(item.get('PROVINCIA', '')).strip().upper()
            
            direccion_f = "\n".join([l for l in [f"{calle} {detalle}".strip(), f"{cp} {ciudad}".strip(), f"({prov})" if prov else ""] if l])
            temas_v = "\n".join([str(t).upper() for t in item.get('TEMARIOS', [])])
            
            lineas_max = max(direccion_f.count('\n') + 1, len(item.get('TEMARIOS', [])))
            altura_fila = max(60, (lineas_max * 12) + 10)
            
            if y_actual + altura_fila > 800:
                page = doc.new_page(width=595, height=842)
                y_actual = 50

            page.insert_textbox(fitz.Rect(x[0], y_actual, x[1] - 5, y_actual + altura_fila), nom, fontsize=10, fontname="helv")
            page.insert_textbox(fitz.Rect(x[1], y_actual, x[2] - 10, y_actual + altura_fila), direccion_f, fontsize=10, fontname="helv")
            page.insert_textbox(fitz.Rect(x[2], y_actual, x[3] - 5, y_actual + altura_fila), tel, fontsize=10, fontname="helv")
            page.insert_textbox(fitz.Rect(x[3], y_actual, 565, y_actual + altura_fila), temas_v, fontsize=10, fontname="helv")
            
            y_actual += altura_fila
            page.draw_line((30, y_actual - 2), (565, y_actual - 2), color=(0.8, 0.8, 0.8), width=0.5)
            y_actual += 8
        doc.save(ruta_salida)
        doc.close()
    except Exception as e:
        print(f"❌ ERROR en hoja de encargo: {e}")
def subir_a_github():
    print(f"\n--- SINCRONIZANDO CON GITHUB ---")
    try:
        # Detecta la carpeta donde está el script
        ruta_repo = os.path.dirname(os.path.abspath(__file__))
        repo = git.Repo(ruta_repo)
        
        # 1. Prepara el archivo HISTORICO.xlsx para subir
        repo.index.add([HISTORICO_FILE])
        
        # 2. Crea una nota interna con la fecha del cambio
        fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        repo.index.commit(f"Actualización automática: {fecha_actual}")
        
        # 3. Empuja el archivo a GitHub
        origen = repo.remotes.origin
        origen.push()
        
        print(f"🚀 HISTORICO.xlsx subido a GitHub con éxito.")
    except Exception as e:
        print(f"⚠️ Error al subir a GitHub: {e}")

def procesar_sistema_gestion():
    try:
        if not os.path.exists(EXCEL_FILE):
            print(f"❌ No se encuentra: {EXCEL_FILE}"); return
        df = pd.read_excel(EXCEL_FILE, dtype=str, engine='openpyxl').fillna("")
        df.columns = [str(c).strip().upper() for c in df.columns]
        if df.empty:
            print("📭 PEDIDOS.xlsx está vacío."); return
        if "TOKEN" not in df.columns: df["TOKEN"] = ""
    except Exception as e:
        print(f"❌ Error al leer PEDIDOS.xlsx: {e}"); return

    if not os.path.exists(OUTPUT_FOLDER): os.makedirs(OUTPUT_FOLDER)
    
    col_nombre = next((c for c in df.columns if "NOMBRE" in c), "NOMBRE")
    col_tel = next((c for c in df.columns if any(p in c for p in ["TEL", "MÓVIL", "TELEFONO"])), "TELEFONO")
    col_dir = next((c for c in df.columns if "DIREC" in c), "DIRECCIÓN")
    col_det = next((c for c in df.columns if "DETALLE" in c), "DETALLE")
    col_cp = next((c for c in df.columns if "CP" in c), "CP")
    col_ciu = next((c for c in df.columns if "CIUDAD" in c), "CIUDAD")
    col_pro = next((c for c in df.columns if "PROV" in c), "PROVINCIA")

    lista_para_hoja = []
    indices_procesados = []

    for index, row in df.iterrows():
        nombre_c = row.get(col_nombre, "").strip()
        if not nombre_c: continue
        
        temas = [c for c in MIS_CODIGOS if str(row.get(c.upper(), "")).strip().upper() not in ["", "0", "NO", "NAN"]]
        
        if temas:
            indices_procesados.append(index)
            token_actual = generar_token_seguro()
            df.at[index, "TOKEN"] = token_actual

            print(f"📦 Procesando PDFs para: {nombre_c}")
            nombre_t, tel_a = nombre_c.title(), str(row.get(col_tel, "")).strip()
            folder_c = os.path.join(OUTPUT_FOLDER, nombre_t.replace(" ", "_"))
            if not os.path.exists(folder_c): os.makedirs(folder_c)

            for cod in temas:
                pdf_orig = os.path.join(BASE_FOLDER, f"{cod}.pdf")
                if os.path.exists(pdf_orig):
                    try:
                        doc_lib = fitz.open(pdf_orig)
                        qr_buffer = io.BytesIO()
                        segno.make(f"{URL_WEB_VERIFICACION}/?id={token_actual}&libro={cod}", error='h').save(qr_buffer, kind='png', scale=5)
                        qr_bytes = qr_buffer.getvalue()

                        for i, pag in enumerate(doc_lib):
                            # QR Esquinas
                            pag.insert_image(fitz.Rect(50, 25, 86, 61), stream=qr_bytes, overlay=True)
                            pag.insert_image(fitz.Rect(pag.rect.width-75, pag.rect.height-54, pag.rect.width-39, pag.rect.height-18), stream=qr_bytes, overlay=True)
                            
                            # Sello Lateral
                            sello = f"{nombre_t} - {tel_a} - {token_actual} - {NOMBRES_LIBROS.get(cod, cod)}"
                            pag.insert_textbox(fitz.Rect(pag.rect.width-25, 220, pag.rect.width-5, 620), sello, fontsize=8, color=(0.5, 0.5, 0.5), rotate=90, overlay=True)
                            
                            # Pie de página
                            if i > 0:
                                texto_pie = f"Verificación de autenticidad: {URL_PARA_PIE}"
                                pag.insert_text(((595 - fitz.get_text_length(texto_pie, fontsize=8)) / 2, 822), texto_pie, fontsize=8, fontname="helv", color=(0.5, 0.5, 0.5))

                            # MARCA DE AGUA CENTRAL (Reintegrada)
                            if i > 0 and (i % 4 == 0):
                                lineas = [nombre_c.upper(), token_actual]
                                f_size = 77
                                max_len = max([fitz.get_text_length(l, fontname="helv", fontsize=f_size) for l in lineas])
                                while max_len > 500 and f_size > 15:
                                    f_size -= 2
                                    max_len = max([fitz.get_text_length(l, fontname="helv", fontsize=f_size) for l in lineas])

                                p_centro = fitz.Point(297.5, 421)
                                m_rot = fitz.Matrix(35)
                                for idx, texto in enumerate(lineas):
                                    long_l = fitz.get_text_length(texto, fontname="helv", fontsize=f_size)
                                    offset_y = (idx - 0.5) * (f_size * 1.2)
                                    p_ini = fitz.Point(297.5 - (long_l / 2), 421 + offset_y)
                                    pag.insert_text(p_ini, texto, fontsize=f_size, fontname="helv", color=(0.68, 0.68, 0.68), fill_opacity=0.3, morph=(p_centro, m_rot), overlay=True)
                        
                        doc_lib.save(os.path.join(folder_c, f"{cod}_{nombre_t.replace(' ','_')}.pdf"))
                        doc_lib.close()
                    except Exception as e: print(f"❌ Error en {cod}: {e}")

            lista_para_hoja.append({
                'OPOSITOR': nombre_c, 'TELÉFONO': tel_a, 'TEMARIOS': temas,
                'DIRECCIÓN': row.get(col_dir, ''), 'DETALLE': row.get(col_det, ''),
                'CP': row.get(col_cp, ''), 'CIUDAD': row.get(col_ciu, ''), 'PROVINCIA': row.get(col_pro, '')
            })

    if indices_procesados:
        generar_pdf_copisteria(lista_para_hoja, os.path.join(OUTPUT_FOLDER, "HOJA_DE_ENCARGO.pdf"))
        df_solo_procesados = df.loc[indices_procesados].copy()
        anexar_a_historico(df_solo_procesados)
        print(f"\n✅ Proceso completado. Watermark central incluida.")
    else:
        print("ℹ️ Nada que procesar.")

if __name__ == "__main__":
    procesar_sistema_gestion()
    subir_a_github()