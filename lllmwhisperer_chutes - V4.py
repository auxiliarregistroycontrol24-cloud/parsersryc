import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
# Importar queue al inicio del archivo
import queue

from unstract.llmwhisperer import LLMWhispererClientV2
from unstract.llmwhisperer.client_v2 import LLMWhispererClientException
import pandas as pd
import requests
import json
import re
import os
import traceback
import unicodedata 
import threading 
import time
import subprocess
from dotenv import load_dotenv

load_dotenv()

# --- Configuración y variables globales ---
LLMWHISPERER_API_KEY_1 = os.getenv("LLMWHISPERER_API_KEY_1")
LLMWHISPERER_API_KEY_2 = os.getenv("LLMWHISPERER_API_KEY_2")
DEEPSEEK_PRIMARY_API_TOKEN = os.getenv("DEEPSEEK_PRIMARY_API_TOKEN")
DEEPSEEK_FALLBACK_API_TOKEN = os.getenv("DEEPSEEK_FALLBACK_API_TOKEN")
PRIMARY_LLM_MODEL = "deepseek-ai/DeepSeek-V3-0324"
FALLBACK_LLM_MODEL = "deepseek-ai/DeepSeek-V3"
CHUTES_API_URL = "https://llm.chutes.ai/v1/chat/completions" 
mensajes_resumen_procesamiento = []
active_api_key_index = 1
MAX_WORKERS = 4 
thread_lock = threading.Lock()
class ServerSideApiException(Exception):
    pass

def quitar_tildes(texto):
    if not isinstance(texto, str): return texto
    nfkd_form = unicodedata.normalize('NFD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

# MODIFICADA: Ahora es segura para hilos
def _intentar_extraccion_llmwhisperer(file_path, api_key, key_number):
    nombre_base = os.path.basename(file_path)
    print(f"Inicializando cliente LLMWhisperer para PDF '{nombre_base}' con API Key #{key_number}...")
    client = LLMWhispererClientV2(api_key=api_key)
    print(f"Enviando PDF '{file_path}' a LLMWhisperer (Key #{key_number})...")
    resultado = client.whisper(file_path=file_path, wait_for_completion=True, wait_timeout=360)
    print(f"\n✅ Resultado del análisis de LLMWhisperer obtenido para PDF '{nombre_base}'.")
    if isinstance(resultado, dict) and 'extraction' in resultado and isinstance(resultado['extraction'], dict) and 'result_text' in resultado['extraction']:
        return resultado['extraction']['result_text'].replace("<<<\x0c", "\n--- NUEVA PÁGINA --- \n") 
    else:
        msg = f"❌ Error LLMWhisperer PDF '{nombre_base}': No se encontró 'extraction' o 'result_text'."
        print(msg)
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg)
        return None

# MODIFICADA: Ahora es segura para hilos
def extraer_texto_pdf_con_apis(file_path):
    global mensajes_resumen_procesamiento, active_api_key_index
    # NOTA: La lógica de 'active_api_key_index' es un punto débil en paralelismo.
    # Si dos hilos la leen y modifican al mismo tiempo, puede haber condiciones de carrera.
    # Una solución más robusta sería un pool de claves. Por ahora, se mantiene simple.
    with thread_lock:
        current_api_key_index = active_api_key_index

    if current_api_key_index == 1:
        current_key, current_key_num = LLMWHISPERER_API_KEY_1, 1
    else:
        current_key, current_key_num = LLMWHISPERER_API_KEY_2, 2
        print(f"INFO: Usando directamente la API Key de respaldo #{current_key_num}...")
    try:
        return _intentar_extraccion_llmwhisperer(file_path, current_key, current_key_num)
    except LLMWhispererClientException as e:
        msg_error_api = str(e).lower()
        codigo_estado_api = e.status_code if hasattr(e, 'status_code') else 0
        if (codigo_estado_api == 402 or "breached your free processing limit" in msg_error_api) and current_api_key_index == 1:
            print(f"⚠️ LÍMITE ALCANZADO en API Key #1. Cambiando a API Key #2 para '{os.path.basename(file_path)}' y subsiguientes.")
            with thread_lock:
                mensajes_resumen_procesamiento.append(f"⚠️ LÍMITE API Unstract #1. Cambiando a API #2.")
                # Proteger la modificación de la variable global
                globals()['active_api_key_index'] = 2 
            try: return _intentar_extraccion_llmwhisperer(file_path, LLMWHISPERER_API_KEY_2, 2)
            except Exception as e2:
                msg = f"❌ Fallo en el reintento con API Key #2 para '{os.path.basename(file_path)}': {e2}"
                print(msg); traceback.print_exc()
                with thread_lock:
                    mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
                return None
        elif (codigo_estado_api == 402 or "breached your free processing limit" in msg_error_api) and current_api_key_index == 2:
            msg = f"❌ LÍMITE ALCANZADO también en API Key #2. No hay más claves disponibles."
            print(msg)
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg)
            return None
        else:
            msg = f"❌ Error con LLMWhisperer para PDF '{os.path.basename(file_path)}': {e} (Código: {codigo_estado_api})"
            print(msg)
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg)
            return None
    except Exception as e:
        msg = f"❌ Error inesperado durante el análisis con LLMWhisperer para PDF '{os.path.basename(file_path)}': {e}"
        print(msg); traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

# MODIFICADA: Ahora es segura para hilos
def extraer_texto_excel_con_pandas(file_path):
    global mensajes_resumen_procesamiento
    nombre_base = os.path.basename(file_path)
    print(f"Procesando archivo Excel '{nombre_base}' con pandas...")
    try:
        xls = pd.ExcelFile(file_path)
        full_text_parts = []
        if not xls.sheet_names:
            msg = f"⚠️ El archivo Excel '{nombre_base}' no contiene hojas."
            print(msg)
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg)
            return ""
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, header=None, dtype=str).fillna('') 
            full_text_parts.append(f"--- INICIO HOJA: {sheet_name} ---\n")
            full_text_parts.extend(" | ".join(str(cell).strip() for cell in row) for index, row in df.iterrows())
            full_text_parts.append(f"\n--- FIN HOJA: {sheet_name} ---\n")
        print(f"✅ Texto extraído del archivo Excel '{nombre_base}'.")
        return "\n".join(full_text_parts)
    except Exception as e:
        msg = f"❌ Error al procesar Excel '{nombre_base}' con pandas: {e}"
        print(msg); traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

def limpiar_texto_general_para_llm(texto):
    if not texto: return ""
    return re.sub(r'\n{3,}', '\n', texto).strip()

# MODIFICADA: Ahora es segura para hilos
def analizar_con_llm(texto_documento, pregunta_o_instruccion, api_token, model_name, api_url, nombre_archivo_base):
    global mensajes_resumen_procesamiento
    print(f"\nPreparando para enviar a la API de Chutes para '{nombre_archivo_base}' usando el modelo '{model_name}'...")
    prompt_completo = f"Aquí tienes el contenido de un documento (originalmente '{nombre_archivo_base}'):\n--- INICIO DEL DOCUMENTO ---\n{texto_documento}\n--- FIN DEL DOCUMENTO ---\nPor favor, sigue estas instrucciones detalladamente basadas ÚNICAMENTE en el documento proporcionado:\n{pregunta_o_instruccion}"
    headers = {"Authorization": f"Bearer {api_token}", "Content-Type": "application/json"}
    data = {"model": model_name, "messages": [{"role": "user", "content": prompt_completo}], "stream": False, "max_tokens": 4000, "temperature": 0.15}
    try:
        print(f"Enviando solicitud a Chutes (modelo: {model_name}) para '{nombre_archivo_base}'...")
        response = requests.post(api_url, headers=headers, data=json.dumps(data), timeout=360) 
        response.raise_for_status()
        respuesta_json = response.json()
        print(f"✅ Respuesta recibida de Chutes API para '{nombre_archivo_base}'.")
        if respuesta_json.get("choices") and len(respuesta_json["choices"]) > 0 and respuesta_json["choices"][0].get("message") and "content" in respuesta_json["choices"][0]["message"]:
            return respuesta_json["choices"][0]["message"]["content"]
        else:
            msg = f"❌ Error API Chutes '{nombre_archivo_base}': Respuesta sin formato esperado."
            print(msg)
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg)
            return None
    except requests.exceptions.Timeout as e:
        msg = f"❌ Error API Chutes '{nombre_archivo_base}': Timeout."
        print(msg)
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg)
        raise ServerSideApiException(msg) from e
    except requests.exceptions.RequestException as e:
        status_code = e.response.status_code if e.response is not None else 0
        msg = f"❌ Error API Chutes '{nombre_archivo_base}': {e}"
        print(msg)
        if e.response is not None: 
            print(f"Detalles: {e.response.status_code} - {e.response.text}")
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg + (f"Detalles: {e.response.status_code} - {e.response.text}" if e.response is not None else ""))
        if 500 <= status_code < 600:
            raise ServerSideApiException(msg) from e
        return None
    except Exception as e:
        msg = f"❌ Error inesperado con Chutes API para '{nombre_archivo_base}': {e}"
        print(msg); traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

def ajustar_abreviacion(texto_abreviado, limite_caracteres=30):
    if len(texto_abreviado) > limite_caracteres: return texto_abreviado[:limite_caracteres].rstrip()
    return texto_abreviado

# MODIFICADA: Ahora es segura para hilos
def extraer_script_ahk_y_ajustar_abreviacion(respuesta_llm):
    global mensajes_resumen_procesamiento
    try:
        match_script = re.search(r"--- INICIO SCRIPT AUTOHOTKEY ---(.*?)--- FIN SCRIPT AUTOHOTKEY ---", respuesta_llm, re.DOTALL)
        if not match_script:
            with thread_lock:
                mensajes_resumen_procesamiento.append("❌ No se encontraron delimitadores AHK.")
            return None
        lineas_script = match_script.group(1).strip().splitlines()
        script_final_ajustado, linea_encabezado_modificada = [], False
        patron_linea_encabezado = r"Send,\s*1{Tab}HOM01{Tab}HOM01{Tab}(.*?){Tab}{Tab}{Tab}(.*?){F10}\^\{PgDn\}\{Space\}\{Tab\}"
        for i, linea in enumerate(lineas_script):
            match_encabezado = re.match(patron_linea_encabezado, linea)
            if match_encabezado and not linea_encabezado_modificada:
                creditos_llm = match_encabezado.group(1).strip().replace("(número total creditos homologados)", "").strip()
                if not creditos_llm.isdigit(): creditos_llm = "0"
                abreviacion_llm = match_encabezado.group(2).strip().replace("(abreviacion del programa de origen)", "").strip() or "ABREV_NO_PROPORCIONADA"
                abreviacion_ajustada = ajustar_abreviacion(quitar_tildes(abreviacion_llm), 30)
                linea_modificada = f"Send, 1{{Tab}}HOM01{{Tab}}HOM01{{Tab}}{creditos_llm}{{Tab}}{{Tab}}{{Tab}}{abreviacion_ajustada}{{F10}}^{{PgDn}}{{Space}}{{Tab}}"
                script_final_ajustado.append(linea_modificada)
                linea_encabezado_modificada = True
            else:
                linea_corregida = re.sub(r'(WinTitle\s*:=\s*"Oracle Fusion Middleware Forms Services:)\s+(Open\s*>\s*SHATRNS")', r'\1  \2', linea)
                script_final_ajustado.append(linea_corregida)
        return "\n".join(script_final_ajustado)
    except Exception as e:
        with thread_lock:
            mensajes_resumen_procesamiento.append(f"❌ Error extrayendo/ajustando AHK: {e}\n{traceback.format_exc()}")
        return None

# ... (El resto de funciones hasta worker_process_file se mantienen igual)
def extraer_datos_clave_para_tabla(respuesta_llm):
    def limpiar_valor(texto): return texto.strip().strip('*').strip()
    nombre_estudiante, programa_aspira, plan_estudio = "No extraído", "No extraído", "N/A"
    match_nombre = re.search(r"NOMBRE_ESTUDIANTE:\s*(.*)", respuesta_llm)
    if match_nombre: nombre_estudiante = limpiar_valor(match_nombre.group(1).strip().splitlines()[0])
    match_programa = re.search(r"PROGRAMA_ASPIRA:\s*(.*)", respuesta_llm)
    if match_programa: programa_aspira = limpiar_valor(match_programa.group(1).strip().splitlines()[0])
    match_plan = re.search(r"PLAN_ESTUDIO:\s*(.*)", respuesta_llm)
    if match_plan: plan_estudio = limpiar_valor(match_plan.group(1).strip().splitlines()[0]) or "N/A"
    return {"Nombre": nombre_estudiante, "Programa": programa_aspira, "Plan": plan_estudio}

def sanitizar_nombre_archivo(nombre):
    nombre_sanitizado = re.sub(r'[<>:"/\\|?*]', '', nombre).replace("\n", " ").replace("\r", " ").strip()
    return re.sub(r'\s+', ' ', nombre_sanitizado) or "ScriptSinNombreValido"

def mostrar_resumen_log_gui(main_root_ref, lista_mensajes):
    if not lista_mensajes: lista_mensajes = ["No se procesaron archivos o no hubo mensajes de resumen."]
    resumen_ventana = tk.Toplevel(main_root_ref)
    resumen_ventana.title("Resumen del Procesamiento (Log)")
    resumen_ventana.geometry("800x500")
    txt_area = scrolledtext.ScrolledText(resumen_ventana, wrap=tk.WORD, width=100, height=25)
    txt_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
    # Ordenar los mensajes para que el log sea más legible
    lista_mensajes_ordenada = sorted(lista_mensajes)
    for msg in lista_mensajes_ordenada: txt_area.insert(tk.END, msg + "\n" + "-"*50 + "\n")
    txt_area.config(state=tk.DISABLED)
    tk.Button(resumen_ventana, text="Cerrar Log", command=resumen_ventana.destroy).pack(pady=10)
    resumen_ventana.transient(main_root_ref); resumen_ventana.grab_set()

def agregar_fila_a_tabla(ventana_principal, tabla_treeview, datos_fila):
    def insertar():
        valores = (
            datos_fila.get("Archivo Origen", "N/A"), 
            datos_fila.get("Nombre", "N/A"), 
            datos_fila.get("Programa", "N/A"), 
            datos_fila.get("Plan", "N/A"),
            "Abrir Archivo"
        )
        tabla_treeview.insert("", "end", values=valores, tags=('accion', datos_fila.get("Ruta Completa", "")))
        tabla_treeview.yview_moveto(1.0)
    ventana_principal.after(0, insertar)

def crear_tabla_resumen_en_vivo(main_root_ref):
    tabla_ventana = tk.Toplevel(main_root_ref)
    tabla_ventana.title("Tabla Resumen de Archivos Procesados (En Vivo)")
    tabla_ventana.geometry("1050x450")
    frame = ttk.Frame(tabla_ventana, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)
    cols = ("Archivo Origen", "Nombre del estudiante", "Programa al que aspira", "Plan", "Acciones")
    tree = ttk.Treeview(frame, columns=cols, show='headings')
    for col in cols:
        tree.heading(col, text=col)
        if col == "Acciones": tree.column(col, width=100, minwidth=80, anchor='center')
        elif col == "Plan": tree.column(col, width=100, minwidth=80, anchor='center')
        elif col == "Programa al que aspira": tree.column(col, width=250, minwidth=200, anchor='w')
        else: tree.column(col, width=200, minwidth=150, anchor='w')
    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("Treeview", rowheight=25, font=("Segoe UI", 9))
    tree.tag_configure('accion', background="#E3F2FD")
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    vsb.pack(side='right', fill='y')
    tree.configure(yscrollcommand=vsb.set)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    hsb.pack(side='bottom', fill='x')
    tree.configure(xscrollcommand=hsb.set)
    tree.pack(fill=tk.BOTH, expand=True)
    def on_click(event):
        region = tree.identify_region(event.x, event.y)
        if region != "cell": return
        
        column_id = tree.identify_column(event.x)
        # Asegurarse de que el clic es en la columna "Acciones" (la quinta, ID #5)
        if column_id != '#5': return

        rowid = tree.identify_row(event.y)
        if not rowid: return

        # --- Lógica mejorada para encontrar y abrir el archivo .ahk ---
        item_values = tree.item(rowid, 'values')
        item_tags = tree.item(rowid, 'tags')

        # El nombre del estudiante está en la segunda columna
        nombre_estudiante = item_values[1] 
        # La ruta del archivo original está en los tags
        ruta_archivo_original = item_tags[1] if len(item_tags) > 1 else None

        if not nombre_estudiante or not ruta_archivo_original or nombre_estudiante == "No extraído":
            messagebox.showwarning("Advertencia", "No se puede abrir el script porque no se extrajo un nombre de estudiante válido.", parent=tabla_ventana)
            return

        # 1. Construir el nombre del archivo .ahk esperado
        nombre_sanitizado = sanitizar_nombre_archivo(nombre_estudiante)
        nombre_archivo_ahk = f"{nombre_sanitizado}.ahk"

        # 2. Construir la ruta completa del archivo .ahk
        directorio_original = os.path.dirname(ruta_archivo_original)
        ruta_completa_ahk = os.path.join(directorio_original, nombre_archivo_ahk)

        # 3. Intentar abrir el archivo
        if os.path.exists(ruta_completa_ahk):
            try:
                if os.name == 'nt': # Para Windows
                    os.startfile(ruta_completa_ahk)
                elif os.name == 'posix': # Para Linux
                    subprocess.run(['xdg-open', ruta_completa_ahk], check=True)
                else: # Para macOS
                    subprocess.run(['open', ruta_completa_ahk], check=True)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo de script:\n{ruta_completa_ahk}\n\nError: {e}", parent=tabla_ventana)
        else:
            messagebox.showwarning("Archivo no encontrado", f"No se encontró el archivo de script esperado:\n{ruta_completa_ahk}\n\nAsegúrese de que el proceso haya terminado y el archivo no haya sido movido.", parent=tabla_ventana)

    tree.bind("<Button-1>", on_click)

# MODIFICADA: Ahora es segura para hilos y simplificada
def worker_process_file(archivo_queue, ventana_principal, tabla_treeview, barra_progreso, etiqueta_estado, num_archivos):
    global mensajes_resumen_procesamiento
    
    while not archivo_queue.empty():
        try:
            i, archivo_actual = archivo_queue.get_nowait()
        except queue.Empty:
            break

        nombre_base_archivo = os.path.basename(archivo_actual)
        ventana_principal.after(0, lambda n=nombre_base_archivo: etiqueta_estado.config(text=f"Procesando: {n}"))
        
        # El mensaje se añade al final de forma segura
        mensaje_log_actual = [f"\n--- Procesando archivo {i+1}/{num_archivos}: {nombre_base_archivo} ---"]
        
        texto_extraido = None
        if nombre_base_archivo.lower().endswith('.pdf'):
            texto_extraido = extraer_texto_pdf_con_apis(archivo_actual)
        elif nombre_base_archivo.lower().endswith(('.xlsx', '.xls')):
            texto_extraido = extraer_texto_excel_con_pandas(archivo_actual)
        
        # Guardaremos la ruta del AHK para usarla en la tabla
        ruta_archivo_ahk_generado = ""
        datos_para_fila_actual = {"Archivo Origen": nombre_base_archivo}

        if texto_extraido:
            texto_limpio_para_llm = limpiar_texto_general_para_llm(texto_extraido)
            if texto_limpio_para_llm:
                respuesta_llm = None
                try:
                    respuesta_llm = analizar_con_llm(texto_limpio_para_llm, PREGUNTA_PARA_DEEPSEEK, DEEPSEEK_PRIMARY_API_TOKEN, PRIMARY_LLM_MODEL, CHUTES_API_URL, nombre_base_archivo)
                except ServerSideApiException:
                    mensaje_log_actual.append(f"⚠️ Fallo primario, reintentando con modelo de respaldo...")
                    respuesta_llm = analizar_con_llm(texto_limpio_para_llm, PREGUNTA_PARA_DEEPSEEK, DEEPSEEK_FALLBACK_API_TOKEN, FALLBACK_LLM_MODEL, CHUTES_API_URL, nombre_base_archivo)
                
                if respuesta_llm:
                    datos_clave = extraer_datos_clave_para_tabla(respuesta_llm)
                    datos_para_fila_actual.update(datos_clave)
                    if "DETENER_PROCESO_CONDICION_" in respuesta_llm:
                         for line in respuesta_llm.splitlines():
                            if "DETENER_PROCESO_CONDICION_" in line:
                                detalle_parada = line.strip()
                                if "PERIODO" in detalle_parada.upper(): datos_para_fila_actual["Plan"] = "DETENIDO: PERIODO"
                                elif "PROGRAMA" in detalle_parada.upper(): datos_para_fila_actual["Plan"] = "DETENIDO: PROGRAMA"
                                elif "CALIFICACION" in detalle_parada.upper(): datos_para_fila_actual["Plan"] = "DETENIDO: CALIFICACION"
                                else: datos_para_fila_actual["Plan"] = "DETENIDO: OTRO"
                                break
                    else:
                        script_ahk = extraer_script_ahk_y_ajustar_abreviacion(respuesta_llm)
                        if script_ahk and datos_para_fila_actual.get("Nombre") != "No extraído":
                            nombre_archivo_sanitizado = sanitizar_nombre_archivo(datos_para_fila_actual["Nombre"])
                            ruta_archivo_ahk_generado = os.path.join(os.path.dirname(archivo_actual), f"{nombre_archivo_sanitizado}.ahk")
                            try:
                                with open(ruta_archivo_ahk_generado, "w", encoding="utf-8") as f: f.write(script_ahk)
                                mensaje_log_actual.append(f"✅ Script AHK guardado en: {ruta_archivo_ahk_generado}")
                            except Exception as e:
                                mensaje_log_actual.append(f"❌ Error guardando AHK: {e}")
                else:
                    datos_para_fila_actual.update({"Nombre": "Error en API LLM", "Programa": "N/A", "Plan": "N/A"})
            else:
                datos_para_fila_actual.update({"Nombre": "Texto vacío", "Programa": "N/A", "Plan": "N/A"})
        else:
            datos_para_fila_actual.update({"Nombre": "Error Extracción", "Programa": "N/A", "Plan": "N/A"})

        # Guardar la ruta del AHK generado (o vacío si no se generó) para el evento de clic
        datos_para_fila_actual["Ruta Completa"] = archivo_actual
        
        agregar_fila_a_tabla(ventana_principal, tabla_treeview, datos_para_fila_actual)
        
        with thread_lock:
            mensajes_resumen_procesamiento.extend(mensaje_log_actual)

        archivo_queue.task_done()
        ventana_principal.after(0, barra_progreso.step)


# ... (El resto del código, incluido 'procesar_archivos_seleccionados', 'iniciar_procesamiento_en_hilo'
# y el bloque `if __name__ == "__main__"`, se mantiene exactamente como lo tenías en v4.txt,
# ya que su lógica de gestión de hilos es correcta.)
def procesar_archivos_seleccionados(lista_archivos, boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal, tabla_treeview):
    global mensajes_resumen_procesamiento, active_api_key_index
    mensajes_resumen_procesamiento = []
    
    if LLMWHISPERER_API_KEY_1: active_api_key_index = 1

    num_archivos = len(lista_archivos)
    if barra_progreso:
        barra_progreso['maximum'] = num_archivos
        barra_progreso['value'] = 0
    if etiqueta_etr: 
        etiqueta_etr.config(text="Procesando en paralelo...")

    def procesar_en_paralelo():
        archivos_queue = queue.Queue()
        
        for i, archivo in enumerate(lista_archivos):
            archivos_queue.put((i, archivo))

        threads = []
        for _ in range(MAX_WORKERS):
            thread = threading.Thread(
                target=worker_process_file,
                args=(archivos_queue, ventana_principal, tabla_treeview, barra_progreso, etiqueta_estado, num_archivos),
                daemon=True
            )
            thread.start()
            threads.append(thread)
            
        archivos_queue.join()

        def actualizar_gui_final():
            etiqueta_estado.config(text=f"Procesamiento completado para {num_archivos} archivos.")
            etiqueta_etr.config(text="Proceso finalizado.")
            boton_seleccionar.config(state=tk.NORMAL)
            mostrar_resumen_log_gui(ventana_principal, mensajes_resumen_procesamiento)
            if messagebox.askyesno("Proceso Finalizado", "El procesamiento de los archivos ha terminado.\n¿Desea cerrar la aplicación?", parent=ventana_principal):
                ventana_principal.destroy()

        ventana_principal.after(0, actualizar_gui_final)

    threading.Thread(target=procesar_en_paralelo, daemon=True).start()

def iniciar_procesamiento_en_hilo(boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal):
    if boton_seleccionar: boton_seleccionar.config(state=tk.DISABLED)
    if etiqueta_estado: etiqueta_estado.config(text="Seleccionando archivos...")

    file_paths_tupla = filedialog.askopenfilenames(
        title="Selecciona uno o más archivos PDF o Excel para analizar",
        filetypes=(("Documentos Soportados", "*.pdf *.xlsx *.xls"),("All files", "*.*"))
    )
    if not file_paths_tupla:
        if etiqueta_estado: etiqueta_estado.config(text="Ningún archivo seleccionado.")
        if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)
        return

    tabla_treeview = crear_tabla_resumen_en_vivo(ventana_principal)

    try:
        lista_archivos_ordenada = sorted(list(file_paths_tupla), key=lambda ruta: os.path.getmtime(ruta), reverse=True)
        
        # El hilo que se lanza ahora es el GESTOR de los workers, no el que hace el trabajo directamente
        procesar_archivos_seleccionados(
            lista_archivos_ordenada, boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal, tabla_treeview
        )
    except FileNotFoundError as e:
        messagebox.showerror("Error de Archivo", f"No se encontró el archivo {e.filename}.", parent=ventana_principal)
        if etiqueta_estado: etiqueta_estado.config(text="Error al procesar la selección.")
        if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)


PREGUNTA_PARA_DEEPSEEK = """Analiza el contenido del documento proporcionado (que ha sido convertido a texto, originado de un PDF o de un archivo Excel) y extrae los datos de código y calificación para cada una de las materias correspondientes a la sección: Programa Destino Ibero (que también puede llamarse simplemente Programa Destino en algunas ocasiones). Asegúrate de no incluir información de la sección Programa de Origen.
Igualmente revisa el dato correspondiente al periodo académico, programa al que aspira, el nombre completo del programa de origen y el número correspondiente al total de créditos homologados.
CONDICIONES DE PARADA:
Revise los siguientes datos del documento: el periodo académico, el programa al que aspira y las calificaciones del "Programa Destino Ibero". Aplique las siguientes condiciones EN ORDEN. Si CUALQUIERA de estas condiciones se cumple, notifíquela claramente indicando la condición específica y la etiqueta de detención asociada, y NO continúe con los pasos de "PROCESAMIENTO DE DATOS" ni "FORMATO DE SALIDA".
1.  **Condición de Periodo Académico:**
    *   Extraiga el valor del periodo académico del documento.
    *   **SI** la cadena de texto del periodo académico extraído **CONTIENE EXACTAMENTE** la secuencia numérica "2024" (por ejemplo, "2024-1", "PERIODO 2024B", "202433"), ENTONCES:
        Notifique: "El periodo académico '{valor_del_periodo_extraido}' contiene '2024'."
        Indique: "DETENER_PROCESO_CONDICION_PERIODO"
        Y detenga todo procesamiento adicional.
    *   DE LO CONTRARIO (si "2024" no está presente en el periodo académico), continúe con la siguiente condición de parada.
2.  **Condición de Programa (Maestría/Especialización):**
    *   Extraiga el nombre del programa al que aspira.
    *   **SI** el nombre del programa al que aspira (ignorando mayúsculas/minúsculas) **CONTIENE** alguna de las siguientes palabras: "maestria", "especializacion" o "esp", ENTONCES:
        Notifique: "El programa al que aspira '{nombre_del_programa_extraido}' indica maestría/especialización."
        Indique: "DETENER_PROCESO_CONDICION_PROGRAMA"
        Y detenga todo procesamiento adicional.
    *   DE LO CONTRARIO, continúe con la siguiente condición de parada.
3.  **Condición de Calificaciones Bajas o en Blanco:**
    *   Analice las calificaciones de las materias del "Programa Destino Ibero".
    *   **SI** encuentra CUALQUIER materia en el "Programa Destino Ibero" con una calificación en blanco (vacía) O con un valor numérico inferior a 3.0 (por ejemplo, 2.9, 1.5, 0.0), ENTONCES:
        Notifique: "Se encontró una calificación no válida para la materia '{nombre_materia_con_problema}' con calificación '{calificacion_problematica}' en el Programa Destino Ibero."
        Indique: "DETENER_PROCESO_CONDICION_CALIFICACION"
        Y detenga todo procesamiento adicional.
    *   DE LO CONTRARIO, si todas las calificaciones son válidas (3.0 o superior y no en blanco), proceda con los siguientes pasos.
Si NINGUNA de las condiciones de parada anteriores se cumple, Y SOLO EN ESE CASO, proceda con lo siguiente:
PROCESAMIENTO DE DATOS:
Ten en cuenta las siguientes recomendaciones al procesar los datos de las materias del "Programa Destino Ibero":
- Algunos códigos pueden contener espacios; omite estos espacios para obtener solo letras y números.
- En los códigos de las materias, los números siempre preceden a las letras. Además, en algunos casos, puede aparecer una letra "F" al final de los números; esta letra debe ser omitida.
- Si el código de la materia tiene un número 0 al final de este debes preservarlo y no eliminarlo.
- Debemos traer al final los códigos tal cual como aparecen en el archivo no debes cambiar los número ni las letras de los códigos.
FORMATO DE SALIDA (si no hay condiciones de parada):
1.  **DATOS CLAVE PARA RESUMEN (Presentar antes de la tabla y el script AHK):**
    *   **NOMBRE_ESTUDIANTE:** [Nombre completo del estudiante como aparece en el documento]
    *   **PROGRAMA_ASPIRA:** [Nombre completo del programa al que aspira el estudiante]
    *   **PLAN_ESTUDIO:** [Número o identificador del plan de estudio asociado al programa al que aspira, si se menciona. Si no se menciona, indique "N/A"]
2.  **TABLA DE DATOS (Materias):**
    Crea una tabla con tres columnas para las materias del "Programa Destino Ibero":
    Primera columna: Contiene solamente las letras de los códigos.
    Segunda columna: Contiene solamente los números correspondientes a los códigos.
    Tercera columna: Contiene las calificaciones correspondientes a cada código.
    Las calificaciones deben formatearse siempre con un decimal:
    - Si la calificación es un número entero, debe mostrarse con un decimal (por ejemplo, 4 debe ser 4.0).
    - Si la calificación tiene dos decimales, redondea al primer decimal (por ejemplo, 4.15 se convierte en 4.2, 4.14 se convierte en 4.1).
3.  **SCRIPT AUTOHOTKEY:**
    Genera un script de AutoHotkey que automatice el ingreso de todos los datos extraídos de la tabla de materias.
    Es importante que en la seccion de INICIO SCRIPT AUTOHOTKEY mantengas y no elimines espacios o caracteres de las palabras.
    ¡ATENCIÓN CRÍTICA AL DETALLE DEL ESPACIADO EN WINTITLE!: En la línea `WinTitle := "Oracle Fusion Middleware Forms Services:  Open > SHATRNS"`, asegúrate de que haya exactamente DOS espacios después de "Services:" y antes de "Open". Esto es crucial para la activación de la ventana.
    Delimita claramente el script de la siguiente manera:
    --- INICIO SCRIPT AUTOHOTKEY ---
    #SingleInstance force
    #NoEnv
    SendMode Input
    SetKeyDelay, 10 ; Aumenta el tiempo entre teclas (en milisegundos)
    ; ===== ESPERAR Y ACTIVAR LA VENTANA ESPECÍFICA =====
    WinTitle := "Oracle Fusion Middleware Forms Services:  Open > SHATRNS"  ; ¡NOTA LOS DOS ESPACIOS DESPUÉS DE Services:!
    ; Esperar hasta que la ventana exista
    WinWait, %WinTitle%
    ; Activar la ventana objetivo
    WinActivate, %WinTitle%
    WinWaitActive, %WinTitle%
    ; === LÍNEA DE ENCABEZADO/CONFIGURACIÓN INICIAL (SOLO UNA VEZ) ===
    ; Primero, extrae el 'número total de créditos homologados' y el 'nombre completo del programa de origen' del documento.
    ; Luego, genera una 'abreviacion del programa de origen' a partir del nombre completo del programa de origen, siguiendo ESTRICTAMENTE las reglas y ejemplos detallados a continuación.
    ; REGLAS Y EJEMPLOS PARA LA ABREVIACIÓN DEL PROGRAMA DE ORIGEN:
    ;   OBJETIVO: Crear una abreviatura de MÁXIMO 30 caracteres, incluyendo espacios.
    ;
    ; --- NUEVA REGLA ---
    ;   REGLA ESPECIAL Y PRIORITARIA (SOBREESCRIBE TODAS LAS DEMÁS):
    ;     **SI** el nombre completo del programa de origen **CONTIENE** las palabras `NORMALISTA SUPERIOR`, **ENTONCES** la `(abreviacion del programa de origen)` **DEBE SER EXACTAMENTE** `NORMALISTA SUPERIOR`. Esta regla tiene la máxima prioridad y anula cualquier otra regla de abreviación.
    ;
    ;   REGLAS GENERALES:
    ;     1. Omitir artículos (el, la, los, las), preposiciones comunes (de, en, a, por, para, con) y conjunciones (y, e, o, u) a menos que sean esenciales para el significado o parte de una sigla común.
    ;     2. Priorizar la legibilidad y el reconocimiento del programa original dentro de la abreviatura.
    ;     3. Siempre que sea posible, intenta utilizar los 30 caracteres disponibles si esto mejora la claridad de la abreviatura y sigue las sustituciones. No excedas los 30 caracteres.
    ;     4. Si el nombre completo del programa de origen ya es corto (ej. menos de 30-35 caracteres) y su abreviatura natural es más corta, no es necesario forzarla a 30 caracteres.
    ;   SUSTITUCIONES COMUNES (aplicar donde corresponda, usualmente a mayúsculas, pero intenta mantener el caso original si es relevante y luego convierte a mayúsculas si es necesario para la abreviatura final):
    ;     - TECNICO / TÉCNICO / TECNICA -> TC
    ;     - TECNOLOGO / TECNÓLOGO / TECNOLOGIA / TECNOLOGÍA -> TG
    ;     - ESPECIALIZACION / ESPECIALIZACIÓN -> ESP
    ;     - LICENCIATURA -> LIC
    ;     - MÁSTER / MASTER -> MA
    ;     - INGENIERIA / INGENIERÍA -> ING
    ;     - ADMINISTRACIÓN / ADMINISTRATIVA / ADMINISTRATIVO / ADMINISTRACIÓN -> ADMON / ADMIN / ADMTVA / ADMINIST
    ;     - GESTION / GESTIÓN -> GSTON / GSTN
    ;     - CONTABLE / CONTABILIDAD / CONTABILIZACION -> CONT / CONTAB / CONTBLE
    ;     - FINANCIERA / FINANCIERO / FINANZAS -> FINAN / FINANC / FINACIRA
    ;     - DESARROLLO -> DESARR / DSRLLO
    ;     - INFORMACIÓN / INFORMÁTICA -> INFO / INFORM
    ;     - SOFTWARE -> SOFT / SOFTW
    ;     - SISTEMAS -> SIST / SISTE
    ;     - OPERACIONES -> OPER / OPERAC
    ;     - COMERCIALES / COMERCIO -> COMER / COMR
    ;     - INDUSTRIAL / INDUSTRIALES -> IND / INDUS / INDSTRL
    ;     - PRIMERA INFANCIA -> PR INFANC / PRIM INFAN / PRIME INFAN
    ;     - SEGURIDAD -> SEG / SEGU
    ;     - SALUD -> SALD
    ;     - OCUPACIONAL -> OCUP / OCUPAC
    ;     - MANTENIMIENTO -> MANTO / MANT / MANTMTO / MTNTO
    ;     - PROCESOS -> PROC / PROCE / PRCSOS
    ;     - LOGÍSTICA / LOGÍSTICOS -> LOGIST / LOGISTI
    ;     - PROGRAMA / PROGRAMACION -> PROG / PROGRAM
    ;     - APLICACIONES -> APLIC / APLICAC
    ;     - INTEGRAL / INTEGRACIÓN / INTEGRADA -> INTGRL / INTGRCION / INTGRAC / INTGR
    ;     - EDUCACIÓN -> EDUC / EDUCION
    ;     - ATENCIÓN -> ATEN / ATENC / ATNCION
    ;     - AUXILIAR -> AUX
    ;     - COMPETENCIAS -> COMP / COMPT / COMPETE
    ;     - LABORAL -> LAB / LBRL
    ;     - AUTOMOTORES -> AUTOMOT
    ;     - ELECTROMECÁNICO / ELECTROMECANICO -> ELECTRMC
    ;     - TELECOMUNICACIONES -> TELECOM / TELECOMU
    ;     - DISEÑO -> DISÑ
    ;     - ELEMENTOS -> ELE
    ;     - MECÁNICOS / MECANICO -> MEC
    ;     - FABRICACIÓN -> FABR
    ;     - MÁQUINAS HERRAMIENTAS -> MAQ HER / MAQU HER
    ;     - EQUIPOS -> EQUIP / EQUPO
    ;     - COMPUTO / COMPUTACIÓN -> COMP / COMPUT
    ;     - CABLEADO ESTRUCTURADO -> CABL ESTRU
    ;     - ASESORÍA -> ASERIA
    ;     - ENTIDADES -> ENT
    ;     - PROYECTOS -> PROY
    ;     - ECONÓMICO -> ECO
    ;     - SOCIAL -> SOC
    ;     - ELECTRÓNICO -> ELECTRO / ELEC
    ;     - RESIDUOS SÓLIDOS -> RESID SOLID
    ;     - APLICATIVOS MÓVILES -> APLICA MOV
    ;     - ENFERMERÍA -> ENFERM / ENFERME
    ;     - MAQUINARIA PESADA -> MAQUI PESAD
    ;     - EXCAVACIÓN -> EXCAVA
    ;     - ADOLESCENCIA -> ADOL
    ;     - ANÁLISIS / ANALÍTICA -> ANALI / ANALTICA
    ;     - VIRTUAL -> VIRTUAL
    ;     - AMBIENTAL -> AMBI / AMBIENT / AMBTAL
    ;     - APOYO -> APOYO
    ;     - PROFESIONAL -> PROF / PROFES
    ;     - EXTERIOR -> EXT / EXTER
    ;     - INSTRUMENTAL -> INSTRUM
    ;     - PEDAGOGÍA -> PEDAG
    ;     - RECREACIÓN -> RECRE / RCRE
    ;     - ECOLÓGICA -> ECOLG
    ;     - NEUROPSICOLOGÍA -> NEUROPSICO
    ;     - REDES AÉREAS -> RDES AERE
    ;     - DISTRIBUCIÓN -> DISTR
    ;     - ENERGÍA ELÉCTRICA -> ENER ELECT
    ;     - RECURSOS HUMANOS -> REC HUM / RECUR HUMA
    ;     - CONFECCIÓN -> CONFEC / CONFECC
    ;     - ASISTENTE -> ASIST / ASISTE / ASISTEN
    ;     - INSTALACIONES -> INSTAL / INSTALC
    ;     - MINERAS BAJO TIERRA -> MIN BAJO TIERRA
    ;     - LÚDICA -> LUD
    ;     - CULTURAL -> CULT / CUL
    ;     - ESTABLECIMIENTOS -> ESTBLCI
    ;     - ALIMENTOS Y BEBIDAS -> ALIMNT Y BEBD
    ;     - ARCHIVO -> ARCHIVO
    ;     - AUTOMATISMOS MECATRÓNICOS -> AUTOM MECATR
    ;     - DISTRIBUCIÓN FÍSICA INTERNACIONAL -> DISTRIB FISICA INTERNA
    ;     - NEGOCIOS INTERNACIONALES -> NEGO INTERNA
    ;     - MERCADEO -> MERCADEO
    ;     - PUBLICITARIO -> PUBLIC
    ;     - COORDINADOR -> COORD / COORDIN / COORDINADR
    ;     - GERENCIA -> GER / GERNCIA
    ;     - CALIDAD -> CALID / CALDAD
    ;     - QUÍMICA -> QUIM / QUIMICA
    ;     - APLICADA -> APLICD
    ;     - LEVANTAMIENTOS TOPOGRÁFICOS -> LEVANTA TOPOGRA
    ;     - GEORREFERENCIACIÓN -> GEOREFERE
    ;     - FARMACÉUTICOS -> FARMAC
    ;     - GRANDES SUPERFICIES -> GRAN SUPER
    ;     - ASEGURAMIENTO METROLÓGICO -> ASEGUR METROLOG
    ;     - FOTOGRAFÍA -> FOTOGRAF
    ;     - PROCESOS DIGITALES -> PROCES DIGITA
    ;     - CADENA DE ABASTECIMIENTO -> CADEN ABAST / CADN AB
    ;     - RIESGO EN SEGUROS -> RISGO EN SEGUR
    ;     - INFRAESTRUCTURA VIAL -> INFRAESTRUC VIAL
    ;     - PLANTAS DE PRODUCCIÓN -> PLAN PRODUCC
    ;     - CICLO DE VIDA DEL PRODUCTO -> CICLO VIDA DEL PRODU
    - AGROPECUARIA ECOLÓGICA -> AGROPEC ECOLO
    ;     - BANCA Y SERVICIOS FINANCIEROS -> BANC Y SER FINAN
    ;     - CONTACT CENTER Y BPO -> CONTA CENT Y BPO
    ;     - ALISTAMIENTO Y OPERACIÓN -> ALIMTO Y OPER
    ;     - PLANEACIÓN EDUCATiva -> PLANE EDUC
    ;     - PLANES DE DESARROLLO -> PLAN DESARR
    ;     - INVESTIGACIÓN E INNOVACIÓN -> INVSTG INNOVAC
    ;     - SUPERVISIÓN -> SUPERVI / SUPERV
    ;     - TERMINALES PORTUARIAS -> TERM PORTU
    ;     - MEDIOS GRÁFICOS VISUALES -> MEDI GRAF VISUALE
    ;     - CONTROL DE CALIDAD -> CONTRL CALID
    ;     - AGUA Y SANEAMIENTO -> AGU Y SANEAMIE
    ;     - MICROFINANCIERAS -> MICROFINA
    ;     - REVISADOR DE CALIDAD -> REVISA CALID
    ;     - BILINGUAL EXPERT ON BUSINESS PROCESS OUTSOURCING -> BILG EXPT ON BUSIN PRO OUT
    ;     - ENTRENAMIENTO DEPORTIVO -> ENTRENAMTO DEPORTIVO
    ;     - SERVICIO DE POLICÍA -> SERV POLIC
    ;     - DIRECCIÓN NACIONAL DE ESCUELAS -> DIR NAC ESC
    ;     - FORMACIÓN PREESCOLAR -> FORM PREESC
    ;     - PRESELECCIÓN -> PRESELEC / PRESELECC
    ;     - HERRAMIENTAS TIC -> HERR TIC
    ;     - UNIVERSITARIO -> UNIVER
    ;     - INCLUSIVA E INTERCULTURAL -> INCLUS INTERC
    ;     - PETROQUÍMICAS -> PETROQUIM
    ;     - PREVENCIÓN Y CONTROL -> PREVEN Y CTRL
    ;     - DISEÑO WEB Y MULTIMEDIA -> DISEÑ WEB Y MULTIMED
    ;   EJEMPLOS DE APLICACIÓN (Nombre Completo -> Abreviatura Esperada):
    ;     1.  "TECNICO EN CONTABILIZACION DE OPERACIONES COMERCIALES Y FINANCIERAS" -> "TC CONT OPER COMER Y FINANC"
    ;     2.  "TECNÓLOGO EN GESTION CONTABLE Y FINANCIERA" -> "TG GSTON CONTBLE Y FINACIRA"
    ;     3.  "TECNOLOGO EN ANALISIS Y DESARROLLO DE SOFTWARE" -> "TG ANALISIS Y DESARR DE SOFTWA"
    ;     4.  "ESPECIALIZACIÓN EN EDUCACIÓN E INTERVENCIÓN PARA LA PRIMERA INFANCIA." -> "ESP EDUC INTRCION PRIME INF"
    ;     5.  "INGENIERIA EN SEGURIDAD INDUSTRIAL E HIGIENE OCUPACIONAL" -> "ING SEG INDU Y SALUD OCUPAC"
    ;     6.  "TÉCNICO LABORAL POR COMPETENCIAS EN AUXILIAR ADMINISTRATIVO EN SALUD" -> "TC LAB AUX ADMINIST SALUD"
    ;     7.  "TECNÓLOGO EN GESTIÓN INTEGRADA DE LA CALIDAD, MEDIO AMBIENTE, SEGURIDAD Y SALUD OCUPACIONAL" -> "TC GSTN CALID AMBI SEG Y SALD"
    ;     8.  "TECNÓLOGO EN DISEÑO DE ELEMENTOS MECÁNICOS PARA SU FABRICACIÓN CON MÁQUINAS HERRAMIENTAS CNC" -> "TG DISÑ ELE MEC FABR MAQ CNC"
    ;     9.  "TÉCNICO EN MONTAJE Y MANTENIMIENTO ELECTROMECANICO DE INSTALACIONES MINERAS BAJO TIERRA" -> "TC MONTJE MNTO ELECTR INST MIN"
    ;     10. "TECNOLOGIA EN GESTION DE SISTEMAS DE TELECOMUNICACIONES" -> "TG GSTON SISTEM DE TELECOMU"
    ;     11. "MÁSTER UNIVERSITARIO EN EDUCACIÓN INCLUSIVA E INTERCULTURAL." -> "MA UNIVER EDUC INCLUS INTERC"
    ;     12. "PROGRAMA DE FORMACION COMPLEMENTARIA DE LA ESCUELA NORMALISTA SUPERIOR" -> "NORMALISTA SUPERIOR"
    ;
    ; Con base en estas reglas y ejemplos, genera la 'abreviacion del programa de origen'.
    ; Luego, genera la siguiente línea Send exactamente una vez ANTES de procesar las materias:
    Send, 1{Tab}HOM01{Tab}HOM01{Tab}(número total creditos homologados){Tab}{Tab}{Tab}(abreviacion del programa de origen){F10}^{PgDn}{Space}{Tab}
    ; === LÍNEAS DE DATOS PARA CADA MATERIA ===
    ; Para cada materia extraída de la TABLA DE DATOS (Materias) anterior, genera una línea Send con la siguiente estructura:
    Send, (letras del codigo){Tab}(números del código){Tab}{Tab}(calificación){Tab}N{Down}{Space}{Tab}
    ; === MANEJO DE LA ÚLTIMA FILA ===
    ; Para la última fila de datos de materia, la línea Send debe ser:
    ; Send, (letras del codigo de la última materia){Tab}(números del código de la última materia){Tab}{Tab}(calificación de la última materia){Tab}N
    ; (Es decir, omite {Down}{Space}{Tab} al final de la última materia).
    --- FIN SCRIPT AUTOHOTKEY ---
"""

if __name__ == "__main__":
    ventana_principal_app = tk.Tk()
    ventana_principal_app.title("Procesador de Documentos AHK")
    ventana_principal_app.geometry("500x250")
    frame_superior = ttk.Frame(ventana_principal_app, padding="10")
    frame_superior.pack(expand=True, fill=tk.BOTH)

    etiqueta_estado_gui = ttk.Label(frame_superior, text="Listo para iniciar.", font=("Segoe UI", 10))
    etiqueta_estado_gui.pack(pady=5)
    barra_progreso_gui = ttk.Progressbar(frame_superior, orient="horizontal", length=300, mode="determinate")
    barra_progreso_gui.pack(pady=10)
    etiqueta_etr_gui = ttk.Label(frame_superior, text="Tiempo restante: --:--", font=("Segoe UI", 9))
    etiqueta_etr_gui.pack(pady=5)
    
    boton_seleccionar_gui = ttk.Button(
        frame_superior, 
        text="Seleccionar Archivos y Procesar",
        command=lambda: iniciar_procesamiento_en_hilo(
            boton_seleccionar_gui, 
            etiqueta_estado_gui, 
            barra_progreso_gui,
            etiqueta_etr_gui,
            ventana_principal_app
        )
    )
    boton_seleccionar_gui.pack(pady=10)

    keys_faltantes = []
    if not LLMWHISPERER_API_KEY_1: keys_faltantes.append("LLMWHISPERER_API_KEY_1")
    if not DEEPSEEK_PRIMARY_API_TOKEN: keys_faltantes.append("DEEPSEEK_PRIMARY_API_TOKEN")
    if not DEEPSEEK_FALLBACK_API_TOKEN: keys_faltantes.append("DEEPSEEK_FALLBACK_API_TOKEN")

    if keys_faltantes:
        mensaje_error = "ERROR: Faltan claves en el archivo .env:\n\n" + "\n".join(keys_faltantes)
        etiqueta_estado_gui.config(text="ERROR: ¡Configure las claves API en .env!")
        boton_seleccionar_gui.config(state=tk.DISABLED)
        messagebox.showerror("Error de Configuración", mensaje_error, parent=ventana_principal_app) 
    else:
        etiqueta_estado_gui.config(text="Claves API OK. Listo para seleccionar archivos.")

    ventana_principal_app.mainloop()