import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
# Para PDF con LLMWhisperer
from unstract.llmwhisperer import LLMWhispererClientV2
from unstract.llmwhisperer.client_v2 import LLMWhispererClientException
# Para Excel
import pandas as pd
# Comunes
import requests
import json
import re
import os
import traceback
import unicodedata 
import threading 
import time

# --- Configuración de APIs ---
LLMWHISPERER_API_KEY_1 = "EJkhlXLOq97DOJky3NlG7uiNHz3TRaIG1uQnSH7U8Rg"
LLMWHISPERER_API_KEY_2 = "Gu68hhTHA7OhurrXRct8MaZT8UgkaSq_XN9N4kvlnYU"
CHUTES_API_TOKEN = "cpk_f93ac3cb60654df98d05a23657613b0a.1ed5a86f4f755f1ab10a4afd4b0cc914.KYyQjp2L2dylm5V4UtDTVlZAmz1SPUc7"

DEEPSEEK_MODEL = "deepseek-ai/DeepSeek-V3-0324"
DEEPSEEK_API_URL = "https://llm.chutes.ai/v1/chat/completions"

# --- Listas y Bandera Global ---
mensajes_resumen_procesamiento = []
datos_para_tabla_resumen = [] 
active_api_key_index = 1

# --- Funciones de Normalización y Extracción ---
def quitar_tildes(texto):
    if not isinstance(texto, str):
        return texto
    nfkd_form = unicodedata.normalize('NFD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def seleccionar_archivos_soportados():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Selecciona uno o más archivos PDF o Excel para analizar",
        filetypes=(
            ("Documentos Soportados", "*.pdf *.xlsx *.xls"),
            ("PDF files", "*.pdf"),
            ("Excel files (*.xlsx, *.xls)", "*.xlsx;*.xls"),
            ("All files", "*.*")
        )
    )
    root.destroy()
    return file_paths

def _intentar_extraccion_llmwhisperer(file_path, api_key, key_number):
    """Función auxiliar que realiza una única llamada a la API."""
    nombre_base = os.path.basename(file_path)
    print(f"Inicializando cliente LLMWhisperer para PDF '{nombre_base}' con API Key #{key_number}...")
    client = LLMWhispererClientV2(api_key=api_key)
    print(f"Enviando PDF '{file_path}' a LLMWhisperer (Key #{key_number})...")
    resultado = client.whisper(file_path=file_path, wait_for_completion=True, wait_timeout=360)
    print(f"\n✅ Resultado del análisis de LLMWhisperer obtenido para PDF '{nombre_base}'.")

    if isinstance(resultado, dict) and 'extraction' in resultado and \
       isinstance(resultado['extraction'], dict) and 'result_text' in resultado['extraction']:
        texto_extraido = resultado['extraction']['result_text']
        return texto_extraido.replace("<<<\x0c", "\n\n--- NUEVA PÁGINA --- \n\n") 
    else:
        msg = f"❌ Error LLMWhisperer PDF '{nombre_base}': No se encontró 'extraction' o 'result_text'."
        print(msg)
        mensajes_resumen_procesamiento.append(msg)
        return None

def extraer_texto_pdf_con_apis(file_path):
    global mensajes_resumen_procesamiento, active_api_key_index

    # Determinar qué clave usar
    if active_api_key_index == 1:
        current_key = LLMWHISPERER_API_KEY_1
        current_key_num = 1
    else: # Ya estamos en la clave de respaldo
        current_key = LLMWHISPERER_API_KEY_2
        current_key_num = 2
        print(f"INFO: Usando directamente la API Key de respaldo #{current_key_num}...")

    try:
        # Intento con la clave activa actual
        return _intentar_extraccion_llmwhisperer(file_path, current_key, current_key_num)
    except LLMWhispererClientException as e:
        msg_error_api = str(e).lower()
        codigo_estado_api = e.status_code if hasattr(e, 'status_code') else 0

        if (codigo_estado_api == 402 or "breached your free processing limit" in msg_error_api) and active_api_key_index == 1:
            print(f"⚠️ LÍMITE ALCANZADO en API Key #1. Cambiando a API Key #2 para '{os.path.basename(file_path)}' y subsiguientes.")
            mensajes_resumen_procesamiento.append(f"⚠️ LÍMITE API Unstract #1. Cambiando a API #2.")
            active_api_key_index = 2 
            try:
                # Reintentar para el archivo actual con la clave 2
                return _intentar_extraccion_llmwhisperer(file_path, LLMWHISPERER_API_KEY_2, 2)
            except Exception as e2:
                msg = f"❌ Fallo en el reintento con API Key #2 para '{os.path.basename(file_path)}': {e2}"
                print(msg); traceback.print_exc(); mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
                return None
        elif (codigo_estado_api == 402 or "breached your free processing limit" in msg_error_api) and active_api_key_index == 2:
            msg = f"❌ LÍMITE ALCANZADO también en API Key #2. No hay más claves disponibles."
            print(msg); mensajes_resumen_procesamiento.append(msg)
            return None
        else:
            msg = f"❌ Error con LLMWhisperer para PDF '{os.path.basename(file_path)}': {e} (Código: {codigo_estado_api})"
            print(msg); mensajes_resumen_procesamiento.append(msg)
            return None
            
    except Exception as e:
        msg = f"❌ Error inesperado durante el análisis con LLMWhisperer para PDF '{os.path.basename(file_path)}': {e}"
        print(msg); traceback.print_exc(); mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

def extraer_texto_excel_con_pandas(file_path):
    global mensajes_resumen_procesamiento
    nombre_base = os.path.basename(file_path)
    print(f"Procesando archivo Excel '{nombre_base}' con pandas...")
    try:
        xls = pd.ExcelFile(file_path)
        full_text_parts = []
        if not xls.sheet_names:
            msg = f"⚠️ El archivo Excel '{nombre_base}' no contiene hojas."
            print(msg); mensajes_resumen_procesamiento.append(msg)
            return ""

        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, header=None, dtype=str).fillna('') 
            full_text_parts.append(f"--- INICIO HOJA: {sheet_name} ---\n")
            for index, row in df.iterrows():
                full_text_parts.append(" | ".join(str(cell).strip() for cell in row))
            full_text_parts.append(f"\n--- FIN HOJA: {sheet_name} ---\n")
        
        print(f"✅ Texto extraído del archivo Excel '{nombre_base}'.")
        return "\n".join(full_text_parts)
    except Exception as e:
        msg = f"❌ Error al procesar Excel '{nombre_base}' con pandas: {e}"
        print(msg); traceback.print_exc(); mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

# --- Funciones de DeepSeek API y Procesamiento de Respuesta ---
def limpiar_texto_general_para_llm(texto):
    if not texto: return ""
    return re.sub(r'\n{3,}', '\n\n', texto).strip()

def analizar_con_deepseek(texto_documento, pregunta_o_instruccion, api_token, model_name, api_url, nombre_archivo_base):
    global mensajes_resumen_procesamiento
    print(f"\nPreparando para enviar a la API de DeepSeek para '{nombre_archivo_base}'...")
    prompt_completo = f"Aquí tienes el contenido de un documento (originalmente '{nombre_archivo_base}'):\n--- INICIO DEL DOCUMENTO ---\n{texto_documento}\n--- FIN DEL DOCUMENTO ---\n\nPor favor, sigue estas instrucciones detalladamente basadas ÚNICAMENTE en el documento proporcionado:\n{pregunta_o_instruccion}"
    headers = {"Authorization": f"Bearer {api_token}", "Content-Type": "application/json"}
    data = {"model": model_name, "messages": [{"role": "user", "content": prompt_completo}], "stream": False, "max_tokens": 4000, "temperature": 0.15}
    try:
        print(f"Enviando solicitud a DeepSeek (modelo: {model_name}) para '{nombre_archivo_base}'...")
        response = requests.post(api_url, headers=headers, data=json.dumps(data), timeout=360) 
        response.raise_for_status()
        respuesta_json = response.json()
        print(f"✅ Respuesta recibida de DeepSeek para '{nombre_archivo_base}'.")
        if respuesta_json.get("choices") and len(respuesta_json["choices"]) > 0 and respuesta_json["choices"][0].get("message") and "content" in respuesta_json["choices"][0]["message"]:
            return respuesta_json["choices"][0]["message"]["content"]
        else:
            msg = f"❌ Error DeepSeek '{nombre_archivo_base}': Respuesta sin formato esperado."
            print(msg); mensajes_resumen_procesamiento.append(msg)
            return None
    except requests.exceptions.Timeout:
        msg = f"❌ Error DeepSeek API '{nombre_archivo_base}': Timeout."
        print(msg); mensajes_resumen_procesamiento.append(msg)
        return None
    except requests.exceptions.RequestException as e:
        msg = f"❌ Error DeepSeek API '{nombre_archivo_base}': {e}"
        print(msg); 
        if e.response is not None: 
            print(f"Detalles: {e.response.status_code} - {e.response.text}")
            mensajes_resumen_procesamiento.append(msg + (f"Detalles: {e.response.status_code} - {e.response.text}" if e.response is not None else ""))
        return None
    except json.JSONDecodeError:
        msg = f"❌ Error decodificando JSON de DeepSeek para '{nombre_archivo_base}'."
        print(msg); mensajes_resumen_procesamiento.append(msg)
        return None
    except Exception as e:
        msg = f"❌ Error inesperado con DeepSeek para '{nombre_archivo_base}': {e}"
        print(msg); traceback.print_exc(); mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

def ajustar_abreviacion(texto_abreviado, limite_caracteres=30):
    global mensajes_resumen_procesamiento 
    abreviatura_original_copia = texto_abreviado 
    if len(texto_abreviado) > limite_caracteres:
        texto_ajustado = texto_abreviado[:limite_caracteres].rstrip() 
        log_msg = f"INFO: Abreviación original: '{abreviatura_original_copia}' ({len(abreviatura_original_copia)}), truncada a: '{texto_ajustado}' ({len(texto_ajustado)})"
        print(log_msg); mensajes_resumen_procesamiento.append(log_msg)
        return texto_ajustado
    else:
        return texto_abreviado

def extraer_script_ahk_y_ajustar_abreviacion(respuesta_deepseek):
    global mensajes_resumen_procesamiento
    try:
        match_script = re.search(r"--- INICIO SCRIPT AUTOHOTKEY ---(.*?)--- FIN SCRIPT AUTOHOTKEY ---", respuesta_deepseek, re.DOTALL)
        if not match_script:
            print("❌ No se encontraron los delimitadores del script AutoHotkey en la respuesta.")
            mensajes_resumen_procesamiento.append("❌ No se encontraron delimitadores AHK.")
            return None

        script_content_bruto = match_script.group(1).strip()
        lineas_script = script_content_bruto.splitlines()
        script_final_ajustado = []
        linea_encabezado_modificada = False
        
        abreviacion_llm = "ABREV_POR_DEFECTO"; creditos_llm = "0" 
        patron_linea_encabezado = r"Send,\s*1{Tab}HOM01{Tab}HOM01{Tab}(.*?){Tab}{Tab}{Tab}(.*?){F10}\^\{PgDn\}\{Space\}\{Tab\}"

        for i, linea in enumerate(lineas_script):
            if i == 0 and not linea.startswith("#SingleInstance force"):
                print("⚠️ ADVERTENCIA: Script AHK extraído, pero NO comienza con '#SingleInstance force'.")
                mensajes_resumen_procesamiento.append("⚠️ AHK no inicia con #SingleInstance force.")
            
            match_encabezado = re.match(patron_linea_encabezado, linea)
            if match_encabezado and not linea_encabezado_modificada:
                creditos_extraidos_str = match_encabezado.group(1).strip()
                abreviacion_extraida_str = match_encabezado.group(2).strip()

                creditos_llm = creditos_extraidos_str.replace("(número total creditos homologados)", "").strip()
                if not creditos_llm.isdigit(): creditos_llm = "0" 

                abreviacion_llm = abreviacion_extraida_str.replace("(abreviacion del programa de origen)", "").strip()
                if not abreviacion_llm: abreviacion_llm = "ABREV_NO_PROPORCIONADA"
                
                mensajes_resumen_procesamiento.append(f"INFO (AHK): Abreviación LLM: '{abreviacion_llm}', Créditos LLM: '{creditos_llm}'")

                abreviacion_sin_tildes = quitar_tildes(abreviacion_llm)
                abreviacion_ajustada = ajustar_abreviacion(abreviacion_sin_tildes, 30) 
                
                linea_modificada = f"Send, 1{{Tab}}HOM01{{Tab}}HOM01{{Tab}}{creditos_llm}{{Tab}}{{Tab}}{{Tab}}{abreviacion_ajustada}{{F10}}^{{PgDn}}{{Space}}{{Tab}}"
                script_final_ajustado.append(linea_modificada)
                linea_encabezado_modificada = True
                mensajes_resumen_procesamiento.append(f"INFO (AHK): Línea encabezado AHK ajustada con abrev: '{abreviacion_ajustada}'.")
            else:
                linea_corregida = linea.replace('WinTitle := "Oracle Fusion Middleware Forms Services: Open > SHATRNS"', 'WinTitle := "Oracle Fusion Middleware Forms Services:  Open > SHATRNS"')
                linea_corregida = re.sub(r'(WinTitle\s*:=\s*"Oracle Fusion Middleware Forms Services:)\s+(Open\s*>\s*SHATRNS")', r'\1  \2', linea_corregida)
                script_final_ajustado.append(linea_corregida)
        
        if not linea_encabezado_modificada:
            print("⚠️ ADVERTENCIA: No se encontró o no se pudo procesar la línea de encabezado Send del AHK como se esperaba.")
            mensajes_resumen_procesamiento.append("⚠️ No se procesó línea encabezado AHK.")

        script_final_str = "\n".join(script_final_ajustado)
        print("✅ Script AHK extraído y procesado.")
        return script_final_str

    except Exception as e:
        print(f"❌ Error al extraer y ajustar script AHK: {e}"); traceback.print_exc(); mensajes_resumen_procesamiento.append(f"❌ Error extrayendo/ajustando AHK: {e}\n{traceback.format_exc()}"); return None

def extraer_datos_clave_para_tabla(respuesta_deepseek):
    def limpiar_valor(texto):
        return texto.strip().strip('*').strip()

    nombre_estudiante = "No extraído"; programa_aspira = "No extraído"; plan_estudio = "No extraído"

    match_nombre = re.search(r"NOMBRE_ESTUDIANTE:\s*(.*)", respuesta_deepseek)
    if match_nombre:
        nombre_estudiante = limpiar_valor(match_nombre.group(1).strip().splitlines()[0])
    else: 
        match_nombre_fb = re.search(r"El nombre de la estudiante es\s*([A-Za-z\sÁÉÍÓÚÑáéíóúñ]+)\.?$", respuesta_deepseek, re.MULTILINE | re.IGNORECASE)
        if match_nombre_fb: nombre_estudiante = limpiar_valor(match_nombre_fb.group(1).strip().splitlines()[0])

    match_programa = re.search(r"PROGRAMA_ASPIRA:\s*(.*)", respuesta_deepseek)
    if match_programa:
        programa_aspira = limpiar_valor(match_programa.group(1).strip().splitlines()[0])
    
    match_plan = re.search(r"PLAN_ESTUDIO:\s*(.*)", respuesta_deepseek)
    if match_plan:
        plan_estudio = limpiar_valor(match_plan.group(1).strip().splitlines()[0])
        if not plan_estudio: plan_estudio = "N/A" 
    else:
        plan_estudio = "N/A" 

    print(f"INFO (Tabla Resumen): Nombre: '{nombre_estudiante}', Programa: '{programa_aspira}', Plan: '{plan_estudio}'")
    return {"Nombre": nombre_estudiante, "Programa": programa_aspira, "Plan": plan_estudio}


def sanitizar_nombre_archivo(nombre):
    nombre_sanitizado = re.sub(r'[<>:"/\\|?*]', '', nombre).replace("\n", " ").replace("\r", " ")
    nombre_sanitizado = re.sub(r'\s+', ' ', nombre_sanitizado).strip()
    return nombre_sanitizado if nombre_sanitizado else "ScriptSinNombreValido"

# --- GUI Functions ---
def mostrar_resumen_log_gui(main_root_ref, lista_mensajes): 
    if not lista_mensajes: lista_mensajes = ["No se procesaron archivos o no hubo mensajes de resumen."]
    resumen_ventana = tk.Toplevel(main_root_ref) 
    resumen_ventana.title("Resumen del Procesamiento de Archivos (Log)")
    window_width = 800; window_height = 500 
    screen_width = resumen_ventana.winfo_screenwidth(); screen_height = resumen_ventana.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2); center_y = int(screen_height/2 - window_height / 2 - 60) 
    resumen_ventana.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    txt_area = scrolledtext.ScrolledText(resumen_ventana, wrap=tk.WORD, width=100, height=25)
    txt_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
    for mensaje in lista_mensajes: txt_area.insert(tk.END, mensaje + "\n--------------------------------------------------\n")
    txt_area.config(state=tk.DISABLED)
    btn_cerrar_log = tk.Button(resumen_ventana, text="Cerrar Log", command=resumen_ventana.destroy)
    btn_cerrar_log.pack(pady=10)
    resumen_ventana.protocol("WM_DELETE_WINDOW", resumen_ventana.destroy)
    resumen_ventana.transient(main_root_ref) 
    resumen_ventana.grab_set() 
    main_root_ref.wait_window(resumen_ventana) 

def mostrar_tabla_resumen_gui(main_root_ref, datos_tabla): 
    if not datos_tabla:
        print("No hay datos para mostrar en la tabla de resumen.")
        mensajes_resumen_procesamiento.append("INFO: No se generaron datos para la tabla de resumen.")
        return

    datos_tabla_ordenados = sorted(datos_tabla, key=lambda item: item.get("Programa", "").lower())

    tabla_ventana = tk.Toplevel(main_root_ref)
    tabla_ventana.title("Tabla Resumen de Archivos Procesados")
    window_width = 1050; window_height = 450
    screen_width = tabla_ventana.winfo_screenwidth(); screen_height = tabla_ventana.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2); center_y = int(screen_height/2 - window_height / 2 + 60)  
    tabla_ventana.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    frame = ttk.Frame(tabla_ventana, padding="10"); frame.pack(fill=tk.BOTH, expand=True)

    cols = ("Archivo Origen", "Nombre del estudiante", "Programa al que aspira", "Plan", "Acciones")
    tree = ttk.Treeview(frame, columns=cols, show='headings')

    for col in cols:
        tree.heading(col, text=col)
        if col == "Archivo Origen": tree.column(col, width=200, minwidth=150, anchor='w')
        elif col == "Nombre del estudiante": tree.column(col, width=200, minwidth=150, anchor='w')
        elif col == "Programa al que aspira": tree.column(col, width=250, minwidth=200, anchor='w') 
        elif col == "Plan": tree.column(col, width=100, minwidth=80, anchor='center')
        elif col == "Acciones": tree.column(col, width=100, minwidth=80, anchor='center')

    tree.pack(fill=tk.BOTH, expand=True)

    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview); vsb.pack(side='right', fill='y')
    tree.configure(yscrollcommand=vsb.set)

    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview); hsb.pack(side='bottom', fill='x')
    tree.configure(xscrollcommand=hsb.set)

    # Insertar filas
    for idx, item in enumerate(datos_tabla_ordenados):
        valores_fila = (
            item.get("Archivo Origen", "N/A"), 
            item.get("Nombre", "N/A"), 
            item.get("Programa", "N/A"), 
            item.get("Plan", "N/A"),
            "Abrir Archivo"
        )
        tree.insert("", "end", values=valores_fila, tags=(item.get("Ruta Completa", ""),))

    def on_click(event):
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = tree.identify_column(event.x)
        rowid = tree.identify_row(event.y)
        if column == '#5':  # Columna de acciones
            ruta_completa = tree.item(rowid, 'tags')[0]
            if ruta_completa and os.path.exists(ruta_completa):
                try:
                    if os.name == 'nt':  # Windows
                        os.startfile(ruta_completa)
                    elif os.name == 'posix':
                        subprocess.run(['xdg-open', ruta_completa], check=True)  # Linux
                    else:
                        subprocess.run(['open', ruta_completa], check=True)  # macOS
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    tree.bind("<Button-1>", on_click)

    btn_cerrar_tabla_final = tk.Button(tabla_ventana, text="Cerrar Todo y Salir", command=main_root_ref.destroy) 
    btn_cerrar_tabla_final.pack(pady=10)

    tabla_ventana.protocol("WM_DELETE_WINDOW", main_root_ref.destroy) 
    tabla_ventana.transient(main_root_ref)
    tabla_ventana.grab_set()
    main_root_ref.wait_window(tabla_ventana)

def procesar_archivos_seleccionados(lista_archivos, boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal):
    global mensajes_resumen_procesamiento, datos_para_tabla_resumen, active_api_key_index
    
    mensajes_resumen_procesamiento = []; datos_para_tabla_resumen = []
    
    # Resetear el índice de la clave si las claves están configuradas
    if LLMWHISPERER_API_KEY_1 != "TU_LLMWHISPERER_API_KEY_1":
        active_api_key_index = 1
    else: # Si las claves son placeholders, no se usará la API
        active_api_key_index = 3 # Un índice inválido para no intentar

    num_archivos = len(lista_archivos)
    tiempos_procesamiento = [] 
    if barra_progreso: barra_progreso['maximum'] = num_archivos; barra_progreso['value'] = 0
    if etiqueta_etr: etiqueta_etr.config(text="Tiempo restante: Calculando...")
    msg_seleccion = f"Procesando {num_archivos} archivos..."
    print(msg_seleccion); mensajes_resumen_procesamiento.append(msg_seleccion.strip())
    if etiqueta_estado: etiqueta_estado.config(text=msg_seleccion)

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
    ;     - AGROPECUARIA ECOLÓGICA -> AGROPEC ECOLO
    ;     - BANCA Y SERVICIOS FINANCIEROS -> BANC Y SER FINAN
    ;     - CONTACT CENTER Y BPO -> CONTA CENT Y BPO
    ;     - ALISTAMIENTO Y OPERACIÓN -> ALIMTO Y OPER
    ;     - PLANEACIÓN EDUCATIVA -> PLANE EDUC
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
    ;     9.  "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR EN EDUCACIÓN PARA LA PRIMERA INFANCIA" -> "TC LAB COMPT AUX EDU PRIM INF"
    ;     10. "ESPECIALIZACIÓN EN GERENCIA DE LA CALIDAD EN SALUD" -> "ESP GERNCIA DE CALDAD EN SALUD"
    ;     11. "TÉCNICO EN MONTAJE Y MANTENIMIENTO ELECTROMECANICO DE INSTALACIONES MINERAS BAJO TIERRA" -> "TC MONTJE MNTO ELECTR INST MIN"
    ;     12. "TECNOLOGIA EN GESTION DE SISTEMAS DE TELECOMUNICACIONES" -> "TG GSTON SISTEM DE TELECOMU"
    ;     13. "MÁSTER UNIVERSITARIO EN EDUCACIÓN INCLUSIVA E INTERCULTURAL." -> "MA UNIVER EDUC INCLUS INTERC"
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
    for i, archivo_actual in enumerate(lista_archivos):
        start_time = time.time()
        nombre_base_archivo = os.path.basename(archivo_actual)
        if etiqueta_estado: etiqueta_estado.config(text=f"Procesando {i+1}/{num_archivos}: {nombre_base_archivo}")
        if barra_progreso: barra_progreso['value'] = i 

        mensajes_resumen_procesamiento.append(f"--- Procesando archivo {i+1}/{num_archivos}: {nombre_base_archivo} ---")
        print(f"\n--- Procesando archivo {i+1}/{num_archivos}: {nombre_base_archivo} ---")
        
        texto_extraido = None
        if nombre_base_archivo.lower().endswith('.pdf'):
            # --- CORRECCIÓN: Llamar a la función correcta ---
            texto_extraido = extraer_texto_pdf_con_apis(archivo_actual)
        elif nombre_base_archivo.lower().endswith(('.xlsx', '.xls')):
            texto_extraido = extraer_texto_excel_con_pandas(archivo_actual)
        else:
            msg = f"⚠️ Tipo de archivo no soportado para '{nombre_base_archivo}'. Solo PDF, XLSX, XLS."
            print(msg); mensajes_resumen_procesamiento.append(msg);
            if barra_progreso: barra_progreso['value'] = i + 1 
            continue

        if texto_extraido:
            texto_limpio_para_llm = limpiar_texto_general_para_llm(texto_extraido)
            if not texto_limpio_para_llm:
                msg_vacio = f"INFO: Texto extraído para '{nombre_base_archivo}' vacío post-limpieza."
                print(msg_vacio); mensajes_resumen_procesamiento.append(msg_vacio)
                datos_para_tabla_resumen.append({"Archivo Origen": nombre_base_archivo, "Nombre": "No procesado (texto vacío)", "Programa": "N/A", "Plan": "N/A"})
                # (resto del manejo de texto vacío)
            else:
                # (resto del procesamiento normal)
                print(f"\nPregunta/Instrucción para DeepSeek para '{nombre_base_archivo}'")
            
                respuesta_deepseek = analizar_con_deepseek(
                    texto_limpio_para_llm, PREGUNTA_PARA_DEEPSEEK, CHUTES_API_TOKEN,
                    DEEPSEEK_MODEL, DEEPSEEK_API_URL, nombre_base_archivo
                )

                if respuesta_deepseek:
                    if "DETENER_PROCESO_CONDICION_" in respuesta_deepseek:
                        msg_parada = f"⚠️ CONDICIÓN DE PARADA para '{nombre_base_archivo}':"
                        print(f"\n{msg_parada}"); mensajes_resumen_procesamiento.append(msg_parada)
                        datos_clave_parada = extraer_datos_clave_para_tabla(respuesta_deepseek)
                        datos_clave_parada["Archivo Origen"] = nombre_base_archivo
                        for line in respuesta_deepseek.splitlines():
                            if "DETENER_PROCESO_CONDICION_" in line:
                                detalle_parada = f"  -> {line.strip()}"; print(detalle_parada); mensajes_resumen_procesamiento.append(detalle_parada)
                                if "PERIODO" in detalle_parada.upper(): datos_clave_parada["Plan"] = "DETENIDO: PERIODO 2024"
                                elif "PROGRAMA" in detalle_parada.upper(): datos_clave_parada["Plan"] = "DETENIDO: PROGRAMA"
                                elif "CALIFICACION" in detalle_parada.upper(): datos_clave_parada["Plan"] = "DETENIDO: CALIFICACION"
                                else: datos_clave_parada["Plan"] = "DETENIDO: OTRA CONDICION"
                                break 
                        datos_para_tabla_resumen.append(datos_clave_parada)
                    else:
                        datos_extraidos_tabla = extraer_datos_clave_para_tabla(respuesta_deepseek)
                        datos_extraidos_tabla["Archivo Origen"] = nombre_base_archivo
                        datos_extraidos_tabla["Ruta Completa"] = archivo_actual  # Esta línea es nueva
                        datos_para_tabla_resumen.append(datos_extraidos_tabla) 
                        script_ahk_contenido_final = extraer_script_ahk_y_ajustar_abreviacion(respuesta_deepseek)

                        if datos_extraidos_tabla["Nombre"] != "No extraído" and script_ahk_contenido_final : 
                            if datos_extraidos_tabla["Nombre"] == "NombreEstudianteNoEncontrado" : 
                                msg_nombre_err = f"⚠️ ADVERTENCIA: No se pudo determinar nombre de estudiante para '{nombre_base_archivo}'. Usando nombre genérico para AHK."
                                print(msg_nombre_err); mensajes_resumen_procesamiento.append(msg_nombre_err)
                                nombre_base_doc_sin_ext = os.path.splitext(nombre_base_archivo)[0]
                                nombre_archivo_sanitizado = sanitizar_nombre_archivo(nombre_base_doc_sin_ext + "_AHK_Script")
                            else:
                                nombre_archivo_sanitizado = sanitizar_nombre_archivo(datos_extraidos_tabla["Nombre"])
                            
                            carpeta_doc_original = os.path.dirname(archivo_actual)
                            nombre_archivo_ahk = f"{nombre_archivo_sanitizado}.ahk"
                            ruta_archivo_ahk = os.path.join(carpeta_doc_original, nombre_archivo_ahk)
                            try:
                                with open(ruta_archivo_ahk, "w", encoding="utf-8") as f: f.write(script_ahk_contenido_final)
                                msg_guardado_ok = f"✅ Script AHK para '{nombre_base_archivo}' (Est: {datos_extraidos_tabla['Nombre']}) guardado en: {ruta_archivo_ahk}"
                                print(f"\n{msg_guardado_ok}"); mensajes_resumen_procesamiento.append(msg_guardado_ok)
                            except Exception as e:
                                msg_guardado_err = f"❌ Error guardando AHK '{ruta_archivo_ahk}': {e}"
                                print(f"\n{msg_guardado_err}"); traceback.print_exc(); mensajes_resumen_procesamiento.append(msg_guardado_err + f"\n{traceback.format_exc()}")
                        else:
                            if datos_extraidos_tabla["Nombre"] == "No extraído" or datos_extraidos_tabla["Nombre"] == "NombreEstudianteNoEncontrado":
                                msg_no_nombre = f"❌ No se pudo extraer nombre de estudiante de DeepSeek para '{nombre_base_archivo}' para el AHK."
                                print(msg_no_nombre); mensajes_resumen_procesamiento.append(msg_no_nombre)
                            if not script_ahk_contenido_final:
                                msg_no_ahk = f"❌ No se pudo extraer/ajustar script AHK de DeepSeek para '{nombre_base_archivo}'."
                                print(msg_no_ahk); mensajes_resumen_procesamiento.append(msg_no_ahk)
                            mensajes_resumen_procesamiento.append(f"No se generó archivo AHK para '{nombre_base_archivo}'.")
                else: # Corresponde a if respuesta_deepseek:
                    msg_no_resp_ds = f"No se pudo obtener respuesta de DeepSeek para '{nombre_base_archivo}'."
                    print(f"\n{msg_no_resp_ds}"); mensajes_resumen_procesamiento.append(msg_no_resp_ds)
                    datos_para_tabla_resumen.append({"Archivo Origen": nombre_base_archivo, "Nombre": "Error DeepSeek", "Programa": "N/A", "Plan": "N/A"})
        else: # Corresponde a if texto_extraido:
            msg_no_extraccion = f"No se pudo extraer texto del documento '{nombre_base_archivo}'."
            print(f"\n{msg_no_extraccion}"); mensajes_resumen_procesamiento.append(msg_no_extraccion)
            datos_para_tabla_resumen.append({"Archivo Origen": nombre_base_archivo, "Nombre": "Error Extracción Texto", "Programa": "N/A", "Plan": "N/A"})
        
        end_time = time.time()
        duracion_archivo = end_time - start_time
        tiempos_procesamiento.append(duracion_archivo)
        
        tiempo_promedio = sum(tiempos_procesamiento) / len(tiempos_procesamiento)
        archivos_restantes = num_archivos - (i + 1)
        etr_segundos = tiempo_promedio * archivos_restantes
        
        if etiqueta_etr:
            mins, secs = divmod(int(etr_segundos), 60)
            etr_texto = f"Tiempo restante: {mins:02d}:{secs:02d} min"
            etiqueta_etr.config(text=etr_texto)
        
        if barra_progreso: barra_progreso['value'] = i + 1 
        
        msg_fin_archivo = f"--- Fin del procesamiento para: {nombre_base_archivo} (Duración: {duracion_archivo:.2f}s) ---"
        print(msg_fin_archivo); mensajes_resumen_procesamiento.append(msg_fin_archivo)
        mensajes_resumen_procesamiento.append("--------------------------------------------------")

    # Fin del bucle
    if etiqueta_estado: etiqueta_estado.config(text=f"Procesamiento completado para {num_archivos} archivos.")
    if etiqueta_etr: etiqueta_etr.config(text="Tiempo restante: 00:00 min")
    if barra_progreso: barra_progreso['value'] = num_archivos
    if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL) 

    if mensajes_resumen_procesamiento:
        mostrar_resumen_log_gui(ventana_principal, mensajes_resumen_procesamiento)
    if datos_para_tabla_resumen:
        mostrar_tabla_resumen_gui(ventana_principal, datos_para_tabla_resumen)
    
    if messagebox.askyesno("Proceso Finalizado", "El procesamiento de los archivos ha terminado.\n¿Desea cerrar la aplicación?", parent=ventana_principal):
        ventana_principal.destroy()
    else:
        if etiqueta_estado: etiqueta_estado.config(text="Listo para seleccionar nuevos archivos.")
        if barra_progreso: barra_progreso['value'] = 0
        if etiqueta_etr: etiqueta_etr.config(text="Tiempo restante: --:--")
        if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)


def iniciar_procesamiento_en_hilo(boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal):
    if boton_seleccionar: boton_seleccionar.config(state=tk.DISABLED)
    if etiqueta_estado: etiqueta_estado.config(text="Seleccionando archivos...")

    file_paths_tupla = filedialog.askopenfilenames(
        title="Selecciona uno o más archivos PDF o Excel para analizar",
        filetypes=(("Documentos Soportados", "*.pdf *.xlsx *.xls"),("PDF files", "*.pdf"),("Excel files (*.xlsx, *.xls)", "*.xlsx;*.xls"),("All files", "*.*"))
    )
    
    if file_paths_tupla:
        lista_archivos_desordenada = list(file_paths_tupla)
        try:
            print("Ordenando archivos por fecha de modificación (del más antiguo al más reciente)...")
            lista_archivos_desordenada.sort(key=lambda ruta: os.path.getmtime(ruta))
            print("Archivos ordenados correctamente.")
            lista_archivos_ordenada = lista_archivos_desordenada

            hilo_procesamiento = threading.Thread(
                target=procesar_archivos_seleccionados, 
                args=(lista_archivos_ordenada, boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal),
                daemon=True # El hilo de fondo terminará si la ventana principal se cierra
            )
            hilo_procesamiento.start()

        except FileNotFoundError as e:
            msg_error_orden = f"Error al ordenar archivos: no se encontró el archivo {e.filename}."
            print(msg_error_orden)
            messagebox.showerror("Error de Archivo", msg_error_orden, parent=ventana_principal)
            if etiqueta_estado: etiqueta_estado.config(text="Error al procesar la selección.")
            if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)
    else:
        if etiqueta_estado: etiqueta_estado.config(text="Ningún archivo seleccionado.")
        if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)

# --- Flujo Principal con GUI Inicial ---
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
    
    # Placeholder check
    if LLMWHISPERER_API_KEY_1 == "TU_LLMWHISPERER_API_KEY_1" or CHUTES_API_TOKEN == "TU_CHUTES_API_TOKEN" or \
       not LLMWHISPERER_API_KEY_1 or not CHUTES_API_TOKEN:
        etiqueta_estado_gui.config(text="ERROR: ¡Configure las claves API en el script primero!")
        boton_seleccionar_gui.config(state=tk.DISABLED)
        messagebox.showerror("Error de Configuración", 
                             "Por favor, configura 'LLMWHISPERER_API_KEY_1' y 'CHUTES_API_TOKEN' en el script.",
                             parent=ventana_principal_app) 
    else:
        etiqueta_estado_gui.config(text="Claves API OK. Listo para seleccionar archivos.")

    ventana_principal_app.mainloop()