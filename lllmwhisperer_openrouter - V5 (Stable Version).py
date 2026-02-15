import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import queue
import asyncio
import shutil

# aiohttp ya no es necesario, se elimina
# import aiohttp 

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
import google.generativeai as genai

load_dotenv()

# --- Configuración y variables globales ---
LLMWHISPERER_API_KEY_1 = os.getenv("LLMWHISPERER_API_KEY_1")
LLMWHISPERER_API_KEY_2 = os.getenv("LLMWHISPERER_API_KEY_2")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_API_KEY_2 = os.getenv("OPENROUTER_API_KEY_2")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") # Se mantiene por si se quisiera usar en un worker

OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"

# --- 1. CONFIGURACIÓN DE WORKERS Y MODELOS (CAMBIO PRINCIPAL) ---
# <<< CAMBIO: Se define una lista de modelos. Cada modelo será asignado a un worker.
# Puedes añadir más modelos aquí para tener más workers.
WORKER_MODELS = [
    "deepseek/deepseek-chat-v3-0324:free",  # Worker 1 usará este
    "moonshotai/kimi-k2:free",                # Worker 2 usará este
    "tencent/hunyuan-a13b-instruct:free",                # Worker 3 usará este
    "meta-llama/llama-3.3-70b-instruct:free",                # Worker 3 usará este
    "qwen/qwen3-235b-a22b:free",                # Worker 4 usará este
    "qwen/qwen-2.5-72b-instruct:free",                # Worker 5 usará este
]

# <<< CAMBIO: MAX_WORKERS ahora es dinámico, basado en cuántos modelos configures.
MAX_WORKERS = len(WORKER_MODELS)

YOUR_SITE_URL = "https://github.com/tu_usuario/tu_proyecto"
YOUR_SITE_NAME = "Procesador AHK"

PROGRAMAS_VERDES = {
    "administración de empresas", "administracion de empresas", "ingenieria de sistemas", "ingenieria de software",
    "ingeniería de software", "de software", "ingenieria industrial", "ingeniería industrial", "ingenieria en ciencia de datos", "ingeniería en ciencia de datos",
    "licenciatura en ciencias sociales", "licenciatura en educación infantil", "licenciatura en educacion infantil",
    "licenciatura en matemáticas", "licenciatura en matematicas", "licenciatura en infantil", "mercadeo y publicidad",
    "negocios internacionales", "derecho", "administraciòn financiera virtual", "administración financiera virtual",
    "administracion financiera virtual", "administracion financiera", "administración financiera", "financiera virtual", "administracion en salud",
    "administración en salud", "administración en salud virtual", "especializacion en auditoria en salud",
    "especialización en auditoria en salud", "esp en marketing digital vir"
}

PALABRAS_CLAVE_POSGRADO = {"maestria", "especializacion", "esp", "posgrado"}

ABREVIATURAS_PREDEFINIDAS = {
    "TECNICO EN CONTABILIZACION DE OPERACIONES COMERCIALES Y FINANCIERAS": "TC CONTAB OPERA COMER Y FINANC",
    "TECNÓLOGO EN GESTION CONTABLE Y FINANCIERA": "TG EN GSTON CONTBLE Y FINACIRA",
    "TECNOLOGO EN ANALISIS Y DESARROLLO DE SOFTWARE": "TG ANALISIS Y DESARR DE SOFTWA",
    "TECNOLOGO EN ANALISIS Y DESARROLLO DE SISTEMAS DE INFORMACIÓN": "TG ANALISIS Y DESARR SIST INFO",
    "Tecnología en Gestión Bancaria y de Entidades Financieras": "TG EN GSTON BANCARIA Y FINANCI",
    "TECNÓLOGO EN GESTION CONTABLE Y DE INFORMACION FINANCIERA": "TG EN GSTON CONTBLE INFO FINAN",
    "TÉCNICO EN SEGURIDAD VIAL, CONTROL DE TRANSITO Y TRANSPORTE": "TC SEGU VIAL CTRL TRANSIT TRAN",
    "TECNOLOGIA EN GESTION DE LA SEGURIDAD Y SALUD EN EL TRABAJO": "TG EN GSTON SEG SLUD EN TRABA",
    "TECNOLOGO EN GESTIÓN INTEGRADA DE LA CALIDAD, MEDIO AMBIENTE, SEGURIDAD Y SALUD OCUPACIONAL": "TG GSTN CALID AMBI SEG Y SALD",
    "TECNÓLOGO EN GESTIÓN DEL TALENTO HUMANO": "TG EN GSTON DEL TLENTO HUMANO",
    "TÉCNICO EN SOLDADURA DE PRODUCTOS METÁLICOS": "TC EN SOLDADUR PRODUCT METALIC",
    "TÉCNICO EN ATENCIÓN INTEGRAL A LA PRIMERA INFANCIA": "TC ATENCION INTGRL PRIME INFAN",
    "TECNICO LABORAL EN EDUCACION PARA LA PRIMERA INFANCIA": "TC EN EDUCION PRIMERA INFANCIA",
    "Especialización en educación e intervención para la primera infancia.": "ESP EDUCION INTRCION PRIME INF",
    "TECNOLOGO EN CONTABILIDAD Y FINANZAS": "TG EN CONTABILIDAD Y FINANZAS",
    "TÉCNICO INTEGRACION DE CONTENIDOS DIGITALES": "TC INTGRCION CNTENIDOS DIGITAL",
    "TÉCNICO EN ASISTENCIA ADMINISTRATIVA": "TC EN ASTENCIA ADMINISTRATIVA",
    "TECNICO EN ASISTENCIA ADMINISTRATIVA": "TC EN ASTENCIA ADMINISTRATIVA",
    "TECNÓLOGO EN GESTIÓN ADMINISTRATIVA": "TG GESTION ADMINISTRATIVA",
    "TECNOLOGIA EN GESTIÓN ADMINISTRATIVA": "TG GESTION ADMINISTRATIVA",
    "TECNOLOGÍA EN GESTIÓN ADMINISTRATIVA": "TG GESTION ADMINISTRATIVA",
    "TECNOLOGÍA EN GESTIÓN INTEGRADA DE LA CALIDAD, MEDIO AMBIENTE, SEGURIDAD Y SALUD OCUPACIONAL": "TG GSTN CALID AMBI SEG Y SALD",
    "TECNOLOGO EN ACTIVIDAD FISICA": "TG EN ACTIVIDAD FISICA",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR ADMINISTRATIVO EN SALUD": "TC LABRAL AUX ADMINIST SALUD",
    "TÉCNICO LABORAL POR COMPETENCIAS EN LOGÍSTICA Y ADMINISTRACIÓN DE ALMACENES": "TC LABRAL LOGISTI Y ADM ALMACE",
    "TECNICO LABORAL POR COMPETENCIAS EN EDUCACION INICIAL Y RECREACION": "TC EN EDUCION INCIAL RECRCION",
    "TECNOLOGÍA EN MANTENIMIENTO ELECTROMECÁNICO INDUSTRIAL": "TG EN MANTEMTO ELECTRMC INDUS",
    "TECNICO EN SISTEMAS AGROPECUARIOS ECOLOGICOS": "TC EN SISTEMA AGROPECU ECOLOGI",
    "TÉCNICO EN PROGRAMACION DE APLICACIONES PARA DISPOSITIVOS MOVILES": "TC PROGRAM APLICAC DISPOS MOVI",
    "TÉCNICO EN PROGRAMACIÓN DE SOFTWARE": "TC PROGRAMACIÓN DE SOFTWARE",
    "TECNICO EN PROGRAMACION DE SOFTWARE": "TC PROGRAMACIÓN DE SOFTWARE",
    "TÉCNICO EN PROGRAMACION DE SOFTWARE": "TC PROGRAMACIÓN DE SOFTWARE",
    "Tecnología en Gestión Financiera y Tesorería": "TG EN GSTON FINACIRA Y TESORIA",
    "TECNOLOGÍA EN COORDINACIÓN DE PROCESOS LOGÍSTICOS": "TG EN COORDN PRCSOS LOGISTIC",
    "TÉCNICO EN SERVICIOS COMERCIALES Y FINANCIEROS": "TC EN SERVICI COMERC Y FINANCI",
    "TECNÓLOGO EN GESTIÓN INTEGRADA DE LA CALIDAD, MEDIO AMBIENTE, SEGURIDAD Y SALUD OCUPACIONAL": "TG GSTN CALID AMBI SEG Y SALD",
    "TECNÓLOGO EN DISEÑO DE ELEMENTOS MECÁNICOS PARA SU FABRICACIÓN CON MÁQUINAS HERRAMIENTAS CNC": "TG DISÑ ELE MEC MAQU HER CNC",
    "TECNÓLOGO EN MANTENIMIENTO DE EQUIPOS DE COMPUTO, DISEÑO E INSTALACION DE CABLEADO ESTRUCTURADO": "TG MANTO EQUIP COMP CABL ESTRU",
    "INGENIERIA EN SEGURIDAD INDUSTRIAL E HIGIENE OCUPACIONAL": "ING EN SEG INDU Y SALUD OCUPAC",
    "TECNICO EN ASESORÍA COMERCIAL Y OPERACIONES DE ENTIDADES FINANCIERAS": "TC EN ASERIA COMR OPER ENT FIN",
    "TECNICO LABORAL EN AUXILIAR ADMINISTRATIVO EN SALUD": "TC LAB EN AUX ADMINIST EN SALD",
    "TECNÓLOGO EN GESTION DE PROYECTOS DE DESARROLLO ECONOMICO Y SOCIAL": "TG GSTN PROY DESRRLL ECO SOC",
    "TÉCNICO PROFESIONAL EN MANTENIMIENTO ELECTRÓNICO INDUSTRIAL": "TC EN MNTMTO ELECTRO INDSTRL",
    "TÉCNICO EN MANEJO INTEGRAL DE RESIDUOS SOLIDOS": "TC MANEJO INTGRL RESIDU SOLID",
    "TG EN GESTION DE PROCESOS ADMINISTRATIVOS DE SALUD": "TG GSTON PROCE ADMINI DE SALUD",
    "TECNOLOGÍA EN DESARROLLO DE SOFTWARE Y APLICATIVOS MÓVILES": "TG EN DSRLLO SOFT Y APLICA MOV",
    "TECNICO LABORAL POR COMPETENCIA EN AUXILIAR EN ENFERMERIA": "TC LAB POR COMP EN AUX ENFERME",
    "TÉCNICO EN OPERACIÓN DE MAQUINARIA PESADA PARA EXCAVACIÓN": "TC OPERAC MAQUI PESAD EXCAVA",
    "Especialización en Desarrollo integral de la infancia y la adolescencia": "ESP EN DESRRLL INTG INF Y ADOL",
    "ESPECIALIZACIÓN EN ADMINISTRACIÓN DE LA INFORMÁTICA EDUCATIVA": "ESP ADMON DE LA INFORM EDUC",
    "TÉCNICO EN SEGURIDAD Y SALUD EN EL TRABAJO": "TC EN SEGU Y SALD EN EL TRABAJ",
    "TECNÓLOGO EN DISEÑO, IMPLEMENTACIÓN Y MANTENIMIENTO DE SISTEMAS DE TELECOMUNICACIONES": "TG DISÑ IMPLEM MANT SIS Y TELE",
    "TECNÓLOGO EN SALUD OCUPACIONAL": "TG EN SALUD OCUPACIONAL",
    "Tecnología en Coordinación de Servicios Hoteleros": "TG EN COORDIN DE SERVIC HOTELR",
    "TÉCNICO LABORAL POR COMPETENCIAS EN GESTOR COMUNITARIO Y SOCIAL": "TC LBRL COMP GSTR COMN Y SOCIL",
    "TECNOLOGO EN GESTIÓN ADMINISTRATIVA DEL SECTOR SALUD": "TG EN GSTON ADMTIVA SECT SALUD",
    "ESPECIALIZACION EN BIGDATA Y ANALITICA VIRTUAL": "ESP BIGDATA Y ANALTICA VIRTUAL",
    "TECNÓLOGO EN SISTEMAS DE GESTIÓN AMBIENTAL": "TG EN SISTEMA GSTON AMBIENTAL",
    "TECNICO EN APOYO ADMINISTRATIVO EN SALUD": "TC EN APOYO ADMINTIVO EN SALD",
    "TECNICO PROFESONAL EN SEGURIDAD Y SALUD EN EL TRABAJO": "TC PROF SEG Y SALD EN EL TRABJ",
    "TÉCNICO DE OPERACIONES DE COMERCIO EXTERIOR": "TC DE OPERAC COMER EXTERIOR",
    "TECNÓLOGO EN MANTENIMIENTO ELECTRÓNICO E INSTRUMENTAL INDUSTRIAL": "TG MNTMTO ELECT E INSTRUM INDU",
    "ESPECIALIZACIÓN EN PEDAGOGÍA DE LA RECREACIÓN ECOLÓGICA": "ESP EN PEDAG DE LA RECRE ECOLG",
    "Especialización en Neuropsicología de la educación": "ESP NEUROPSICO DE LA EDUCACION",
    "TÉCNICO EN MONTAJE Y MANTENimiento DE REDES AEREAS DE DISTRIBUCIÓN DE ENERGÍA ELECTRICA": "TC MONT MTNTO RDES AERE DISTR",
    "TECNOLOGÍA EN GESTIÓN DE SALUD OCUPACIONAL, SEGURIDAD Y MEDIO AMBIENTE": "TG GSTN SALD OCUP SEG MED AMB",
    "TÉCNICO EN IMPLEMENTACION Y MANTENIMIENTO DE EQUIPOS ELECTRONICOS INDUSTRIALES": "TC IMPL MANTNTO EQU ELEC INDUS",
    "TECNICO LABORAL POR COMPETENCIAS AUXILIAR EN RECURSOS HUMANOS": "TC LAB COMPETE AUX EN REC HUMA",
    "TÉCNICO EN MANTENIMIENTO DE MAQUINAS DE CONFECCIÓN INDUSTRIAL": "TC MTNTO MAQUI CONFECC INDUSTR",
    "TECNOLOGÍA EN MANTENIMIENTO MECATRÓNICO DE AUTOMOTORES": "TG EN MTNTO MECATRNICO AUTOMOT",
    "TÉCNICO LABORAL POR COMPETENCIAS EN ATENCIÓN INTEGRAL A LA PRIMERA INFANCIA": "TC COMP ATNCN INTGRL PRIM INFA",
    "TECNICO LABORAL EN ATENCIÓN INTEGRAL A LA PRIMERA INFANCIA": "TC LAB ATNCON INTGRL PRIM INFA",
    "TECNICO LABORAL POR COMPETENCIAS EN CONTABLE Y FINANCIERO": "TC LAB POR COMP CONTABL Y FINA",
    "TECNICO PROFESIONAL EN OPERACIONES CONTABLES": "TC PROF EN OPERACIO CONTABLES",
    "TECNICO LABORAL POR COMPETENCIAS EN SEGURIDAD OCUPACIONAL": "TC LAB POR COMP EN SEG OCUPACI",
    "TÉCNICO EN PRESELECCION DE TALENTO HUMANO MEDIADO POR HERRAMIENTAS TIC": "TC PRESELECC TLTO HUM HERR TIC",
    "TÉCNICO EN INSTALACION Y MANTENIMIENTO DE EQUIPOS PARA INSTRUMENTACION INDUSTRIAL": "TC INSTL MNTO EQUIP INSTRU IND",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILILAR ADMINISTRATIVO": "TC LAB POR COMP EN AUX ADMON",
    "TÉCNICO EN MANTENIMIENTO DE EQUIPOS DE CÓMPUTO": "TG EN MANTENIMTO EQUIP COMP",
    "TECNICO LABORAL POR COMPETENCIAS EN ELECTRICISTA INDUSTRIAL": "TC LAB POR COMP EN ELECTR IND",
    "TÉCNICO EN INTEGRACIÓN DE OPERACIONES LOGISTICAS": "TC EN INTGRAC DE OPERA LOGISTI",
    "TÉCNICO EN MONTAJE Y MANTENIMIENTO ELECTROMECANico DE INSTALACIONES MINERAS BAJO TIERRA": "TC MONTJE MNTO ELECTR INST MIN",
    "ESPECIALIZACIÓN EN LÚDICA Y RECREACIÓN PARA EL DESARROLLO SOCIAL Y CULTURAL": "ESP LUD RCRE DESARR SOCI Y CUL",
    "TECNÓLOGO GESTION PARA ESTABLECIMIENTOS DE ALIMENTOS Y BEBIDAS": "TG GSTON ESTBLCI ALIMNT Y BEBD",
    "TÉCNICO EN INSTALACIONES ELECTRICAS RESIDENCIALES": "TC EN INSTALC ELECTRIC RESIDEN",
    "TÉCNICO EN ASISTENCIA EN ORGANIZACIÓN DE ARCHIVO": "TC ASISTEN ORGANIZ DE ARCHIVO",
    "TECNOLOGÍA EN DISEÑO E INTEGRACIÓN DE AUTOMATISMOS MECATRÓNICOS": "TC DISEÑ INTGRC AUTOM MECATR",
    "TECNOLOGÍA EN DISTRIBUCIÓN FÍSICA INTERNACIONAL": "TG EN DISTRIBUC FISICA INTERNA",
    "TECNOLOGA EN COMERCIO EXTERIOR Y NEGOCIOS INTERNACIONALES": "TG EN COM EXTER Y NEGO INTERNA",
    "Tecnologo en mercadeo y diseño publicitario": "TG EN MERCADEO Y DISEÑO PUBLIC",
    "Tecnologo Coordinador de Procesos Logísticos": "TG COORDINADR PROCES LOGISTIC",
    "ESPECIALIZACION EN GERENCIA DE LA CALIDAD EN SALUD": "ESP GERNCIA DE CALDAD EN SALUD",
    "TECNOLOGÍA EN GESTION DE REDES DE DATOS": "TG GSTON DE REDES DATOS",
    "TECNOLOGÍA EN COMPUTACIÓN Y DESARROLLO DE SOFTWARE": "TG COMPUTACI Y DESARR DE SOFTW",
    "TECNOLOGÍA EN GESTIÓN DE LA PRODUCCIÓN INDUSTRIAL": "TG GSTON DE LA PRODUC INDUSTRI",
    "TECNOLOGÍA EN QUÍMICA APLICADA A LA INDUSTRIA": "TG QUIMICA APLICD A LA INDUSTR",
    "TECNÓLOGÍA EN LEVANTAMIENTOS TOPOGRAFICOS Y GEORREFERENCIACIÓN": "TG LEVANTA TOPOGRA Y GEOREFERE",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR EN EDUCACIÓN PARA LA PRIMERA INFANCIA": "TC LAB POR COMP EDU PRIME INFA",
    "TÉCNICO EN AUXILIAR EN EDUCACIÓN INTEGRAL PARA LA PRIMERA INFANCIA": "TC AUX EDUC INTGRL PRIME INFAN",
    "TECNICO LABORAL EN AUXILIAR DE SERVICIOS FARMACEUTICOS": "TC LAB EN AUX DE SERVIC FARMAC",
    "TECNICO LABORAL AUXILIAR EN SEGURIDAD Y SALUD EN EL TRABAJO": "TC LAB SEG Y SALUD EN EL TRAB",
    "Técnico en operaciones de logística comercial en grandes superficies": "TC EN OPER LOG COME GRAN SUPER",
    "TECNOLOGÍA EN ASEGURAMIENTO METROLÓGICO INDUSTRIAL": "TG ASEGUR METROLOG INDUSTRIAL",
    "TECNOLOGÍA EN FOTOGRAFÍA Y PROCESOS DIGITALES": "TG EN FOTOGRAF Y PROCES DIGITA",
    "Técnico Desarrollo de Operaciones Logísticas": "TC DESARR DE OPERACI LOGISTICS",
    "TECNOLOGÍA EN PROCESOS DE LA INDUSTRIA QUÍMICA": "TG EN PROCES DE INDUST QUIMICA",
    "TÉCNICO PROFESIONAL EN PROCESO ADMINISTRATIVO EN SALUD": "TC PROF PROCE ADMTVO EN SALUD",
    "Tecnología en Gestión de la Propiedad Horizontal": "TG GSTON DE LA PROPIED HORIZO",
    "TÉCNICO EN MECANICO DE MAQUINARIA INDUSTRIAL": "TC MECANICO MAQUIN INDUSTRIAL",
    "TÉCNICO LABORAL POR COMPETENCIAS EN LOGÍSTICA Y PRODUCCIÓN": "TC LAB POR COMP LOGIST Y PRODU",
    "TÉCNICO LABORAL POR COMPETENCIAS EN AUXILIAR CONTABLE Y FINANCIERO": "TC LAB COMPET AUX CONT Y FINAN",
    "TECNICO LABORAL EN AUXILIAR CONTABLE Y FINANCIERO": "TC LAB EN AUX CONTABL Y FINANC",
    "Técnico Desarrollo de Operaciones Logísticas en la Cadena de Abastecimiento": "TC DESARR OPER LOG CADEN ABAST",
    "Tecnología en Gestión Integral del Riesgo en Seguros": "TG GSTON INTGRL RISGO EN SEGUR",
    "ESPECIALIZACIÓN EN ATENCIÓN INTEGRAL A LA PRIMERA INFANCIA": "ESP ATCION INTGRL PRIMER INFAN",
    "TECNICO LABORAL POR COMPETENCIAS EN SEGURIDAD Y SALUD EN EL TRABAJO": "TC LAB COMPET SEG SLUD TRBJ",
    "TÉCNICO LABORAL POR COMPETENCIAS EN AUXILIAR DE PREESCOLAR": "TC LAB POR COMP AUX PREESC",
    "TECNICO PROFESIONAL EN SERVICIOS ADMINISTRATIVOS DE SALUD": "TC PROF SERVIC ADMIN SALUD",
    "ADMINISTRACIÓN DE EMPRESAS NIVEL I TÉCNICO PROFESIONAL EN PROCESOS EMPRESARIALES": "TC PROF EN PROCES EMPRESAR",
    "TECNOLOGÍA EN CONSTRUCCIÓN DE INFRAESTRUCTURA VIAL": "TG CONSTRU INFRAESTRUC VIAL",
    "TECNÓLOGO EN GESTIÓN DE RECURSOS EN PLANTAS DE PRODUCCIÓN": "TG GSTON RECUR EN PLAN PRODUCC",
    "TECNÓLOGO EN IMPLEMENTACION DE INFRAESTRUCTURA DE TECNOLOGIAS DE LA INFORMACION Y LAS COMUNICACIONES": "TG IMPL INFRA TECNO INFO COMUN",
    "TÉCNICO LABORAL EN AUXILIAR DE EDUCACIÓN DE LA PRIMERA INFANCIA": "TC LAB EN AUX EDUC PRIMER INFA",
    "Técnico Laboral por competencias en Apoyo a la Primera Infancia": "TC LAB COMP APOY PRIMER INFAN",
    "TECNOLOGÍA EN GESTIÓN DEL CICLO DE VIDA DEL PRODUCTO": "TG GSTON CICLO VIDA DEL PRODU",
    "TECNOLOGÍA EN PRODUCCIÓN AGROPECUARIA ECOLÓGICA": "TG EN PRODUCC AGROPEC ECOLO",
    "Técnico Laboral por Competencias en Banca y Servicios Financieros": "TC LAB COMPET BANC Y SER FINAN",
    "TÉCNICO PROFESIONAL EN ATENCIÓN INTEGRAL A LA PRIMERA INFANCIA": "TC PROF ATNCIN INTGRL PRIM INF",
    "TÉCNICO EN OPERACION DE SERVICIOS EN CONTACT CENTER Y BPO": "TC OPER SERV CONTA CENT Y BPO",
    "TÉCNICO EN ALISTAMIENTO Y OPERACIÓN DE MAQUINARIA PARA LA PRODUCCIÓN INDUSTRIAL": "TC ALIMTO Y OPER MAQU PRO IND",
    "Especialización en Planeación Educativa y Planes de Desarrollo": "ESP EN PLANE EDUC PLAN DESARR",
    "ESPECIALIZACIÓN EN INVESTIGACIÓN E INNOVACIÓN EDUCATIVA": "ESP EN INVSTG INNOVAC EDUCATIV",
    "TECNOLOGO EN SUPERVISIÓN EN PROCESOS DE CONFECCIÓN": "TG EN SUPERVI PROCES DE CONFEC",
    "TECNICO LABORAL EN ATENCION A LA PRIMERA INFANCIA": "TC LAB ATNCION PRIMERA INFANCI",
    "TECNICO LABORAL POR COMPETENCIAS EN TRABAJO SOCIAL Y COMUNITARIO": "TC LAB COMPT TRAB SOCIAL Y COM",
    "TECNICO LABORAL POR COMPETENCIAS EN ASISTENCIA ADMINISTRATIVA": "TC LAB POR COMP ASISTEN ADMIN",
    "TECNICO LABORAL POR COMPETENCIAS EN ASISTENTE ADMINISTRATIVO": "TC LAB POR COMPET ASISTE ADMIN",
    "ESPECIALIZACION EN METODOS Y TECNICAS DE INVESTIGACION EN CIENCIAS SOCIALES": "ESP MET Y TEC INVST CIENC SOCI",
    "TÉCNICO LABORAL POR COMPETENCIAS EN ATENCIÓN A LA PRIMERA INFANCIA": "TC LAB COMP ATNCIN PRIMR INFA",
    "ESPECIALIZACIÓN EN EVALUACIÓN E INTERVENCIÓN PSICOEDUCATIVA": "ESP EVALUA E INTERVE PSICOEDU",
    "TECNOLOGO EN GESTION DE OPERACIONES EN TERMINALES PORTUARIAS": "TG GSTON OPERACIO TERM PORTU",
    "TECNICO LABORAL COMO AUXILIAR EN EL DESARROLLO DE LA PRIMERA INFANCIA": "TC LAB AUX DESARR PRIME INFANC",
    "TECNÓLOGO EN DESARROLLO DE MEDIOS GRAFICOS VISUALES": "TG DESARR DE MEDI GRAF VISUALE",
    "TÉCNICO LABORAL ASISTENTE EN DESARROLLO DE SOFTWARE": "TC LAB ASISTE DEASRR DE SOFTW",
    "TÉCNICO LABORAL POR COMPETENCIAS AUXILIAR EN ATENCIÓN A LA PRIMERA INFANCIA": "TC LAB COMP AUX ATNCN PRIM INF",
    "TECNÓLOGO EN CONTROL DE CALIDAD EN LA INDUSTRIA DE ALIMENTOS": "TG CONTRL CALID IND ALIMENTOS",
    "TECNÓLOGO EN SUPERVISIÓN DE REDES DE DISTRIBUCIÓN DE ENERGÍA ELÉCTRICA": "TG SUPER RED DISTR ENER ELECT",
    "TECNICO LABORAL POR COMPETENCIAS COMO AUXILIAR EN TALENTO HUMANO, SEGURIDAD Y SALUD EN EL TRABAJO": "TC LAB COMP AUX TL HUM SST",
    "TECNICO EN MANTENIMIENTO DE EQUIPOS DE REFRIGERACIÓN, VENTILACIÓN Y CLIMATIZACIÓN": "TC MANTO EQUPO REFRG VENT CLI",
    "TECNICO LABORAL POR COMPETENCIAS EN ASISTENCIA EN ATENCIÓN A LA PRIMERA INFANCIA": "TC LAB COMPT ASIS ATNC PRM INF",
    "TÉCNICO LABORAL EN AUXILIAR EN SEGURIDAD OCUPACIONAL": "TC LAB AUX SEGUR OCUPACIONAL",
    "TECNOLOGÍA EN GESTIÓN DE SISTEMAS DE INFORMACIÓN Y REDES DE COMPUTO": "TG GSTON SIST INF Y REDES COMP",
    "TÉCNICO LABORAL POR COMPETENCIAS EN ASISTENTE EN ATENCIÓN INTEGRAL A LA PRIMERA INFANCIA": "TC LAB COM ATEN INTGR PRI INFA",
    "TECNICO LABORAL POR COMPETENCIAS EN SEGURIDAD OCUPACIONAL Y LABORAL": "TC LAB COMP SEGU OCUP Y LAB",
    "TÉCNICO LABORAL POR COMPETENCIAS EN CONTABILIDAD Y FINANZAS": "TC LAB COMPT CONTABI Y FINANZ",
    "TECNÓLOGO EN SUPERVISIÓN EN SISTEMAS DE AGUA Y SANEAMIENTO": "TG SUPERV SISTE AGU Y SANEAMIE",
    "TECNICO LABORAL EN ÁNALISIS Y SISTEMAS DE INFORMACIÓN": "TC LAB ANALI Y SISTEMAS INFORM",
    "TÉCNICO EN SERVICIOS Y OPERACIONES MICROFINANCIERAS": "TC EN SERV Y OPERAC MICROFINA",
    "TÉCNICO LABORAL POR COMPETENCIAS COMO REVISADOR DE CALIDAD": "TC LAB POR COMPET REVISA CALID",
    "TECNOLOGO EN IMPLEMENTACION DE REDES Y SERVICIOS DE TELECOMUNICACIONES": "TG IMPLTC REDE Y SERV DE TELEC",
    "TECNÓLOGO EN IMPLEMENTACIÓN DE INFRAESTRUCTURA DE TECNOLOGÍAS DE LA INFORMACIÓN Y LAS COMUNICACIONES": "TG IMPLEM INFRA TEC INFOR COM",
    "TÉCNICO EN BILINGUAL EXPERT ON BUSINESS PROCESS OUTSOURCING": "TC BILG EXPT ON BUSIN PRO OUT",
    "TECNOLOGIA EN GESTION DE SISTEMAS DE TELECOMUNICACIONES": "TG GSTON SISTEM DE TELECOMU",
    "TÉCNICO DESARROLLO DE OPERACIONES LOGÍSTICA EN LA CADENA DE ABASTECIMIENTO": "TC DESAR OPRC LOGST CADN AB",
    "TÉCNICO LABORAL POR COMPETENCIAS EN EDUCACIÓN PARA LA PRIMERA INFANCIA": "TC LAB COMPT EDUC PRIM INFAN",
    "TECNOLOGO EN GESTION BANCARIA Y DE ENTIDADES FINANCIERAS": "TG GSTON BANC ENTIDA FINANCI",
    "Tecnología en Gestión de Proyectos de Desarrollo Económico y Social": "TG GSTON PROYE DESAR ECON SOC",
    "TÉCNICO LABORAL AUXILIAR EN CUIDADO DE NIÑOS - PRIMERA INFANCIA": "TC LAB AUX CUID NIÑOS PRIM INF",
    "TECNÓLOGO EN CONTROL DE CALIDAD EN DE ALIMENTOS": "TG CTRL CALIDAD DE ALIMENTOS",
    "TECNICO EN ATENCION INTEGRAL A LA PRIMERA INFANCIA": "TC ATNCION INTGRL PRIME INFAN",
    "TECNOLOGÍA EN GESTIÓN CONTABLE Y DE INFORMACIÓN FINANCIERA": "TG GSTON CONTABLE INFO FINAN",
    "TECNICO LABORAL POR COMPETENCIAS EN TRABAJO SOCIAL Y COMUNITARIO SGC": "TC LAB COMPT TRA SOC COM SGC",
    "Tecnología en Gestión Bancaria y Entidades Financieras": "TG GSTON BANC Y ENTID FINANCIE",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR EN SERVICIOS FARMACEUTICOS": "TC LAB COMPT AUX SER FARMACEU",
    "TÉCNICO EN PRESELECCION DE TALENTO HUMANO MEDIADO POR HERRAMIENTAS TIC": "TC PRESELEC TLTO HUM MED TIC",
    "Técnico Laboral por Competencias en: AUXILIAR DE PREESCOLAR": "TC LAB COMP AUX PREESCOLAR",
    "TECNOLOGÍA EN AUTOMATIZACIÓN INDUSTRIAL": "TG AUTOMATIZAC INDUSTRIAL",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR DE EDUCACIÓN PARA LA PRIMERA INFANCIA": "TC LAB COMPT AUX EDU PRIM INF",
    "TECNÓLOGO EN ENTRENAMIENTO DEPORTIVO": "TG ENTRENAMTO DEPORTIVO",
    "Técnico Contabilización de Operaciones Comerciales y Financieras": "TC CONTAB OPER COMER Y FINAN",
    "TÉCNICO EN OPERACIONES DE COMERCIO EXTERIOR": "TC OPERACIO COMERC EXTERIOR",
    "TECNICO LABORAL POR COMPETENCIS EN AUXILIAR DE EDUCACIÓN PARA LA PRIMERA INFANCIA": "TC LAB COMPT AUX EDU PRIM INFA",
    "Técnico Laboral en AUXILIAR CONTABLE Y FINANCIERO": "TC LAB AUX CONTAB Y FINANCIERO",
    "TECNÓLOGO GESTION CONTABLE Y DE INFORMACION FINANCIERA": "TG GSTON CONTABL Y INFO FINANC",
    "TÉCNICO EN DESARROLLO DE OPERACIONES EN LA CADENA DE ABASTECIMIENTO": "TC DESARR OPERAC CADEN ABASTC",
    "TÉCNICO EN OPERACIÓN DE SERVICIOS EN CONTACT CENTER Y BPO": "TC OPERAC SERV CONTC CENT BPO",
    "TECNÓLOGO EN MANTENIMIENTO MECATRÓNICO DE AUTOMOTORES": "TG MANTO MECATRCO DE AUTOMTO",
    "TÉCNICO LABORAL POR COMPETENCIAS EN ATENCION A LA PRIMERA INFANCIA": "TC LAB COMPT ATNC PRIMER INFAN",
    "LICENCIATURA EN PEDAGOGÍA INFANTIL VIRTUAL": "LIC PEDAGOGIA INFANTIL VIRTUAL",
    "TÉCNICO PROFESIONAL EN SERVICIO DE POLICÍA DE LA DIRECCIÓN NACIONAL DE ESCUELAS": "TC PROF SERV POLIC DIR NAC ESC",
    "TÉCNICO EN CONTABILIZACION DE OPERACIONES COMERCIALES Y FINANCIERA": "TC CONT OPERAC COMER Y FINAN",
    "TÉCNICO LABORAL POR COMPETENCIAS EN INSTALADOR DE REDES DE TELECOMUNICACIONES": "TC LAB COMPT INSTALD RED COMU",
    "TÉCNICO LABORAL EN FORMACIÓN PREESCOLAR Y RECREACIÓN INFANTIL": "TC LAB FORM PREESC RECRE INFAN",
    "Especialización en Desarollo integral de la infancia y la adolescencia": "ESP DESAR INTGRL INFA Y ADOLES",
    "ESPECIALIZACIÓN EN APLICACIÓN DE TIC PARA LA ENSEÑANZA": "ESP APLICA TIC PARA LA ENSEÑAN",
    "PRESELECCION DE TALENTO HUMANO MEDIADO POR HERRAMIENTAS TIC": "PRESELEC TALTO HUM HERRA TIC",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR EN SEGURIDAD OCUPACIONAL": "TC LAB COMPT AUX SEG OCUPACI",
    "TÉCNICO LABORAL POR COMPETENCIAS AUXILIAR EN SISTEMAS": "TC LAB COMPT AUX EN SISTEMAS",
    "TÉCNICO LABORAL AUXILIAR DE EDUCACIÓN PARA LA PRIMERA INFANCIA": "TC LAB AUX EDUC PRIMER INFAN",
    "Máster Universitario en Educación Inclusiva e Intercultural.": "MA UNIVER EDUC INCLUS INTERC",
    "TÉCNICO EN DISEÑO E INTEGRACION MULTIMEDIA": "TC DISEÑO INTEGR MULTIMEDIA",
    "TÉCNICO EN CONTABILIZACIÓN DE OPERaciones COMERCIALES Y FINANCIERAS": "TC CONTA OPERAC COMER FINA",
    "TECNOLOGÍA EN MANTENIMIENTO DE SISTEMAS ELECTROMECÁNICOS": "TG MANTMTO SISTE ELECTRMECA",
    "TECNÓLOGÍA EN GESTION DE LA SEGURIDAD Y SALUD EN EL TRABAJO": "TG GSTON SEGURD Y SALUD TRABAJ",
    "Técnico Laboral en Manejo y Aplicación de Sistemas Informáticos y Bases de Datos": "TC LAB MAN APLIC SISTE INF BAS",
    "TECNICO LABORAL EN ATENCION INTEGRAL A LA PRIMERA INFANCIA": "TC LAB ATNCION INTGRL PRIM INF",
    "TECNÓLOGIA EN CONTROL DE BIOPROCESOS INDUSTRIALES": "TG CTRL DE BIOPROCES INDUSTRI",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR DE ENFERMERIA": "TC LAB COMPT AUX DE ENFERMER",
    "ESPECIALIZACION EN PEDAGOGIA AMBIENTAL A DISTANCIA": "ESP PEDAGOGIA AMB A DISTANC",
    "TÉCNICO LABORAL POR COMPETENCIAS EN ASISTENTE SOCIAL Y COMUNITARIO CON ÉNFASIS EN ATENCIÓN A LA PRIMERA": "TC LAB COMPT ASIS SOC COMUN",
    "TÉCNICO LABORAL POR COMPETENCIAS AUXILIAR EN ENFERMERIA": "TC LAB COMPT AUX EN ENFERME",
    "TÉCNICO EN VENTAS DE PRODUCTOS Y SERVICIOS": "TC VENTAS PRODUCT Y SERVIC",
    "TECNICO LABORAL POR COMPETENCIAS EN INSTALACIONES ELÉCTRICAS": "TC LAB COMPT INSTALAC ELECTR",
    "TÉCNICO EN INSTALACION DE SISTEMAS ELECTRICOS RESIDENCIALES Y COMERCIALES": "TC INSTALC SISTE ELECT RESID",
    "TECNÓLOGO EN PROCESOS PARA LA COMERCIALIZACION INTERNACIONAL": "TG PROCES COMERCI INTERNA",
    "TÉCNICO EN MANTENIMIENTO ELÉCTRICO Y CONTROL ELECTRÓNICO EN AUTOMOTORES": "TC MANTMTO ELECTR AUTOMO",
    "TECNICA PROFESIONAL EN SOPORTE DE SISTEMAS EN INFORMÁTICA": "TC PROFE SOPORT SISTE INFOR",
    "TECNÓLOGO EN GESTION FINANCIERA Y DE TESORERIA": "TG GSTON FINANC Y TESORERIA",
    "TECNICO LABORAL POR COMPETENCIAS EN ASISTENTE DE PREESCOLAR": "TC LAB COMPT ASISTE PREESCOL",
    "TÉCNICO PROFESional EN ATENCIÓN A LA PRIMERA INFANCIA": "TC PROFESI ATNCON PRIMER INFA",
    "Tecnologia en Gestión de Recursos Naturales": "TG GSTON RECURSOS NATURALES",
    "TECNOLOGÍA EN OPERACIÓN DE PLANTAS PETROQUÍMICAS": "TG OPERACIÓN PLNTS PETROQUIM",
    "TECNOLOGO EN ADMINISTRACIÓN EN SERVICIOS DE SALUD": "TG ADMON SERVICIOS DE SALUD",
    "TÉCNICO LABORAL EN AUXILIAR EN EDUCACIÓN PARA LA PRIMERA INFANCIA": "TC LAB AUX EDUC PRIMER INFAN",
    "TÉCNICO LABORAL POR COMPETENCIAS FORMACIÓN Y ATENCIÓN A LA PRIMERA INFANCIA": "TC LAB COMPT FORM ATNC PRIM",
    "TECNOLOGIA EN GESTION DE ANALITICA Y BIG DATA": "TG GSTON ANALITICA Y BIG DATA",
    "TECNOLOGIA EN GESTION CONTABLE Y FINANCIERA VIRTUAL": "TG GSTON CONTBLE Y FINAN VIR",
    "Técnico en el Riesgo Crediticio y su Administración": "TC RIESGO CREDIT Y SU ADMON",
    "TÉCNICO PROFESIONAL EN SERVICIO DE POLICIA ESECU": "TC PROFES SERV POLIC ESECU",
    "TÉCNICO EN CONSERVACION DE RECURSOS NATURALES": "TC EN CONSERV RECURS NATURAL",
    "TECNOLOGÍA EN PREVENCIÓN Y CONTROL AMBIENTAL": "TG PREVENCION Y CTRL AMBIENTAL",
    "PROGRAMA TÉCNICO PROFESIONAL DISEÑO WEB Y MULTIMEDIA": "TC PROFES DISEÑ WEB Y MULTIMED",
    "TECNOLOGÍA GESTIÓN DE EMPRESas AGROPECUARIAS": "TG GSTON EMPRESAS AGROPECUAR",
    "TECNÓLOGO EN GESTIÓN DE LA SEGURIDAD Y SALUD EN EL TRABAJO": "TG GSTON SEGUR Y SALUD TRABAJO",
    "TECNOLOGO EN ANÁLISIS Y DESARROLLO DE SOFTWARE": "TG ANALISIS Y DESARRO DE SOFTW",
    "TECNICO EN ASESORIA COMERCIAL Y OPERACIONES DE ENTIDADES FINANCIERAS": "TC EN ASERIA COMR OPER ENT FIN",
    "TÉCNICO LABORAL POR COMPETENCIAS EN ATENCION INTEGRAL A LA PRIMERA INFANCIA": "TC COMP ATNCN INTGRL PRIM INFA",
    "TECNOLOGO EN SUPERVISIÓN DE REDES DE DISTRIBUCIÓN DE ENERGÍA ELÉCTRICA": "TG SUPERV RED DISTR ENER ELECT",
    "TÉCNICA PROFESIONAL EN SOPORTE DE SISTEMAS EN INFORMATIC": "TC PROFES SOPORT SISTE INFORMA",
    "TÉCNICO EN ASESORÍA COMERCIAL Y OPERACIONES DE ENTIDADES FINANCIERA": "TC EN ASERIA COMR OPER ENT FIN",
    "ESPECIALISTA EN ATENCION INTEGRAL A LA PRIMERA INFANCIA": "ESP ATNCION INTGRL PRIME INFA",
    "TECNICO LABORAL POR COMPETENCIAS AUXILIAR EN SEGURIDAD OCUPACIONAL Y LABORAL": "TC LAB COMP SEGU OCUP Y LAB",
    "Especialización en Pedagogía de la Lúdica para el desarrollo cultural": "ESP PEDAGO LUD DESARR CULTURAL",
    "TECNOLOGO EN GESTION DE LA SEGURIDAD Y SALUD EN EL TRABAJO": "TG GSTON SEGU Y SALUD TRABAJO",
    "ESP GERENCIA FINANCIERAV VIR": "ESP GERENCIA FINANCIERAV VIR",
    "tecnologo en analisis y desarrollo de software": "TG ANALISIS Y DESARROLLO SOFTW",
    "tecnologo en gestion contable y de informacion financiera": "TG GSTON CONTABL Y INFO FINANC"
}


mensajes_resumen_procesamiento = []
# <<< CAMBIO: unprocessed_files_paths ahora almacenará diccionarios para un mejor seguimiento de fallos.
unprocessed_files_paths = [] 
processed_data_cache = [] 
active_api_key_index = 1
active_openrouter_key_index = 1
thread_lock = threading.Lock()

# <<< CAMBIO: Se elimina openrouter_rate_limited ya que no hay fallback. Cada worker es independiente.
# openrouter_rate_limited = threading.Event()

class RateLimitException(Exception):
    pass
class ServerSideApiException(Exception):
    pass

def normalizar_texto_para_busqueda(texto):
    if not isinstance(texto, str): return ""
    return texto.lower().strip().replace('.', '')

ABREVIATURAS_NORMALIZADAS = {normalizar_texto_para_busqueda(k): v for k, v in ABREVIATURAS_PREDEFINIDAS.items()}

def quitar_tildes(texto):
    if not isinstance(texto, str): return texto
    nfkd_form = unicodedata.normalize('NFD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

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

def extraer_texto_pdf_con_apis(file_path):
    global mensajes_resumen_procesamiento, active_api_key_index
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
        if (codigo_estado_api == 402 or "breached your free processing limit" in msg_error_api):
            if current_api_key_index == 1:
                print(f"⚠️ LÍMITE ALCANZADO en API Key #1. Cambiando a API Key #2 para '{os.path.basename(file_path)}' y subsiguientes.")
                with thread_lock:
                    mensajes_resumen_procesamiento.append(f"⚠️ LÍMITE API Unstract #1. Cambiando a API #2.")
                    globals()['active_api_key_index'] = 2
                try:
                    return _intentar_extraccion_llmwhisperer(file_path, LLMWHISPERER_API_KEY_2, 2)
                except Exception as e2:
                    msg = f"❌ Fallo en el reintento con API Key #2 para '{os.path.basename(file_path)}': {e2}"
                    print(msg)
                    traceback.print_exc()
                    with thread_lock:
                        mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
                    return None
            else: # current_api_key_index == 2
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
        print(msg)
        traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
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
        print(msg)
        traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

def limpiar_texto_general_para_llm(texto):
    if not texto: return ""
    return re.sub(r'\n{3,}', '\n', texto).strip()

def analizar_con_llm(texto_documento, pregunta_o_instruccion, model_name, api_url, nombre_archivo_base, force_api_key_num=None):
    global mensajes_resumen_procesamiento, active_openrouter_key_index

    with thread_lock:
        current_api_key_num = force_api_key_num if force_api_key_num else active_openrouter_key_index
        if current_api_key_num == 1:
            api_token = OPENROUTER_API_KEY
        elif current_api_key_num == 2 and OPENROUTER_API_KEY_2:
            api_token = OPENROUTER_API_KEY_2
        else:
            msg = f"❌ No hay una clave de API de OpenRouter válida (intentando usar la #{current_api_key_num}) para '{nombre_archivo_base}'."
            print(msg)
            mensajes_resumen_procesamiento.append(msg)
            return None

    print(f"\nPreparando para enviar a OpenRouter para '{nombre_archivo_base}' usando modelo '{model_name}' (API Key #{current_api_key_num})...")
    prompt_completo = f"Aquí tienes el contenido de un documento (originalmente '{nombre_archivo_base}'):\n--- INICIO DEL DOCUMENTO ---\n{texto_documento}\n--- FIN DEL DOCUMENTO ---\nPor favor, sigue estas instrucciones detalladamente basadas ÚNICAMENTE en el documento proporcionado:\n{pregunta_o_instruccion}"

    headers = {
        "Authorization": f"Bearer {api_token}", "Content-Type": "application/json",
        "HTTP-Referer": YOUR_SITE_URL, "X-Title": YOUR_SITE_NAME
    }
    data = {
        "model": model_name, "messages": [{"role": "user", "content": prompt_completo}],
        "stream": False, "max_tokens": 4000, "temperature": 0.15
    }

    try:
        print(f"Enviando solicitud a OpenRouter (modelo: {model_name}, Key #{current_api_key_num}) para '{nombre_archivo_base}'...")
        response = requests.post(api_url, headers=headers, data=json.dumps(data), timeout=360)
        response.raise_for_status()
        respuesta_json = response.json()
        print(f"✅ Respuesta recibida de OpenRouter API para '{nombre_archivo_base}'.")
        if respuesta_json.get("choices") and len(respuesta_json["choices"]) > 0 and respuesta_json["choices"][0].get("message") and "content" in respuesta_json["choices"][0]["message"]:
            return respuesta_json["choices"][0]["message"]["content"]
        else:
            msg = f"❌ Error API OpenRouter '{nombre_archivo_base}': Respuesta sin formato esperado."
            print(msg)
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg)
            return None
    except requests.exceptions.Timeout as e:
        msg = f"❌ Error API OpenRouter '{nombre_archivo_base}': Timeout."
        print(msg)
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg)
        raise ServerSideApiException(msg) from e
    except requests.exceptions.RequestException as e:
        status_code = e.response.status_code if e.response is not None else 0
        details = f"Detalles: {status_code} - {e.response.text}" if e.response is not None else ""
        msg = f"❌ Error API OpenRouter '{nombre_archivo_base}' (Modelo: {model_name}, Key: #{current_api_key_num}): {e}"
        
        print(msg)
        print(details)
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + "\n" + details)
        
        if status_code == 429:
            raise RateLimitException(msg) from e
        if 500 <= status_code < 600:
            raise ServerSideApiException(msg) from e
        return None
    except Exception as e:
        msg = f"❌ Error inesperado con OpenRouter API para '{nombre_archivo_base}': {e}"
        print(msg)
        traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

# --- Se eliminan las funciones analizar_con_chutes y analizar_con_chutes_sync ---
# <<< CAMBIO: Se mantiene la función de Gemini, pero ya no se usa en el flujo principal del worker.
# Podría ser asignada a un worker si se añade "gemini-1.5-pro" a la lista WORKER_MODELS.
def analizar_con_gemini(texto_documento, pregunta_o_instruccion, nombre_archivo_base):
    global mensajes_resumen_procesamiento
    if not GEMINI_API_KEY:
        msg = f"❌ API de Gemini no configurada para '{nombre_archivo_base}'. Fallback no disponible."
        print(msg)
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg)
        return None

    print(f"\n⚡️ Intentando con Google Gemini para '{nombre_archivo_base}'...")
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
        ]
        generation_config = {"temperature": 0.15, "max_output_tokens": 4000}
        model = genai.GenerativeModel(
            model_name="gemini-1.5-pro-latest", generation_config=generation_config, safety_settings=safety_settings
        )
        prompt_completo = f"Aquí tienes el contenido de un documento (originalmente '{nombre_archivo_base}'):\n--- INICIO DEL DOCUMENTO ---\n{texto_documento}\n--- FIN DEL DOCUMENTO ---\nPor favor, sigue estas instrucciones detalladamente basadas ÚNICAMENTE en el documento proporcionado:\n{pregunta_o_instruccion}"
        response = model.generate_content(prompt_completo)
        
        if not response.parts:
            block_reason = response.prompt_feedback.block_reason if response.prompt_feedback else "No especificada"
            msg = f"❌ La respuesta de Gemini para '{nombre_archivo_base}' fue bloqueada. Razón: {block_reason}"
            print(msg)
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg)
            return None

        print(f"✅ Respuesta recibida de Google Gemini API para '{nombre_archivo_base}'.")
        return response.text
    except Exception as e:
        msg = f"❌ Error fatal durante el análisis con Google Gemini para '{nombre_archivo_base}': {e}"
        print(msg)
        traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None


def extraer_datos_completos(respuesta_llm):
    def limpiar_valor(texto): return texto.strip().strip('*').strip()
    
    datos = {
        "Nombre": "No extraído", "Programa": "No extraído", "Plan": "N/A",
        "Programa Origen": "No extraído", "Abreviacion Sugerida": "ABREV_NO_PROPORCIONADA",
        "Creditos": "0", "Script Template": None
    }
    if not respuesta_llm: return datos

    match_nombre = re.search(r"NOMBRE_ESTUDIANTE:\s*(.*)", respuesta_llm)
    if match_nombre: datos["Nombre"] = limpiar_valor(match_nombre.group(1).splitlines()[0])
    
    match_programa = re.search(r"PROGRAMA_ASPIRA:\s*(.*)", respuesta_llm)
    if match_programa: datos["Programa"] = limpiar_valor(match_programa.group(1).splitlines()[0])

    match_plan = re.search(r"PLAN_ESTUDIO:\s*(.*)", respuesta_llm)
    if match_plan: datos["Plan"] = limpiar_valor(match_plan.group(1).strip().splitlines()[0]) or "N/A"

    match_origen = re.search(r"NOMBRE_PROGRAMA_ORIGEN:\s*(.*)", respuesta_llm)
    if match_origen: datos["Programa Origen"] = limpiar_valor(match_origen.group(1).strip().splitlines()[0])
    
    match_abrev = re.search(r"ABREVIACION_SUGERIDA:\s*(.*)", respuesta_llm)
    if match_abrev: datos["Abreviacion Sugerida"] = limpiar_valor(match_abrev.group(1).strip().splitlines()[0])

    match_creditos = re.search(r"CREDITOS_HOMOLOGADOS:\s*(.*)", respuesta_llm)
    if match_creditos:
        creditos_str = limpiar_valor(match_creditos.group(1).strip().splitlines()[0])
        datos["Creditos"] = creditos_str if creditos_str.isdigit() else "0"

    match_script = re.search(r"--- INICIO SCRIPT AUTOHOTKEY ---(.*?)--- FIN SCRIPT AUTOHOTKEY ---", respuesta_llm, re.DOTALL)
    if match_script:
        datos["Script Template"] = match_script.group(1).strip()
    return datos

def sanitizar_nombre_archivo(nombre):
    nombre_sanitizado = re.sub(r'[<>:"/\\|?*]', '', nombre).replace("\n", " ").replace("\r", " ").strip()
    return re.sub(r'\s+', ' ', nombre_sanitizado) or "ScriptSinNombreValido"

def save_unprocessed_files(files_to_save, parent_window):
    # <<< CAMBIO: Ahora `files_to_save` es una lista de diccionarios.
    if not files_to_save:
        messagebox.showinfo("Información", "No hay archivos no procesados para guardar.", parent=parent_window)
        return

    dest_folder = filedialog.askdirectory(title="Seleccione una carpeta para guardar los archivos no procesados")
    if not dest_folder:
        return 

    # Extraer solo las rutas de los diccionarios
    file_paths = [item['path'] for item in files_to_save]
    copied_count = 0
    errors = []
    for file_path in file_paths:
        try:
            shutil.copy(file_path, dest_folder)
            copied_count += 1
        except Exception as e:
            errors.append(f"No se pudo copiar '{os.path.basename(file_path)}': {e}")
    
    summary_message = f"{copied_count} de {len(file_paths)} archivos guardados exitosamente en:\n{dest_folder}"
    if errors:
        summary_message += "\n\nOcurrieron los siguientes errores:\n" + "\n".join(errors)
        messagebox.showerror("Error al Guardar", summary_message, parent=parent_window)
    else:
        messagebox.showinfo("Proceso Completado", summary_message, parent=parent_window)

# <<< NUEVA FUNCIÓN: Para guardar los resultados en un archivo CSV
def guardar_resultados_csv(datos_a_guardar, directorio_salida, ventana_padre):
    if not datos_a_guardar:
        print("No hay datos en caché para guardar en el archivo de resultados.")
        return

    try:
        # Usar una marca de tiempo para crear un nombre de archivo único
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        nombre_archivo = f"Resultados_Procesamiento_{timestamp}.csv"
        ruta_completa_salida = os.path.join(directorio_salida, nombre_archivo)

        # Preparar los datos: Usar una copia para no modificar el caché original
        datos_para_df = [dict(row) for row in datos_a_guardar]
        
        # Eliminar datos que no son relevantes para el CSV final
        for row in datos_para_df:
            row.pop("Script Template", None)
            row.pop("Ruta Completa", None)
            row.pop("Abreviacion Sugerida", None)

        df = pd.DataFrame(datos_para_df)
        
        # Reordenar y renombrar columnas para que coincidan con la tabla de la GUI
        column_map = {
            "Archivo Origen": "Archivo Origen",
            "Modelo Usado": "Modelo Usado",
            "Nombre": "Nombre del estudiante",
            "Programa": "Programa al que aspira",
            "Plan": "Plan",
            "Nivel": "Nivel",
            "Programa Origen": "Programa Origen",
            "Creditos": "Créditos Homologados"
        }
        
        # Filtrar solo las columnas que existen en el DataFrame
        df_columnas_existentes = [col for col in column_map.keys() if col in df.columns]
        df = df[df_columnas_existentes]
        df.rename(columns=column_map, inplace=True)

        # Guardar en CSV con codificación utf-8-sig para compatibilidad con Excel
        df.to_csv(ruta_completa_salida, index=False, encoding='utf-8-sig')
        
        print(f"✅ Resultados guardados exitosamente en: {ruta_completa_salida}")
        messagebox.showinfo(
            "Resultados Guardados",
            f"Un resumen de los resultados ha sido guardado en:\n\n{ruta_completa_salida}",
            parent=ventana_padre
        )

    except Exception as e:
        print(f"❌ Error al guardar el archivo de resultados CSV: {e}")
        traceback.print_exc()
        messagebox.showerror(
            "Error al Guardar Resultados",
            f"No se pudo guardar el archivo de resumen CSV.\n\nError: {e}",
            parent=ventana_padre
        )

def mostrar_resumen_log_gui(main_root_ref, lista_mensajes, unprocessed_list):
    if not lista_mensajes:
        lista_mensajes = ["No se procesaron archivos o no hubo mensajes de resumen."]

    resumen_ventana = tk.Toplevel(main_root_ref)
    resumen_ventana.title("Resumen del Procesamiento (Log)")
    resumen_ventana.geometry("800x600")

    text_frame = ttk.Frame(resumen_ventana)
    text_frame.pack(padx=10, pady=(10, 0), fill=tk.BOTH, expand=True)
    
    txt_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, width=100, height=30)
    txt_area.pack(fill=tk.BOTH, expand=True)

    lista_mensajes_ordenada = sorted(lista_mensajes)
    for msg in lista_mensajes_ordenada:
        txt_area.insert(tk.END, msg + "\n" + "-"*50 + "\n")
    txt_area.config(state=tk.DISABLED)

    button_frame = ttk.Frame(resumen_ventana)
    button_frame.pack(pady=10, fill=tk.X, padx=10)

    save_button = ttk.Button(
        button_frame, 
        text=f"Guardar {len(unprocessed_list)} Archivos No Procesados",
        command=lambda: save_unprocessed_files(unprocessed_list, resumen_ventana)
    )
    save_button.pack(side=tk.LEFT, padx=(0, 10), expand=True, fill=tk.X)
    if not unprocessed_list:
        save_button.config(state=tk.DISABLED)
    
    close_button = ttk.Button(button_frame, text="Cerrar Log", command=resumen_ventana.destroy)
    close_button.pack(side=tk.RIGHT, expand=True, fill=tk.X)

    resumen_ventana.transient(main_root_ref)
    resumen_ventana.grab_set()

def agregar_fila_a_tabla(ventana_principal, tabla_treeview, datos_fila):
    def insertar():
        with thread_lock:
            # Eliminar cualquier entrada previa para este archivo para evitar duplicados en reintentos
            if tabla_treeview.exists(datos_fila["Ruta Completa"]):
                tabla_treeview.delete(datos_fila["Ruta Completa"])

            # Agregar o actualizar en el caché de datos procesados
            # Esto asegura que `show_last_results_table` tenga la última información
            # (ya sea éxito o el último fallo)
            ruta_a_buscar = datos_fila["Ruta Completa"]
            # Elimina cualquier registro anterior del mismo archivo
            processed_data_cache[:] = [d for d in processed_data_cache if d.get("Ruta Completa") != ruta_a_buscar]
            processed_data_cache.append(datos_fila)
            
        programa_original = datos_fila.get("Programa", "N/A")
        programa_normalizado = programa_original.lower().strip()
        tag_color = 'verde' if programa_normalizado in PROGRAMAS_VERDES else 'naranja'
        programa_con_indicador = f"{programa_original.strip()} ●"
        
        nivel_estudio = datos_fila.get("Nivel", "Pregrado")
        
        valores = (
            datos_fila.get("Archivo Origen", "N/A"), datos_fila.get("Modelo Usado", "N/A"),
            datos_fila.get("Nombre", "N/A"),
            programa_con_indicador, datos_fila.get("Plan", "N/A"),
            nivel_estudio, "Abrir Archivo"
        )
        tags_finales = ('accion', datos_fila.get("Ruta Completa", ""), tag_color)
        
        # <<< CAMBIO: Usar la ruta completa del archivo como un ID único (iid) para la fila.
        # Esto permite encontrar y eliminar la fila fácilmente si se necesita reintentar.
        tabla_treeview.insert("", "end", values=valores, tags=tags_finales, iid=datos_fila["Ruta Completa"])
        tabla_treeview.yview_moveto(1.0)
    ventana_principal.after(0, insertar)

def crear_tabla_resumen_en_vivo(main_root_ref, retry_callback):
    tabla_ventana = tk.Toplevel(main_root_ref)
    tabla_ventana.title("Tabla Resumen de Archivos Procesados (En Vivo)")
    tabla_ventana.geometry("1300x500") # Ancho aumentado y alto para el botón
    
    # Frame principal que contiene la tabla y los botones
    main_frame = ttk.Frame(tabla_ventana, padding="10")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Frame para la tabla y las barras de scroll
    tree_frame = ttk.Frame(main_frame)
    tree_frame.pack(fill=tk.BOTH, expand=True)
    
    cols = ("Archivo Origen", "Modelo Usado", "Nombre del estudiante", "Programa al que aspira", "Plan", "Nivel", "Acciones")
    tree = ttk.Treeview(tree_frame, columns=cols, show='headings')

    def sort_by_column(tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)
        tv.heading(col, command=lambda: sort_by_column(tv, col, not reverse))

    for col in cols:
        tree.heading(col, text=col, command=lambda _col=col: sort_by_column(tree, _col, False))
        if col == "Acciones": tree.column(col, width=100, minwidth=80, anchor='center')
        elif col == "Plan": tree.column(col, width=100, minwidth=80, anchor='center')
        elif col == "Nivel": tree.column(col, width=100, minwidth=80, anchor='center')
        elif col == "Programa al que aspira": tree.column(col, width=250, minwidth=200, anchor='w')
        elif col == "Modelo Usado": tree.column(col, width=150, minwidth=120, anchor='w')
        else: tree.column(col, width=200, minwidth=150, anchor='w')

    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
    style.configure("Treeview", rowheight=25, font=("Segoe UI", 9))
    
    tree.tag_configure('verde', background='#D5F5E3')
    tree.tag_configure('naranja', background='#FAE5D3')
    tree.tag_configure('accion', background="#E3F2FD")
    
    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    vsb.pack(side='right', fill='y')
    tree.configure(yscrollcommand=vsb.set)
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    hsb.pack(side='bottom', fill='x')
    tree.configure(xscrollcommand=hsb.set)
    tree.pack(fill=tk.BOTH, expand=True)

    # Frame para los botones debajo de la tabla
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=(10, 0))

    # <<< CAMBIO: Se añade el botón de reintento
    retry_button = ttk.Button(
        button_frame,
        text="Reintentar Archivos Fallidos",
        command=retry_callback,
        state=tk.DISABLED
    )
    retry_button.pack(side=tk.RIGHT, padx=5)

    def on_click(event):
        if tree.identify_region(event.x, event.y) != "cell": return
        if tree.identify_column(event.x) != '#7': return
        rowid = tree.identify_row(event.y)
        if not rowid: return
        item_tags = tree.item(rowid, 'tags')
        ruta_archivo_original = item_tags[1] if len(item_tags) > 1 else None
        if not ruta_archivo_original:
            messagebox.showwarning("Advertencia", "No se pudo encontrar la ruta del archivo original.", parent=tabla_ventana)
            return
        if os.path.exists(ruta_archivo_original):
            try:
                if os.name == 'nt': os.startfile(ruta_archivo_original)
                elif os.name == 'posix': subprocess.run(['xdg-open', ruta_archivo_original], check=True)
                else: subprocess.run(['open', ruta_archivo_original], check=True)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}", parent=tabla_ventana)
        else:
            messagebox.showwarning("Archivo no encontrado", f"No se encontró el archivo:\n{ruta_archivo_original}", parent=tabla_ventana)
    
    tree.bind("<Button-1>", on_click)
    
    # <<< CAMBIO: Devolver el árbol y el botón para poder controlarlos desde fuera
    return tree, retry_button

def determinar_nivel_academico(texto_documento):
    """
    Determina el nivel académico buscando palabras clave de posgrado en todo el documento.
    Esta versión es más robusta contra errores de OCR en las etiquetas.
    """
    texto_lower = texto_documento.lower()
    
    # Busca cualquiera de las palabras clave de posgrado en TODO el texto.
    # Se usan límites de palabra (\b) para evitar coincidencias parciales (ej. "especial" en lugar de "especializacion").
    if any(re.search(r'\b' + p + r'\b', texto_lower) for p in PALABRAS_CLAVE_POSGRADO):
        return "Posgrado"
    else:
        return "Pregrado"

# --- 2. MODIFICACIÓN DEL WORKER (CAMBIO PRINCIPAL) ---
# <<< CAMBIO: La función ahora acepta `model_a_usar` para saber qué modelo ejecutar.
def worker_process_file(archivo_queue, ventana_principal, tabla_treeview, barra_progreso, etiqueta_estado, num_archivos, model_a_usar, worker_id, is_retry=False):
    global mensajes_resumen_procesamiento, unprocessed_files_paths

    while not archivo_queue.empty():
        try:
            # <<< CAMBIO: La cola ahora puede contener solo la ruta o más info para reintentos.
            queue_item = archivo_queue.get_nowait()
            if is_retry:
                i, archivo_actual, model_a_usar = queue_item
            else:
                i, archivo_actual = queue_item
        except queue.Empty:
            break

        nombre_base_archivo = os.path.basename(archivo_actual)
        ventana_principal.after(0, lambda n=nombre_base_archivo, m=model_a_usar.split('/')[0]: etiqueta_estado.config(text=f"Procesando: {n} (con {m})"))
        mensaje_log_actual = [f"\n--- [Worker #{worker_id} | Modelo: {model_a_usar}] Procesando archivo {nombre_base_archivo} ---"]
        
        texto_extraido = None
        if nombre_base_archivo.lower().endswith('.pdf'):
            texto_extraido = extraer_texto_pdf_con_apis(archivo_actual)
        elif nombre_base_archivo.lower().endswith(('.xlsx', '.xls')):
            texto_extraido = extraer_texto_excel_con_pandas(archivo_actual)

        datos_para_fila_actual = {
            "Archivo Origen": nombre_base_archivo,
            "Ruta Completa": archivo_actual,
            "Modelo Usado": model_a_usar
        }
        
        is_processed_successfully = False

        if texto_extraido:
            texto_limpio_para_llm = limpiar_texto_general_para_llm(texto_extraido)
            if texto_limpio_para_llm:
                nivel_academico = determinar_nivel_academico(texto_limpio_para_llm)
                datos_para_fila_actual["Nivel"] = nivel_academico
                mensaje_log_actual.append(f"INFO: Nivel académico clasificado como: {nivel_academico}.")
                prompt_a_usar = PROMPT_ANALISIS_POSTGRADO if nivel_academico == "Posgrado" else PROMPT_ANALISIS_DOCUMENTO
                if nivel_academico == "Posgrado": mensaje_log_actual.append("INFO: Usando prompt especializado para Postgrado.")
                
                respuesta_llm = None
                
                try:
                    time.sleep(1) # Pequeña pausa para no saturar la API
                    mensaje_log_actual.append(f"INFO: Intentando análisis con el modelo asignado: {model_a_usar}.")
                    respuesta_llm = analizar_con_llm(texto_limpio_para_llm, prompt_a_usar, model_a_usar, OPENROUTER_API_URL, nombre_base_archivo)
                except (RateLimitException, ServerSideApiException) as e:
                    mensaje_log_actual.append(f"❌ FALLO DE API con el modelo {model_a_usar}: {e}. El archivo no será procesado por este worker.")
                    respuesta_llm = None
                except Exception as e:
                    mensaje_log_actual.append(f"❌ ERROR INESPERADO con el modelo {model_a_usar}: {e}")
                    respuesta_llm = None
                
                if respuesta_llm:
                    datos_completos = extraer_datos_completos(respuesta_llm)
                    datos_para_fila_actual.update(datos_completos)
                    
                    if "DETENER_PROCESO_CONDICION_" in respuesta_llm:
                        for line in respuesta_llm.splitlines():
                            if "DETENER_PROCESO_CONDICION_" in line:
                                detalle = line.strip().upper()
                                if "PERIODO" in detalle: datos_para_fila_actual["Plan"] = "DETENIDO: PERIODO"
                                elif "PROGRAMA" in detalle: datos_para_fila_actual["Plan"] = "DETENIDO: PROGRAMA"
                                elif "CALIFICACION" in detalle: datos_para_fila_actual["Plan"] = "DETENIDO: CALIFICACION"
                                else: datos_para_fila_actual["Plan"] = "DETENIDO: OTRO"
                                break
                    else:
                        if datos_completos["Script Template"] and datos_completos.get("Nombre") != "No extraído":
                            programa_origen = datos_completos.get("Programa Origen", "")
                            abreviatura = ""

                            if len(programa_origen) <= 30:
                                abreviatura = programa_origen
                                mensaje_log_actual.append(f"INFO: Abreviatura: Usando nombre de programa de origen original (<=30 caracteres).")
                            else:
                                nombre_origen_norm = normalizar_texto_para_busqueda(programa_origen)
                                if nombre_origen_norm in ABREVIATURAS_NORMALIZADAS:
                                    abreviatura = ABREVIATURAS_NORMALIZADAS[nombre_origen_norm]
                                    mensaje_log_actual.append(f"INFO: Abreviatura: Encontrada en la lista predefinida.")
                                else:
                                    abreviatura = datos_completos["Abreviacion Sugerida"]
                                    mensaje_log_actual.append(f"INFO: Abreviatura: Usando sugerencia del LLM (no encontrada en predefinidas).")
                            
                            abreviatura = quitar_tildes(abreviatura.strip())[:30]
                            encabezado = f"Send, 1{{Tab}}HOM01{{Tab}}HOM01{{Tab}}{datos_completos['Creditos']}{{Tab}}{{Tab}}{{Tab}}{abreviatura}{{F10}}^{{PgDn}}{{Space}}{{Tab}}"
                            script_final = datos_completos["Script Template"].replace('(LINEA_ENCABEZADO_AQUI)', encabezado)
                            
                            nombre_sanitizado = sanitizar_nombre_archivo(datos_completos["Nombre"])
                            ruta_ahk = os.path.join(os.path.dirname(archivo_actual), f"{nombre_sanitizado}.ahk")
                            try:
                                with open(ruta_ahk, "w", encoding="utf-8") as f: f.write(script_final)
                                mensaje_log_actual.append(f"✅ Script AHK guardado en: {ruta_ahk}")
                                is_processed_successfully = True
                            except Exception as e:
                                mensaje_log_actual.append(f"❌ Error guardando AHK: {e}")
                        else:
                            mensaje_log_actual.append(f"❌ No se generó script: faltan datos clave (Nombre/Plantilla).")
                else:
                    mensaje_log_actual.append(f"❌ Falló el API del modelo {model_a_usar}. No se pudo procesar '{nombre_base_archivo}'.")
                    datos_para_fila_actual.update({"Nombre": f"Error en API ({model_a_usar.split('/')[0]})", "Programa": "N/A", "Plan": "N/A"})
            else:
                datos_para_fila_actual.update({"Nombre": "Texto vacío", "Programa": "N/A", "Plan": "N/A"})
        else:
            datos_para_fila_actual.update({"Nombre": "Error Extracción", "Programa": "N/A", "Plan": "N/A"})

        if not is_processed_successfully:
            with thread_lock:
                # <<< CAMBIO: Almacenar un diccionario con detalles del fallo.
                failure_info = {
                    "path": archivo_actual,
                    "failed_model": "EXTRACTION_FAILURE" if not texto_extraido else model_a_usar
                }
                # Evitar duplicados en la lista de no procesados
                if not any(d['path'] == failure_info['path'] for d in unprocessed_files_paths):
                    unprocessed_files_paths.append(failure_info)
        
        agregar_fila_a_tabla(ventana_principal, tabla_treeview, datos_para_fila_actual)
        with thread_lock:
            mensajes_resumen_procesamiento.extend(mensaje_log_actual)
        archivo_queue.task_done()
        ventana_principal.after(0, barra_progreso.step)

def get_next_model(failed_model):
    """Obtiene el siguiente modelo de la lista para el reintento."""
    try:
        current_index = WORKER_MODELS.index(failed_model)
        next_index = (current_index + 1) % len(WORKER_MODELS)
        return WORKER_MODELS[next_index]
    except ValueError:
        # Si el modelo que falló no está en la lista (p.ej. fallo de extracción),
        # se devuelve el primer modelo como opción por defecto.
        return WORKER_MODELS[0]

def procesar_archivos_seleccionados(lista_archivos, boton_seleccionar, boton_ver_tabla, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal, tabla_treeview, retry_button, is_retry=False):
    global unprocessed_files_paths, processed_data_cache, original_file_list_for_saving
    
    if not is_retry:
        mensajes_resumen_procesamiento.clear()
        unprocessed_files_paths.clear() 
        processed_data_cache.clear() 
        active_api_key_index = 1
        active_openrouter_key_index = 1
        boton_ver_tabla.config(state=tk.DISABLED)
        # Guardar la lista original para saber dónde guardar el CSV final
        original_file_list_for_saving = lista_archivos
    
    num_archivos = len(lista_archivos)
    if barra_progreso:
        barra_progreso['maximum'] = num_archivos
        barra_progreso['value'] = 0
    if etiqueta_etr:
        etiqueta_etr.config(text=f"Procesando con {MAX_WORKERS} workers...")

    def procesar_en_paralelo():
        archivos_queue = queue.Queue()
        if is_retry:
            # Para reintentos, la lista_archivos contiene tuplas (índice, ruta, modelo)
            for item in lista_archivos:
                archivos_queue.put(item)
        else:
            # Para el procesamiento inicial, solo contiene rutas
            for i, archivo in enumerate(lista_archivos):
                archivos_queue.put((i, archivo))
        
        threads = []
        num_workers_a_usar = min(MAX_WORKERS, archivos_queue.qsize())

        for i in range(num_workers_a_usar):
            model_name = WORKER_MODELS[i % len(WORKER_MODELS)]
            thread = threading.Thread(
                target=worker_process_file,
                args=(archivos_queue, ventana_principal, tabla_treeview, barra_progreso, etiqueta_estado, num_archivos, model_name, i + 1, is_retry),
                daemon=True
            )
            threads.append(thread)
            thread.start()

        for t in threads:
            t.join()

        def actualizar_gui_final():
            etiqueta_estado.config(text=f"Procesamiento completado para {num_archivos} archivos.")
            etiqueta_etr.config(text="Proceso finalizado.")
            boton_seleccionar.config(state=tk.NORMAL)
            boton_ver_tabla.config(state=tk.NORMAL)
            
            # <<< CAMBIO: Activar el botón de reintento si hay fallos.
            if unprocessed_files_paths:
                retry_button.config(state=tk.NORMAL)
                messagebox.showwarning("Proceso con Fallos", f"{len(unprocessed_files_paths)} archivo(s) no se procesaron correctamente.\nPuede usar el botón 'Reintentar Archivos Fallidos' en la ventana de resultados para intentarlo de nuevo.", parent=ventana_principal)
            else:
                retry_button.config(state=tk.DISABLED)

            # <<< CAMBIO: Llamada para guardar el CSV automáticamente
            if original_file_list_for_saving:
                directorio_salida = os.path.dirname(original_file_list_for_saving[0])
                guardar_resultados_csv(processed_data_cache, directorio_salida, ventana_principal)

            mostrar_resumen_log_gui(ventana_principal, mensajes_resumen_procesamiento, unprocessed_files_paths)
            
        ventana_principal.after(0, actualizar_gui_final)

    threading.Thread(target=procesar_en_paralelo, daemon=True).start()

def iniciar_procesamiento_en_hilo(boton_seleccionar, boton_ver_tabla, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal):
    
    # Esta variable se usará para pasar el botón de reintento entre funciones
    retry_button_ref = None

    def iniciar_reintento():
        global unprocessed_files_paths, mensajes_resumen_procesamiento
        
        files_to_retry_info = list(unprocessed_files_paths)
        if not files_to_retry_info:
            messagebox.showinfo("Información", "No hay archivos fallidos para reintentar.", parent=ventana_principal)
            return

        # Limpiar la lista de fallos antes de reintentar
        unprocessed_files_paths.clear()
        mensajes_resumen_procesamiento.append("\n" + "="*20 + " INICIANDO REINTENTO " + "="*20 + "\n")
        
        # Desactivar botones durante el reintento
        if retry_button_ref: retry_button_ref.config(state=tk.DISABLED)
        boton_seleccionar.config(state=tk.DISABLED)
        
        # Preparar la cola para el reintento
        archivos_para_reintentar = []
        for i, file_info in enumerate(files_to_retry_info):
            path = file_info['path']
            failed_model = file_info['failed_model']
            next_model = get_next_model(failed_model)
            
            # Eliminar la fila fallida anterior de la tabla
            if tabla_treeview.exists(path):
                tabla_treeview.delete(path)
                
            archivos_para_reintentar.append((i, path, next_model))

        procesar_archivos_seleccionados(
            archivos_para_reintentar, boton_seleccionar, boton_ver_tabla, etiqueta_estado, 
            barra_progreso, etiqueta_etr, ventana_principal, tabla_treeview, retry_button_ref, is_retry=True
        )

    if boton_seleccionar: boton_seleccionar.config(state=tk.DISABLED)
    if etiqueta_estado: etiqueta_estado.config(text="Seleccionando archivos...")
    file_paths = filedialog.askopenfilenames(
        title="Selecciona uno o más archivos PDF o Excel",
        filetypes=(("Documentos Soportados", "*.pdf *.xlsx *.xls"),("Todos los archivos", "*.*"))
    )
    if not file_paths:
        if etiqueta_estado: etiqueta_estado.config(text="Ningún archivo seleccionado.")
        if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)
        return
    
    # <<< CAMBIO: Crear la tabla y obtener el botón de reintento.
    tabla_treeview, retry_button_ref = crear_tabla_resumen_en_vivo(ventana_principal, iniciar_reintento)
    
    try:
        lista_ordenada = sorted(list(file_paths), key=lambda r: os.path.getmtime(r), reverse=True)
        procesar_archivos_seleccionados(
            lista_ordenada, boton_seleccionar, boton_ver_tabla, etiqueta_estado, barra_progreso, 
            etiqueta_etr, ventana_principal, tabla_treeview, retry_button_ref, is_retry=False
        )
    except FileNotFoundError as e:
        messagebox.showerror("Error de Archivo", f"No se encontró el archivo {e.filename}.", parent=ventana_principal)
        if etiqueta_estado: etiqueta_estado.config(text="Error al procesar.")
        if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)

def show_last_results_table(main_root_ref):
    if not processed_data_cache:
        messagebox.showinfo("Sin Resultados", "Aún no se ha procesado ningún archivo en esta sesión.", parent=main_root_ref)
        return

    # La función de reintento no tiene sentido al recrear una tabla estática,
    # por lo que pasamos una función vacía (lambda: None).
    tabla_recreada, _ = crear_tabla_resumen_en_vivo(main_root_ref, lambda: None)
    toplevel_window = tabla_recreada.winfo_toplevel()
    toplevel_window.title("Última Tabla de Resultados Guardada")

    for datos_fila in processed_data_cache:
        programa_original = datos_fila.get("Programa", "N/A")
        programa_normalizado = programa_original.lower().strip()
        tag_color = 'verde' if programa_normalizado in PROGRAMAS_VERDES else 'naranja'
        programa_con_indicador = f"{programa_original.strip()} ●"
        
        nivel_estudio = datos_fila.get("Nivel", "Pregrado")
        
        valores = (
            datos_fila.get("Archivo Origen", "N/A"), datos_fila.get("Modelo Usado", "N/A"),
            datos_fila.get("Nombre", "N/A"),
            programa_con_indicador, datos_fila.get("Plan", "N/A"),
            nivel_estudio, "Abrir Archivo"
        )
        tags_finales = ('accion', datos_fila.get("Ruta Completa", ""), tag_color)
        # Usar el iid también aquí para consistencia
        tabla_recreada.insert("", "end", values=valores, tags=tags_finales, iid=datos_fila.get("Ruta Completa"))

PROMPT_ANALISIS_DOCUMENTO = """Analiza el contenido del documento proporcionado (que ha sido convertido a texto, originado de un PDF o de un archivo Excel) y extrae los datos de código y calificación para cada una de las materias correspondientes a la sección: Programa Destino Ibero (que también puede llamarse simplemente Programa Destino en algunas ocasiones). Asegúrate de no incluir información de la sección Programa de Origen.
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
- Debemos traer al final los códigos tal cual como aparecen en el archivo no debes cambiar los números ni las letras de los códigos.
FORMATO DE SALIDA (si no hay condiciones de parada):
1.  **DATOS CLAVE PARA RESUMEN (Presentar antes de la tabla y el script AHK):**
    *   **NOMBRE_ESTUDIANTE:** [Nombre completo del estudiante como aparece en el documento]
    *   **PROGRAMA_ASPIRA:** [Nombre completo del programa al que aspira el estudiante]
    *   **PLAN_ESTUDIO:** [Número o identificador del plan de estudio asociado al programa al que aspira, si se menciona. Indique "N/A"]
    *   **NOMBRE_PROGRAMA_ORIGEN:** [Nombre completo y exacto del programa de origen extraído del documento]
    *   **CREDITOS_HOMOLOGADOS:** [Número total de créditos homologados extraído del documento]
    *   **ABREVIACION_SUGERIDA:** [Genera una abreviatura para el NOMBRE_PROGRAMA_ORIGEN siguiendo las reglas de sustitución detalladas más abajo. Máximo 30 caracteres.]

2.  **TABLA DE DATOS (Materias):**
    Crea una tabla con tres columnas para las materias del "Programa Destino Ibero":
    Primera columna: Contiene solamente las letras de los códigos.
    Segunda columna: Contiene solamente los números correspondientes a los códigos.
    Tercera columna: Contiene las calificaciones correspondientes a cada código.
    Las calificaciones deben formatearse siempre con un decimal:
    - Si la calificación es un número entero, debe mostrarse con un decimal (por ejemplo, 4 debe ser 4.0).
    - Si la calificación tiene dos decimales, redondea al primer decimal (por ejemplo, 4.15 se convierte en 4.2, 4.14 se convierte en 4.1).

3.  **PLANTILLA SCRIPT AUTOHOTKEY:**
    Genera una plantilla de script de AutoHotkey. ¡ATENCIÓN CRÍTICA AL DETALLE DEL ESPACIADO EN WINTITLE!
    En WinTitle debe tener 2 espacios luego de: (Services:)
    Delimita claramente el script de la siguiente manera:
    --- INICIO SCRIPT AUTOHOTKEY ---
    #SingleInstance force
    #NoEnv
    SendMode Input
    SetKeyDelay, 10
    WinTitle := "Oracle Fusion Middleware Forms Services:  Open > SHATRNS"
    WinWait, %WinTitle%
    WinActivate, %WinTitle%
    WinWaitActive, %WinTitle%
    ; La siguiente línea es un placeholder que será reemplazado por el código Python.
    (LINEA_ENCABEZADO_AQUI)
    ; Para cada materia extraída de la TABLA DE DATOS (Materias) anterior, genera una línea Send con la siguiente estructura:
    Send, (letras del codigo){Tab}(números del código){Tab}{Tab}(calificación){Tab}N{Down}{Space}{Tab}
    ; Para la última materia, omite {Down}{Space}{Tab} al final.
    --- FIN SCRIPT AUTOHOTKEY ---

---
REGLAS Y EJEMPLOS PARA GENERAR 'ABREVIACION_SUGERIDA':
    OBJETIVO: Crear una abreviatura de MÁXIMO 30 caracteres.

    REGLA ESPECIAL Y PRIORITARIA:
      - **SI** el nombre completo del programa de origen **CONTIENE** las palabras `NORMALISTA SUPERIOR`, **ENTONCES** la `ABREVIACION_SUGERIDA` **DEBE SER EXACTAMENTE** `NORMALISTA SUPERIOR`.

    REGLAS GENERALES:
      1. Omitir artículos (el, la, los, las), preposiciones comunes (de, en, a, por, para, con) y conjunciones (y, e, o, u) a menos que sean esenciales.
      2. Priorizar la legibilidad. No excedas los 30 caracteres.

    SUSTITUCIONES COMUNES:
      - TECNICO / TÉCNICO / TECNICA -> TC
      - TECNOLOGO / TECNÓLOGO / TECNOLOGIA / TECNOLOGÍA -> TG
      - ESPECIALIZACION / ESPECIALIZACIÓN -> ESP
      - LICENCIATURA -> LIC
      - MÁSTER / MASTER -> MA
      - INGENIERIA / INGENIERÍA -> ING
      - ADMINISTRACIÓN / ADMINISTRATIVA / ADMINISTRATIVO -> ADMON / ADMIN
      - GESTION / GESTIÓN -> GSTON / GSTN
      - CONTABLE / CONTABILIDAD / CONTABILIZACION -> CONT / CONTAB
      - FINANCIERA / FINANCIERO / FINANZAS -> FINAN / FINANC
      - DESARROLLO -> DESARR / DSRLLO
      - INFORMACIÓN / INFORMÁTICA -> INFO / INFORM
      - SOFTWARE -> SOFT / SOFTW
      - SISTEMAS -> SIST / SISTE
      - PRIMERA INFANCIA -> PR INFANC / PRIM INFAN

    EJEMPLOS DE APLICACIÓN:
      1.  "TECNICO EN CONTABILIZACION DE OPERACIONES COMERCIALES Y FINANCIERAS" -> "TC CONT OPER COMER Y FINANC"
      2.  "TECNÓLOGO EN GESTION CONTABLE Y FINANCIERA" -> "TG GSTON CONTBLE Y FINACIRA"
      3.  "TECNOLOGO EN ANALISIS Y DESARROLLO DE SOFTWARE" -> "TG ANALISIS Y DESARR DE SOFTWA"
"""
PROMPT_ANALISIS_POSTGRADO = """Analiza el contenido del documento proporcionado (que ha sido convertido a texto, originado de un PDF o de un archivo Excel) y extrae los datos de código y calificación para cada una de las materias correspondientes a la sección: Programa Destino Ibero (que también puede llamarse simplemente Programa Destino en algunas ocasiones). Asegúrate de no incluir información de la sección Programa de Origen.
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
2.  **Condición de Calificaciones Bajas o en Blanco:**
    *   Analice las calificaciones de las materias del "Programa Destino Ibero".
    *   **SI** encuentra CUALQUIER materia en el "Programa Destino Ibero" con una calificación en blanco (vacía) O con un valor numérico inferior a 3.5 (por ejemplo 3.0, 3.4, 2.9, 1.5, 0.0), ENTONCES:
        Notifique: "Se encontró una calificación no válida para la materia '{nombre_materia_con_problema}' con calificación '{calificacion_problematica}' en el Programa Destino Ibero."
        Indique: "DETENER_PROCESO_CONDICION_CALIFICACION"
        Y detenga todo procesamiento adicional.
    *   DE LO CONTRARIO, si todas las calificaciones son válidas (3.5 o superior y no en blanco), proceda con los siguientes pasos.
Si NINGUNA de las condiciones de parada anteriores se cumple, Y SOLO EN ESE CASO, proceda con lo siguiente:
PROCESAMIENTO DE DATOS:
Ten en cuenta las siguientes recomendaciones al procesar los datos de las materias del "Programa Destino Ibero":
- Algunos códigos pueden contener espacios; omite estos espacios para obtener solo letras y números.
- En los códigos de las materias, los números siempre preceden a las letras. Además, en algunos casos, puede aparecer una letra "F" al final de los números; esta letra debe ser omitida.
- Si el código de la materia tiene un número 0 al final de este debes preservarlo y no eliminarlo.
- Debemos traer al final los códigos tal cual como aparecen en el archivo no debes cambiar los números ni las letras de los códigos.
FORMATO DE SALIDA (si no hay condiciones de parada):
1.  **DATOS CLAVE PARA RESUMEN (Presentar antes de la tabla y el script AHK):**
    *   **NOMBRE_ESTUDIANTE:** [Nombre completo del estudiante como aparece en el documento]
    *   **PROGRAMA_ASPIRA:** [Nombre completo del programa al que aspira el estudiante]
    *   **PLAN_ESTUDIO:** [Número o identificador del plan de estudio asociado al programa al que aspira, si se menciona. Indique "N/A"]
    *   **NOMBRE_PROGRAMA_ORIGEN:** [Nombre completo y exacto del programa de origen extraído del documento]
    *   **CREDITOS_HOMOLOGADOS:** [Número total de créditos homologados extraído del documento]
    *   **ABREVIACION_SUGERIDA:** [Genera una abreviatura para el NOMBRE_PROGRAMA_ORIGEN siguiendo las reglas de sustitución detalladas más abajo. Máximo 30 caracteres.]

2.  **TABLA DE DATOS (Materias):**
    Crea una tabla con tres columnas para las materias del "Programa Destino Ibero":
    Primera columna: Contiene solamente las letras de los códigos.
    Segunda columna: Contiene solamente los números correspondientes a los códigos.
    Tercera columna: Contiene las calificaciones correspondientes a cada código.
    Las calificaciones deben formatearse siempre con un decimal:
    - Si la calificación es un número entero, debe mostrarse con dos decimales (por ejemplo, 4 debe ser 4.00).
    - Si la calificación tiene un decimal, agrega un 0 a la derecha (por ejemplo, 4.1 se convierte en 4.10, 3.6 se convierte en 3.60).

3.  **PLANTILLA SCRIPT AUTOHOTKEY:**
    Genera una plantilla de script de AutoHotkey. ¡ATENCIÓN CRÍTICA AL DETALLE DEL ESPACIADO EN WINTITLE!
    En WinTitle debe tener 2 espacios luego de: (Services:)
    Delimita claramente el script de la siguiente manera:
    --- INICIO SCRIPT AUTOHOTKEY ---
    #SingleInstance force
    #NoEnv
    SendMode Input
    SetKeyDelay, 10
    WinTitle := "Oracle Fusion Middleware Forms Services:  Open > SHATRNS"
    WinWait, %WinTitle%
    WinActivate, %WinTitle%
    WinWaitActive, %WinTitle%
    ; La siguiente línea es un placeholder que será reemplazado por el código Python.
    (LINEA_ENCABEZADO_AQUI)
    ; Para cada materia extraída de la TABLA DE DATOS (Materias) anterior, genera una línea Send con la siguiente estructura:
    Send, (letras del codigo){Tab}(números del código){Tab}{Tab}(calificación){Tab}P{Down}{Space}{Tab}
    ; Para la última materia, omite {Down}{Space}{Tab} al final.
    --- FIN SCRIPT AUTOHOTKEY ---

---
REGLAS Y EJEMPLOS PARA GENERAR 'ABREVIACION_SUGERIDA':
    OBJETIVO: Crear una abreviatura de MÁXIMO 30 caracteres.

    REGLA ESPECIAL Y PRIORITARIA:
      - **SI** el nombre completo del programa de origen **CONTIENE** las palabras `NORMALISTA SUPERIOR`, **ENTONCES** la `ABREVIACION_SUGERIDA` **DEBE SER EXACTAMENTE** `NORMALISTA SUPERIOR`.

    REGLAS GENERALES:
      1. Omitir artículos (el, la, los, las), preposiciones comunes (de, en, a, por, para, con) y conjunciones (y, e, o, u) a menos que sean esenciales.
      2. Priorizar la legibilidad. No excedas los 30 caracteres.

    SUSTITUCIONES COMUNES:
      - TECNICO / TÉCNICO / TECNICA -> TC
      - TECNOLOGO / TECNÓLOGO / TECNOLOGIA / TECNOLOGÍA -> TG
      - ESPECIALIZACION / ESPECIALIZACIÓN -> ESP
      - LICENCIATURA -> LIC
      - MÁSTER / MASTER -> MA
      - INGENIERIA / INGENIERÍA -> ING
      - ADMINISTRACIÓN / ADMINISTRATIVA / ADMINISTRATIVO -> ADMON / ADMIN
      - GESTION / GESTIÓN -> GSTON / GSTN
      - CONTABLE / CONTABILIDAD / CONTABILIZACION -> CONT / CONTAB
      - FINANCIERA / FINANCIERO / FINANZAS -> FINAN / FINANC
      - DESARROLLO -> DESARR / DSRLLO
      - INFORMACIÓN / INFORMÁTICA -> INFO / INFORM
      - SOFTWARE -> SOFT / SOFTW
      - SISTEMAS -> SIST / SISTE
      - PRIMERA INFANCIA -> PR INFANC / PRIM INFAN

    EJEMPLOS DE APLICACIÓN:
      1.  "TECNICO EN CONTABILIZACION DE OPERACIONES COMERCIALES Y FINANCIERAS" -> "TC CONT OPER COMER Y FINANC"
      2.  "TECNÓLOGO EN GESTION CONTABLE Y FINANCIERA" -> "TG GSTON CONTBLE Y FINACIRA"
      3.  "TECNOLOGO EN ANALISIS Y DESARROLLO DE SOFTWARE" -> "TG ANALISIS Y DESARR DE SOFTWA"
"""
# Variable global para almacenar la lista de archivos original
original_file_list_for_saving = []

if __name__ == "__main__":
    ventana_principal_app = tk.Tk()
    ventana_principal_app.title("Procesador de Documentos AHK")
    ventana_principal_app.geometry("500x300")
    frame_superior = ttk.Frame(ventana_principal_app, padding="10")
    frame_superior.pack(expand=True, fill=tk.BOTH)

    etiqueta_estado_gui = ttk.Label(frame_superior, text="Listo para iniciar.", font=("Segoe UI", 10))
    etiqueta_estado_gui.pack(pady=5)
    barra_progreso_gui = ttk.Progressbar(frame_superior, orient="horizontal", length=300, mode="determinate")
    barra_progreso_gui.pack(pady=10)
    etiqueta_etr_gui = ttk.Label(frame_superior, text="Tiempo restante: --:--", font=("Segoe UI", 9))
    etiqueta_etr_gui.pack(pady=5)
    
    botones_frame = ttk.Frame(frame_superior)
    botones_frame.pack(pady=10)

    boton_ver_tabla_gui = ttk.Button(
        botones_frame,
        text="Ver Última Tabla de Resultados",
        command=lambda: show_last_results_table(ventana_principal_app)
    )
    boton_ver_tabla_gui.pack(pady=5)
    boton_ver_tabla_gui.config(state=tk.DISABLED)

    boton_seleccionar_gui = ttk.Button(
        botones_frame,
        text="Seleccionar Archivos y Procesar",
        command=lambda: iniciar_procesamiento_en_hilo(
            boton_seleccionar_gui, boton_ver_tabla_gui, etiqueta_estado_gui, barra_progreso_gui,
            etiqueta_etr_gui, ventana_principal_app
        )
    )
    boton_seleccionar_gui.pack(pady=5)

    keys_faltantes = []
    if not LLMWHISPERER_API_KEY_1: keys_faltantes.append("LLMWHISPERER_API_KEY_1")
    if not OPENROUTER_API_KEY: keys_faltantes.append("OPENROUTER_API_KEY")
    if not GEMINI_API_KEY: keys_faltantes.append("GEMINI_API_KEY")

    if not OPENROUTER_API_KEY_2:
        print("ADVERTENCIA: No se encontró OPENROUTER_API_KEY_2 en .env. Fallback secundario de OpenRouter inactivo.")
    
    if keys_faltantes:
        mensaje_error = "ERROR: Faltan claves obligatorias en el archivo .env:\n\n" + "\n".join(keys_faltantes)
        etiqueta_estado_gui.config(text="ERROR: ¡Configure las claves API en .env!")
        boton_seleccionar_gui.config(state=tk.DISABLED)
        messagebox.showerror("Error de Configuración", mensaje_error, parent=ventana_principal_app)
    else:
        # <<< CAMBIO: El mensaje de inicio refleja la nueva configuración de workers.
        etiqueta_estado_gui.config(text=f"Claves OK. Listo para procesar con {MAX_WORKERS} workers.")

    ventana_principal_app.mainloop()