import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
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
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions"
PRIMARY_LLM_MODEL = "openrouter/cypher-alpha:free"
FALLBACK_LLM_MODEL = "google/gemma-3n-e2b-it:free"
YOUR_SITE_URL = "https://github.com/tu_usuario/tu_proyecto"
YOUR_SITE_NAME = "Procesador AHK"

# --- Diccionario de abreviaturas predefinidas en Python ---
ABREVIATURAS_PREDEFINIDAS = {
    "TECNICO EN CONTABILIZACION DE OPERACIONES COMERCIALES Y FINANCIERAS": "TC CONTAB OPERA COMER Y FINANC",
    "TECNÓLOGO EN GESTION CONTABLE Y FINANCIERA": "TG EN GSTON CONTBLE Y FINACIRA",
    "TECNOLOGO EN ANALISIS Y DESARROLLO DE SOFTWARE": "TG ANALISIS Y DESARR DE SOFTWA",
    "TECNOLOGO EN ANALISIS Y DESARROLLO DE SISTEMAS DE INFORMACIÓN": "TG ANALISIS Y DESARR SIST INFO",
    "Tecnología en Gestión Bancaria y de Entidades Financieras": "TG EN GSTON BANCARIA Y FINANCI",
    "TECNÓLOGO EN GESTION CONTABLE Y DE INFORMACION FINANCIERA": "TG EN GSTON CONTBLE INFO FINAN",
    "TÉCNICO EN SEGURIDAD VIAL, CONTROL DE TRANSITO Y TRANSPORTE": "TC SEGU VIAL CTRL TRANSIT TRAN",
    "TECNOLOGIA EN GESTION DE LA SEGURIDAD Y SALUD EN EL TRABAJO": "TG EN GSTON SEG SLUD EN TRABA",
    "TECNÓLOGO EN GESTIÓN DEL TALENTO HUMANO": "TG EN GSTON DEL TLENTO HUMANO",
    "TÉCNICO EN SOLDADURA DE PRODUCTOS METÁLICOS": "TC EN SOLDADUR PRODUCT METALIC",
    "TÉCNICO EN ATENCIÓN INTEGRAL A LA PRIMERA INFANCIA": "TC ATENCION INTGRL PRIME INFAN",
    "TECNICO LABORAL EN EDUCACION PARA LA PRIMERA INFANCIA": "TC EN EDUCION PRIMERA INFANCIA",
    "Especialización en educación e intervención para la primera infancia.": "ESP EDUCION INTRCION PRIME INF",
    "TECNOLOGO EN CONTABILIDAD Y FINANZAS": "TG EN CONTABILIDAD Y FINANZAS",
    "TÉCNICO INTEGRACION DE CONTENIDOS DIGITALES": "TC INTGRCION CNTENIDOS DIGITAL",
    "TÉCNICO EN ASISTENCIA ADMINISTRATIVA": "TC EN ASTENCIA ADMINISTRATIVA",
    "TECNICO LABORAL POR COMPETENCIAS EN AUXILIAR ADMINISTRATIVO EN SALUD": "TC LABRAL AUX ADMINIST SALUD",
    "TÉCNICO LABORAL POR COMPETENCIAS EN LOGÍSTICA Y ADMINISTRACIÓN DE ALMACENES": "TC LABRAL LOGISTI Y ADM ALMACE",
    "TECNICO LABORAL POR COMPETENCIAS EN EDUCACION INICIAL Y RECREACION": "TC EN EDUCION INCIAL RECRCION",
    "TECNOLOGÍA EN MANTENIMIENTO ELECTROMECÁNICO INDUSTRIAL": "TG EN MANTEMTO ELECTRMC INDUS",
    "TECNICO EN SISTEMAS AGROPECUARIOS ECOLOGICOS": "TC EN SISTEMA AGROPECU ECOLOGI",
    "TÉCNICO EN PROGRAMACION DE APLICACIONES PARA DISPOSITIVOS MOVILES": "TC PROGRAM APLICAC DISPOS MOVI",
    "Tecnología en Gestión Financiera y Tesorería": "TG EN GSTON FINACIRA Y TESORIA",
    "TECNOLOGÍA EN COORDINACIÓN DE PROCESOS LOGÍSTICOS": "TG EN COORDN PRCSOS LOGISTIC",
    "TÉCNICO EN SERVICIOS COMERCIALES Y FINANCIEROS": "TC EN SERVICI COMERC Y FINANCI",
    "TECNÓLOGO EN GESTIÓN INTEGRADA DE LA CALIDAD, MEDIO AMBIENTE, SEGURIDAD Y SALUD OCUPACIONAL": "TC GSTN CALID AMBI SEG Y SALD",
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
    "TÉCNICO EN MONTAJE Y MANTENIMIENTO DE REDES AEREAS DE DISTRIBUCIÓN DE ENERGÍA ELECTRICA": "TC MONT MTNTO RDES AERE DISTR",
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
    "TÉCNICO EN MONTAJE Y MANTENIMIENTO ELECTROMECANICO DE INSTALACIONES MINERAS BAJO TIERRA": "TC MONTJE MNTO ELECTR INST MIN",
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
    "TÉCNICO EN CONTABILIZACIÓN DE OPERACIONES COMERCIALES Y FINANCIERAS": "TC CONTA OPERAC COMER FINA",
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
    "TÉCNICO PROFESIONAL EN ATENCIÓN A LA PRIMERA INFANCIA": "TC PROFESI ATNCON PRIMER INFA",
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
    "TECNOLOGÍA GESTIÓN DE EMPRESAS AGROPECUARIAS": "TG GSTON EMPRESAS AGROPECUAR",
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
    # Añadir duplicados comunes para robustez
    "tecnologo en analisis y desarrollo de software": "TG ANALISIS Y DESARROLLO SOFTW", 
    "tecnologo en gestion contable y de informacion financiera": "TG GSTON CONTABL Y INFO FINANC"
}

mensajes_resumen_procesamiento = []
active_api_key_index = 1
MAX_WORKERS = 4
thread_lock = threading.Lock()
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
    # (Sin cambios)
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
    # (Sin cambios)
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
        if (codigo_estado_api == 402 or "breached your free processing limit" in msg_error_api) and current_api_key_index == 1:
            print(f"⚠️ LÍMITE ALCANZADO en API Key #1. Cambiando a API Key #2 para '{os.path.basename(file_path)}' y subsiguientes.")
            with thread_lock:
                mensajes_resumen_procesamiento.append(f"⚠️ LÍMITE API Unstract #1. Cambiando a API #2.")
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

def extraer_texto_excel_con_pandas(file_path):
    # (Sin cambios)
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

def analizar_con_llm(texto_documento, pregunta_o_instruccion, api_token, model_name, api_url, nombre_archivo_base):
    # (Sin cambios)
    global mensajes_resumen_procesamiento
    print(f"\nPreparando para enviar a la API de OpenRouter para '{nombre_archivo_base}' usando el modelo '{model_name}'...")
    prompt_completo = f"Aquí tienes el contenido de un documento (originalmente '{nombre_archivo_base}'):\n--- INICIO DEL DOCUMENTO ---\n{texto_documento}\n--- FIN DEL DOCUMENTO ---\nPor favor, sigue estas instrucciones detalladamente basadas ÚNICAMENTE en el documento proporcionado:\n{pregunta_o_instruccion}"

    headers = {
        "Authorization": f"Bearer {api_token}",
        "Content-Type": "application/json",
        "HTTP-Referer": YOUR_SITE_URL,
        "X-Title": YOUR_SITE_NAME
    }
    data = {
        "model": model_name,
        "messages": [{"role": "user", "content": prompt_completo}],
        "stream": False,
        "max_tokens": 4000,
        "temperature": 0.15
    }

    try:
        print(f"Enviando solicitud a OpenRouter (modelo: {model_name}) para '{nombre_archivo_base}'...")
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
        msg = f"❌ Error API OpenRouter '{nombre_archivo_base}': {e}"
        print(msg)
        if e.response is not None:
            print(f"Detalles: {e.response.status_code} - {e.response.text}")
            with thread_lock:
                mensajes_resumen_procesamiento.append(msg + (f"Detalles: {e.response.status_code} - {e.response.text}" if e.response is not None else ""))
        if 500 <= status_code < 600:
            raise ServerSideApiException(msg) from e
        return None
    except Exception as e:
        msg = f"❌ Error inesperado con OpenRouter API para '{nombre_archivo_base}': {e}"
        print(msg); traceback.print_exc()
        with thread_lock:
            mensajes_resumen_procesamiento.append(msg + f"\n{traceback.format_exc()}")
        return None

def extraer_tabla_materias(respuesta_llm):
    materias = []
    match = re.search(r"TABLA DE DATOS \(Materias\):(.*)", respuesta_llm, re.DOTALL | re.IGNORECASE)
    if not match:
        return materias
    
    bloque_tabla = match.group(1).strip()
    for linea in bloque_tabla.splitlines():
        linea = linea.strip()
        if not linea:
            continue
        
        partes = [p.strip() for p in linea.split('|')]
        letras = ""
        numeros = ""
        calificacion = ""

        # --- LÓGICA FINAL Y ROBUSTA ---

        # CASO 1: El formato es perfecto (3 columnas con datos).
        if len(partes) == 3 and partes[0] and partes[1]:
            letras = partes[0]
            numeros = partes[1]
            calificacion = partes[2]
        
        # CASO 2 (Fallback): El LLM juntó el código en la primera columna.
        elif len(partes) >= 2 and partes[0]:
            codigo_completo = partes[0]
            calificacion = partes[-1] # La calificación es la última parte
            
            # Intentamos dividir el código combinado
            match_codigo = re.match(r'^([A-Z]+)(\d+)$', codigo_completo)
            if match_codigo:
                letras = match_codigo.group(1)
                numeros = match_codigo.group(2)

        # --- Verificación final antes de agregar ---
        if letras and numeros and calificacion:
            materias.append({"letras": letras, "numeros": numeros, "calificacion": calificacion})
        else:
            # Esta advertencia solo aparecerá si ambos métodos fallan.
            print(f"ADVERTENCIA: No se pudo extraer la información completa de la línea: '{linea}'")
            
    return materias

def extraer_datos_completos(respuesta_llm):
    # (Sin cambios)
    def limpiar_valor(texto): return texto.strip().strip('*').strip()
    
    datos = {
        "Nombre": "No extraído",
        "Programa": "No extraído",
        "Plan": "N/A",
        "Programa Origen": "No extraído",
        "Abreviacion Sugerida": "ABREV_NO_PROPORCIONADA",
        "Creditos": "0",
        "Tabla Materias": [],
    }

    if not respuesta_llm: return datos

    match_nombre = re.search(r"NOMBRE_ESTUDIANTE:\s*(.*)", respuesta_llm)
    if match_nombre: datos["Nombre"] = limpiar_valor(match_nombre.group(1).splitlines()[0])
    
    match_programa = re.search(r"PROGRAMA_ASPIRA:\s*(.*)", respuesta_llm)
    if match_programa: datos["Programa"] = limpiar_valor(match_programa.group(1).splitlines()[0])

    match_plan = re.search(r"PLAN_ESTUDIO:\s*(.*)", respuesta_llm)
    if match_plan: datos["Plan"] = limpiar_valor(match_plan.group(1).strip().splitlines()[0]) or "N/A"

    match_origen = re.search(r"NOMBRE_PROGRAMA_ORIGEN:\s*(.*)", respuesta_llm)
    if match_origen: datos["Programa Origen"] = limpiar_valor(match_origen.group(1).splitlines()[0])
    
    match_abrev = re.search(r"ABREVIACION_SUGERIDA:\s*(.*)", respuesta_llm)
    if match_abrev: datos["Abreviacion Sugerida"] = limpiar_valor(match_abrev.group(1).strip().splitlines()[0])

    match_creditos = re.search(r"CREDITOS_HOMOLOGADOS:\s*(.*)", respuesta_llm)
    if match_creditos:
        creditos_str = limpiar_valor(match_creditos.group(1).strip().splitlines()[0])
        datos["Creditos"] = creditos_str if creditos_str.isdigit() else "0"

    datos["Tabla Materias"] = extraer_tabla_materias(respuesta_llm)
    
    return datos

def sanitizar_nombre_archivo(nombre):
    # (Sin cambios)
    nombre_sanitizado = re.sub(r'[<>:"/\\|?*]', '', nombre).replace("\n", " ").replace("\r", " ").strip()
    return re.sub(r'\s+', ' ', nombre_sanitizado) or "ScriptSinNombreValido"

def mostrar_resumen_log_gui(main_root_ref, lista_mensajes):
    # (Sin cambios)
    if not lista_mensajes: lista_mensajes = ["No se procesaron archivos o no hubo mensajes de resumen."]
    resumen_ventana = tk.Toplevel(main_root_ref)
    resumen_ventana.title("Resumen del Procesamiento (Log)")
    resumen_ventana.geometry("800x500")
    txt_area = scrolledtext.ScrolledText(resumen_ventana, wrap=tk.WORD, width=100, height=25)
    txt_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
    lista_mensajes_ordenada = sorted(lista_mensajes)
    for msg in lista_mensajes_ordenada: txt_area.insert(tk.END, msg + "\n" + "-"*50 + "\n")
    txt_area.config(state=tk.DISABLED)
    tk.Button(resumen_ventana, text="Cerrar Log", command=resumen_ventana.destroy).pack(pady=10)
    resumen_ventana.transient(main_root_ref); resumen_ventana.grab_set()

def agregar_fila_a_tabla(ventana_principal, tabla_treeview, datos_fila):
    # (Sin cambios)
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
    # (Sin cambios)
    tabla_ventana = tk.Toplevel(main_root_ref)
    tabla_ventana.title("Tabla Resumen de Archivos Procesados (En Vivo)")
    tabla_ventana.geometry("1050x450")
    frame = ttk.Frame(tabla_ventana, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)
    cols = ("Archivo Origen", "Nombre del estudiante", "Programa al que aspira", "Plan", "Acciones")
    tree = ttk.Treeview(frame, columns=cols, show='headings')

    def sort_by_column(treeview, col, reverse):
        data_list = [(treeview.set(k, col).lower(), k) for k in treeview.get_children('')]
        data_list.sort(reverse=reverse)
        for index, (val, k) in enumerate(data_list):
            treeview.move(k, '', index)
        treeview.heading(col, command=lambda: sort_by_column(treeview, col, not reverse))
        arrow = " 🔼" if reverse else " 🔽"
        for c in cols:
            if c != col and c != "Acciones":
                treeview.heading(c, text=c)
        treeview.heading(col, text=col + arrow)

    for col in cols:
        tree.heading(col, text=col)
        if col == "Acciones": 
            tree.column(col, width=100, minwidth=80, anchor='center')
        else:
            tree.heading(col, text=col, command=lambda _col=col: sort_by_column(tree, _col, False))
            if col == "Plan": 
                tree.column(col, width=100, minwidth=80, anchor='center')
            elif col == "Programa al que aspira": 
                tree.column(col, width=250, minwidth=200, anchor='w')
            else:
                tree.column(col, width=200, minwidth=150, anchor='w')

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
        if column_id != '#5': return
        rowid = tree.identify_row(event.y)
        if not rowid: return
        item_values = tree.item(rowid, 'values')
        item_tags = tree.item(rowid, 'tags')
        nombre_estudiante = item_values[1]
        ruta_archivo_original = item_tags[1] if len(item_tags) > 1 else None
        if not nombre_estudiante or not ruta_archivo_original or nombre_estudiante == "No extraído":
            messagebox.showwarning("Advertencia", "No se puede abrir el script porque no se extrajo un nombre de estudiante válido.", parent=tabla_ventana)
            return
        nombre_sanitizado = sanitizar_nombre_archivo(nombre_estudiante)
        nombre_archivo_ahk = f"{nombre_sanitizado}.ahk"
        directorio_original = os.path.dirname(ruta_archivo_original)
        ruta_completa_ahk = os.path.join(directorio_original, nombre_archivo_ahk)
        if os.path.exists(ruta_completa_ahk):
            try:
                if os.name == 'nt': os.startfile(ruta_completa_ahk)
                elif os.name == 'posix': subprocess.run(['xdg-open', ruta_completa_ahk], check=True)
                else: subprocess.run(['open', ruta_completa_ahk], check=True)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo de script:\n{ruta_completa_ahk}\n\nError: {e}", parent=tabla_ventana)
        else:
            messagebox.showwarning("Archivo no encontrado", f"No se encontró el archivo de script esperado:\n{ruta_completa_ahk}\n\nAsegúrese de que el proceso haya terminado y el archivo no haya sido movido.", parent=tabla_ventana)
    
    tree.bind("<Button-1>", on_click)
    return tree

def worker_process_file(archivo_queue, ventana_principal, tabla_treeview, barra_progreso, etiqueta_estado, num_archivos):
    # (Sin cambios hasta la construcción del script)
    global mensajes_resumen_procesamiento

    while not archivo_queue.empty():
        try:
            i, archivo_actual = archivo_queue.get_nowait()
        except queue.Empty:
            break

        nombre_base_archivo = os.path.basename(archivo_actual)
        ventana_principal.after(0, lambda n=nombre_base_archivo: etiqueta_estado.config(text=f"Procesando: {n}"))
        mensaje_log_actual = [f"\n--- Procesando archivo {i+1}/{num_archivos}: {nombre_base_archivo} ---"]
        texto_extraido = None
        if nombre_base_archivo.lower().endswith('.pdf'):
            texto_extraido = extraer_texto_pdf_con_apis(archivo_actual)
        elif nombre_base_archivo.lower().endswith(('.xlsx', '.xls')):
            texto_extraido = extraer_texto_excel_con_pandas(archivo_actual)

        datos_para_fila_actual = {"Archivo Origen": nombre_base_archivo}

        if texto_extraido:
            print("\n" + "#"*60)
            print(f"DEBUG: TEXTO CRUDO EXTRAÍDO POR OCR PARA '{nombre_base_archivo}'")
            print(texto_extraido)
            print("#"*60 + "\n")
            texto_limpio_para_llm = limpiar_texto_general_para_llm(texto_extraido)
            if texto_limpio_para_llm:
                respuesta_llm = None
                try:
                    respuesta_llm = analizar_con_llm(texto_limpio_para_llm, PROMPT_ANALISIS_DOCUMENTO, OPENROUTER_API_KEY, PRIMARY_LLM_MODEL, OPENROUTER_API_URL, nombre_base_archivo)
                except ServerSideApiException:
                    mensaje_log_actual.append(f"⚠️ Fallo con modelo primario, reintentando con modelo de respaldo...")
                    respuesta_llm = analizar_con_llm(texto_limpio_para_llm, PROMPT_ANALISIS_DOCUMENTO, OPENROUTER_API_KEY, FALLBACK_LLM_MODEL, OPENROUTER_API_URL, nombre_base_archivo)

                if respuesta_llm:
                    print("\n" + "="*60)
                    print(f"DEBUG: RESPUESTA COMPLETA DEL LLM PARA '{nombre_base_archivo}'")
                    print(respuesta_llm)
                    print("="*60 + "\n")
                    datos_completos = extraer_datos_completos(respuesta_llm)
                    datos_para_fila_actual.update(datos_completos)
                    
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
                        tabla_materias = datos_completos.get("Tabla Materias", [])
                        if tabla_materias and datos_completos.get("Nombre") != "No extraído":
                            nombre_origen_norm = normalizar_texto_para_busqueda(datos_completos["Programa Origen"])
                            abreviatura_final = ABREVIATURAS_NORMALIZADAS.get(nombre_origen_norm, datos_completos["Abreviacion Sugerida"])
                            abreviatura_final = quitar_tildes(abreviatura_final)
                            abreviatura_final = abreviatura_final[:30]
                            
                            linea_encabezado = f"Send, 1{{Tab}}HOM01{{Tab}}HOM01{{Tab}}{datos_completos['Creditos']}{{Tab}}{{Tab}}{{Tab}}{abreviatura_final}{{F10}}^{{PgDn}}{{Space}}{{Tab}}"
                            
                            lineas_materias_script = []
                            for materia in tabla_materias:
                                if materia.get("letras") and materia.get("numeros"):
                                    linea = f"Send, {materia['letras']}{{Tab}}{materia['numeros']}{{Tab}}{{Tab}}{materia['calificacion']}{{Tab}}N{{Down}}{{Space}}{{Tab}}"
                                    lineas_materias_script.append(linea)

                            if lineas_materias_script:
                                lineas_materias_script[-1] = lineas_materias_script[-1].rsplit('{Down}', 1)[0]

                            script_final_parts = [
                                "#SingleInstance force", "#NoEnv", "SendMode Input", "SetKeyDelay, 10",
                                'WinTitle := "Oracle Fusion Middleware Forms Services:  Open > SHATRNS"',
                                "WinWait, %WinTitle%", "WinActivate, %WinTitle%", "WinWaitActive, %WinTitle%",
                                linea_encabezado
                            ]
                            script_final_parts.extend(lineas_materias_script)
                            script_final = "\n".join(script_final_parts)
                            
                            nombre_archivo_sanitizado = sanitizar_nombre_archivo(datos_completos["Nombre"])
                            ruta_archivo_ahk_generado = os.path.join(os.path.dirname(archivo_actual), f"{nombre_archivo_sanitizado}.ahk")
                            try:
                                with open(ruta_archivo_ahk_generado, "w", encoding="utf-8") as f: f.write(script_final)
                                mensaje_log_actual.append(f"✅ Script AHK guardado en: {ruta_archivo_ahk_generado}")
                            except Exception as e:
                                mensaje_log_actual.append(f"❌ Error guardando AHK: {e}")
                        else:
                            mensaje_log_actual.append(f"❌ No se generó script para '{nombre_base_archivo}': tabla de materias vacía o nombre no extraído.")
                            if not tabla_materias:
                                 with thread_lock:
                                    mensajes_resumen_procesamiento.append(f"❌ No se encontró o no se pudo procesar la TABLA DE DATOS para '{nombre_base_archivo}'.")

                else:
                    datos_para_fila_actual.update({"Nombre": "Error en API LLM", "Programa": "N/A", "Plan": "N/A"})
            else:
                datos_para_fila_actual.update({"Nombre": "Texto vacío", "Programa": "N/A", "Plan": "N/A"})
        else:
            datos_para_fila_actual.update({"Nombre": "Error Extracción", "Programa": "N/A", "Plan": "N/A"})

        datos_para_fila_actual["Ruta Completa"] = archivo_actual
        agregar_fila_a_tabla(ventana_principal, tabla_treeview, datos_para_fila_actual)
        with thread_lock:
            mensajes_resumen_procesamiento.extend(mensaje_log_actual)
        archivo_queue.task_done()
        ventana_principal.after(0, barra_progreso.step)

# --- FUNCIÓN MODIFICADA PARA ELIMINAR LA VENTANA EMERGENTE ---
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
            # --- CAMBIO AQUÍ: El bloque if messagebox.askyesno ha sido eliminado ---
            # La aplicación ahora permanecerá abierta hasta que el usuario la cierre manualmente.

        ventana_principal.after(0, actualizar_gui_final)

    threading.Thread(target=procesar_en_paralelo, daemon=True).start()

def iniciar_procesamiento_en_hilo(boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal):
    # (Sin cambios)
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
        procesar_archivos_seleccionados(
            lista_archivos_ordenada, boton_seleccionar, etiqueta_estado, barra_progreso, etiqueta_etr, ventana_principal, tabla_treeview
        )
    except FileNotFoundError as e:
        messagebox.showerror("Error de Archivo", f"No se encontró el archivo {e.filename}.", parent=ventana_principal)
        if etiqueta_estado: etiqueta_estado.config(text="Error al procesar la selección.")
        if boton_seleccionar: boton_seleccionar.config(state=tk.NORMAL)


PROMPT_ANALISIS_DOCUMENTO = PROMPT_ANALISIS_DOCUMENTO = """Analiza el contenido del documento proporcionado, que es un acta de homologación. Tu tarea principal es reconstruir la tabla de materias basándote en las columnas de la derecha, bajo el encabezado "PROGRAMA DESTINO IBERO".

INSTRUCCIONES CRÍTICAS PARA LA TABLA:
1.  **Enfoque Exclusivo:** Concéntrate ÚNICAMENTE en las columnas "CÓDIGO", "DENOMINACIÓN" y "CALIF" que pertenecen a "PROGRAMA DESTINO IBERO". IGNORA por completo la información de "PROGRAMA DE ORIGEN".
2.  **Tolerancia a Formato:** El texto extraído del PDF (OCR) puede estar muy desalineado. Las filas pueden estar rotas en varias líneas. Tu misión es reconstruir la tabla lógicamente.
3.  **Asociación Lógica:** Si encuentras un CÓDIGO de "Programa Destino" en una línea y su CALIFICACIÓN correspondiente aparece una o dos líneas más abajo, ASÓCIALOS. Asume que pertenecen a la misma materia homologada. No omitas una materia solo porque sus datos no están en una única línea perfecta.

CONDICIONES DE PARADA:
- Si el periodo académico contiene "2024", detén el proceso y notifica: DETENER_PROCESO_CONDICION_PERIODO.
- Si el programa al que aspira es una "maestria" o "especializacion", detén el proceso y notifica: DETENER_PROCESO_CONDICION_PROGRAMA.
- Si encuentras una calificación en "Programa Destino" que es menor a 3.0, detén el proceso y notifica: DETENER_PROCESO_CONDICION_CALIFICACION.

Si NINGUNA condición de parada se cumple, extrae la siguiente información:

FORMATO DE SALIDA:
1.  **DATOS CLAVE PARA RESUMEN:**
    *   NOMBRE_ESTUDIANTE: [Nombre completo del estudiante]
    *   PROGRAMA_ASPIRA: [Nombre completo del programa al que aspira]
    *   PLAN_ESTUDIO: [Número del plan de estudio, o "N/A"]
    *   NOMBRE_PROGRAMA_ORIGEN: [Nombre completo del programa de origen]
    *   CREDITOS_HOMOLOGADOS: [Número total de créditos homologados]
    *   ABREVIACION_SUGERIDA: [Abreviatura del NOMBRE_PROGRAMA_ORIGEN según las reglas]

2.  **TABLA DE DATOS (Materias):**
    Presenta la tabla reconstruida de "PROGRAMA DESTINO IBERO". Cada materia en una nueva línea con el formato exacto, separando letras y números del código:
    LetrasCodigo | NumerosCodigo | CalificacionFormateada
    (Ejemplo CORRECTO: PSPP | 22180 | 3.4)
    (Ejemplo INCORRECTO: MUV21600 | | 4.5)

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
    if not OPENROUTER_API_KEY: keys_faltantes.append("OPENROUTER_API_KEY")

    if keys_faltantes:
        mensaje_error = "ERROR: Faltan claves en el archivo .env:\n\n" + "\n".join(keys_faltantes)
        etiqueta_estado_gui.config(text="ERROR: ¡Configure las claves API en .env!")
        boton_seleccionar_gui.config(state=tk.DISABLED)
        messagebox.showerror("Error de Configuración", mensaje_error, parent=ventana_principal_app)
    else:
        etiqueta_estado_gui.config(text="Claves API OK. Listo para seleccionar archivos.")

    ventana_principal_app.mainloop()