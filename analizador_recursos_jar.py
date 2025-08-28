import streamlit as st
import zipfile
import os
from io import BytesIO
import streamlit as st
import shutil
from zipfile import ZipFile
import tempfile
import subprocess
from docx import Document
from docx.shared import Pt
from docx.shared import RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
import xml.etree.ElementTree as ET
import gspread
import time  # Importar el modulo time
import logging
import inspect
import ast
import difflib
import glob
import base64
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from lxml import etree
import re
from datetime import date
from collections import OrderedDict
from pathlib import Path
import pandas as pd
from io import BytesIO
import io
import html

#  Variable global para el servicio detectado
SERVICIO_GLOBAL = None

def obtener_mapeo_nombres(jar_path):
    """Lee ExportInfo del JAR y arma un mapeo jarentryname -> instanceId con extension"""
    mapping = {}
    try:
        with zipfile.ZipFile(jar_path, "r") as jar:
            if "ExportInfo" in jar.namelist():
                with jar.open("ExportInfo") as f:
                    xml_content = f.read()

                root = ET.fromstring(xml_content)
                ns = {"imp": "http://www.bea.com/wli/config/importexport"}

                # Diccionario para mapear typeId -> extension
                type_extensions = {
                    "WSDL": ".wsdl",
                    "ProxyService": ".proxy",
                    "Pipeline": ".pipeline",
                    "XMLSchema": ".xsd",
                    "BusinessService": ".bix",
                    "XSLT": ".xsl",
                    "XQuery": ".xquery",
                    "JCA": ".jca",
                }

                for item in root.findall("imp:exportedItemInfo", ns):
                    instance_id = item.attrib.get("instanceId")
                    type_id = item.attrib.get("typeId")

                    # Asignar extension según typeId
                    extension = type_extensions.get(type_id, "")
                    instance_id_ext = instance_id + extension

                    props = item.find("imp:properties", ns)
                    if props is not None:
                        for prop in props.findall("imp:property", ns):
                            if prop.attrib.get("name") == "jarentryname":
                                jar_name = prop.attrib["value"]
                                mapping[jar_name] = instance_id_ext
    except Exception as e:
        st.warning(f"No se pudo leer ExportInfo: {e}")
    return mapping

#  Funcion para obtener el nombre del servicio desde el JAR
def obtener_nombre_servicio(file_list: list) -> str:
    """
    Busca la carpeta con 'EXP' o 'exp' y obtiene el nombre del pipeline principal.
    Retorna el nombre del servicio sin la extension (.pipeline).
    """
    global SERVICIO_GLOBAL
    for ruta in file_list:
        partes = ruta.split("/")

        # 1️⃣ Buscar folder con EXP
        for parte in partes:
            if "EXP" in parte.upper():  # acepta EXP o exp
                # 2️⃣ Buscar el pipeline dentro de esa carpeta
                if ruta.endswith(".pipeline"):
                    servicio = os.path.basename(ruta).replace(".pipeline", "")
                    SERVICIO_GLOBAL = servicio
                    return servicio
    return None


#  Transforma los artefactos y separa bien la ruta (sin el archivo)
def transformar_datos(file_list: list) -> pd.DataFrame:
    registros = []

    for ruta in file_list:
        # Saltar entradas que son directorios
        if ruta.endswith("/"):
            continue

        partes = ruta.split("/")
        artefacto = partes[-1]
        ruta_solo_directorio = "/".join(partes[:-1]) + ("/" if len(partes) > 1 else "")

        servicio = SERVICIO_GLOBAL if SERVICIO_GLOBAL else "Desconocido"

        # Ajuste de extensiones mostrado al usuario
        if artefacto.endswith(".XMLSchema"):
            artefacto = artefacto[:-10] + ".xsd"
        elif artefacto.endswith(".XSLT"):
            artefacto = artefacto[:-5] + ".xsl"
        elif artefacto.endswith(".WSDL"):
            artefacto = artefacto[:-5] + ".wsdl"
        elif artefacto.endswith(".Pipeline"):
            artefacto = artefacto[:-9] + ".pipeline"
        elif artefacto.endswith(".ProxyService"):
            artefacto = artefacto[:-13] + ".proxy"
        elif artefacto.endswith(".BusinessService"):
            artefacto = artefacto[:-16] + ".bix"

        # Omitir ExportInfo
        if "ExportInfo" in artefacto:
            continue

        registros.append({
            "Servicio": servicio,
            "Ruta": ruta_solo_directorio,  # <-- ahora sin el archivo
            "Artefacto": artefacto
        })

    return pd.DataFrame(registros)

# ================= STREAMLIT APP =================

st.title("🗂️ Normalizador de artefactos OSB desde JAR")

archivo = st.file_uploader("▶️ Carga tu archivo JAR", type=["jar"])

if archivo:
    with tempfile.TemporaryDirectory() as tmpdir:
        jar_path = os.path.join(tmpdir, archivo.name)
        with open(jar_path, "wb") as f:
            f.write(archivo.getvalue())

        # Abrir el JAR como ZIP
        with zipfile.ZipFile(jar_path, "r") as jar:
            file_list = jar.namelist()

        #  Aplicar mapeo ExportInfo (si existe)
        mapping = obtener_mapeo_nombres(jar_path)
        file_list_legibles = [mapping.get(f, f) for f in file_list]

        # Detectar servicio (pipeline dentro de carpeta EXP/exp)
        servicio_detectado = obtener_nombre_servicio(file_list_legibles)
        if servicio_detectado:
            st.success(f"✅ Servicio detectado: **{servicio_detectado}**")
            SERVICIO_GLOBAL = servicio_detectado
        else:
            st.warning("⚠️ No se encontro servicio con carpeta EXP o pipeline asociado.")
            
            # Opcion de nombre de servicio manual
            usar_manual = st.checkbox("✏️ Ingresar nombre del servicio manualmente")
            if usar_manual:
                servicio_manual = st.text_input("Nombre del servicio")
                if servicio_manual:
                    SERVICIO_GLOBAL = servicio_manual

        #  Boton para ejecutar la transformacion
        if st.button("🚀 Ejecutar transformacion"):
            df_transformado = transformar_datos(file_list)

            #  Ordenar el dataframe por la columna 'Ruta' (ascendente)
            df_transformado = df_transformado.sort_values(by="Ruta", ascending=True).reset_index(drop=True)

            st.subheader("🚀 Datos transformados (ordenados por Ruta)")
            st.dataframe(df_transformado)

            # ======================
            #  Excel normal
            # ======================
            output_normal = io.BytesIO()
            with pd.ExcelWriter(output_normal, engine="openpyxl") as writer:
                df_transformado.to_excel(writer, index=False, sheet_name="Artefactos")
            output_normal.seek(0)

            # st.download_button(
                # label="Descargar Excel Normal",
                # data=output_normal.getvalue(),
                # file_name="artefactos_normalizados.xlsx",
                # mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            # )

            # ======================
            #  Excel en una fila con \n
            # ======================
            df_unico = pd.DataFrame({
                "Servicio": ["\n".join(df_transformado["Servicio"].astype(str))],
                "Ruta": ["\n".join(df_transformado["Ruta"].astype(str))],
                "Artefacto": ["\n".join(df_transformado["Artefacto"].astype(str))]
            })
            
            st.markdown(
                df_unico.to_html(index=False).replace("\\n", "<br>"),
                unsafe_allow_html=True
            )
            
            output_mejorado = io.BytesIO()
            with pd.ExcelWriter(output_mejorado, engine="openpyxl") as writer:
                df_unico.to_excel(writer, index=False, sheet_name="Artefactos")
            output_mejorado.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel",
                data=output_mejorado.getvalue(),
                file_name="artefactos_unica_fila.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )