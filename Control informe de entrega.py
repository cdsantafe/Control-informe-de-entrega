# -*- coding: utf-8 -*-
"""
Tablero de Auditoría y Control de Asignación de Flota
Cervecería Quilmes · Operación Santa Fe

Ejecutar:    streamlit run "Control informe de entrega.py"
Requiere:    pip install streamlit pandas openpyxl requests
"""
from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Auditoría de Flota · Quilmes Santa Fe", page_icon="🚛", layout="wide")

# ════════════════════════════════════════════════════════════════════════════
# CONSTANTES DE NEGOCIO Y REGLAS
# ════════════════════════════════════════════════════════════════════════════
URL_FLOTA = (
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRlajPJLs9j3Kl6WU0-xb7xxuPfV1jLwgp3tYSMBerrkqmZHHwnTnriwHKlliGjH2gKDNwBOycnUJRo/pub?output=xlsx"
)

UNIDAD_RESTRINGIDA = "AG622AO"
ZONAS_AG622AO = ["SAN JUSTO", "SAN JAVIER", "EMILIA", "VIDELA"]

UNIDADES_CENTRO = {"AH474WT", "PAP418"}
CENTRO_MAX_OB = 10

POOL_SMK = {"AH243PW", "AH528YU", "PAC686", "NXF323"}
SMK_MAX_CLIENTES = 5

UNIDADES_GAS = {"AE514OY", "AE514OZ"}
GAS_MAX_CLI = 15
GAS_MAX_OB_BASE = 6
GAS_MAX_OB_VARIOS_CLI = 7  # Máximo de OB si tiene más de 3 clientes
GAS_EXCEPCION_CLI = 5

CABLES_BAJOS_PROHIBIDAS = {
    "NIF946", "PJX924", "NHH822", "OTW502", "PPN082", "MPS763", "AE514OY", "AE514OZ",
}
ZONAS_CABLES_BAJOS = {
    "Barranquitas": r"\bBARRANQUITAS\b",
    "Loyola": r"\bLOYOLA\b",
    "San Lorenzo": r"\bSAN LORENZO\b",
    "B° Roma": r"\bB\.?\s?ROMA\b|\bBARRIO ROMA\b",
    "Centenario": r"\bCENTENARIO\b",
    "Yapeyú": r"\bYAPEYU\b",
    "Santa Rosa": r"\bSANTA ROSA\b",
}

PUEBLOS_DEFAULT = [
    "Coronda", "Gálvez", "Laguna Paiva", "Costa", "Sauce Viejo", "San Carlos",
    "San Justo", "San Javier", "Emilia", "Videla",
]

RE_CENTRO = re.compile(r"\b(CENTRO|PUERTO)\b")
RE_SMK = re.compile(r"SMK|MAYORISTA|SUPERMERC")
RE_BALANCIN = re.compile(r"BALANC")

ORDEN_SEVERIDAD = {"Crítico": 0, "Alto": 1, "Medio": 2, "": 3}
CLAVES_PATENTE = {"PATENTE", "DOMINIO", "PAT"}

ALIAS_INFORME = {
    "planilla": ["PLANILLA"],
    "patente": ["PATENTE", "DOMINIO"],
    "zona": ["ZONA"],
    "cli
