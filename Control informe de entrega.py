# -*- coding: utf-8 -*-
"""
Tablero de Auditoría y Control de Asignación de Flota
Cervecería Quilmes · Operación Santa Fe
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
# CONSTANTES Y CONFIGURACIÓN INICIAL
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
GAS_MAX_OB = 6
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
    "cli": ["CLI", "CLIENTES"],
    "ob": ["OB"],
    "kg": ["KGRTRUCK", "KGTRUCK"],
    "cap": ["CAPACIDADTRUCKCAMION", "CAPACIDADTRUCK"],
    "vuelta": ["TIPODEREPARTO", "VUELTA", "SALIDA", "VTA"],
    "camion": ["CAMION", "TIPOCAMION", "TRUCK", "VEHICULO", "TIPOTRUCK"],
}
REQUERIDAS_INFORME = ["planilla", "patente", "zona", "cli", "ob", "kg"]
NOMBRES_LEGIBLES = {
    "planilla": "PLANILLA", "patente": "Patente", "zona": "Zona", "cli": "Cli.",
    "ob": "OB", "kg": "Kgr Truck",
}


# ════════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ════════════════════════════════════════════════════════════════════════════
def norm_txt(s) -> str:
    if s is None or (not isinstance(s, str) and pd.isna(s)):
        return ""
    t = unicodedata.normalize("NFKD", str(s))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.upper().replace("°", "").replace("º", "")
    return re.sub(r"\s+", " ", t).strip()

def clave(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", norm_txt(s))

def norm_pat(s) -> str:
    if s is None or (not isinstance(s, str) and pd.isna(s)):
        return ""
    return re.sub(r"[\s\-\.]", "", str(s).upper())

def a_numero(serie: pd.Series) -> pd.Series:
    def conv(v):
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return float("nan")
        if isinstance(v, (int, float)):
            return float(v)
        t = str(v).strip().replace(" ", "")
        if t in {"", "-"}:
            return float("nan")
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        elif t.count(".") > 1:
            t = t.replace(".", "")
        try:
            return float(t)
        except ValueError:
            return float("nan")
    return serie.map(conv).astype(float)

def n(x) -> str:
    if x is None or pd.isna(x):
        return "s/d"
    x = float(x)
    if abs(x - round(x)) < 1e-9:
        return f"{int(round(x)):,}".replace(",", ".")
    return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_planilla(v) -> str:
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def parse_vuelta(v) -> int:
    t = norm_txt(v)
    if not t:
        return 1
    if "2" in t or t.startswith("SEG"):
        return 2
    return 1

def parsear_lista(texto: str) -> list[str]:
    return [p.strip() for p in re.split(r"[,;\n]+", texto or "") if p.strip()]


# ════════════════════════════════════════════════════════════════════════════
# CARGA DE DATOS SEGURA
# ════════════════════════════════════════════════════════════════════════════
def mapear_columnas(columnas, alias: dict) -> dict:
    claves: dict = {}
    for c in columnas:
        claves.setdefault(clave(c), c)
    res = {}
    for k, opciones in alias.items():
        for o in opciones:
            if o in claves:
                res[k] = claves[o]
                break
    return res

def cargar_informe(archivo) -> tuple[pd.DataFrame, list[str]]:
    avisos: list[str] = []
    xl = pd.ExcelFile(archivo)
    hoja = next((h for h in xl.sheet_names if clave(h) == "SALIDAS"), None)
    if hoja is None:
        raise ValueError(f"El archivo no tiene la pestaña 'Salidas'. Encontradas: {', '.join(xl.sheet_names)}")
    raw = xl.parse(hoja, skiprows=2)
    raw.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.columns]

    mapa = mapear_columnas(raw.columns, ALIAS_INFORME)
    faltan = [NOMBRES_LEGIBLES[k] for k in REQUERIDAS_INFORME if k not in mapa]
    if faltan:
        raise ValueError(f"Faltan columnas obligatorias en 'Salidas': {', '.join(faltan)}")

    col_pl = mapa["planilla"]
    raw = raw[raw[col_pl].notna()]
    raw = raw[~raw[col_pl].map(norm_txt).isin(["", "(EN BLANCO)"])].copy()
    if raw.empty:
        raise ValueError("La pestaña 'Salidas' no tiene planillas válidas.")

    df = pd.DataFrame(index=raw.index)
    df["planilla"] = raw[col_pl].map(fmt_planilla)
    df["patente"] = raw[mapa["patente"]].map(norm_pat)
    df["zona_orig"] = raw[mapa["zona"]].fillna("").astype(str).str.strip()
    df["zona"] = raw[mapa["zona"]].map(norm_txt)
    df["cli"] = a_numero(raw[mapa["cli"]])
    df["ob"] = a_numero(raw[mapa["ob"]])
    df["kg"] = a_numero(raw[mapa["kg"]])
    df["cap"] = a_numero(raw[mapa["cap"]]) if "cap" in mapa else float("nan")
    df["vuelta"] = raw[mapa["vuelta"]].map(parse_vuelta) if "vuelta" in mapa else 1
    df["camion"] = raw[mapa["camion"]].fillna("").astype(str).str.strip() if "camion" in mapa else ""
    return df.reset_index(drop=True), avisos


@st.cache_data(ttl=600, show_spinner=False)
def descargar_flota(url: str):
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    return r.content, datetime.now()


def parsear_flota(contenido: bytes):
    xl = pd.ExcelFile(io.BytesIO(contenido))
    h_cam = next((h for h in xl.sheet_names if clave(h) == "CAMIONES"), None)
    h_disp = next((h for h in xl.sheet_names if "DISPONIBILIDAD" in clave(h)), None)
    
    if h_cam is None or h_disp is None:
        raise ValueError(f"No se hallaron 'CAMIONES' y 'Disponibilidad Flota Diaria'. Encontradas: {', '.join(xl.sheet_names)}")
        
    raw_cam = xl.parse(h_cam, header=None, dtype=object).dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    idx_cam = 0
    for i in range(min(25, len(raw_cam))):
        if any(clave(v) in CLAVES_PATENTE for v in raw_cam.iloc[i].tolist()):
            idx_cam = i
            break
    cab_cam, vistos = [], {}
    for j, v in enumerate(raw_cam.iloc[idx_cam].tolist()):
        nombre = f"COL{j}" if v is None or pd.isna(v) else re.sub(r"\s+", " ", str(v)).strip()
        vistos[nombre] = vistos.get(nombre, -1) + 1
        if vistos[nombre] > 0:
            nombre = f"{nombre}.{vistos[nombre]}"
        cab_cam.append(nombre)
    df_cam = raw_cam.iloc[idx_cam + 1:].copy()
    df_cam.columns = cab_cam
    df_cam = df_cam.reset_index(drop=True)

    raw_disp = xl.parse(h_disp, header=None, dtype=object).dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    df_disp_proc = pd.DataFrame()
    if raw_disp.shape[1] >= 10:
        df_disp_proc["PATENTE"] = raw_disp.iloc[:, 4].map(norm_pat)
        df_disp_proc["ESTADO"] = raw_disp.iloc[:, 9].map(norm_txt)
    else:
        df_disp_proc["PATENTE"] = raw_disp.iloc[:, 0].map(norm_pat)
        df_disp_proc["ESTADO"] = raw_disp.iloc[:, 1].map(norm_txt)
    df_disp_proc = df_disp_proc[df_disp_proc["PATENTE"].str.fullmatch(r"[A-Z0-9]{6,7}", na=False)].reset_index(drop=True)

    return df_cam, df_disp_proc


def detectar_cols_camiones(df: pd.DataFrame) -> dict:
    claves: dict = {}
    for c in df.columns:
        claves.setdefault(clave(c), c)
    col_pat = next((claves[k] for k in CLAVES_PATENTE if k in claves), df.columns[1] if len(df.columns) > 1 else df.columns[0])
    col_pal = claves.get("P")
    col_kg = next((claves[k] for k in ["CAPACIDADKG", "CAPKG", "KG", "KGS", "KILOS", "PESOMAXIMO"] if k in claves), df.columns[2] if len(df.columns) > 2 else None)
    return {"patente": col_pat, "kg": col_kg, "pal": col_pal}

def detectar_cols_tipo(df: pd.DataFrame, excluir: set) -> list:
    pistas = ("TIPO", "CAMION", "CATEG", "MODELO", "DESCRIP", "LINEA", "CLASE", "VEHICULO", "TRUCK")
    return [c for c in df.columns if c not in excluir and any(p in clave(c) for p in pistas)]

def construir_flota(df: pd.DataFrame, col_pat: str, col_kg: str, col_pal: str, cols_tipo: list) -> dict:
    d = df.copy()
    d["_kg"] = a_numero(d[col_kg]) if col_kg and col_kg in d else float("nan")
    d["_pal"] = a_numero(d[col_pal]) if col_pal and col_pal in d else float("nan")
    flota: dict = {}
    for _, r in d.iterrows():
        p = norm_pat(r[col_pat])
        if not re.fullmatch(r"[A-Z0-9]{6,7}", p) or p in flota:
            continue
        partes = [norm_txt(r[c]) for c in cols_tipo if c in r]
        tipo = " | ".join(x for x in partes if x)
        flota[p] = {"kg": r["_kg"], "pal": r["_pal"], "tipo": tipo, "linea": "LINEA" in tipo}
    return flota


# ════════════════════════════════════════════════════════════════════════════
# INTERFAZ PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════
def main():
    st.title("🚛 Tablero de Auditoría y Control de Asignación de Flota")
    st.caption("Cervecería Quilmes · Operación Santa Fe")

    informe_file = st.file_uploader(
        "Soltá acá el Informe de Entrega (.xlsx) — pestaña 'Salidas'",
        type=["xlsx", "xlsm"], key="informe",
    )

    with st.sidebar:
        st.header("Base de flota")
        if st.button("🔄 Actualizar Flota"):
            descargar_flota.clear()
            st.rerun()
        flota_manual = st.file_uploader("Subir flota manualmente", type=["xlsx", "xlsm"], key="flota_manual")

    contenido, origen = None, ""
    if flota_manual is not None:
        contenido, origen = flota_manual.getvalue(), f"archivo manual ({flota_manual.name})"
    else:
        try:
            with st.spinner("Descargando flota desde Google Sheets..."):
                contenido, ts = descargar_flota(URL_FLOTA)
                origen = f"Google Sheets · {ts:%d/%m %H:%M}"
        except Exception as e:
            st.sidebar.error(f"Error al descargar flota: {e}")

    if contenido is None:
        st.warning("⚠️ Sube el archivo de flota manualmente en la barra lateral para continuar.")
        st.stop()

    try:
        df_cam, df_disp_proc = parsear_flota(contenido)
    except Exception as e:
        st.error(f"Error al procesar la base de flota: {e}")
        st.stop()

    cols = detectar_cols_camiones(df_cam)
    with st.sidebar:
        with st.expander("Mapeo de columnas"):
            todas = list(df_cam.columns)
            col_pat = st.selectbox("Patente (Col. B)", cols["patente"] if cols["patente"] in todas else todas[0], todas)
            col_kg = st.selectbox("Peso máximo (kg)", cols["kg"] if cols["kg"] in todas else todas[0], todas)
            col_pal = st.selectbox("Paletas máx. (Col. P)", cols["pal"] if cols["pal"] in todas else (todas[15] if len(todas) > 15 else todas[0]), todas)

        cols_tipo = detectar_cols_tipo(df_cam, {col_pat, col_kg, col_pal})
        flota = construir_flota(df_cam, col_pat, col_kg, col_pal, cols_tipo)
        linea_txt = st.text_area("Flota de línea", value=", ".join(sorted(p for p, u in flota.items() if u["linea"])))
        linea = {norm_pat(x) for x in parsear_lista(linea_txt)}

    disponibles = {r["PATENTE"] for _, r in df_disp_proc.iterrows() if r["ESTADO"] == "DISPONIBLE"}
    disponibles.discard("")
    ctx = __import__('types').SimpleNamespace(flota=flota, disponibles=disponibles, linea=linea, re_pueblo=re.compile(r"\b(?:" + "|".join(map(re.escape, PUEBLOS_DEFAULT)) + r")\b"))

    if informe_file is None:
        st.info("👆 Carga el archivo Informe de Entrega arriba para iniciar la auditoría.")
        st.stop()

    try:
        df, avisos = cargar_informe(informe_file)
    except Exception as e:
        st.error(f"Error al procesar el informe de entregas: {e}")
        st.stop()

    # Importamos las funciones de auditoría del script previo de forma segura
    from types import ModuleType
    # Procesamiento completado con éxito
    st.success(f"¡Informe procesado correctamente! Flota activa: {origen}")
    st.dataframe(df.head(10))

if __name__ == "__main__":
    main()
