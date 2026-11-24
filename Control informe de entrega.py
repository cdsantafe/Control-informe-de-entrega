# -*- coding: utf-8 -*-
"""
Tablero de Auditoría y Control de Asignación de Flota
Cervecería Quilmes · Operación Santa Fe

Ejecutar:   streamlit run app.py
Requiere:   pip install streamlit pandas openpyxl requests
"""
from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import requests
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import streamlit as st

# ════════════════════════════════════════════════════════════════════════════
# CONSTANTES DE NEGOCIO (editar acá si cambian las reglas)
# ════════════════════════════════════════════════════════════════════════════
URL_FLOTA = (
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRlajPJLs9j3Kl6WU0-xb7xxuPfV1jLwgp3tYSMBerrkqmZHHwnTnriwHKlliGjH2gKDNwBOycnUJRo/pub?output=xlsx"
)

UNIDAD_RESTRINGIDA = "AG622AO"
# Zonas "Interior": las cubre la unidad destinada AG622AO
ZONAS_INTERIOR = ["SAN JUSTO", "SAN JAVIER", "EMILIA", "VIDELA"]

# Flota de línea (11 unidades): tienen prioridad en zonas de línea / pueblos.
# Solo si ninguna está disponible y libre se puede usar una unidad de Santa Fe / Santo Tomé.
LINEA_FIJA = [
    "AD413RY", "AA932SE", "LWT087", "NJK122", "KKF571", "NNF606",
    "NMH398", "PME394", "NHH822", "AB071NV", "AA639XH",
]

# Paraná: zona exclusiva de estas unidades (SMK / mayoristas). No están limitadas a Paraná.
PARANA_UNIDADES = {"AB071NV", "AA639XH", "AD413RY", "AF520BS", "AH243PW", "AH243YU"}
RE_PARANA = re.compile(r"\bPARANA\b")

UNIDADES_CENTRO = {"AH474WT", "PAP418"}
CENTRO_MAX_OB = 10

POOL_SMK = {"AH243PW", "AH528YU", "PAC686", "NXF323"}
SMK_MAX_CLIENTES = 5  # aplica a planillas SMK con menos de 5 clientes

UNIDADES_GAS = {"AE514OY", "AE514OZ"}
GAS_MAX_CLI = 15
GAS_MAX_OB = 6
GAS_EXCEPCION_CLI = 5  # no aplica límite si Cli. < 5

CABLES_BAJOS_PROHIBIDAS = {
    "NIF946", "PJX924", "NHH822", "OTW502", "PPN082", "MPS763", "AE514OY", "AE514OZ",
    "AF520BS",  # unidad "alta"
}
ZONAS_CABLES_BAJOS = {
    "Barranquitas": r"\bBARRANQUITAS\b",
    "Loyola": r"\bLOYOLA\b",
    "San Lorenzo": r"\bSAN LORENZO\b",
    "B° Roma": r"\bB\.?\s?ROMA\b|\bBARRIO ROMA\b",
    "Centenario": r"\bCENTENARIO\b",
    "Yapeyú": r"\bYAPEYU\b",
    "Santa Rosa": r"\bSANTA ROSA\b(?! DE CALCHINES)",  # barrio; "Santa Rosa de Calchines" es zona de línea
    "Periferia": r"\bPERIFERIA\b",
}

# Zonas de línea / pueblos (Interior se trata aparte y es exclusiva de AG622AO)
# Se pueden combinar entre sí (ej. CAYASTA/HELVECIA/STA ROSA/COSTA) y siguen siendo zona de línea.
PUEBLOS_DEFAULT = [
    "Coronda", "Gálvez", "Laguna Paiva", "Costa", "Sauce Viejo", "San Carlos",
    "Cayasta", "Helvecia", "Santa Rosa de Calchines", "Sta Rosa", "Calchines",
]

RE_CENTRO = re.compile(r"\b(CENTRO|PUERTO)\b")
RE_SMK = re.compile(r"SMK|MAYORISTA|SUPERMERC")
RE_BALANCIN = re.compile(r"BALANC")

ORDEN_SEVERIDAD = {"Crítico": 0, "Alto": 1, "Medio": 2, "": 3}

CLAVES_PATENTE = ("PATENTE", "DOMINIO", "MATRICULA", "PAT")

ALIAS_INFORME = {
    "planilla": ["PLANILLA"],
    "patente": ["PATENTE", "DOMINIO"],
    "zona": ["ZONA"],
    "cli": ["CLI", "CLIENTES"],
    "ob": ["OB"],
    "kg": ["KGRTRUCK", "KGTRUCK"],
    "cap": ["CAPACIDADTRUCKCAMION", "CAPACIDADTRUCK"],
    "vuelta": ["VUELTA", "TIPODEREPARTO", "SALIDA", "VTA"],
    "camion": ["CAMION", "TIPOCAMION", "TRUCK", "VEHICULO", "TIPOTRUCK"],
}
REQUERIDAS_INFORME = ["planilla", "patente", "zona", "cli", "ob", "kg"]
NOMBRES_LEGIBLES = {
    "planilla": "PLANILLA", "patente": "Patente", "zona": "Zona", "cli": "Cli.",
    "ob": "OB", "kg": "Kgr Truck",
}


# ════════════════════════════════════════════════════════════════════════════
# UTILIDADES DE TEXTO / NÚMEROS
# ════════════════════════════════════════════════════════════════════════════
def norm_txt(s) -> str:
    """Mayúsculas, sin acentos, sin ° / º y con espacios colapsados."""
    if s is None or (not isinstance(s, str) and pd.isna(s)):
        return ""
    t = unicodedata.normalize("NFKD", str(s))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = t.upper().replace("°", "").replace("º", "")
    return re.sub(r"\s+", " ", t).strip()


def norm_zona(s) -> str:
    """Normaliza nombres de zona: sin distinguir mayúsculas/acentos y 'Sta.' -> 'Sta'."""
    return re.sub(r"\bSTA\.\s*", "STA ", norm_txt(s))


def clave(s) -> str:
    """Clave alfanumérica para comparar encabezados ('Cli.' -> 'CLI')."""
    return re.sub(r"[^A-Z0-9]", "", norm_txt(s))


def norm_pat(s) -> str:
    if s is None or (not isinstance(s, str) and pd.isna(s)):
        return ""
    return re.sub(r"[\s\-\.]", "", str(s).upper())


def a_numero(serie: pd.Series) -> pd.Series:
    """Convierte a float tolerando textos ('1.234,5', '', '-') y celdas vacías."""
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
    """Formato legible de números (miles con punto, sin decimales inútiles)."""
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
    if "2" in t or "SEG" in t:  # "2", "2da", "2° reparto", "Segunda", "Segundo reparto"…
        return 2
    return 1


def parsear_lista(texto: str) -> list[str]:
    return [p.strip() for p in re.split(r"[,;\n]+", texto or "") if p.strip()]


# ════════════════════════════════════════════════════════════════════════════
# CARGA: INFORME DE ENTREGA (pestaña Salidas)
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
        raise ValueError(
            f"El archivo no tiene la pestaña 'Salidas'. Pestañas encontradas: {', '.join(xl.sheet_names)}"
        )
    raw = xl.parse(hoja, skiprows=2)
    raw.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.columns]

    mapa = mapear_columnas(raw.columns, ALIAS_INFORME)
    faltan = [NOMBRES_LEGIBLES[k] for k in REQUERIDAS_INFORME if k not in mapa]
    if faltan:
        raise ValueError(
            f"Faltan columnas obligatorias en 'Salidas': {', '.join(faltan)}. "
            f"Columnas detectadas: {', '.join(map(str, raw.columns))}"
        )

    # Descarta filas sin planilla o "(en blanco)"
    col_pl = mapa["planilla"]
    raw = raw[raw[col_pl].notna()]
    raw = raw[~raw[col_pl].map(norm_txt).isin(["", "(EN BLANCO)"])].copy()
    if raw.empty:
        raise ValueError("La pestaña 'Salidas' no tiene planillas válidas luego de la limpieza.")

    df = pd.DataFrame(index=raw.index)
    df["planilla"] = raw[col_pl].map(fmt_planilla)
    df["patente"] = raw[mapa["patente"]].map(norm_pat)
    df["zona_orig"] = raw[mapa["zona"]].fillna("").astype(str).str.strip()
    df["zona"] = raw[mapa["zona"]].map(norm_zona)
    df["cli"] = a_numero(raw[mapa["cli"]])
    df["ob"] = a_numero(raw[mapa["ob"]])
    df["kg"] = a_numero(raw[mapa["kg"]])

    if "cap" in mapa:
        df["cap"] = a_numero(raw[mapa["cap"]])
    else:
        df["cap"] = float("nan")
        avisos.append("No se encontró la columna 'Capacidad Truck Camion'; se usa solo la pestaña CAMIONES.")

    if "vuelta" in mapa:
        df["vuelta"] = raw[mapa["vuelta"]].map(parse_vuelta)
    else:
        df["vuelta"] = 1
        avisos.append(
            "No se encontró la columna de vuelta ('Vuelta' o 'TIPO DE REPARTO') en Salidas: se asume Primera Vuelta "
            "para todas las planillas, por lo que el control de patentes duplicadas y de unidades libres por vuelta es aproximado."
        )

    if "camion" in mapa:
        df["camion"] = raw[mapa["camion"]].fillna("").astype(str).str.strip()
    else:
        df["camion"] = ""

    df = df.reset_index(drop=True)
    if "vuelta" in mapa:
        # valores originales -> vuelta asignada, para que se pueda verificar la lectura
        vistos = raw[mapa["vuelta"]].fillna("").astype(str).str.strip().drop_duplicates()
        df.attrs["mapa_vuelta"] = {v or "(vacío)": parse_vuelta(v) for v in vistos}
        df.attrs["col_vuelta"] = mapa["vuelta"]
    return df, avisos


# ════════════════════════════════════════════════════════════════════════════
# CARGA: BASE DE FLOTA
# ════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=900, show_spinner="Descargando flota desde Google Sheets…")
def descargar_flota(url: str):
    r = requests.get(url, timeout=45)
    r.raise_for_status()
    return r.content, datetime.now()


def _leer_hoja_con_encabezado(xl: pd.ExcelFile, hoja: str) -> pd.DataFrame:
    """Lee la hoja detectando automáticamente la fila de encabezados (la que contiene 'Patente')."""
    raw = xl.parse(hoja, header=None, dtype=object)
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if raw.empty:
        raise ValueError(f"La pestaña '{hoja}' está vacía.")
    idx = 0
    for i in range(min(25, len(raw))):
        if any(clave(v) in CLAVES_PATENTE for v in raw.iloc[i].tolist()):
            idx = i
            break
    cabecera, vistos = [], {}
    for j, v in enumerate(raw.iloc[idx].tolist()):
        if isinstance(v, (pd.Timestamp, datetime)):
            nombre = v.strftime("%d/%m/%Y")
        elif v is None or pd.isna(v):
            nombre = f"COL{j}"
        else:
            nombre = re.sub(r"\s+", " ", str(v)).strip()
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre}.{vistos[nombre]}"
        else:
            vistos[nombre] = 0
        cabecera.append(nombre)
    df = raw.iloc[idx + 1:].copy()
    df.columns = cabecera
    return df.reset_index(drop=True)


def _leer_disponibilidad(xl: pd.ExcelFile, hoja: str) -> pd.DataFrame:
    """Lee 'Disponibilidad flota Diaria' conservando la posición de las columnas.
    Las columnas se rotulan con su letra de Excel (E · Matrícula, J · Estado, etc.)."""
    raw = xl.parse(hoja, header=None, dtype=object)
    if raw.empty:
        raise ValueError(f"La pestaña '{hoja}' está vacía.")
    idx = None
    for i in range(min(25, len(raw))):
        if any(clave(v) in CLAVES_PATENTE for v in raw.iloc[i].tolist()):
            idx = i
            break
    etiquetas = []
    for j in range(raw.shape[1]):
        h = raw.iloc[idx, j] if idx is not None else None
        txt = "" if h is None or pd.isna(h) else re.sub(r"\s+", " ", str(h)).strip()
        letra = get_column_letter(j + 1)
        etiquetas.append(f"{letra} · {txt}" if txt else letra)
    df = raw.iloc[(idx + 1 if idx is not None else 0):].copy()
    df.columns = etiquetas
    return df.reset_index(drop=True)


@st.cache_data(show_spinner="Procesando base de flota…")
def parsear_flota(contenido: bytes):
    xl = pd.ExcelFile(io.BytesIO(contenido))
    h_cam = next((h for h in xl.sheet_names if clave(h) == "CAMIONES"), None)
    h_disp = next((h for h in xl.sheet_names if clave(h) == "DISPONIBILIDADFLOTADIARIA"), None)
    if h_disp is None:
        h_disp = next((h for h in xl.sheet_names if "DISPONIBILIDAD" in clave(h)), None)
    faltan = [nom for nom, h in (("CAMIONES", h_cam), ("Disponibilidad flota Diaria", h_disp)) if h is None]
    if faltan:
        raise ValueError(
            f"El archivo de flota no tiene la/s pestaña/s: {', '.join(faltan)}. "
            f"Pestañas encontradas: {', '.join(xl.sheet_names)}"
        )
    return _leer_hoja_con_encabezado(xl, h_cam), _leer_disponibilidad(xl, h_disp)


def detectar_cols_camiones(df: pd.DataFrame) -> dict:
    claves: dict = {}
    for c in df.columns:
        claves.setdefault(clave(c), c)
    col_pat = next((claves[k] for k in CLAVES_PATENTE if k in claves), None)
    col_pal = claves.get("P")
    pref_kg = ["CAPACIDADKG", "CAPKG", "KG", "KGS", "KILOS", "PESOMAXIMO", "PESOMAX",
               "CAPACIDADMAXIMA", "CAPACIDADMAX", "CARGAMAXIMA", "CAPACIDAD", "PESO"]
    col_kg = next((claves[k] for k in pref_kg if k in claves), None)
    if col_kg is None:
        col_kg = next((c for k, c in claves.items() if ("KG" in k or "PESO" in k) and k != "P"), None)
    return {"patente": col_pat, "kg": col_kg, "pal": col_pal}


def detectar_cols_tipo(df: pd.DataFrame, excluir: set) -> list:
    pistas = ("TIPO", "CAMION", "CATEG", "MODELO", "DESCRIP", "LINEA", "CLASE", "VEHICULO", "TRUCK")
    return [c for c in df.columns if c not in excluir and any(p in clave(c) for p in pistas)]


def construir_flota(df: pd.DataFrame, col_pat: str, col_kg: str, col_pal: str, cols_tipo: list) -> dict:
    d = df.copy()
    d["_kg"] = a_numero(d[col_kg])
    d["_pal"] = a_numero(d[col_pal])
    flota: dict = {}
    for _, r in d.iterrows():
        p = norm_pat(r[col_pat])
        if not re.fullmatch(r"[A-Z0-9]{6,7}", p) or p in flota:
            continue
        partes = [norm_txt(r[c]) for c in cols_tipo]
        tipo = " | ".join(x for x in partes if x)
        flota[p] = {"kg": r["_kg"], "pal": r["_pal"], "tipo": tipo, "linea": "LINEA" in tipo}
    return flota


ESTADOS_UTILIZABLES = ("DISPONIBLE", "RESERVA")  # "Reserva" se usa igual que "Disponible"


def es_utilizable(estado) -> bool:
    # "NO DISPONIBLE" no empieza con DISPONIBLE, por lo que queda afuera
    return norm_txt(estado).startswith(ESTADOS_UTILIZABLES)


def construir_disponibles(df: pd.DataFrame, col_pat: str, col_estado: str) -> set:
    disp = set()
    for p, e in zip(df[col_pat], df[col_estado]):
        pat = norm_pat(p)
        if re.fullmatch(r"[A-Z0-9]{6,7}", pat) and es_utilizable(e):
            disp.add(pat)
    return disp


def construir_estados(df: pd.DataFrame, col_pat: str, col_estado: str) -> dict:
    """patente -> estado tal cual figura en el parte diario (si repite, prevalece el utilizable)."""
    estados: dict = {}
    for p, e in zip(df[col_pat], df[col_estado]):
        pat = norm_pat(p)
        if not re.fullmatch(r"[A-Z0-9]{6,7}", pat):
            continue
        txt = "" if e is None or (not isinstance(e, str) and pd.isna(e)) else str(e).strip()
        txt = txt or "(sin estado)"
        if pat not in estados or es_utilizable(txt):
            estados[pat] = txt
    return estados


# ════════════════════════════════════════════════════════════════════════════
# MOTOR DE REGLAS
# ════════════════════════════════════════════════════════════════════════════
class Ctx:
    def __init__(self, flota: dict, disponibles: set, linea: set, pueblos: list[str]):
        self.flota = flota
        self.disponibles = disponibles
        self.linea = linea
        self.estados: dict = {}  # patente -> estado en el parte diario (lo completa main)
        self.ocup: dict = {1: {}, 2: {}}  # patente -> índices de planillas, por vuelta (lo completa auditar)
        self.planillas: list = []          # planillas del informe (lo completa auditar)
        pueblos_n = [norm_zona(p) for p in pueblos if norm_zona(p)]
        self.re_pueblo = (
            re.compile(r"\b(?:" + "|".join(re.escape(p) for p in pueblos_n) + r")\b") if pueblos_n else None
        )


def es_zona_linea(zona: str, ctx: Ctx) -> bool:
    if any(z in zona for z in ZONAS_INTERIOR):
        return False
    return bool(ctx.re_pueblo and ctx.re_pueblo.search(zona))


def lineas_utilizables(p: dict, ctx: Ctx) -> tuple[list[str], list[tuple[str, str, str]]]:
    """Unidades de línea DISPONIBLES (o en RESERVA) que podrían cubrir la planilla `p`.
    - libres: no están asignadas en la vuelta de la planilla.
    - reasignables: están asignadas en esa vuelta, pero solo a planillas que NO son de zona de línea (enroque posible).
    Una unidad de línea asignada a otra zona de línea se considera bien usada y no cuenta."""
    ocup = ctx.ocup.get(p["vuelta"], {})
    libres, reasignables = [], []
    for c in sorted(ctx.linea):
        if c not in ctx.disponibles or c not in ctx.flota:
            continue
        if evaluar(p, c, ctx, verificar_disp=False):
            continue
        ocupantes = [ctx.planillas[j] for j in ocup.get(c, [])]
        if not ocupantes:
            libres.append(c)
        elif all(not es_zona_linea(q["zona"], ctx) for q in ocupantes):
            reasignables.append((c, ", ".join(q["planilla"] for q in ocupantes), ocupantes[0]["zona_orig"] or "s/d"))
    return libres, reasignables


def evaluar(p: dict, pat: str, ctx: Ctx, verificar_disp: bool = True) -> list[tuple[str, str]]:
    """Devuelve la lista de (severidad, detalle) de reglas rotas si la planilla `p` la cubre la patente `pat`."""
    out: list[tuple[str, str]] = []
    if not pat:
        return [("Crítico", "La planilla no tiene patente asignada.")]

    u = ctx.flota.get(pat)
    es_actual = pat == p["patente"]
    zona, zona_o = p["zona"], p["zona_orig"] or "s/d"
    cli, ob, kg = p["cli"], p["ob"], p["kg"]

    # 1) Disponibilidad ------------------------------------------------------
    if verificar_disp and pat not in ctx.disponibles:
        est = ctx.estados.get(pat)
        detalle = f"figura como '{est}'" if est else "no figura"
        out.append(("Alto", f"Disponibilidad: la patente {pat} {detalle} en el parte diario (no está DISPONIBLE ni en RESERVA)."))

    # 2) Capacidad física ----------------------------------------------------
    if u is None:
        out.append(("Medio", f"Capacidad: la patente {pat} no figura en la pestaña CAMIONES, no se puede validar peso ni paletas."))
    max_kg = u["kg"] if u is not None and pd.notna(u["kg"]) else (p["cap"] if es_actual else float("nan"))
    max_pal = u["pal"] if u is not None else float("nan")
    if pd.notna(kg) and pd.notna(max_kg) and kg > max_kg:
        out.append(("Alto", f"Sobrepeso: Kgr Truck {n(kg)} kg supera el máximo de la unidad ({n(max_kg)} kg, exceso {n(kg - max_kg)} kg)."))
    if pd.notna(ob) and pd.notna(max_pal) and ob > max_pal:
        out.append(("Alto", f"Exceso de paletas: OB {n(ob)} supera la capacidad de la unidad ({n(max_pal)} paletas)."))

    # 3) Restricciones geográficas y de ruta ---------------------------------
    es_interior = any(z in zona for z in ZONAS_INTERIOR)
    es_parana = bool(RE_PARANA.search(zona))
    es_pueblo = bool(ctx.re_pueblo and ctx.re_pueblo.search(zona)) and not es_interior
    tipo = (u["tipo"] if u is not None else "")
    if es_actual:
        tipo = f"{tipo} | {norm_txt(p['camion'])}"

    if pat == UNIDAD_RESTRINGIDA:
        if not es_interior:
            out.append(("Crítico", f"Incumplimiento crítico: AG622AO está destinada exclusivamente a zonas Interior (San Justo, San Javier, Emilia/Videla) y está asignada a '{zona_o}'."))
    else:
        if es_interior:
            out.append(("Alto", f"Interior: la zona '{zona_o}' es exclusiva de AG622AO y {pat} no puede cubrirla."))
        if es_parana and pat not in PARANA_UNIDADES:
            out.append(("Alto", f"Paraná: zona exclusiva de {', '.join(sorted(PARANA_UNIDADES))}; {pat} no está habilitada."))
        if es_pueblo and ctx.linea and pat not in ctx.linea:
            libres, reasig = lineas_utilizables(p, ctx)
            if libres or reasig:
                partes = []
                if libres:
                    partes.append(f"libres en la {p['vuelta']}ª vuelta: {', '.join(libres[:4])}")
                if reasig:
                    partes.append("asignadas a zonas que no son de línea: " +
                                  "; ".join(f"{c} (planilla {pl}, {z})" for c, pl, z in reasig[:3]))
                out.append(("Alto", f"Zona de línea ('{zona_o}') cubierta por {pat}, que no es unidad de línea. "
                                    f"Hay unidades de línea disponibles → {' | '.join(partes)}. "
                                    "Una unidad de Santa Fe/Santo Tomé solo se admite si no queda ninguna de línea disponible."))

    if not es_pueblo and RE_CENTRO.search(zona) and pat not in UNIDADES_CENTRO:
        motivos = []
        if RE_BALANCIN.search(tipo):
            motivos.append("es camión balancín")
        if RE_SMK.search(tipo):
            motivos.append("es camión de supermercado/mayorista")
        if pd.notna(ob) and ob > CENTRO_MAX_OB:
            motivos.append(f"OB {n(ob)} > {CENTRO_MAX_OB}")
        if pd.notna(max_pal) and max_pal > CENTRO_MAX_OB:
            motivos.append(f"capacidad de {n(max_pal)} paletas > {CENTRO_MAX_OB}")
        if motivos:
            out.append(("Alto", f"Centro/Puerto: es zona preferencial de AH474WT y PAP418; otra unidad solo es válida si no es balancín ni SMK/mayorista y tiene ≤ {CENTRO_MAX_OB} paletas. {pat} {'; '.join(motivos)}."))

    if RE_SMK.search(zona) and pd.notna(cli) and cli < SMK_MAX_CLIENTES and pat not in POOL_SMK:
        out.append(("Alto", f"SMK Mayorista con {n(cli)} clientes (< {SMK_MAX_CLIENTES}): debe cubrirlo el pool {', '.join(sorted(POOL_SMK))}."))

    if pat in UNIDADES_GAS and not (pd.notna(cli) and cli < GAS_EXCEPCION_CLI):
        if pd.notna(cli) and cli > GAS_MAX_CLI:
            out.append(("Alto", f"Camión de gas: {n(cli)} clientes supera el máximo de {GAS_MAX_CLI}."))
        if pd.notna(ob) and ob > GAS_MAX_OB:
            out.append(("Alto", f"Camión de gas: OB {n(ob)} supera el máximo de {GAS_MAX_OB} paletas."))

    if pat in CABLES_BAJOS_PROHIBIDAS:
        hits = [nom for nom, rx in ZONAS_CABLES_BAJOS.items() if re.search(rx, zona)]
        if es_pueblo:  # combinada con otras zonas de línea (ej. SANTA ROSA/COSTA): es Calchines, no el barrio
            hits = [h for h in hits if h != "Santa Rosa"]
        if hits:
            out.append(("Alto", f"Cables bajos: la zona '{zona_o}' ({', '.join(hits)}) tiene prohibida la unidad {pat}."))

    return out


# ════════════════════════════════════════════════════════════════════════════
# AUDITORÍA COMPLETA
# ════════════════════════════════════════════════════════════════════════════
def auditar(df: pd.DataFrame, ctx: Ctx) -> pd.DataFrame:
    planillas = df.to_dict("records")
    ocup: dict = {1: {}, 2: {}}
    for i, p in enumerate(planillas):
        if p["patente"]:
            ocup[p["vuelta"]].setdefault(p["patente"], []).append(i)

    ctx.ocup = ocup
    ctx.planillas = planillas
    filas = []
    for i, p in enumerate(planillas):
        try:
            probs = evaluar(p, p["patente"], ctx)
            dup = ocup[p["vuelta"]].get(p["patente"], [])
            dup_txt = ""
            if p["patente"] and len(dup) > 1:
                otras = ", ".join(planillas[j]["planilla"] for j in dup if j != i)
                dup_txt = f"{p['patente']} también está en la/s planilla/s {otras} de la {p['vuelta']}ª vuelta."
        except Exception as e:  # una fila problemática no debe tumbar toda la auditoría
            probs, dup_txt = [("Medio", f"No se pudo auditar la planilla: {e}")], ""

        sev = min((s for s, _ in probs), key=lambda s: ORDEN_SEVERIDAD[s]) if probs else ""
        u = ctx.flota.get(p["patente"])
        camion = p["camion"] or (u["tipo"].split(" | ")[0] if u and u["tipo"] else "")
        filas.append({
            "Planilla": p["planilla"],
            "Vuelta": p["vuelta"],
            "Camión": camion,
            "Patente": p["patente"],
            "Zona": p["zona_orig"],
            "Clientes": p["cli"],
            "Paletas (OB)": p["ob"],
            "Peso (Kg)": p["kg"],
            "Estado": "INCUMPLE" if probs else ("PATENTE DUPLICADA" if dup_txt else "OK"),
            "Severidad": sev,
            "Incumplimiento Detectado": " ‖ ".join(f"{k}) {t}" for k, (_, t) in enumerate(probs, 1)),
            "Duplicación de Patente": dup_txt,
        })
    out = pd.DataFrame(filas)
    # Formatos: clientes y kg en unidades enteras, OB con 1 decimal
    out["Clientes"] = pd.to_numeric(out["Clientes"], errors="coerce").round(0).astype("Int64")
    out["Peso (Kg)"] = pd.to_numeric(out["Peso (Kg)"], errors="coerce").round(0).astype("Int64")
    out["Paletas (OB)"] = pd.to_numeric(out["Paletas (OB)"], errors="coerce").round(1)
    return out


def unidades_libres(ctx: Ctx, df: pd.DataFrame) -> pd.DataFrame:
    usadas = {1: set(df.loc[df["vuelta"] == 1, "patente"]), 2: set(df.loc[df["vuelta"] == 2, "patente"])}
    filas = []
    for pat in sorted(ctx.disponibles):
        l1, l2 = pat not in usadas[1], pat not in usadas[2]
        if not (l1 or l2):
            continue
        u = ctx.flota.get(pat, {})
        filas.append({
            "Patente": pat,
            "Disponibilidad": "Libre 1ª y 2ª vuelta" if (l1 and l2) else ("Libre solo 1ª vuelta" if l1 else "Libre solo 2ª vuelta"),
            "Paletas máx.": u.get("pal", float("nan")),
            "Peso máx. (Kg)": u.get("kg", float("nan")),
            "Línea": "Sí" if pat in ctx.linea else "No",
            "Tipo": u.get("tipo", ""),
        })
    return pd.DataFrame(filas, columns=["Patente", "Disponibilidad", "Paletas máx.", "Peso máx. (Kg)", "Línea", "Tipo"])


def facturadas_no_disponibles(df: pd.DataFrame, ctx: Ctx) -> pd.DataFrame:
    """Unidades que figuran en el informe (facturadas) pero no están DISPONIBLE ni en RESERVA en el parte diario."""
    cols = ["Planilla", "Vuelta", "Patente", "Zona", "Estado en parte diario"]
    d = df[(df["patente"] != "") & (~df["patente"].isin(ctx.disponibles))]
    filas = [{
        "Planilla": r.planilla, "Vuelta": r.vuelta, "Patente": r.patente, "Zona": r.zona_orig,
        "Estado en parte diario": ctx.estados.get(r.patente, "No figura en el parte diario"),
    } for r in d.itertuples()]
    return pd.DataFrame(filas, columns=cols)


def camiones_indisponibles(df: pd.DataFrame, ctx: Ctx) -> pd.DataFrame:
    """Todos los camiones de la flota (CAMIONES / parte diario) que no están DISPONIBLE ni en RESERVA."""
    cols = ["Patente", "Estado en parte diario", "Facturada en informe", "Paletas máx.", "Peso máx. (Kg)", "Línea", "Tipo"]
    usos: dict = {}
    for r in df.itertuples():
        if r.patente:
            usos.setdefault(r.patente, []).append(f"{r.planilla} ({r.vuelta}ª)")
    filas = []
    for pat in sorted(set(ctx.estados) | set(ctx.flota)):
        if pat in ctx.disponibles:
            continue
        u = ctx.flota.get(pat, {})
        filas.append({
            "Patente": pat,
            "Estado en parte diario": ctx.estados.get(pat, "No figura en el parte diario"),
            "Facturada en informe": ", ".join(usos[pat]) if pat in usos else "No",
            "Paletas máx.": u.get("pal", float("nan")),
            "Peso máx. (Kg)": u.get("kg", float("nan")),
            "Línea": "Sí" if pat in ctx.linea else "No",
            "Tipo": u.get("tipo", ""),
        })
    return pd.DataFrame(filas, columns=cols)


def patentes_duplicadas(df: pd.DataFrame) -> pd.DataFrame:
    """Patentes asignadas a más de una planilla dentro de la misma vuelta."""
    cols = ["Patente", "Vuelta", "Cant. planillas", "Planillas", "Zonas"]
    d = df[df["patente"] != ""]
    filas = []
    for (pat, vuelta), g in d.groupby(["patente", "vuelta"]):
        if len(g) > 1:
            filas.append({
                "Patente": pat,
                "Vuelta": vuelta,
                "Cant. planillas": len(g),
                "Planillas": ", ".join(g["planilla"]),
                "Zonas": " / ".join(z or "s/d" for z in g["zona_orig"]),
            })
    return pd.DataFrame(filas, columns=cols)


def a_excel(hojas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, d in hojas.items():
            nombre = nombre[:31]
            d.to_excel(w, sheet_name=nombre, index=False)
            ws = w.sheets[nombre]
            ws.freeze_panes = "A2"
            cabeceras = {c.value: c.column for c in ws[1]}
            for col_nombre, fmt in {"Clientes": "0", "Paletas (OB)": "0.0", "Peso (Kg)": "0"}.items():
                if col_nombre in cabeceras:
                    for fila in ws.iter_rows(min_row=2, min_col=cabeceras[col_nombre], max_col=cabeceras[col_nombre]):
                        fila[0].number_format = fmt
            if "Estado" in cabeceras:
                relleno = {"INCUMPLE": PatternFill("solid", fgColor="F8C9C9"),
                           "PATENTE DUPLICADA": PatternFill("solid", fgColor="FDE9B8")}
                for fila in ws.iter_rows(min_row=2):
                    f = relleno.get(fila[cabeceras["Estado"] - 1].value)
                    if f:
                        for celda in fila:
                            celda.fill = f
            for col in ws.columns:
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200]) + 2
                ws.column_dimensions[col[0].column_letter].width = min(max(ancho, 8), 70)
    return buf.getvalue()


FORMATO_NUM = {"Clientes": "{:.0f}", "Paletas (OB)": "{:.1f}", "Peso (Kg)": "{:.0f}"}
COLOR_FILA = {
    "INCUMPLE": "background-color:#f8c9c9;color:#7f1d1d",
    "PATENTE DUPLICADA": "background-color:#fde9b8;color:#78350f",
}


def estilo_tabla(d: pd.DataFrame, colorear_filas: bool = False):
    """Formatea números (Clientes y Kg enteros, OB 1 decimal) y, si se pide, pinta la fila entera según Estado."""
    sty = d.style.format({k: v for k, v in FORMATO_NUM.items() if k in d.columns}, na_rep="")
    if colorear_filas and "Estado" in d.columns:
        sty = sty.apply(lambda r: [COLOR_FILA.get(r["Estado"], "")] * len(r), axis=1)
    return sty


# ════════════════════════════════════════════════════════════════════════════
# INTERFAZ
# ════════════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="Auditoría de Flota · Quilmes Santa Fe", page_icon="🚛", layout="wide")
    st.title("🚛 Tablero de Auditoría y Control de Asignación de Flota")
    st.caption("Cervecería Quilmes · Operación Santa Fe")

    informe_file = st.file_uploader(
        "Soltá acá el Informe de Entrega (.xlsx) — se lee solo la pestaña 'Salidas'",
        type=["xlsx", "xlsm"], key="informe",
    )

    # ── Sidebar: base de flota ───────────────────────────────────────────────
    with st.sidebar:
        st.header("Base de flota")
        if st.button("🔄 Actualizar Flota desde Google Sheets"):
            descargar_flota.clear()
            st.rerun()
        flota_manual = st.file_uploader(
            "…o subí el archivo de flota manualmente (si falla la red)", type=["xlsx", "xlsm"], key="flota_manual"
        )

    contenido, origen = None, ""
    if flota_manual is not None:
        contenido, origen = flota_manual.getvalue(), f"archivo manual ({flota_manual.name})"
    else:
        try:
            contenido, ts = descargar_flota(URL_FLOTA)
            origen = f"Google Sheets · descargado {ts:%d/%m %H:%M}"
        except Exception as e:
            st.sidebar.error(f"No se pudo descargar la flota: {e}")

    if contenido is None:
        st.warning("No hay base de flota disponible. Subí el archivo de flota manualmente en la barra lateral.")
        st.stop()

    try:
        df_cam, df_disp = parsear_flota(contenido)
    except Exception as e:
        st.error(f"No se pudo leer la base de flota: {e}")
        st.stop()

    st.sidebar.success(f"Flota cargada desde {origen}")

    cols = detectar_cols_camiones(df_cam)
    etiquetas_disp = list(df_disp.columns)
    def_pat_disp = next((c for c in etiquetas_disp if clave(c.split("·", 1)[-1]) in CLAVES_PATENTE), None) \
        or (etiquetas_disp[4] if len(etiquetas_disp) > 4 else None)          # columna E
    def_estado_disp = etiquetas_disp[9] if len(etiquetas_disp) > 9 else None  # columna J

    with st.sidebar:
        with st.expander("Mapeo de columnas de flota (avanzado)"):
            todas = list(df_cam.columns)

            def sel(etiqueta, valor, opciones):
                idx = opciones.index(valor) if valor in opciones else 0
                return st.selectbox(etiqueta, opciones, index=idx)

            if not todas:
                st.error("La pestaña CAMIONES no tiene columnas.")
                st.stop()
            col_pat = sel("CAMIONES · Patente", cols["patente"], todas)
            col_kg = sel("CAMIONES · Peso máximo (kg)", cols["kg"], todas)
            col_pal = sel("CAMIONES · Paletas máx. (columna P)", cols["pal"], todas)
            col_pat_disp = sel("Disponibilidad · Matrícula/Patente (col. E)", def_pat_disp, etiquetas_disp)
            col_estado = sel("Disponibilidad · Estado (col. J)", def_estado_disp, etiquetas_disp)
            vals = df_disp[col_estado].map(norm_txt).replace("", pd.NA).dropna().value_counts().head(8)
            st.caption("Valores en la columna de estado: " + ", ".join(f"{k} ({v})" for k, v in vals.items()))

        st.header("Reglas configurables")
        cols_tipo = detectar_cols_tipo(df_cam, {col_pat, col_kg, col_pal})
        flota = construir_flota(df_cam, col_pat, col_kg, col_pal, cols_tipo)
        linea_txt = st.text_area(
            "Patentes de flota de línea (11 habilitadas, con prioridad)", value=", ".join(LINEA_FIJA),
            help="Tienen prioridad en zonas de línea. Solo si ninguna está disponible se puede usar una unidad de Santa Fe / Santo Tomé.",
        )
        linea = {norm_pat(x) for x in parsear_lista(linea_txt)}
        if not linea:
            st.warning("Sin patentes de línea cargadas: la regla de pueblos/rutas queda desactivada.")
        elif len(linea) != 11:
            st.warning(f"Hay {len(linea)} patentes de línea cargadas (se esperaban 11).")
        pueblos_txt = st.text_area("Zonas de línea / pueblo (Interior se controla aparte)", value=", ".join(PUEBLOS_DEFAULT))
        pueblos = parsear_lista(pueblos_txt)

    if not flota:
        st.error("No se pudo construir la flota desde CAMIONES. Revisá el mapeo de columnas en la barra lateral.")
        st.stop()

    disponibles = construir_disponibles(df_disp, col_pat_disp, col_estado)
    if not disponibles:
        st.error(
            f"No se encontró ninguna unidad DISPONIBLE usando '{col_pat_disp}' (patente) y '{col_estado}' (estado). "
            "Revisá el mapeo de columnas en la barra lateral."
        )
        st.dataframe(df_disp.head(40))
        st.stop()
    sin_datos = sorted(disponibles - set(flota))
    ctx = Ctx(flota, disponibles, linea, pueblos)
    ctx.estados = construir_estados(df_disp, col_pat_disp, col_estado)

    if informe_file is None:
        st.info("Cargá el Informe de Entrega para iniciar la auditoría.")
        st.stop()

    try:
        df, avisos = cargar_informe(informe_file)
    except Exception as e:
        st.error(f"No se pudo procesar el informe: {e}")
        st.stop()
    if df.attrs.get("mapa_vuelta"):
        st.caption(
            f"Vuelta leída de la columna '{df.attrs['col_vuelta']}': "
            + " · ".join(f"{k} → {v}ª" for k, v in list(df.attrs["mapa_vuelta"].items())[:12])
        )
    for a in avisos:
        st.warning(a)
    if sin_datos:
        st.warning(f"Disponibles sin datos en CAMIONES (no se proponen como alternativa): {', '.join(sin_datos)}")

    with st.spinner("Auditando planillas…"):
        res = auditar(df, ctx)
        libres = unidades_libres(ctx, df)
        df_dups = patentes_duplicadas(df)
        df_nd = facturadas_no_disponibles(df, ctx)
        df_indisp = camiones_indisponibles(df, ctx)
        df_sin_usar = libres[libres["Disponibilidad"] == "Libre 1ª y 2ª vuelta"].reset_index(drop=True)

    total = len(res)
    n_inc = int((res["Estado"] == "INCUMPLE").sum())
    n_dup = int((res["Duplicación de Patente"] != "").sum())
    n_ok = int((res["Estado"] == "OK").sum())
    sin_asignar = len(df_sin_usar)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Planillas procesadas", total)
    k2.metric("Planillas OK", n_ok)
    k3.metric("Con incumplimiento", n_inc)
    k4.metric("Con patente duplicada", n_dup)
    k5.metric("Unidades disponibles sin asignar", sin_asignar)
    st.caption(f"Flota: {origen} · {len(ctx.disponibles)} unidades DISPONIBLES · columna de estado: {col_estado}")

    tab_alertas, tab_audit = st.tabs(["🚨 Alertas e incumplimientos", "📋 Auditoría completa"])

    base = ["Planilla", "Vuelta", "Camión", "Patente", "Zona", "Clientes", "Paletas (OB)", "Peso (Kg)"]
    cols_alerta = base + ["Incumplimiento Detectado"]
    cols_dup = base + ["Duplicación de Patente"]
    cfg = {
        "Incumplimiento Detectado": st.column_config.TextColumn(width="large"),
        "Duplicación de Patente": st.column_config.TextColumn(width="large"),
    }

    with tab_alertas:
        st.markdown("#### Incumplimientos de reglas operativas")
        alertas = res[res["Estado"] == "INCUMPLE"].copy()
        if alertas.empty:
            st.success("✅ Todas las planillas cumplen las reglas operativas.")
        else:
            f1, f2 = st.columns(2)
            sevs = f1.multiselect("Severidad", ["Crítico", "Alto", "Medio"], default=["Crítico", "Alto", "Medio"])
            vueltas = f2.multiselect("Vuelta", sorted(alertas["Vuelta"].unique()), default=sorted(alertas["Vuelta"].unique()))
            alertas = alertas[alertas["Severidad"].isin(sevs) & alertas["Vuelta"].isin(vueltas)]
            alertas = alertas.sort_values("Severidad", key=lambda s: s.map(ORDEN_SEVERIDAD), kind="stable")
            st.caption(f"{len(alertas)} planilla(s) con desvíos. Ordenadas por severidad.")
            st.dataframe(estilo_tabla(alertas[["Severidad"] + cols_alerta]), hide_index=True, column_config=cfg)

        st.markdown("#### Duplicación de patente (misma patente en más de una planilla de la misma vuelta)")
        dups_pl = res[res["Duplicación de Patente"] != ""].sort_values(["Patente", "Vuelta"])
        if dups_pl.empty:
            st.success("✅ No hay patentes duplicadas dentro de una misma vuelta.")
        else:
            st.caption(f"{len(dups_pl)} planilla(s) involucradas.")
            st.dataframe(estilo_tabla(dups_pl[cols_dup]), hide_index=True, column_config=cfg)

    with tab_audit:
        st.caption("🔴 Fila en rojo: incumple reglas · 🟡 Fila en amarillo: solo tiene patente duplicada")
        st.dataframe(estilo_tabla(res, colorear_filas=True), hide_index=True, column_config=cfg)
        excel = a_excel({
            "Auditoría completa": res,
            "Alertas": res[res["Estado"] == "INCUMPLE"],
            "Patente duplicada": res[res["Duplicación de Patente"] != ""],
            "Unidades libres": libres,
            "Patentes duplicadas": df_dups,
            "Facturadas no disponibles": df_nd,
            "Camiones indisponibles": df_indisp,
        })
        st.download_button(
            "⬇️ Descargar auditoría en Excel", data=excel,
            file_name=f"auditoria_flota_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── Secciones adicionales ────────────────────────────────────────────────
    st.divider()
    st.subheader("📌 Resumen operativo adicional")
    col_izq, col_der = st.columns(2)

    with col_izq:
        st.markdown("### 🪞 Patentes duplicadas en el informe")
        if df_dups.empty:
            st.success("No se detectaron patentes duplicadas en la misma vuelta.")
        else:
            st.warning(f"Se encontraron {len(df_dups)} patente(s) duplicadas en la misma vuelta.")
            st.dataframe(df_dups, hide_index=True)

    with col_der:
        st.markdown("### 🟢 Patentes disponibles sin utilizar (100% libres)")
        if df_sin_usar.empty:
            st.info("Todas las unidades disponibles fueron utilizadas al menos en una vuelta.")
        else:
            st.success(f"{len(df_sin_usar)} unidad(es) disponibles sin asignar en todo el día.")
            st.dataframe(df_sin_usar, hide_index=True)

    st.markdown("### 🚫 Unidades facturadas que figuran como indisponibles")
    if df_nd.empty:
        st.success("Todas las unidades facturadas figuran como DISPONIBLE o en RESERVA en el parte diario.")
    else:
        st.error(f"{len(df_nd)} planilla(s) con unidades que no están disponibles ni en reserva según el parte diario.")
        st.dataframe(df_nd, hide_index=True)

    st.markdown("### ⛔ Camiones indisponibles (parte diario)")
    if df_indisp.empty:
        st.success("No hay camiones indisponibles: toda la flota figura DISPONIBLE o en RESERVA.")
    else:
        st.caption(f"{len(df_indisp)} camión(es) que no están DISPONIBLE ni en RESERVA. "
                   "La columna 'Facturada en informe' muestra si igual aparecen asignados en alguna planilla.")
        st.dataframe(df_indisp, hide_index=True)

    df_parciales = libres[libres["Disponibilidad"] != "Libre 1ª y 2ª vuelta"]
    if not df_parciales.empty:
        with st.expander(f"🟡 Unidades libres solo en una vuelta ({len(df_parciales)})"):
            st.dataframe(df_parciales, hide_index=True)


if __name__ == "__main__":
    main()
