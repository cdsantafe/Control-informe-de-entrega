# -*- coding: utf-8 -*-
"""
Tablero de Auditoría y Control de Asignación de Flota
Cervecería Quilmes · Operación Santa Fe

Ejecutar:    streamlit run app.py
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

# ════════════════════════════════════════════════════════════════════════════
# CONSTANTES DE NEGOCIO
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
# UTILIDADES DE TEXTO / NÚMEROS
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

    col_pl = mapa["planilla"]
    raw = raw[raw[col_pl].notna()]
    raw = raw[~raw[col_pl].map(norm_txt).isin(["", "(EN BLANCO)"])].copy()
    if raw.empty:
        raise ValueError("La pestaña 'Salidas' no tiene planillas válidas luego de la limpieza.")

    df = pd.DataFrame(index=raw.index)
    df["planilla"] = raw[col_pl].map(fmt_planilla)
    df["patente"] = raw[mapa["patente"]].map(norm_pat)
    df["zona_orig"] = raw[mapa["zona"]].fillna("").astype(str).str.strip()
    df["zona"] = raw[mapa["zona"]].map(norm_txt)
    df["cli"] = a_numero(raw[mapa["cli"]])
    df["ob"] = a_numero(raw[mapa["ob"]])
    df["kg"] = a_numero(raw[mapa["kg"]])

    if "cap" in mapa:
        df["cap"] = a_numero(raw[mapa["cap"]])
    else:
        df["cap"] = float("nan")

    if "vuelta" in mapa:
        df["vuelta"] = raw[mapa["vuelta"]].map(parse_vuelta)
    else:
        df["vuelta"] = 1

    if "camion" in mapa:
        df["camion"] = raw[mapa["camion"]].fillna("").astype(str).str.strip()
    else:
        df["camion"] = ""

    return df.reset_index(drop=True), avisos


# ════════════════════════════════════════════════════════════════════════════
# CARGA: BASE DE FLOTA (CAMIONES y Disponibilidad Flota Diaria)
# ════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=900, show_spinner="Descargando flota desde Google Sheets…")
def descargar_flota(url: str):
    r = requests.get(url, timeout=45)
    r.raise_for_status()
    return r.content, datetime.now()


@st.cache_data(show_spinner="Procesando base de flota…")
def parsear_flota(contenido: bytes):
    xl = pd.ExcelFile(io.BytesIO(contenido))
    h_cam = next((h for h in xl.sheet_names if clave(h) == "CAMIONES"), None)
    h_disp = next((h for h in xl.sheet_names if "DISPONIBILIDAD" in clave(h)), None)
    
    if h_cam is None or h_disp is None:
        raise ValueError(
            f"No se encontraron las pestañas necesarias ('CAMIONES' y 'Disponibilidad Flota Diaria'). Encontradas: {', '.join(xl.sheet_names)}"
        )
        
    # 1. Pestaña CAMIONES
    raw_cam = xl.parse(h_cam, header=None, dtype=object).dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    idx_cam = 0
    for i in range(min(25, len(raw_cam))):
        if any(clave(v) in CLAVES_PATENTE for v in raw_cam.iloc[i].tolist()):
            idx_cam = i
            break
    cab_cam, vistos = [], {}
    for j, v in enumerate(raw_cam.iloc[idx_cam].tolist()):
        nombre = f"COL{j}" if v is None or pd.isna(v) else re.sub(r"\s+", " ", str(v)).strip()
        if nombre in vistos:
            vistos[nombre] += 1
            nombre = f"{nombre}.{vistos[nombre]}"
        else:
            vistos[nombre] = 0
        cab_cam.append(nombre)
    df_cam = raw_cam.iloc[idx_cam + 1:].copy()
    df_cam.columns = cab_cam
    df_cam = df_cam.reset_index(drop=True)

    # 2. Pestaña Disponibilidad Flota Diaria (Matrícula en Columna E / índice 4, Estado en Columna J / índice 9)
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
    col_pat = next((claves[k] for k in CLAVES_PATENTE if k in claves), None)
    col_pal = claves.get("P") # Columna P asignada a paletas máx.
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
    d["_kg"] = a_numero(d[col_kg]) if col_kg in d else float("nan")
    d["_pal"] = a_numero(d[col_pal]) if col_pal in d else float("nan")
    flota: dict = {}
    for _, r in d.iterrows():
        p = norm_pat(r[col_pat])
        if not re.fullmatch(r"[A-Z0-9]{6,7}", p) or p in flota:
            continue
        partes = [norm_txt(r[c]) for c in cols_tipo if c in r]
        tipo = " | ".join(x for x in partes if x)
        flota[p] = {"kg": r["_kg"], "pal": r["_pal"], "tipo": tipo, "linea": "LINEA" in tipo}
    return flota


def construir_disponibles_directo(df_disp_proc: pd.DataFrame) -> set:
    disp = {r["PATENTE"] for _, r in df_disp_proc.iterrows() if r["ESTADO"] == "DISPONIBLE"}
    disp.discard("")
    return disp


# ════════════════════════════════════════════════════════════════════════════
# MOTOR DE REGLAS
# ════════════════════════════════════════════════════════════════════════════
class Ctx:
    def __init__(self, flota: dict, disponibles: set, linea: set, pueblos: list[str]):
        self.flota = flota
        self.disponibles = disponibles
        self.linea = linea
        pueblos_n = [norm_txt(p) for p in pueblos if norm_txt(p)]
        self.re_pueblo = (
            re.compile(r"\b(?:" + "|".join(re.escape(p) for p in pueblos_n) + r")\b") if pueblos_n else None
        )


def evaluar(p: dict, pat: str, ctx: Ctx, verificar_disp: bool = True) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    if not pat:
        return [("Crítico", "La planilla no tiene patente asignada.")]

    u = ctx.flota.get(pat)
    es_actual = pat == p["patente"]
    zona, zona_o = p["zona"], p["zona_orig"] or "s/d"
    cli, ob, kg = p["cli"], p["ob"], p["kg"]

    # 1) Disponibilidad
    if verificar_disp and pat not in ctx.disponibles:
        out.append(("Alto", f"Disponibilidad: la patente {pat} no figura como DISPONIBLE en el parte diario de flota."))

    # 2) Capacidad física
    if u is None:
        out.append(("Medio", f"Capacidad: la patente {pat} no figura en la pestaña CAMIONES, no se puede validar peso ni paletas."))
    max_kg = u["kg"] if u is not None and pd.notna(u["kg"]) else (p["cap"] if es_actual else float("nan"))
    max_pal = u["pal"] if u is not None else float("nan")
    if pd.notna(kg) and pd.notna(max_kg) and kg > max_kg:
        out.append(("Alto", f"Sobrepeso: Kgr Truck {n(kg)} kg supera el máximo de la unidad ({n(max_kg)} kg)."))
    if pd.notna(ob) and pd.notna(max_pal) and ob > max_pal:
        out.append(("Alto", f"Exceso de paletas: OB {n(ob)} supera la capacidad de la unidad ({n(max_pal)} paletas)."))

    # 3) Restricciones geográficas y operativas
    es_pueblo = bool(ctx.re_pueblo and ctx.re_pueblo.search(zona))
    tipo = (u["tipo"] if u is not None else "")
    if es_actual:
        tipo = f"{tipo} | {norm_txt(p['camion'])}"

    # RESTRICCIÓN AG622AO
    if pat == UNIDAD_RESTRINGIDA:
        if not any(z in zona for z in ZONAS_AG622AO):
            out.append(("Crítico", f"Incumplimiento crítico: AG622AO solo puede usarse en San Justo, San Javier o Emilia/Videla y está asignada a '{zona_o}'."))
    elif es_pueblo and ctx.linea and pat not in ctx.linea:
        out.append(("Alto", f"Pueblo/ruta: la zona '{zona_o}' solo puede cubrirla una de las unidades habilitadas de línea y {pat} no lo es."))

    # EXCLUSIVIDAD CENTRO
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
            out.append(("Alto", f"Centro/Puerto: zona preferencial de AH474WT y PAP418; {pat} incumple ({'; '.join(motivos)})."))

    # SMK MAYORISTA
    if RE_SMK.search(zona) and pd.notna(cli) and cli < SMK_MAX_CLIENTES and pat not in POOL_SMK:
        out.append(("Alto", f"SMK Mayorista con {n(cli)} clientes (< {SMK_MAX_CLIENTES}): debe cubrirlo el pool {', '.join(sorted(POOL_SMK))}."))

    # CAMIONES DE GAS
    if pat in UNIDADES_GAS and not (pd.notna(cli) and cli < GAS_EXCEPCION_CLI):
        if pd.notna(cli) and cli > GAS_MAX_CLI:
            out.append(("Alto", f"Camión de gas: {n(cli)} clientes supera el máximo de {GAS_MAX_CLI}."))
        if pd.notna(ob) and ob > GAS_MAX_OB:
            out.append(("Alto", f"Camión de gas: OB {n(ob)} supera el máximo de {GAS_MAX_OB} paletas."))

    # CABLES BAJOS
    if pat in CABLES_BAJOS_PROHIBIDAS:
        hits = [nom for nom, rx in ZONAS_CABLES_BAJOS.items() if re.search(rx, zona)]
        if hits:
            out.append(("Alto", f"Cables bajos: la zona '{zona_o}' ({', '.join(hits)}) tiene prohibida la unidad {pat}."))

    return out


# ════════════════════════════════════════════════════════════════════════════
# MOTOR DE ALTERNATIVAS
# ════════════════════════════════════════════════════════════════════════════
def sugerir(i: int, planillas: list[dict], ocup: dict, ctx: Ctx, max_alt: int = 3) -> str:
    p = planillas[i]
    v, otra = p["vuelta"], (2 if p["vuelta"] == 1 else 1)
    cands = []
    for cand in ctx.disponibles:
        u = ctx.flota.get(cand)
        if cand == p["patente"] or u is None or pd.isna(u["kg"]) or pd.isna(u["pal"]):
            continue
        if evaluar(p, cand, ctx, verificar_disp=False):
            continue
        ocup_v = ocup[v].get(cand, [])
        ocup_o = ocup[otra].get(cand, [])
        info = f"{n(u['pal'])} pal / {n(u['kg'])} kg"
        if not ocup_v and not ocup_o:
            cands.append((1, u["pal"], u["kg"], f"{cand} · 100% libre (1ª y 2ª vuelta) · {info}"))
        elif not ocup_v:
            pl = ", ".join(planillas[j]["planilla"] for j in ocup_o)
            cands.append((2, u["pal"], u["kg"], f"{cand} · libre en {v}ª vuelta (ocupada en {otra}ª, planilla {pl}) · {info}"))
        else:
            if len(ocup_v) == 1 and p["patente"]:
                q = planillas[ocup_v[0]]
                if not evaluar(q, p["patente"], ctx, verificar_disp=False) and p["patente"] in ctx.disponibles:
                    cands.append((3, u["pal"], u["kg"],
                                  f"{cand} · requiere enroque con planilla {q['planilla']} ({q['zona_orig'] or 's/d'}), "
                                  f"donde {p['patente']} sí cumple · {info}"))
    if not cands:
        return "Sin alternativas válidas en la flota disponible."
    cands.sort(key=lambda x: (x[0], x[1], x[2]))
    return " | ".join(f"{k}) {c[3]}" for k, c in enumerate(cands[:max_alt], 1))


# ════════════════════════════════════════════════════════════════════════════
# AUDITORÍA COMPLETA
# ════════════════════════════════════════════════════════════════════════════
def auditar(df: pd.DataFrame, ctx: Ctx) -> pd.DataFrame:
    planillas = df.to_dict("records")
    ocup: dict = {1: {}, 2: {}}
    for i, p in enumerate(planillas):
        if p["patente"]:
            ocup[p["vuelta"]].setdefault(p["patente"], []).append(i)

    filas = []
    for i, p in enumerate(planillas):
        try:
            probs = evaluar(p, p["patente"], ctx)
            dup = ocup[p["vuelta"]].get(p["patente"], [])
            if p["patente"] and len(dup) > 1:
                otras = ", ".join(planillas[j]["planilla"] for j in dup if j != i)
                probs.append(("Medio", f"Doble asignación: la patente también está en la/s planilla/s {otras} de la misma vuelta."))
            alt = sugerir(i, planillas, ocup, ctx) if probs else ""
        except Exception as e:
            probs, alt = [("Medio", f"No se pudo auditar la planilla: {e}")], ""

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
            "Estado": "INCUMPLE" if probs else "OK",
            "Severidad": sev,
            "Incumplimiento Detectado": " ‖ ".join(f"{k}) {t}" for k, (_, t) in enumerate(probs, 1)),
            "Alternativas Sugeridas": alt,
        })
    return pd.DataFrame(filas)


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


def a_excel(hojas: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nombre, d in hojas.items():
            nombre = nombre[:31]
            d.to_excel(w, sheet_name=nombre, index=False)
            ws = w.sheets[nombre]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200]) + 2
                ws.column_dimensions[col[0].column_letter].width = min(max(ancho, 8), 70)
    return buf.getvalue()


def estilo_estado(d: pd.DataFrame):
    def color(v):
        return {"OK": "background-color:#d9f2e0;color:#14532d",
                "INCUMPLE": "background-color:#fde2e2;color:#7f1d1d"}.get(v, "")
    sty = d.style
    return sty.map(color, subset=["Estado"]) if hasattr(sty, "map") else sty.applymap(color, subset=["Estado"])


# ════════════════════════════════════════════════════════════════════════════
# INTERFAZ (UI STREAMLIT)
# ════════════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="Auditoría de Flota · Quilmes Santa Fe", page_icon="🚛", layout="wide")
    st.title("🚛 Tablero de Auditoría y Control de Asignación de Flota")
    st.caption("Cervecería Quilmes · Operación Santa Fe")

    informe_file = st.file_uploader(
        "Soltá acá el Informe de Entrega (.xlsx) — se lee solo la pestaña 'Salidas'",
        type=["xlsx", "xlsm"], key="informe",
    )

    with st.sidebar:
        st.header("Base de flota")
        if st.button("🔄 Actualizar Flota desde Google Sheets"):
            descargar_flota.clear()
