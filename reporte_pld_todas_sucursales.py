"""
Reporte PLD — Todas las Sucursales (11 hojas)
=============================================
Genera el archivo Reporte_PLD_TODAS_SUCURSALES_<fecha>.xlsx

Uso:
    1. Edita la sección CONFIGURACIÓN con tus rutas y fechas
    2. Ejecuta: python reporte_pld_todas_sucursales.py
"""

import re, math, json
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ══════════════════════════════════════════════════════════════
# CONFIGURACIÓN — edita solo esta sección en cada nuevo corte
# ══════════════════════════════════════════════════════════════

# Carpeta con los archivos del corte ANTERIOR
CARPETA_ANTERIOR = Path(r"C:\Users\Misael Urbina\OneDrive - CECOM\Documentos\PYTHON\17 MAYO 2026")

# Carpeta con los archivos del corte ACTUAL
CARPETA_ACTUAL = Path(r"C:\Users\Misael Urbina\OneDrive - CECOM\Documentos\PYTHON\24 MAYO 2026")

# Carpeta donde se guardará el reporte generado
CARPETA_SALIDA = Path(r"C:\Users\Misael Urbina\OneDrive - CECOM\Documentos\PYTHON\24 MAYO 2026\COMPARATIVAS")

# Etiquetas de fecha para los encabezados
FECHA_ANTERIOR_CORTA = "17-May"
FECHA_ACTUAL_CORTA   = "24-May"
FECHA_ANTERIOR_LARGA = "17 de Mayo 2026"
FECHA_ACTUAL_LARGA   = "24 de Mayo 2026"

# Rangos de semana para H8/H9 (semana anterior y semana nueva)
SEMANA_ANT_INICIO = pd.Timestamp("2026-05-11")
SEMANA_ANT_FIN    = pd.Timestamp("2026-05-17")
SEMANA_NVA_INICIO = pd.Timestamp("2026-05-18")
SEMANA_NVA_FIN    = pd.Timestamp("2026-05-24")

# YTD — clientes creados desde esta fecha
YTD_INICIO = pd.Timestamp("2026-01-01")

# Sucursales a procesar (el script busca el archivo automáticamente por nombre)
SUCURSALES = [
    "Ajijic", "Allende", "Esmeralda", "Guadalupe",
    "Irapuato", "Mitras", "Neza", "Satelite", "Sendero",
]

# ══════════════════════════════════════════════════════════════
# CONSTANTES (no editar)
# ══════════════════════════════════════════════════════════════

FECHA_ID_VENCIDA = pd.Timestamp("2025-12-31")
HOJA_DATOS       = "KYCTotal"

OPS_STATUS = [
    "Bloqueado por operaciones", "Información actualizada operaciones",
    "Registro completo", "Verificado por operaciones",
]
KYC_STATUS = [
    "Bloqueado por KYC", "Información actualizada Kyc",
    "Pendente Validacion Kyc", "Verificado por Kyc",
]
TODOS_STATUS = OPS_STATUS + KYC_STATUS + ["Registro express"]

CATEGORIAS_NOTAS = [
    "CLIENTE EXCLUIDO SEGOB", "CLIENTE VETADO", "CLIENTE VERIFICADO",
    "ID VENCIDA", "ID NO VALIDA", "ID INCOMPLETA VOLVER A ESCANEAR",
    "FALTA COMPROBANTE DE DOMICILIO", "FALTA OCUPACION", "DATOS DE CONTACTO",
    "SIN HUELLAS DIGITALES", "SIN AVISO FIRMADO", "SIN AVISO FIRMADO - SISTEMA",
    "FALTA NACIONALIDAD", "CODIGO POSTAL DISTINTO", "ERROR DE RFC",
    "OTROS", "SIN NOTA",
]

C = {
    "azul_marino":  "1F3864", "azul_medio":   "2E75B6",
    "verde_oscuro": "375623", "naranja":       "E26B0A",
    "azul_claro":   "BDD7EE", "verde_claro":  "E2EFDA",
    "verde_varia":  "C6EFCE", "rojo_varia":   "FFC7CE",
    "salmon":       "FCE4D6", "rosa":          "F4CCCC",
    "amarillo":     "FFF2CC", "gris_claro":   "F2F2F2",
    "blanco":       "FFFFFF",
}
COLOR_NIVEL = {
    "Corazon": "DEEBF7", "Corazon Plus": "FFF2CC",
    "Diamante": "E2EFDA", "Diamante Plus": "FCE4D6",
    "Espada": "EAF0FB",  "Espada Plus": "E2EFDA",
    "Foliatti Fan": "FFFFFF", "Trebol": "D9EAD3", "SIN NIVEL": "F2F2F2",
}
COLOR_STATUS = {
    "Bloqueado por KYC": "F4CCCC",
    "Bloqueado por operaciones": "FCE4D6",
    "Información actualizada Kyc": "E2EFDA",
    "Información actualizada operaciones": "BDD7EE",
    "Pendente Validacion Kyc": "DEEBF7",
    "Registro completo": "EAF0FB",
    "Registro express": "FFF2CC",
    "Verificado por Kyc": "E2EFDA",
    "Verificado por operaciones": "EAF0FB",
}

# ══════════════════════════════════════════════════════════════
# HELPERS DE ESTILO
# ══════════════════════════════════════════════════════════════

def fill(h):   return PatternFill("solid", fgColor=h)
def font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Arial")
def aln(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def s(cell, bg=None, bold=False, color="000000", size=10,
      h="left", v="center", wrap=False):
    if bg: cell.fill = fill(bg)
    cell.font      = font(bold=bold, color=color, size=size)
    cell.alignment = aln(h=h, v=v, wrap=wrap)

def titulo(ws, txt, fila, ncols, size=14):
    ws.merge_cells(start_row=fila, start_column=1,
                   end_row=fila,   end_column=ncols)
    c = ws.cell(fila, 1, txt)
    s(c, bg=C["azul_marino"], bold=True, color="FFFFFF", size=size)
    ws.row_dimensions[fila].height = 36

def seccion(ws, txt, fila, ncols, bg="azul_marino", size=11):
    ws.merge_cells(start_row=fila, start_column=1,
                   end_row=fila,   end_column=ncols)
    c = ws.cell(fila, 1, txt)
    s(c, bg=C[bg], bold=True, color="FFFFFF", size=size)
    ws.row_dimensions[fila].height = 22

def enc(ws, headers, fila, bg="azul_marino"):
    for j, h in enumerate(headers, 1):
        c = ws.cell(fila, j, h)
        s(c, bg=C[bg], bold=True, color="FFFFFF", size=10,
          h="center" if j > 1 else "left", wrap=True)
    ws.row_dimensions[fila].height = 28

def var_cell(cell, val):
    v = str(val)
    if v.startswith("+") and v != "+":
        cell.fill = fill(C["verde_varia"])
        cell.font = Font(bold=True, color="375623", size=10, name="Arial")
    elif v.startswith("-") and v not in ["-", "0"]:
        cell.fill = fill(C["rojo_varia"])
        cell.font = Font(bold=True, color="9C0006", size=10, name="Arial")
    cell.alignment = aln(h="center")

def var_str(n):
    if n == 0: return "0"
    return f"+{n:,}" if n > 0 else f"{n:,}"

def pct(n, total):
    return f"{n/total*100:.1f}%" if total else "0.0%"

# ══════════════════════════════════════════════════════════════
# LECTURA Y NORMALIZACIÓN
# ══════════════════════════════════════════════════════════════

STATUS_MAPA = {
    "registro express": "Registro express",
    "bloqueado por operaciones": "Bloqueado por operaciones",
    "bloqueo por operaciones": "Bloqueado por operaciones",
    "información actualizada operaciones": "Información actualizada operaciones",
    "informacion actualizada operaciones": "Información actualizada operaciones",
    "registro completo": "Registro completo",
    "verificado por operaciones": "Verificado por operaciones",
    "bloqueado por kyc": "Bloqueado por KYC",
    "bloqueo por kyc": "Bloqueado por KYC",
    "información actualizada kyc": "Información actualizada Kyc",
    "informacion actualizada kyc": "Información actualizada Kyc",
    "pendente validacion kyc": "Pendente Validacion Kyc",
    "pendiente validacion kyc": "Pendente Validacion Kyc",
    "pendiente validación kyc": "Pendente Validacion Kyc",
    "verificado por kyc": "Verificado por Kyc",
}

def limpiar(x):
    if x is None or (isinstance(x, float) and math.isnan(x)): return ""
    return str(x).strip()

def norm_status(x):
    t = re.sub(r"\s+", " ", limpiar(x).lower())
    return STATUS_MAPA.get(t, limpiar(x))

def clasificar_nota(nota):
    t = re.sub(r"\s+", " ", limpiar(nota).upper())
    if not t:                                             return "SIN NOTA"
    if "SEGOB" in t or "EXCLUIDO" in t:                  return "CLIENTE EXCLUIDO SEGOB"
    if "VETADO" in t:                                     return "CLIENTE VETADO"
    if "VERIFICADO" in t:                                 return "CLIENTE VERIFICADO"
    if "ID VENCIDA" in t or "INE VENCIDA" in t:           return "ID VENCIDA"
    if "ID NO VALIDA" in t or "INE NO VALIDA" in t:       return "ID NO VALIDA"
    if "ID INCOMPLETA" in t or "VOLVER A ESCANEAR" in t:  return "ID INCOMPLETA VOLVER A ESCANEAR"
    if "COMPROBANTE" in t or "DOMICILIO" in t:            return "FALTA COMPROBANTE DE DOMICILIO"
    if "OCUPACION" in t or "OCUPACIÓN" in t:              return "FALTA OCUPACION"
    if "CONTACTO" in t or "CORREO" in t or "TELEFONO" in t or "TELÉFONO" in t:
        return "DATOS DE CONTACTO"
    if "HUELLA" in t:                                     return "SIN HUELLAS DIGITALES"
    if "AVISO" in t and "SISTEMA" in t:                   return "SIN AVISO FIRMADO - SISTEMA"
    if "AVISO" in t:                                      return "SIN AVISO FIRMADO"
    if "NACIONALIDAD" in t:                               return "FALTA NACIONALIDAD"
    if "CODIGO POSTAL" in t or "CÓDIGO POSTAL" in t or " CP " in f" {t} ":
        return "CODIGO POSTAL DISTINTO"
    if "RFC" in t:                                        return "ERROR DE RFC"
    return "OTROS"

def es_si(x):
    return limpiar(x).upper() in ["SI", "SÍ", "1", "1.0", "TRUE", "VERDADERO", "YES"]

def buscar_col(df, opts):
    mapa = {str(c).strip().upper().replace("\n", " ").replace("_", " "): c
            for c in df.columns}
    for o in opts:
        r = mapa.get(o.strip().upper().replace("_", " "))
        if r: return r
    return None

def buscar_archivo(carpeta: Path, sucursal: str) -> Path | None:
    """Busca el xlsx de una sucursal en la carpeta sin importar el nombre exacto."""
    nombre_norm = sucursal.upper().replace(" ", "")
    for archivo in sorted(carpeta.glob("*.xlsx")):
        if archivo.name.startswith("~$"):
            continue
        stem_norm = archivo.stem.upper().replace(" ", "").replace("_", "")
        if nombre_norm in stem_norm:
            return archivo
    return None

def leer_df(path: Path, sucursal: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=HOJA_DATOS)
    df.columns = [str(c).strip() for c in df.columns]

    col_id      = buscar_col(df, ["PLAYER_ID", "ID_JUGADOR", "Player ID", "ID"])
    col_status  = buscar_col(df, ["STATUS", "ESTATUS", "ESTADO"])
    col_notas   = buscar_col(df, ["NOTAS", "NOTA", "OBSERVACIONES"])
    col_exp     = buscar_col(df, ["EXPIRATION_IFE", "EXPIRATION IFE", "VENCIMIENTO_IFE"])
    col_aviso   = buscar_col(df, ["AVISO_FIRMADO", "AVISO FIRMADO", "AVISO_PRIVACIDAD"])
    col_huella  = buscar_col(df, ["HUELLA", "HUELLA DIGITAL", "HUELLA_DIGITAL"])
    col_nivel   = buscar_col(df, ["NIVEL_LEALTAD", "NIVEL LEALTAD", "PLAYER_LEVEL_NAME"])
    col_nombre  = buscar_col(df, ["NOMBRE_JUGADOR", "NOMBRE", "NAME"])
    col_created = buscar_col(df, ["DATE_CREATED_NEW_RECORD", "DATE CREATED NEW RECORD",
                                  "FECHA_CREACION", "FECHA CREACION"])
    col_updated = buscar_col(df, ["DATE_CREATED_UPDATE_RECORD", "DATE CREATED UPDATE RECORD",
                                  "FECHA_ACTUALIZACION"])

    if not col_id:     raise ValueError(f"Sin PLAYER_ID en {path.name}")
    if not col_status: raise ValueError(f"Sin STATUS en {path.name}")

    df["_SUCURSAL"] = sucursal
    df["_ID"]       = df[col_id].astype(str).str.strip()
    df["_STATUS"]   = df[col_status].apply(norm_status)
    df["_NOTAS"]    = df[col_notas].apply(limpiar) if col_notas \
                      else pd.Series("", index=df.index)
    df["_NOTA_CAT"] = df["_NOTAS"].apply(clasificar_nota)
    df["_NIVEL"]    = df[col_nivel].apply(limpiar).replace("", "SIN NIVEL") if col_nivel \
                      else pd.Series("SIN NIVEL", index=df.index)
    df["_NOMBRE"]   = df[col_nombre].apply(limpiar) if col_nombre \
                      else pd.Series("", index=df.index)
    df["_EXP"]      = pd.to_datetime(df[col_exp], errors="coerce") if col_exp else pd.NaT
    df["_ID_VENC"]  = df["_EXP"].isna() | (df["_EXP"] <= FECHA_ID_VENCIDA)
    df["_AVISO"]    = df[col_aviso].apply(es_si) if col_aviso \
                      else pd.Series(False, index=df.index)
    df["_HUELLA"]   = df[col_huella].apply(es_si) if col_huella \
                      else pd.Series(False, index=df.index)
    df["_CREATED"]  = pd.to_datetime(df[col_created], errors="coerce") if col_created \
                      else pd.NaT
    df["_UPDATED"]  = pd.to_datetime(df[col_updated], errors="coerce") if col_updated \
                      else pd.NaT

    return df[df["_ID"].str.len() > 0].drop_duplicates(subset=["_ID"], keep="last")

def cnt(df, st):
    return int((df["_STATUS"] == st).sum())

# ══════════════════════════════════════════════════════════════
# BLOQUE DE ESTATUS REUTILIZABLE (usado en H4)
# ══════════════════════════════════════════════════════════════

def bloque_estatus(ws, fila_ini, lbl_total, df_ant, df_act):
    enc(ws, ["ESTATUS / GRUPO", FECHA_ANTERIOR_CORTA,
             FECHA_ACTUAL_CORTA, "VARIACIÓN"], fila_ini)
    ws.row_dimensions[fila_ini].height = 24

    ops_a = sum(cnt(df_ant, st) for st in OPS_STATUS)
    ops_c = sum(cnt(df_act, st) for st in OPS_STATUS)
    kyc_a = sum(cnt(df_ant, st) for st in KYC_STATUS)
    kyc_c = sum(cnt(df_act, st) for st in KYC_STATUS)

    filas = [
        (lbl_total, len(df_ant), len(df_act), C["naranja"], True, "FFFFFF"),
        None,
        ("REGISTRO EXPRESS",
         cnt(df_ant,"Registro express"), cnt(df_act,"Registro express"),
         C["azul_medio"], True, "FFFFFF"),
        None,
        ("GRUPO: OPERACIONES", ops_a, ops_c, C["azul_medio"], True, "FFFFFF"),
        ("      Bloqueado Operaciones",
         cnt(df_ant,"Bloqueado por operaciones"),
         cnt(df_act,"Bloqueado por operaciones"),
         C["azul_claro"], False, "000000"),
        ("      Info Act Operaciones",
         cnt(df_ant,"Información actualizada operaciones"),
         cnt(df_act,"Información actualizada operaciones"),
         C["azul_claro"], False, "000000"),
        ("      Registro Completo",
         cnt(df_ant,"Registro completo"), cnt(df_act,"Registro completo"),
         C["azul_claro"], False, "000000"),
        ("      Verificado Operaciones",
         cnt(df_ant,"Verificado por operaciones"),
         cnt(df_act,"Verificado por operaciones"),
         C["azul_claro"], False, "000000"),
        None,
        ("GRUPO: KYC-PLD", kyc_a, kyc_c, C["verde_oscuro"], True, "FFFFFF"),
        ("      Bloqueado KYC",
         cnt(df_ant,"Bloqueado por KYC"), cnt(df_act,"Bloqueado por KYC"),
         C["verde_claro"], False, "000000"),
        ("      Info Act KYC",
         cnt(df_ant,"Información actualizada Kyc"),
         cnt(df_act,"Información actualizada Kyc"),
         C["verde_claro"], False, "000000"),
        ("      Pendiente Validación KYC",
         cnt(df_ant,"Pendente Validacion Kyc"),
         cnt(df_act,"Pendente Validacion Kyc"),
         C["verde_claro"], False, "000000"),
        ("      Verificado KYC",
         cnt(df_ant,"Verificado por Kyc"), cnt(df_act,"Verificado por Kyc"),
         C["verde_claro"], False, "000000"),
        None,
        ("TOTAL (OPERACIONES + KYC-PLD)",
         ops_a+kyc_a, ops_c+kyc_c, C["naranja"], True, "FFFFFF"),
    ]

    f = fila_ini + 1
    for item in filas:
        if item is None:
            ws.row_dimensions[f].height = 8; f += 1; continue
        lbl, va, vc, bg, bold, col = item
        vr = var_str(vc - va)
        for j, v in enumerate([lbl, va, vc, vr], 1):
            c = ws.cell(f, j, v)
            if j == 4 and not bold:
                var_cell(c, vr)
            else:
                s(c, bg=bg, bold=bold, color=col, size=10,
                  h="left" if j == 1 else "center")
            if j == 4 and bold:
                s(c, bg=bg, bold=True, color=col, size=10, h="center")
        ws.row_dimensions[f].height = 20
        f += 1
    return f

# ══════════════════════════════════════════════════════════════
# H1 — CAMBIO DE ESTATUS
# ══════════════════════════════════════════════════════════════

def hacer_h1(wb, todos_cambios, todos_nuevos):
    ws = wb.create_sheet("H1 - Cambio de Estatus")
    ncols = 5
    titulo(ws, f"RESUMEN DE CAMBIO DE ESTATUS  |  "
               f"{FECHA_ANTERIOR_CORTA} vs {FECHA_ACTUAL_CORTA}", 1, ncols)

    total_c = len(todos_cambios)
    ws.merge_cells("A2:E2")
    c = ws.cell(2, 1, f"Total cambios: {total_c:,}   |   "
                      f"Clientes nuevos: {len(todos_nuevos):,}")
    s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=11)
    ws.row_dimensions[2].height = 20

    # Transiciones globales
    seccion(ws, "▶  TRANSICIONES GLOBALES DE ESTATUS", 4, ncols)
    enc(ws, ["#", f"STATUS ANTERIOR ({FECHA_ANTERIOR_CORTA})",
             f"STATUS NUEVO ({FECHA_ACTUAL_CORTA})",
             "CANTIDAD", "% DEL TOTAL"], 5)

    trans = (todos_cambios
             .groupby(["_STATUS_ANT", "_STATUS_ACT"]).size()
             .reset_index(name="N")
             .sort_values("N", ascending=False)
             .reset_index(drop=True))

    fila = 6
    for idx, row in trans.iterrows():
        bg = COLOR_STATUS.get(row["_STATUS_ANT"], C["blanco"])
        pct_v = f"{row['N']/total_c*100:.1f}%" if total_c else "0.0%"
        for j, v in enumerate([idx+1, row["_STATUS_ANT"], row["_STATUS_ACT"],
                                int(row["N"]), pct_v], 1):
            c = ws.cell(fila, j, v)
            s(c, bg=bg, size=10, h="center" if j in [1, 4, 5] else "left")
        ws.row_dimensions[fila].height = 18
        fila += 1

    ws.cell(fila, 2, "TOTAL CAMBIOS")
    ws.cell(fila, 4, total_c)
    ws.cell(fila, 5, "100.0%")
    for j in range(1, 6):
        s(ws.cell(fila, j), bg=C["azul_marino"], bold=True,
          color="FFFFFF", size=10, h="center" if j != 2 else "left")
    ws.row_dimensions[fila].height = 20
    fila += 2

    # Cambios por sucursal
    seccion(ws, "▶  CAMBIOS POR SUCURSAL", fila, ncols)
    fila += 1
    enc(ws, ["SUCURSAL", f"STATUS ANTERIOR ({FECHA_ANTERIOR_CORTA})",
             f"STATUS NUEVO ({FECHA_ACTUAL_CORTA})", "CANTIDAD", ""], fila)
    fila += 1

    for suc in SUCURSALES:
        df_s = todos_cambios[todos_cambios["_SUCURSAL"] == suc]
        if df_s.empty: continue

        ws.merge_cells(start_row=fila, start_column=1,
                       end_row=fila,   end_column=3)
        c = ws.cell(fila, 1, f"  {suc.upper()}  — {len(df_s):,} cambios")
        ws.cell(fila, 4, len(df_s))
        s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"], size=10)
        s(ws.cell(fila, 4), bg=C["azul_claro"], bold=True,
          color=C["azul_marino"], size=10, h="center")
        ws.row_dimensions[fila].height = 20
        fila += 1

        trans_suc = (df_s.groupby(["_STATUS_ANT", "_STATUS_ACT"]).size()
                     .reset_index(name="N").sort_values("N", ascending=False))
        for _, row in trans_suc.iterrows():
            bg = COLOR_STATUS.get(row["_STATUS_ANT"], C["blanco"])
            for j, v in enumerate(["", row["_STATUS_ANT"], row["_STATUS_ACT"],
                                    int(row["N"]), ""], 1):
                c = ws.cell(fila, j, v)
                s(c, bg=bg, size=10, h="center" if j == 4 else "left")
            ws.row_dimensions[fila].height = 18
            fila += 1
        fila += 1

    # ── Clientes nuevos (sección al final de H1, igual que el reporte de referencia)
    fila += 1

    # Encabezado con total entre paréntesis
    ws.merge_cells(start_row=fila, start_column=1,
                   end_row=fila,   end_column=ncols)
    c = ws.cell(fila, 1,
                f"▶  CLIENTES NUEVOS EN KYCTOTAL_TODAS_SUCURSALES "
                f"{FECHA_ACTUAL_CORTA.upper()}  ({len(todos_nuevos):,} clientes)")
    s(c, bg=C["azul_marino"], bold=True, color="FFFFFF", size=11)
    ws.row_dimensions[fila].height = 24
    fila += 1

    # Encabezados de columna
    enc(ws, ["SUCURSAL", "STATUS", "CANTIDAD", "", ""], fila)
    ws.row_dimensions[fila].height = 22
    fila += 1

    # Colores alternos por sucursal (igual que en el reporte de referencia)
    COLORES_SUCURSAL = [
        ("2E75B6", "FFFFFF"),  # azul medio
        ("E26B0A", "FFFFFF"),  # naranja
        ("70AD47", "FFFFFF"),  # verde
        ("FFC000", "000000"),  # amarillo
        ("4472C4", "FFFFFF"),  # azul
        ("FF0000", "FFFFFF"),  # rojo
        ("7030A0", "FFFFFF"),  # morado
        ("00B0F0", "000000"),  # azul claro
        ("92D050", "000000"),  # verde claro
    ]
    # Color claro para las sub-filas de cada sucursal
    COLORES_CLARO = [
        "BDD7EE",  # azul claro
        "FCE4D6",  # naranja claro
        "E2EFDA",  # verde claro
        "FFEB9C",  # amarillo claro
        "DDEEFF",  # azul claro 2
        "FFD7D7",  # rojo claro
        "E4D5F0",  # morado claro
        "CCECFF",  # azul muy claro
        "D7F0BF",  # verde muy claro
    ]

    suc_idx = 0
    for suc in SUCURSALES:
        df_s = todos_nuevos[todos_nuevos["_SUCURSAL"] == suc]
        if df_s.empty: continue

        color_bg, color_fg = COLORES_SUCURSAL[suc_idx % len(COLORES_SUCURSAL)]
        color_sub = COLORES_CLARO[suc_idx % len(COLORES_CLARO)]

        # Fila de sucursal con total
        ws.merge_cells(start_row=fila, start_column=1,
                       end_row=fila,   end_column=2)
        c = ws.cell(fila, 1, suc.upper())
        ws.cell(fila, 3, len(df_s))
        s(c, bg=color_bg, bold=True, color=color_fg, size=10)
        s(ws.cell(fila, 3), bg=color_bg, bold=True, color=color_fg,
          size=10, h="center")
        for j in [4, 5]:
            ws.cell(fila, j).fill = fill(color_bg)
        ws.row_dimensions[fila].height = 20
        fila += 1

        # Desglose por status (con color claro de la sucursal)
        nagg = (df_s.groupby("_STATUS").size()
                .reset_index(name="N")
                .sort_values("N", ascending=False))
        for _, row in nagg.iterrows():
            ws.cell(fila, 1, "")
            ws.cell(fila, 2, row["_STATUS"])
            ws.cell(fila, 3, int(row["N"]))
            for j in range(1, 6):
                ws.cell(fila, j).fill = fill(color_sub)
            ws.cell(fila, 2).font = font(size=10)
            ws.cell(fila, 2).alignment = aln(h="left")
            ws.cell(fila, 3).font = font(size=10)
            ws.cell(fila, 3).alignment = aln(h="center")
            ws.row_dimensions[fila].height = 18
            fila += 1

        suc_idx += 1

    # Fila TOTAL NUEVOS
    ws.merge_cells(start_row=fila, start_column=1,
                   end_row=fila,   end_column=2)
    ws.cell(fila, 1, "TOTAL NUEVOS")
    ws.cell(fila, 3, len(todos_nuevos))
    for j in range(1, 6):
        s(ws.cell(fila, j), bg=C["azul_marino"], bold=True,
          color="FFFFFF", size=11, h="center" if j == 3 else "left")
    ws.row_dimensions[fila].height = 22

    ws.column_dimensions["A"].width = 40; ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 38; ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 12
    ws.freeze_panes = "A6"

# ══════════════════════════════════════════════════════════════
# H2 — NOTAS BLOQUEADOS
# ══════════════════════════════════════════════════════════════

def hacer_h2(wb, df_act_all):
    ws = wb.create_sheet("H2 - Notas Bloqueados")
    titulo(ws, f"RESUMEN DE NOTAS POR GRUPO  |  {FECHA_ACTUAL_LARGA}  |  "
               f"Todas las sucursales", 1, 6)
    totales_g = []
    fila = 3

    for tg, cbg, sts, etqs, lbl_tot, ctot in [
        ("GRUPO: OPERACIONES", "azul_medio", OPS_STATUS,
         ["Bloq Ops", "Info Act Ops", "Reg Completo", "Verif Ops", "TOTAL OPS"],
         "TOTAL GRUPO OPERACIONES", C["azul_medio"]),
        ("GRUPO: KYC-PLD", "verde_oscuro", KYC_STATUS,
         ["Bloq KYC", "Info Act KYC", "Pend Valid KYC", "Verif KYC", "TOTAL KYC-PLD"],
         "TOTAL GRUPO KYC-PLD", C["verde_oscuro"]),
    ]:
        seccion(ws, tg, fila, 6, bg=cbg); fila += 1
        subs    = [cnt(df_act_all, st) for st in sts]
        tot_g   = sum(subs); totales_g.append(tot_g)
        bg_enc  = C["azul_claro"] if cbg == "azul_medio" else C["verde_claro"]
        col_enc = C["azul_marino"] if cbg == "azul_medio" else C["verde_oscuro"]

        ws.cell(fila, 1, "")
        for j, (et, sub) in enumerate(zip(etqs[:4], subs), 2):
            c = ws.cell(fila, j, f"{et}\n{sub:,}")
            s(c, bg=bg_enc, bold=True, color=col_enc, size=10, h="center", wrap=True)
        c = ws.cell(fila, 6, f"{etqs[4]}\n{tot_g:,}")
        s(c, bg=C["naranja"], bold=True, color="FFFFFF", size=10, h="center", wrap=True)
        ws.row_dimensions[fila].height = 28; fila += 1
        ws.row_dimensions[fila].height = 12; fila += 1

        for j, et in enumerate(["CATEGORÍA DE NOTA"] + etqs[:4] + [etqs[4]], 1):
            c = ws.cell(fila, j, et)
            s(c, bg=C["azul_marino"], bold=True, color="FFFFFF",
              size=10, h="center" if j > 1 else "left")
        ws.row_dimensions[fila].height = 24; fila += 1

        for idx, cat in enumerate(CATEGORIAS_NOTAS):
            conteos  = [int(((df_act_all["_STATUS"] == st) &
                             (df_act_all["_NOTA_CAT"] == cat)).sum()) for st in sts]
            tot_cat  = sum(conteos)
            bg       = bg_enc if idx % 2 == 0 else C["blanco"]
            ws.cell(fila, 1, cat)
            s(ws.cell(fila, 1), bg=bg, size=10)
            for j, v in enumerate(conteos, 2):
                c = ws.cell(fila, j, v if v else "-")
                s(c, bg=bg, size=10, h="center")
            c = ws.cell(fila, 6, tot_cat if tot_cat else "-")
            s(c, bg=bg, bold=True, size=10, h="center")
            ws.row_dimensions[fila].height = 18; fila += 1

        ws.cell(fila, 1, lbl_tot)
        for j, v in enumerate(subs, 2): ws.cell(fila, j, v)
        ws.cell(fila, 6, tot_g)
        for j in range(1, 7):
            s(ws.cell(fila, j), bg=ctot, bold=True, color="FFFFFF",
              size=10, h="center" if j > 1 else "left")
        ws.row_dimensions[fila].height = 22; fila += 2

    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=5)
    ws.cell(fila, 1, "TOTAL GLOBAL (OPERACIONES + KYC-PLD)")
    ws.cell(fila, 6, sum(totales_g))
    s(ws.cell(fila, 1), bg=C["naranja"], bold=True, color="FFFFFF", size=11)
    s(ws.cell(fila, 6), bg=C["naranja"], bold=True, color="FFFFFF",
      size=11, h="center")
    ws.row_dimensions[fila].height = 22
    ws.column_dimensions["A"].width = 38
    for l in ["B", "C", "D", "E", "F"]: ws.column_dimensions[l].width = 18
    ws.freeze_panes = "A6"

# ══════════════════════════════════════════════════════════════
# H3 — DATOS POR SUCURSAL (con sub-filas por nivel de lealtad)
# ══════════════════════════════════════════════════════════════

def hacer_h3(wb, df_act_all, datos_por_sucursal):
    ws = wb.create_sheet("H3 - Datos por Sucursal")
    titulo(ws, f"RESUMEN DE DATOS IMPORTANTES POR SUCURSAL  |  {FECHA_ACTUAL_LARGA}",
           1, 10)

    ws.merge_cells("A2:J2")
    c = ws.cell(2, 1, "ID Vencida: fecha ≤ 31/12/2025")
    s(c, bg=C["azul_medio"], color="FFFFFF", size=10)
    ws.row_dimensions[2].height = 16

    for ini, fin, txt in [(4,5,"ID VENCIDA (≤31/12/25)"),
                          (6,7,"AVISO DE PRIVACIDAD"),
                          (8,9,"HUELLA DIGITAL")]:
        ws.merge_cells(start_row=3, start_column=ini, end_row=3, end_column=fin)
        c = ws.cell(3, ini, txt)
        s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=10, h="center")
    ws.row_dimensions[3].height = 18

    for j, h in enumerate(["SUCURSAL", "NIVEL LEALTAD", "TOTAL", "CANTIDAD", "%",
                            "CON AVISO (SÍ)", "SIN AVISO (NO)",
                            "CON HUELLA", "SIN HUELLA", "% HUELLA"], 1):
        c = ws.cell(4, j, h)
        s(c, bg=C["azul_marino"], bold=True, color="FFFFFF", size=10,
          h="center" if j > 2 else "left", wrap=True)
    ws.row_dimensions[4].height = 24

    fila = 5
    for suc, _, _ in datos_por_sucursal:
        df_suc = df_act_all[df_act_all["_SUCURSAL"] == suc]
        if df_suc.empty: continue

        t = len(df_suc)
        id_v = df_suc["_ID_VENC"].sum()
        av   = df_suc["_AVISO"].sum()
        hu   = df_suc["_HUELLA"].sum()

        # Fila resumen de sucursal
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)
        c = ws.cell(fila, 1, suc.upper())
        s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"], size=10)
        for j, v in enumerate([t, int(id_v), pct(id_v, t),
                                int(av), int(t-av),
                                int(hu), int(t-hu), pct(hu, t)], 3):
            c = ws.cell(fila, j, v)
            s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"],
              size=10, h="center")
        ws.row_dimensions[fila].height = 20; fila += 1

        # Sub-filas por nivel
        datos_nv = (df_suc.groupby("_NIVEL")
                    .agg(TOTAL=("_ID","count"), ID_VENC=("_ID_VENC","sum"),
                         CON_AVISO=("_AVISO","sum"), CON_HUELLA=("_HUELLA","sum"))
                    .reset_index()
                    .sort_values("TOTAL", ascending=False)
                    .reset_index(drop=True))

        for idx2 in range(len(datos_nv)):
            row   = datos_nv.iloc[idx2]
            nivel = str(row["_NIVEL"])
            bg    = COLOR_NIVEL.get(nivel, C["blanco"])
            ws.cell(fila, 1, "")
            ws.cell(fila, 2, nivel)
            s(ws.cell(fila, 2), bg=bg, size=10)
            s(ws.cell(fila, 1), bg=bg)
            for j, v in enumerate([int(row.TOTAL), int(row.ID_VENC),
                                    pct(row.ID_VENC, row.TOTAL),
                                    int(row.CON_AVISO),
                                    int(row.TOTAL - row.CON_AVISO),
                                    int(row.CON_HUELLA),
                                    int(row.TOTAL - row.CON_HUELLA),
                                    pct(row.CON_HUELLA, row.TOTAL)], 3):
                c = ws.cell(fila, j, v)
                s(c, bg=bg, size=10, h="center")
            ws.row_dimensions[fila].height = 18; fila += 1

    # Total general
    t_all  = len(df_act_all)
    id_all = df_act_all["_ID_VENC"].sum()
    av_all = df_act_all["_AVISO"].sum()
    hu_all = df_act_all["_HUELLA"].sum()
    ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)
    ws.cell(fila, 1, "TOTAL GENERAL")
    s(ws.cell(fila, 1), bg=C["naranja"], bold=True, color="FFFFFF", size=10)
    for j, v in enumerate([t_all, int(id_all), pct(id_all, t_all),
                            int(av_all), int(t_all-av_all),
                            int(hu_all), int(t_all-hu_all),
                            pct(hu_all, t_all)], 3):
        c = ws.cell(fila, j, v)
        s(c, bg=C["naranja"], bold=True, color="FFFFFF", size=10, h="center")
    ws.row_dimensions[fila].height = 22

    ws.column_dimensions["A"].width = 16; ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 12
    for l in ["D","E","F","G","H","I","J"]: ws.column_dimensions[l].width = 14
    ws.freeze_panes = "A5"

# ══════════════════════════════════════════════════════════════
# H4 — RESUMEN ESTATUS (global + por sucursal)
# ══════════════════════════════════════════════════════════════

def hacer_h4(wb, df_ant_all, df_act_all, datos_por_sucursal):
    ws = wb.create_sheet("H4 - Resumen Estatus")
    titulo(ws, f"RESUMEN DE ESTATUS  |  {FECHA_ANTERIOR_LARGA} vs {FECHA_ACTUAL_LARGA}",
           1, 4)

    seccion(ws, "RESUMEN GLOBAL", 3, 4)
    ws.row_dimensions[3].height = 22
    fila = bloque_estatus(ws, 4, "TOTAL BASE DE DATOS", df_ant_all, df_act_all)

    fila += 1
    seccion(ws, "DESGLOSE POR SUCURSAL", fila, 4); fila += 1

    for suc, df_ant_suc, df_act_suc in datos_por_sucursal:
        ws.merge_cells(start_row=fila, start_column=1,
                       end_row=fila,   end_column=4)
        c = ws.cell(fila, 1, f"  {suc.upper()}")
        s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"], size=11)
        ws.row_dimensions[fila].height = 22; fila += 1
        fila = bloque_estatus(ws, fila, f"Total {suc}", df_ant_suc, df_act_suc)
        fila += 1

    ws.column_dimensions["A"].width = 44; ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18; ws.column_dimensions["D"].width = 16
    ws.freeze_panes = "A5"

# ══════════════════════════════════════════════════════════════
# H5 — RESUMEN GLOBAL ESTADÍSTICO
# ══════════════════════════════════════════════════════════════

def hacer_h5(wb, df_ant_all, df_act_all, todos_cambios, todos_nuevos):
    ws = wb.create_sheet("H5 - Resumen Global")
    ta, tc = len(df_ant_all), len(df_act_all)

    titulo(ws, f"RESUMEN GLOBAL ESTADÍSTICO  |  Análisis PLD  |  "
               f"{FECHA_ANTERIOR_CORTA} vs {FECHA_ACTUAL_CORTA}", 1, 6, size=15)
    ws.row_dimensions[1].height = 40
    ws.merge_cells("A2:F2")
    c = ws.cell(2, 1, f"Análisis PLD – {len(SUCURSALES)} sucursal(es)  |  "
                      f"{FECHA_ANTERIOR_LARGA} vs {FECHA_ACTUAL_LARGA}")
    s(c, bg=C["azul_medio"], color="FFFFFF", size=11)
    ws.row_dimensions[2].height = 18

    enc(ws, ["INDICADOR", FECHA_ANTERIOR_CORTA, "%",
             FECHA_ACTUAL_CORTA, "%", "VARIACIÓN"], 4)

    ops_a = sum(cnt(df_ant_all, st) for st in OPS_STATUS)
    ops_c = sum(cnt(df_act_all, st) for st in OPS_STATUS)
    kyc_a = sum(cnt(df_ant_all, st) for st in KYC_STATUS)
    kyc_c = sum(cnt(df_act_all, st) for st in KYC_STATUS)

    def fh(fila, lbl, va, vc, bg=None, bold=False, sec=False):
        if sec:
            ws.merge_cells(start_row=fila, start_column=1,
                           end_row=fila, end_column=6)
            c = ws.cell(fila, 1, lbl)
            s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=11)
            ws.row_dimensions[fila].height = 22; return
        if not lbl:
            ws.row_dimensions[fila].height = 8; return
        bg_use = bg or (C["azul_claro"] if fila % 2 == 0 else C["blanco"])
        vr = var_str(vc - va)
        for j, v in enumerate([lbl, va, pct(va, ta), vc, pct(vc, tc), vr], 1):
            c = ws.cell(fila, j, v)
            if j == 6:
                var_cell(c, vr)
                c.font = Font(bold=bold, size=10, name="Arial")
            else:
                s(c, bg=bg_use, bold=bold, size=10,
                  h="left" if j == 1 else "center")
        ws.row_dimensions[fila].height = 22

    fh(5,  "── UNIVERSO TOTAL ────────────────────────", 0, 0, sec=True)
    fh(6,  "Total Clientes en Base", ta, tc)
    fh(7,  "Clientes Nuevos (entre cortes)", 0, len(todos_nuevos))
    fh(8,  "", 0, 0)
    fh(9,  "── ESTATUS ───────────────────────────────", 0, 0, sec=True)
    fh(10, "Registro Express",
       cnt(df_ant_all,"Registro express"), cnt(df_act_all,"Registro express"))
    fh(11, "Grupo Operaciones – SUBTOTAL", ops_a, ops_c, C["azul_claro"], True)
    fh(12, "  Bloqueado por Operaciones",
       cnt(df_ant_all,"Bloqueado por operaciones"),
       cnt(df_act_all,"Bloqueado por operaciones"))
    fh(13, "  Información Actualizada Operaciones",
       cnt(df_ant_all,"Información actualizada operaciones"),
       cnt(df_act_all,"Información actualizada operaciones"))
    fh(14, "  Registro Completo",
       cnt(df_ant_all,"Registro completo"), cnt(df_act_all,"Registro completo"))
    fh(15, "  Verificado por Operaciones",
       cnt(df_ant_all,"Verificado por operaciones"),
       cnt(df_act_all,"Verificado por operaciones"))
    fh(16, "Grupo KYC-PLD – SUBTOTAL", kyc_a, kyc_c, C["verde_claro"], True)
    fh(17, "  Bloqueado por KYC",
       cnt(df_ant_all,"Bloqueado por KYC"), cnt(df_act_all,"Bloqueado por KYC"))
    fh(18, "  Información Actualizada KYC",
       cnt(df_ant_all,"Información actualizada Kyc"),
       cnt(df_act_all,"Información actualizada Kyc"))
    fh(19, "  Pendiente Validación KYC",
       cnt(df_ant_all,"Pendente Validacion Kyc"),
       cnt(df_act_all,"Pendente Validacion Kyc"))
    fh(20, "  Verificado por KYC",
       cnt(df_ant_all,"Verificado por Kyc"), cnt(df_act_all,"Verificado por Kyc"))
    fh(21, "", 0, 0)
    fh(22, "── DOCUMENTACIÓN ─────────────────────────", 0, 0, sec=True)
    fh(23, "ID Vencida (expiración ≤ 31/12/2025)",
       int(df_ant_all["_ID_VENC"].sum()), int(df_act_all["_ID_VENC"].sum()))
    fh(24, "Con Aviso de Privacidad Firmado (SÍ)",
       int(df_ant_all["_AVISO"].sum()), int(df_act_all["_AVISO"].sum()))
    fh(25, "Sin Aviso de Privacidad Firmado (NO)",
       ta-int(df_ant_all["_AVISO"].sum()), tc-int(df_act_all["_AVISO"].sum()))
    fh(26, "Con Huella Digital (SÍ)",
       int(df_ant_all["_HUELLA"].sum()), int(df_act_all["_HUELLA"].sum()))
    fh(27, "Sin Huella Digital (NO)",
       ta-int(df_ant_all["_HUELLA"].sum()), tc-int(df_act_all["_HUELLA"].sum()))
    fh(28, "", 0, 0)
    fh(29, "── MOVIMIENTOS ENTRE CORTES ──────────────", 0, 0, sec=True)
    fh(30, "Total Cambios de Estatus", 0, len(todos_cambios))

    trans_all = (todos_cambios
                 .groupby(["_STATUS_ANT", "_STATUS_ACT"]).size()
                 .reset_index(name="N").sort_values("N", ascending=False))
    for i, (_, row) in enumerate(trans_all.iterrows(), 31):
        fh(i, f"  {row['_STATUS_ANT'][:24]}  →  {row['_STATUS_ACT'][:24]}",
           0, int(row["N"]))

    ultima = 31 + len(trans_all) + 1
    ws.merge_cells(start_row=ultima, start_column=1, end_row=ultima, end_column=6)
    c = ws.cell(ultima, 1,
                "⚠  Nota: Las variaciones negativas en estatus de bloqueo "
                "indican mejora en la regularización de clientes.")
    s(c, bg=C["gris_claro"], size=10, color="595959")
    ws.row_dimensions[ultima].height = 20

    ws.column_dimensions["A"].width = 40; ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 12; ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 12; ws.column_dimensions["F"].width = 16
    ws.freeze_panes = "A5"

# ══════════════════════════════════════════════════════════════
# H6 — DETALLE REGISTRO EXPRESS
# ══════════════════════════════════════════════════════════════

def hacer_h6(wb, todos_cambios, df_act_all):
    ws  = wb.create_sheet("H6 - Detalle Reg Express")
    ncols = 6
    titulo(ws, f"DETALLE: CLIENTES EN REGISTRO EXPRESS QUE CAMBIARON DE ESTATUS  |  "
               f"{FECHA_ACTUAL_LARGA}", 1, ncols)

    reg_exp  = todos_cambios[todos_cambios["_STATUS_ANT"] == "Registro express"].copy()
    total_re = len(reg_exp)

    ws.merge_cells("A2:F2")
    c = ws.cell(2, 1, f"Total clientes: {total_re:,}   |   Corte: {FECHA_ACTUAL_LARGA}")
    s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=11)
    ws.row_dimensions[2].height = 20

    dist     = reg_exp.groupby("_STATUS_ACT").size().sort_values(ascending=False)
    dist_txt = ("   Distribución:  " +
                "  |  ".join([f"{st}: {n:,}" for st, n in dist.items()]))
    ws.merge_cells("A3:F3")
    c = ws.cell(3, 1, dist_txt)
    s(c, bg=C["gris_claro"], size=10, color="595959")
    ws.row_dimensions[3].height = 16

    enc(ws, ["SUCURSAL", "PLAYER ID", "NOMBRE DEL CLIENTE",
             f"STATUS ANTERIOR\n({FECHA_ANTERIOR_CORTA})",
             f"STATUS NUEVO\n({FECHA_ACTUAL_CORTA})",
             "FECHA DE CAMBIO"], 5)
    ws.row_dimensions[5].height = 32
    fila = 6

    for suc in SUCURSALES:
        df_s = reg_exp[reg_exp["_SUCURSAL"] == suc]
        if df_s.empty: continue

        ws.merge_cells(start_row=fila, start_column=1,
                       end_row=fila,   end_column=ncols)
        c = ws.cell(fila, 1, f"  {suc.upper()}  —  {len(df_s):,} clientes")
        s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"], size=10)
        ws.row_dimensions[fila].height = 20; fila += 1

        for nst in df_s["_STATUS_ACT"].unique():
            df_st = df_s[df_s["_STATUS_ACT"] == nst]
            ws.merge_cells(start_row=fila, start_column=1,
                           end_row=fila,   end_column=ncols)
            c = ws.cell(fila, 1,
                        f"    → Nuevo estatus: {nst}  ({len(df_st):,})")
            s(c, bg=COLOR_STATUS.get(nst, C["gris_claro"]), size=10)
            ws.row_dimensions[fila].height = 18; fila += 1

            df_act_suc = df_act_all[df_act_all["_SUCURSAL"] == suc][
                ["_ID", "_NOMBRE", "_UPDATED"]]
            df_m = df_st.merge(df_act_suc, on="_ID", how="left",
                               suffixes=("", "_act"))

            for _, row in df_m.iterrows():
                fecha_str = (row["_UPDATED"].strftime("%d/%m/%Y %H:%M")
                             if pd.notna(row.get("_UPDATED")) else "")
                nombre = row.get("_NOMBRE", row.get("_NOMBRE_act", ""))
                for j, v in enumerate([suc, row["_ID"], nombre,
                                        row["_STATUS_ANT"], row["_STATUS_ACT"],
                                        fecha_str], 1):
                    c = ws.cell(fila, j, v); s(c, size=10)
                ws.row_dimensions[fila].height = 16; fila += 1
        fila += 1

    ws.column_dimensions["A"].width = 14; ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 32; ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 28; ws.column_dimensions["F"].width = 20
    ws.freeze_panes = "A6"

# ══════════════════════════════════════════════════════════════
# H7 — DETALLE CLIENTES NUEVOS
# ══════════════════════════════════════════════════════════════

def hacer_h7(wb, todos_nuevos):
    ws    = wb.create_sheet("H7 - Detalle Clientes Nuevos")
    ncols = 4
    titulo(ws, f"DETALLE DE CLIENTES NUEVOS  |  {FECHA_ACTUAL_LARGA}", 1, ncols)

    ws.merge_cells("A2:D2")
    c = ws.cell(2, 1, f"Clientes en {FECHA_ACTUAL_CORTA} que no estaban en "
                      f"{FECHA_ANTERIOR_CORTA}  —  Total: {len(todos_nuevos):,}")
    s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=11)
    ws.row_dimensions[2].height = 20

    dist_suc = todos_nuevos.groupby("_SUCURSAL").size()
    ws.merge_cells("A3:D3")
    c = ws.cell(3, 1, "  Por sucursal:  " +
                "  |  ".join([f"{k}: {v:,}" for k, v in dist_suc.items()]))
    s(c, bg=C["gris_claro"], size=10, color="595959")
    ws.row_dimensions[3].height = 16

    enc(ws, ["SUCURSAL", "PLAYER ID", "NOMBRE DEL CLIENTE",
             f"STATUS ACTUAL\n({FECHA_ACTUAL_CORTA})"], 5)
    ws.row_dimensions[5].height = 28
    fila = 6

    for suc in SUCURSALES:
        df_s = todos_nuevos[todos_nuevos["_SUCURSAL"] == suc]
        if df_s.empty: continue

        ws.merge_cells(start_row=fila, start_column=1,
                       end_row=fila,   end_column=ncols)
        c = ws.cell(fila, 1, f"  {suc.upper()}  —  {len(df_s):,} clientes nuevos")
        s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"], size=10)
        ws.row_dimensions[fila].height = 20; fila += 1

        for _, row in df_s.sort_values("_STATUS").iterrows():
            bg = COLOR_STATUS.get(row["_STATUS"], C["blanco"])
            for j, v in enumerate([suc, row["_ID"], row["_NOMBRE"],
                                    row["_STATUS"]], 1):
                c = ws.cell(fila, j, v); s(c, bg=bg, size=10)
            ws.row_dimensions[fila].height = 16; fila += 1
        fila += 1

    ws.column_dimensions["A"].width = 14; ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 34; ws.column_dimensions["D"].width = 32
    ws.freeze_panes = "A6"

# ══════════════════════════════════════════════════════════════
# H8 — RESUMEN REGISTROS POR SEMANA
# ══════════════════════════════════════════════════════════════

def hacer_h8(wb, df_ant_all, df_act_all):
    ws    = wb.create_sheet("H8 - Resumen Registros Semana")
    ncols = 11
    titulo(ws, f"RESUMEN DE REGISTROS POR SEMANA  |  "
               f"{FECHA_ANTERIOR_LARGA} vs {FECHA_ACTUAL_LARGA}", 1, ncols)

    sem_ant_lbl = (f"{SEMANA_ANT_INICIO.strftime('%d/%m/%Y')} – "
                   f"{SEMANA_ANT_FIN.strftime('%d/%m/%Y')}")
    sem_nva_lbl = (f"{SEMANA_NVA_INICIO.strftime('%d/%m/%Y')} – "
                   f"{SEMANA_NVA_FIN.strftime('%d/%m/%Y')}")

    COLS_ST = ["Registro express", "Registro completo",
               "Información actualizada operaciones", "Verificado por operaciones",
               "Bloqueado por operaciones", "Pendente Validacion Kyc",
               "Información actualizada Kyc", "Verificado por Kyc", "Bloqueado por KYC"]
    ETQS    = ["Reg express", "Reg completo", "Info Act Ops", "Verif Ops",
               "Bloq Ops", "Pend KYC", "Info KYC", "Verif KYC", "Bloq KYC"]
    headers = ["Sucursal", "Total"] + ETQS

    def filtrar(df, ini, fin):
        m = (df["_CREATED"].notna() &
             (df["_CREATED"].dt.normalize() >= ini) &
             (df["_CREATED"].dt.normalize() <= fin))
        return df[m]

    def tabla(ws, fila, df, ttl, lbl):
        seccion(ws, f"▶  {ttl}  –  {lbl}", fila, ncols); fila += 1
        enc(ws, headers, fila); fila += 1
        tots = {st: 0 for st in COLS_ST}; tg = 0
        for suc in SUCURSALES:
            df_s    = df[df["_SUCURSAL"] == suc]
            t       = len(df_s)
            conteos = [cnt(df_s, st) for st in COLS_ST]
            tg += t
            for st, n in zip(COLS_ST, conteos): tots[st] += n
            for j, v in enumerate([suc, t] +
                                   [v if v else "-" for v in conteos], 1):
                c = ws.cell(fila, j, v)
                s(c, size=10, h="left" if j == 1 else "center")
            ws.row_dimensions[fila].height = 18; fila += 1
        ws.cell(fila, 1, "Total")
        for j, v in enumerate([tg] + [tots[st] for st in COLS_ST], 2):
            ws.cell(fila, j, v)
        for j in range(1, ncols+1):
            s(ws.cell(fila, j), bg=C["azul_marino"], bold=True,
              color="FFFFFF", size=10, h="center" if j > 1 else "left")
        ws.row_dimensions[fila].height = 20; fila += 1
        return fila, tots, tg

    df_sa = filtrar(df_ant_all, SEMANA_ANT_INICIO, SEMANA_ANT_FIN)
    df_sn = filtrar(df_act_all, SEMANA_NVA_INICIO, SEMANA_NVA_FIN)

    fila = 3
    fila, ta, ga = tabla(ws, fila, df_sa, "SEMANA ANTERIOR", sem_ant_lbl)
    fila += 1
    fila, tn, gn = tabla(ws, fila, df_sn, "SEMANA NUEVA",    sem_nva_lbl)
    fila += 1

    # Variación
    seccion(ws, "▶  VARIACIÓN  (Semana Nueva – Semana Anterior)", fila, ncols)
    fila += 1
    enc(ws, headers, fila); fila += 1

    for suc in SUCURSALES:
        dsa = filtrar(df_ant_all[df_ant_all["_SUCURSAL"]==suc],
                      SEMANA_ANT_INICIO, SEMANA_ANT_FIN)
        dsn = filtrar(df_act_all[df_act_all["_SUCURSAL"]==suc],
                      SEMANA_NVA_INICIO, SEMANA_NVA_FIN)
        tv  = len(dsn) - len(dsa)
        cv  = [cnt(dsn, st) - cnt(dsa, st) for st in COLS_ST]
        for j, v in enumerate([suc, var_str(tv)] +
                               [var_str(v) if v != 0 else "-" for v in cv], 1):
            c = ws.cell(fila, j, v)
            if j > 1: var_cell(c, v)
            else:     s(c, size=10)
        ws.row_dimensions[fila].height = 18; fila += 1

    ws.cell(fila, 1, "Total")
    ws.cell(fila, 2, var_str(gn - ga))
    var_cell(ws.cell(fila, 2), var_str(gn - ga))
    for i, st in enumerate(COLS_ST):
        v = tn.get(st, 0) - ta.get(st, 0)
        c = ws.cell(fila, i+3, var_str(v) if v != 0 else "-")
        var_cell(c, var_str(v) if v != 0 else "-")
    s(ws.cell(fila, 1), bg=C["azul_marino"], bold=True,
      color="FFFFFF", size=10)
    ws.row_dimensions[fila].height = 20

    ws.column_dimensions["A"].width = 14; ws.column_dimensions["B"].width = 10
    for l in ["C","D","E","F","G","H","I","J","K"]:
        ws.column_dimensions[l].width = 14
    ws.freeze_panes = "A4"

# ══════════════════════════════════════════════════════════════
# H9 — DETALLE REGISTROS POR SEMANA
# ══════════════════════════════════════════════════════════════

def hacer_h9(wb, df_ant_all, df_act_all):
    ws    = wb.create_sheet("H9 - Detalle Registros Semana")
    ncols = 6
    sem_ant_lbl = (f"{SEMANA_ANT_INICIO.strftime('%d/%m/%Y')} – "
                   f"{SEMANA_ANT_FIN.strftime('%d/%m/%Y')}")
    sem_nva_lbl = (f"{SEMANA_NVA_INICIO.strftime('%d/%m/%Y')} – "
                   f"{SEMANA_NVA_FIN.strftime('%d/%m/%Y')}")

    titulo(ws, f"DETALLE DE REGISTROS POR SEMANA  |  "
               f"{FECHA_ANTERIOR_LARGA} vs {FECHA_ACTUAL_LARGA}", 1, ncols)
    ws.merge_cells("A2:F2")
    c = ws.cell(2, 1, f"Semana anterior: {sem_ant_lbl}   |   "
                      f"Semana nueva: {sem_nva_lbl}")
    s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=11)
    ws.row_dimensions[2].height = 20

    enc(ws, ["SEMANA", "SUCURSAL", "PLAYER ID", "NOMBRE DEL CLIENTE",
             "STATUS ACTUAL",
             f"FECHA CREACIÓN\n(DATE_CREATED_NEW_RECORD)"], 4)
    ws.row_dimensions[4].height = 32
    fila = 5

    def filtrar(df, ini, fin):
        m = (df["_CREATED"].notna() &
             (df["_CREATED"].dt.normalize() >= ini) &
             (df["_CREATED"].dt.normalize() <= fin))
        return df[m]

    for lbl_sem, df_base, ini, fin in [
        (f"ANTERIOR\n{sem_ant_lbl}", df_ant_all, SEMANA_ANT_INICIO, SEMANA_ANT_FIN),
        (f"NUEVA\n{sem_nva_lbl}",    df_act_all, SEMANA_NVA_INICIO, SEMANA_NVA_FIN),
    ]:
        df_sem  = filtrar(df_base, ini, fin)
        cabecera = ("SEMANA ANTERIOR" if "ANTERIOR" in lbl_sem
                    else "SEMANA NUEVA")
        ws.merge_cells(start_row=fila, start_column=1,
                       end_row=fila,   end_column=ncols)
        c = ws.cell(fila, 1,
                    f"▶  {cabecera}  "
                    f"({ini.strftime('%d/%m/%Y')} – {fin.strftime('%d/%m/%Y')})"
                    f"  —  {len(df_sem):,} registros")
        s(c, bg=C["azul_marino"], bold=True, color="FFFFFF", size=11)
        ws.row_dimensions[fila].height = 22; fila += 1

        for suc in SUCURSALES:
            df_suc = df_sem[df_sem["_SUCURSAL"] == suc]
            if df_suc.empty: continue
            ws.merge_cells(start_row=fila, start_column=1,
                           end_row=fila,   end_column=ncols)
            c = ws.cell(fila, 1,
                        f"    {suc.upper()}  —  {len(df_suc):,} registros")
            s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"], size=10)
            ws.row_dimensions[fila].height = 18; fila += 1

            for _, row in df_suc.sort_values("_CREATED").iterrows():
                fecha_str = (row["_CREATED"].strftime("%d/%m/%Y %H:%M")
                             if pd.notna(row["_CREATED"]) else "")
                bg = COLOR_STATUS.get(row["_STATUS"], C["blanco"])
                for j, v in enumerate([lbl_sem, suc, row["_ID"], row["_NOMBRE"],
                                        row["_STATUS"], fecha_str], 1):
                    c = ws.cell(fila, j, v)
                    s(c, bg=bg, size=10, wrap=(j == 1))
                ws.row_dimensions[fila].height = 16; fila += 1
        fila += 1

    ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 38; ws.column_dimensions["D"].width = 32
    ws.column_dimensions["E"].width = 28; ws.column_dimensions["F"].width = 20
    ws.freeze_panes = "A5"

# ══════════════════════════════════════════════════════════════
# H10 — RESUMEN YTD 2026
# ══════════════════════════════════════════════════════════════

def hacer_h10(wb, df_ant_all, df_act_all):
    ws    = wb.create_sheet("H10 - Resumen YTD 2026")
    ncols = 11
    titulo(ws, f"RESUMEN YTD 2026  |  "
               f"Clientes creados desde {YTD_INICIO.strftime('%d/%m/%Y')}",
           1, ncols)
    ws.merge_cells("A2:K2")
    c = ws.cell(2, 1, f"Período: {YTD_INICIO.strftime('%d/%m/%Y')} – "
                      f"{FECHA_ACTUAL_LARGA}")
    s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=11)
    ws.row_dimensions[2].height = 20

    COLS_ST = ["Registro express", "Registro completo",
               "Información actualizada operaciones", "Verificado por operaciones",
               "Bloqueado por operaciones", "Pendente Validacion Kyc",
               "Información actualizada Kyc", "Verificado por Kyc", "Bloqueado por KYC"]
    ETQS    = ["Reg express", "Reg completo", "Info Act Ops", "Verif Ops",
               "Bloq Ops", "Pend KYC", "Info KYC", "Verif KYC", "Bloq KYC"]
    headers = ["Sucursal", "Total"] + ETQS

    def filtrar_ytd(df):
        return df[df["_CREATED"].notna() &
                  (df["_CREATED"].dt.normalize() >= YTD_INICIO)]

    def tabla_ytd(ws, fila, df, ttl):
        seccion(ws, f"▶  {ttl}", fila, ncols); fila += 1
        enc(ws, headers, fila); fila += 1
        tots = {st: 0 for st in COLS_ST}; tg = 0
        for suc in SUCURSALES:
            df_s    = df[df["_SUCURSAL"] == suc]
            t       = len(df_s)
            conteos = [cnt(df_s, st) for st in COLS_ST]
            tg += t
            for st, n in zip(COLS_ST, conteos): tots[st] += n
            for j, v in enumerate([suc, t] +
                                   [v if v else "-" for v in conteos], 1):
                c = ws.cell(fila, j, v)
                s(c, size=10, h="left" if j == 1 else "center")
            ws.row_dimensions[fila].height = 18; fila += 1
        ws.cell(fila, 1, "Total")
        for j, v in enumerate([tg] + [tots[st] for st in COLS_ST], 2):
            ws.cell(fila, j, v)
        for j in range(1, ncols+1):
            s(ws.cell(fila, j), bg=C["azul_marino"], bold=True,
              color="FFFFFF", size=10, h="center" if j > 1 else "left")
        ws.row_dimensions[fila].height = 20; fila += 1
        return fila, tots, tg

    fila = 4
    fila, ta_y, ga_y = tabla_ytd(ws, fila, filtrar_ytd(df_ant_all),
                                  f"BASE ANTERIOR  –  {FECHA_ANTERIOR_CORTA}")
    fila += 1
    fila, tn_y, gn_y = tabla_ytd(ws, fila, filtrar_ytd(df_act_all),
                                  f"BASE NUEVA  –  {FECHA_ACTUAL_CORTA}")
    fila += 1

    seccion(ws, "▶  VARIACIÓN  (Base Nueva – Base Anterior)", fila, ncols)
    fila += 1; enc(ws, headers, fila); fila += 1

    for suc in SUCURSALES:
        da = filtrar_ytd(df_ant_all[df_ant_all["_SUCURSAL"] == suc])
        dn = filtrar_ytd(df_act_all[df_act_all["_SUCURSAL"] == suc])
        tv = len(dn) - len(da)
        cv = [cnt(dn, st) - cnt(da, st) for st in COLS_ST]
        for j, v in enumerate([suc, var_str(tv)] +
                               [var_str(v) if v != 0 else "-" for v in cv], 1):
            c = ws.cell(fila, j, v)
            if j > 1: var_cell(c, v)
            else:     s(c, size=10)
        ws.row_dimensions[fila].height = 18; fila += 1

    ws.cell(fila, 1, "Total")
    ws.cell(fila, 2, var_str(gn_y - ga_y))
    var_cell(ws.cell(fila, 2), var_str(gn_y - ga_y))
    for i, st in enumerate(COLS_ST):
        v = tn_y.get(st, 0) - ta_y.get(st, 0)
        c = ws.cell(fila, i+3, var_str(v) if v != 0 else "-")
        var_cell(c, var_str(v) if v != 0 else "-")
    s(ws.cell(fila, 1), bg=C["azul_marino"], bold=True,
      color="FFFFFF", size=10)
    ws.row_dimensions[fila].height = 20

    ws.column_dimensions["A"].width = 14; ws.column_dimensions["B"].width = 10
    for l in ["C","D","E","F","G","H","I","J","K"]:
        ws.column_dimensions[l].width = 14
    ws.freeze_panes = "A5"

# ══════════════════════════════════════════════════════════════
# H11 — DETALLE YTD 2026
# ══════════════════════════════════════════════════════════════

def hacer_h11(wb, df_act_all):
    ws    = wb.create_sheet("H11 - Detalle YTD 2026")
    ncols = 5
    titulo(ws, f"DETALLE YTD 2026  |  "
               f"Clientes creados desde {YTD_INICIO.strftime('%d/%m/%Y')}  |  "
               f"{FECHA_ACTUAL_LARGA}", 1, ncols)

    df_ytd = df_act_all[
        df_act_all["_CREATED"].notna() &
        (df_act_all["_CREATED"].dt.normalize() >= YTD_INICIO)
    ].copy()

    ws.merge_cells("A2:E2")
    c = ws.cell(2, 1, f"Base: {FECHA_ACTUAL_CORTA}  —  "
                      f"Total: {len(df_ytd):,} clientes")
    s(c, bg=C["azul_medio"], bold=True, color="FFFFFF", size=11)
    ws.row_dimensions[2].height = 20

    dist2 = df_ytd.groupby("_SUCURSAL").size()
    ws.merge_cells("A3:E3")
    c = ws.cell(3, 1, "  Por sucursal:  " +
                "  |  ".join([f"{k}: {v:,}" for k, v in dist2.items()]))
    s(c, bg=C["gris_claro"], size=10, color="595959")
    ws.row_dimensions[3].height = 16

    enc(ws, ["SUCURSAL", "PLAYER ID", "NOMBRE DEL CLIENTE",
             f"STATUS ACTUAL\n({FECHA_ACTUAL_CORTA})",
             f"FECHA CREACIÓN\n(DATE_CREATED_NEW_RECORD)"], 5)
    ws.row_dimensions[5].height = 32
    fila = 6

    for suc in SUCURSALES:
        df_suc = df_ytd[df_ytd["_SUCURSAL"] == suc]
        if df_suc.empty: continue

        ws.merge_cells(start_row=fila, start_column=1,
                       end_row=fila,   end_column=ncols)
        c = ws.cell(fila, 1, f"  {suc.upper()}  —  {len(df_suc):,} clientes")
        s(c, bg=C["azul_claro"], bold=True, color=C["azul_marino"], size=10)
        ws.row_dimensions[fila].height = 20; fila += 1

        for st in TODOS_STATUS:
            df_st = df_suc[df_suc["_STATUS"] == st]
            if df_st.empty: continue
            ws.merge_cells(start_row=fila, start_column=1,
                           end_row=fila,   end_column=ncols)
            c = ws.cell(fila, 1, f"    → {st}  ({len(df_st):,})")
            s(c, bg=COLOR_STATUS.get(st, C["gris_claro"]), size=10)
            ws.row_dimensions[fila].height = 18; fila += 1

            for _, row in df_st.sort_values("_CREATED").iterrows():
                fecha_str = (row["_CREATED"].strftime("%d/%m/%Y %H:%M")
                             if pd.notna(row["_CREATED"]) else "")
                bg = COLOR_STATUS.get(row["_STATUS"], C["blanco"])
                for j, v in enumerate([suc, row["_ID"], row["_NOMBRE"],
                                        row["_STATUS"], fecha_str], 1):
                    c = ws.cell(fila, j, v); s(c, bg=bg, size=10)
                ws.row_dimensions[fila].height = 16; fila += 1
        fila += 1

    ws.column_dimensions["A"].width = 14; ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 34; ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 22
    ws.freeze_panes = "A6"

# ══════════════════════════════════════════════════════════════
# PROCESO PRINCIPAL
# ══════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════
# GENERAR datos.json (para actualización automática de la página)
# ══════════════════════════════════════════════════════════════

def generar_json(df_ant_all, df_act_all, datos_por_sucursal,
                 todos_cambios, todos_nuevos):
    ta = len(df_ant_all); tc = len(df_act_all)
    def c(df, st): return int((df["_STATUS"] == st).sum())

    sucursales = []
    for suc, df_ant, df_act in datos_por_sucursal:
        ta_s = len(df_ant); tc_s = len(df_act)
        nuevos_suc = len(todos_nuevos[todos_nuevos["_SUCURSAL"] == suc]) if "_SUCURSAL" in todos_nuevos.columns else 0
        cambios_suc = len(todos_cambios[todos_cambios["_SUCURSAL"] == suc]) if "_SUCURSAL" in todos_cambios.columns else 0
        sucursales.append({
            "nombre": suc, "total_ant": ta_s, "total_act": tc_s,
            "delta": tc_s - ta_s, "nuevos": int(nuevos_suc), "cambios": int(cambios_suc),
            "re_act":    c(df_act,"Registro express"),
            "bloq_ops":  c(df_act,"Bloqueado por operaciones"),
            "bloq_kyc":  c(df_act,"Bloqueado por KYC"),
            "verif_kyc": c(df_act,"Verificado por Kyc"),
            "pend_kyc":  c(df_act,"Pendente Validacion Kyc"),
            "id_venc_act": int(df_act["_ID_VENC"].sum()),
            "aviso_act":   int(df_act["_AVISO"].sum()),
            "huella_act":  int(df_act["_HUELLA"].sum()),
            "pct_id_venc": round(df_act["_ID_VENC"].sum() / tc_s * 100, 1) if tc_s else 0,
            "pct_aviso":   round(df_act["_AVISO"].sum()   / tc_s * 100, 1) if tc_s else 0,
            "pct_huella":  round(df_act["_HUELLA"].sum()  / tc_s * 100, 1) if tc_s else 0,
            "pct_verif_kyc": round(c(df_act,"Verificado por Kyc") / tc_s * 100, 1) if tc_s else 0,
        })

    top_trans = []
    if len(todos_cambios) > 0 and "_STATUS_ANT" in todos_cambios.columns:
        top_trans = (todos_cambios
                     .groupby(["_STATUS_ANT","_STATUS_ACT"]).size()
                     .reset_index(name="N")
                     .sort_values("N", ascending=False)
                     .head(10)
                     .to_dict("records"))

    datos = {
        "meta": {
            "generado":             datetime.now().strftime("%d/%m/%Y %H:%M"),
            "fecha_anterior_corta": FECHA_ANTERIOR_CORTA,
            "fecha_actual_corta":   FECHA_ACTUAL_CORTA,
            "fecha_anterior_larga": FECHA_ANTERIOR_LARGA,
            "fecha_actual_larga":   FECHA_ACTUAL_LARGA,
            "num_sucursales":       len(datos_por_sucursal),
        },
        "kpis": {
            "total_ant": ta, "total_act": tc, "delta_total": tc - ta,
            "clientes_nuevos": len(todos_nuevos), "cambios_status": len(todos_cambios),
            "ops_ant": sum(c(df_ant_all,s) for s in OPS_STATUS),
            "ops_act": sum(c(df_act_all,s) for s in OPS_STATUS),
            "kyc_ant": sum(c(df_ant_all,s) for s in KYC_STATUS),
            "kyc_act": sum(c(df_act_all,s) for s in KYC_STATUS),
            "re_ant":  c(df_ant_all,"Registro express"),
            "re_act":  c(df_act_all,"Registro express"),
            "bloq_ops_ant": c(df_ant_all,"Bloqueado por operaciones"),
            "bloq_ops_act": c(df_act_all,"Bloqueado por operaciones"),
            "bloq_kyc_ant": c(df_ant_all,"Bloqueado por KYC"),
            "bloq_kyc_act": c(df_act_all,"Bloqueado por KYC"),
            "verif_kyc_ant": c(df_ant_all,"Verificado por Kyc"),
            "verif_kyc_act": c(df_act_all,"Verificado por Kyc"),
            "pend_kyc_ant": c(df_ant_all,"Pendente Validacion Kyc"),
            "pend_kyc_act": c(df_act_all,"Pendente Validacion Kyc"),
            "id_venc_ant": int(df_ant_all["_ID_VENC"].sum()),
            "id_venc_act": int(df_act_all["_ID_VENC"].sum()),
            "aviso_ant": int(df_ant_all["_AVISO"].sum()),
            "aviso_act": int(df_act_all["_AVISO"].sum()),
            "huella_ant": int(df_ant_all["_HUELLA"].sum()),
            "huella_act": int(df_act_all["_HUELLA"].sum()),
        },
        "sucursales": sucursales,
        "top_transiciones": top_trans,
    }

    ruta = CARPETA_SALIDA / "datos.json"
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
    print(f"  ✅ datos.json → {ruta}")

def main():
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*55}")
    print(f"  REPORTE PLD — TODAS LAS SUCURSALES")
    print(f"  {FECHA_ANTERIOR_LARGA}  vs  {FECHA_ACTUAL_LARGA}")
    print(f"{'='*55}\n")

    # 1. Leer archivos
    print("Leyendo archivos...")
    dfs_ant, dfs_act, datos_por_sucursal, errores = [], [], [], []

    for suc in SUCURSALES:
        ruta_ant = buscar_archivo(CARPETA_ANTERIOR, suc)
        ruta_act = buscar_archivo(CARPETA_ACTUAL,   suc)

        if ruta_ant is None:
            msg = f"No encontrado en carpeta anterior"
            print(f"  ⚠  {suc:10}  {msg}")
            errores.append((suc, msg)); continue
        if ruta_act is None:
            msg = f"No encontrado en carpeta actual"
            print(f"  ⚠  {suc:10}  {msg}")
            errores.append((suc, msg)); continue

        try:
            df_ant = leer_df(ruta_ant, suc)
            df_act = leer_df(ruta_act, suc)
            dfs_ant.append(df_ant)
            dfs_act.append(df_act)
            datos_por_sucursal.append((suc, df_ant, df_act))
            delta = len(df_act) - len(df_ant)
            print(f"  ✅ {suc:10}  ant={len(df_ant):>7,}  "
                  f"act={len(df_act):>7,}  Δ={delta:+,}")
        except Exception as e:
            print(f"  ❌ {suc:10}  Error: {e}")
            errores.append((suc, str(e)))

    if not dfs_act:
        print("\n❌ No se pudo leer ningún archivo. Revisa las rutas.")
        return

    # 2. Consolidar
    df_ant_all = pd.concat(dfs_ant, ignore_index=True)
    df_act_all = pd.concat(dfs_act, ignore_index=True)

    comun = df_ant_all[["_SUCURSAL","_ID","_STATUS","_NOTAS","_NOMBRE"]].merge(
        df_act_all[["_SUCURSAL","_ID","_STATUS","_NOTAS","_NOMBRE"]],
        on=["_SUCURSAL","_ID"], suffixes=("_ANT","_ACT"))
    todos_cambios = comun[comun["_STATUS_ANT"] != comun["_STATUS_ACT"]].copy()
    todos_nuevos  = df_act_all[
        ~df_act_all.set_index(["_SUCURSAL","_ID"]).index.isin(
         df_ant_all.set_index(["_SUCURSAL","_ID"]).index)
    ].copy()

    print(f"\n  Total anterior : {len(df_ant_all):,}")
    print(f"  Total actual   : {len(df_act_all):,}")
    print(f"  Cambios status : {len(todos_cambios):,}")
    print(f"  Clientes nuevos: {len(todos_nuevos):,}")

    # 3. Generar hojas
    print("\nGenerando reporte...")
    wb = Workbook(); wb.remove(wb.active)

    pasos = [
        ("H1  - Cambio de Estatus",      lambda: hacer_h1(wb, todos_cambios, todos_nuevos)),
        ("H2  - Notas Bloqueados",        lambda: hacer_h2(wb, df_act_all)),
        ("H3  - Datos por Sucursal",      lambda: hacer_h3(wb, df_act_all, datos_por_sucursal)),
        ("H4  - Resumen Estatus",         lambda: hacer_h4(wb, df_ant_all, df_act_all, datos_por_sucursal)),
        ("H5  - Resumen Global",          lambda: hacer_h5(wb, df_ant_all, df_act_all, todos_cambios, todos_nuevos)),
        ("H6  - Detalle Reg Express",     lambda: hacer_h6(wb, todos_cambios, df_act_all)),
        ("H7  - Detalle Clientes Nuevos", lambda: hacer_h7(wb, todos_nuevos)),
        ("H8  - Resumen Registros Semana",lambda: hacer_h8(wb, df_ant_all, df_act_all)),
        ("H9  - Detalle Registros Semana",lambda: hacer_h9(wb, df_ant_all, df_act_all)),
        ("H10 - Resumen YTD 2026",        lambda: hacer_h10(wb, df_ant_all, df_act_all)),
        ("H11 - Detalle YTD 2026",        lambda: hacer_h11(wb, df_act_all)),
    ]

    for nombre, fn in pasos:
        print(f"  {nombre}...")
        fn()

    # 4. Guardar
    ts     = datetime.now().strftime("%Y%m%d_%H%M")
    nombre = f"Reporte_PLD_TODAS_SUCURSALES_{ts}.xlsx"
    ruta   = CARPETA_SALIDA / nombre
    # Generar datos.json para la página web
    print("\nGenerando datos.json para la página web...")
    generar_json(df_ant_all, df_act_all, datos_por_sucursal,
                 todos_cambios, todos_nuevos)

    wb.save(ruta)

    kb = ruta.stat().st_size // 1024
    print(f"\n✅  {nombre}")
    print(f"    {len(dfs_act)} sucursales  |  {kb:,} KB")
    print(f"    Ubicación: {ruta}")

    if errores:
        ruta_err = CARPETA_SALIDA / "ERRORES.xlsx"
        pd.DataFrame(errores, columns=["SUCURSAL","ERROR"]).to_excel(
            ruta_err, index=False)
        print(f"\n⚠  {len(errores)} error(es) — ver: {ruta_err}")

    print(f"\n{'='*55}")
    print("  PROCESO TERMINADO")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
