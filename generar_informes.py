"""
Genera un informe Excel por proveedor con los recuperos por liquido producto.

Uso:
    python generar_informes.py "Ventas OCT Consignación.xlsx"

Si no se pasa archivo, usa el unico .xlsx (no temporal) de la carpeta actual.

Logica:
    - Lee hoja 'Control Promos' y agrupa por comitente.
    - Descuento neto = -(Importe SUM - Iva SUM)   (col H menos col I, en positivo)
    - Comision proveedor = col K Maestro 11-11 / 100
    - Factor proveedor   = 1 - comision          (lo que efectivamente le pagamos)
    - Cargo P  -> recupero = Descuento neto * factor_proveedor
    - Cargo C  -> recupero = Descuento neto * (Cargo P / (Cargo DP + Cargo P)) * factor_proveedor
    - Cargo D  -> se ignora (lo paga DP).
    - Cruza con hoja 'Ventas' por SKU para armar la visibilidad de ventas.
"""

from __future__ import annotations

import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

LOCALES = {
    "D07": "Aeroparque",
    "D10": "Mendoza",
    "D14": "Tucuman",
    "D20": "Bariloche",
    "D25": "Iguazu",
    "D30": "Cordoba",
    "D35": "Salta",
    "WEB": "WEB",
}


def nombre_local(codigo: str) -> str:
    if not codigo:
        return ""
    return LOCALES.get(codigo.strip().upper(), codigo)


@dataclass
class RecuperoRow:
    fecha: datetime
    local: str
    sku: int
    producto: str
    promo: str
    cargo: str
    cargo_dp: float | None
    cargo_p: float | None
    descuento_bruto: float   # descuento con IVA (positivo)
    iva: float               # IVA del descuento (positivo)
    descuento_neto: float    # descuento_bruto - iva
    comision_pct: float      # fraccion 0-1 (0.45 = 45%)
    recupero: float          # descuento_neto * pct_cargo * (1 - comision_pct)


@dataclass
class VentaRow:
    sku: int
    descripcion: str
    local: str
    unidades: float
    importe: float


@dataclass
class ComitenteData:
    codigo: str
    razon_social: str
    recuperos: list[RecuperoRow] = field(default_factory=list)
    ventas_por_sku_local: dict[tuple[int, str], VentaRow] = field(default_factory=dict)


def sanitize_filename(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return cleaned[:120] or "proveedor"


def pct_recupero_p(cargo: str, dp: float | None, p: float | None) -> float:
    if cargo == "P":
        return 1.0
    if cargo == "C":
        dp = dp or 0.0
        p = p or 0.0
        total = dp + p
        if total <= 0:
            return 0.0
        return p / total
    return 0.0


def leer_control_promos(ws) -> dict[str, ComitenteData]:
    comitentes: dict[str, ComitenteData] = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        local, rubro, com, promo, cod_prod, producto, benef_id, importe, iva, desc, fecha, cargo, cdp, cp = row[:14]
        if not com or cargo not in ("P", "C"):
            continue
        if importe is None or importe == 0:
            continue
        razon = (desc or com or "").strip()
        clave = com.strip() if isinstance(com, str) else str(com)
        if clave not in comitentes:
            comitentes[clave] = ComitenteData(codigo=clave, razon_social=razon)
        else:
            if not comitentes[clave].razon_social and razon:
                comitentes[clave].razon_social = razon
        pct = pct_recupero_p(cargo, cdp, cp)
        descuento_bruto = -float(importe)
        iva_mag = -float(iva or 0)
        descuento_neto = descuento_bruto - iva_mag
        comitentes[clave].recuperos.append(
            RecuperoRow(
                fecha=fecha,
                local=nombre_local(str(local) if local is not None else ""),
                sku=int(cod_prod) if cod_prod is not None else 0,
                producto="",
                promo=str(promo) if promo is not None else "",
                cargo=cargo,
                cargo_dp=cdp,
                cargo_p=cp,
                descuento_bruto=descuento_bruto,
                iva=iva_mag,
                descuento_neto=descuento_neto,
                comision_pct=0.0,
                recupero=descuento_neto * pct,  # se ajusta por comision tras cruzar con Maestro
            )
        )
    return comitentes


def construir_mapas_maestro(
    ws_ventas, ws_maestro, skus_necesarios: set[int]
) -> tuple[dict[int, str], dict[int, float]]:
    """Devuelve (descripcion_por_sku, comision_por_sku). La comision viene en fraccion 0-1."""
    desc_map: dict[int, str] = {}
    com_map: dict[int, float] = {}

    # Primero recorro Ventas: suele traer el nombre comercial actual.
    for i, row in enumerate(ws_ventas.iter_rows(values_only=True)):
        if i == 0:
            continue
        sku = row[11]
        if sku is None:
            continue
        try:
            sku_int = int(sku)
        except (TypeError, ValueError):
            continue
        if sku_int not in skus_necesarios or sku_int in desc_map:
            continue
        descripcion = row[12]
        if descripcion:
            desc_map[sku_int] = str(descripcion)

    # Maestro: comision para todos; descripcion si todavia falta.
    desc_pendientes = skus_necesarios - set(desc_map.keys())
    com_pendientes = set(skus_necesarios)

    for i, row in enumerate(ws_maestro.iter_rows(values_only=True)):
        if i == 0:
            continue
        cod = row[0]
        if cod is None:
            continue
        try:
            cod_int = int(cod)
        except (TypeError, ValueError):
            continue
        if cod_int not in skus_necesarios:
            continue
        if cod_int in desc_pendientes:
            descripcion = row[6]
            if descripcion:
                desc_map[cod_int] = str(descripcion)
                desc_pendientes.discard(cod_int)
        if cod_int in com_pendientes:
            com_raw = row[10]  # columna K 'Comision'
            if com_raw is not None and com_raw != "":
                try:
                    com_map[cod_int] = float(com_raw) / 100.0
                except (TypeError, ValueError):
                    pass
            com_pendientes.discard(cod_int)
        if not desc_pendientes and not com_pendientes:
            break

    return desc_map, com_map


def aplicar_metadatos(
    comitentes: dict[str, ComitenteData],
    desc_map: dict[int, str],
    com_map: dict[int, float],
) -> None:
    for data in comitentes.values():
        for r in data.recuperos:
            r.producto = desc_map.get(r.sku, str(r.sku))
            r.comision_pct = com_map.get(r.sku, 0.0)
            pct_cargo = pct_recupero_p(r.cargo, r.cargo_dp, r.cargo_p)
            r.recupero = r.descuento_neto * pct_cargo * (1.0 - r.comision_pct)


def cargar_ventas_relevantes(ws, comitentes: dict[str, ComitenteData]) -> None:
    sku_a_claves: dict[int, set[str]] = defaultdict(set)
    for clave, data in comitentes.items():
        for r in data.recuperos:
            sku_a_claves[r.sku].add(clave)

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        fecha, ticket, caja, local, rubro, subrubro, familia, especie, com, prov, marca, sku, descripcion, precio, costo, imp_int, unidades, importe, comision, modo = row[:20]
        if sku is None:
            continue
        try:
            sku_int = int(sku)
        except (TypeError, ValueError):
            continue
        claves = sku_a_claves.get(sku_int)
        if not claves:
            continue
        local_s = nombre_local(str(local) if local is not None else "")
        unidades = float(unidades or 0)
        importe = float(importe or 0)
        for clave in claves:
            data = comitentes[clave]
            key = (sku_int, local_s)
            if key not in data.ventas_por_sku_local:
                data.ventas_por_sku_local[key] = VentaRow(
                    sku=sku_int,
                    descripcion=str(descripcion) if descripcion is not None else "",
                    local=local_s,
                    unidades=0.0,
                    importe=0.0,
                )
            v = data.ventas_por_sku_local[key]
            v.unidades += unidades
            v.importe += importe
            if not v.descripcion and descripcion:
                v.descripcion = str(descripcion)


def detectar_periodo(comitentes: dict[str, ComitenteData]) -> str:
    fechas = [r.fecha for data in comitentes.values() for r in data.recuperos if isinstance(r.fecha, datetime)]
    if not fechas:
        return "Periodo"
    fmin, fmax = min(fechas), max(fechas)
    if fmin.year == fmax.year and fmin.month == fmax.month:
        return f"{MESES_ES[fmin.month]} {fmin.year}"
    return f"{fmin.strftime('%d/%m/%Y')} - {fmax.strftime('%d/%m/%Y')}"


# ---------- Formato Excel ----------

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SUBTITLE_FONT = Font(bold=True, size=11, color="1F4E78")
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FONT = Font(bold=True)
MONEY_FMT = '#,##0.00;[Red]-#,##0.00'
INT_FMT = '#,##0'
PCT_FMT = '0.00%'
DATE_FMT = 'dd/mm/yyyy'


def aplicar_header(ws, row_idx: int, cols: list[str]):
    for c, val in enumerate(cols, 1):
        cell = ws.cell(row=row_idx, column=c, value=val)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def autofit(ws, min_w=10, max_w=50):
    for col in ws.columns:
        col_letter = None
        max_len = 0
        for cell in col:
            if cell.column_letter and not col_letter:
                col_letter = cell.column_letter
            val = cell.value
            if val is None:
                continue
            s = val.strftime("%d/%m/%Y") if isinstance(val, datetime) else str(val)
            if len(s) > max_len:
                max_len = len(s)
        if col_letter:
            ws.column_dimensions[col_letter].width = max(min_w, min(max_w, max_len + 2))


def escribir_resumen(ws, data: ComitenteData, periodo: str):
    ws["A1"] = f"Informe de Recuperos - {data.razon_social}"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:F1")

    ws["A2"] = f"Periodo: {periodo}"
    ws["A2"].font = SUBTITLE_FONT
    ws.merge_cells("A2:F2")

    ws["A3"] = f"Codigo comitente: {data.codigo}"
    ws.merge_cells("A3:F3")

    def _sumar(cargo: str, attr: str) -> float:
        return sum(getattr(r, attr) for r in data.recuperos if r.cargo == cargo)

    cargos_info = [
        ("P", "P - Proveedor", "100% a cargo del proveedor, neto de comision"),
        ("C", "C - Compartido", "Porcion segun % Cargo P, neto de comision"),
    ]
    cargos_presentes = [c for c in cargos_info if any(r.cargo == c[0] for r in data.recuperos)]

    aplicar_header(ws, 5, [
        "Tipo de cargo", "Descuento con IVA", "IVA", "Descuento sin IVA",
        "A recuperar del proveedor", "Observacion",
    ])

    fila = 6
    for codigo, etiqueta, obs in cargos_presentes:
        ws.cell(row=fila, column=1, value=etiqueta)
        ws.cell(row=fila, column=2, value=_sumar(codigo, "descuento_bruto")).number_format = MONEY_FMT
        ws.cell(row=fila, column=3, value=_sumar(codigo, "iva")).number_format = MONEY_FMT
        ws.cell(row=fila, column=4, value=_sumar(codigo, "descuento_neto")).number_format = MONEY_FMT
        ws.cell(row=fila, column=5, value=_sumar(codigo, "recupero")).number_format = MONEY_FMT
        ws.cell(row=fila, column=6, value=obs)
        for c in range(1, 7):
            ws.cell(row=fila, column=c).border = BORDER
        fila += 1

    ws.cell(row=fila, column=1, value="TOTAL")
    ws.cell(row=fila, column=2, value=sum(r.descuento_bruto for r in data.recuperos)).number_format = MONEY_FMT
    ws.cell(row=fila, column=3, value=sum(r.iva for r in data.recuperos)).number_format = MONEY_FMT
    ws.cell(row=fila, column=4, value=sum(r.descuento_neto for r in data.recuperos)).number_format = MONEY_FMT
    ws.cell(row=fila, column=5, value=sum(r.recupero for r in data.recuperos)).number_format = MONEY_FMT
    for c in range(1, 7):
        cell = ws.cell(row=fila, column=c)
        cell.fill = TOTAL_FILL
        cell.font = TOTAL_FONT
        cell.border = BORDER

    nota_row = fila + 2
    ws.cell(row=nota_row, column=1, value=(
        "Nota: el recupero se calcula sobre el descuento sin IVA y se ajusta por la "
        "comision del producto (Maestro col K): el proveedor absorbe solo la fraccion "
        "(1 - comision). En promos compartidas (C) se aplica ademas Cargo P / (Cargo DP + Cargo P)."
    ))
    ws.cell(row=nota_row, column=1).font = Font(italic=True, color="595959")
    ws.merge_cells(start_row=nota_row, start_column=1, end_row=nota_row, end_column=6)

    for col, w in zip("ABCDEF", (18, 20, 15, 20, 24, 55)):
        ws.column_dimensions[col].width = w


def escribir_detalle_recuperos(ws, data: ComitenteData):
    cols = [
        "Fecha", "Local", "SKU", "Producto", "Promocion", "Cargo",
        "% Cargo DP", "% Cargo P", "% Recupero Proveedor",
        "Descuento con IVA", "IVA", "Descuento sin IVA",
        "% Comision", "A recuperar",
    ]
    aplicar_header(ws, 1, cols)

    rows_sorted = sorted(
        data.recuperos,
        key=lambda r: (r.fecha if isinstance(r.fecha, datetime) else datetime.min, r.local, r.sku),
    )
    for i, r in enumerate(rows_sorted, 2):
        pct_prov = pct_recupero_p(r.cargo, r.cargo_dp, r.cargo_p)
        valores = [
            r.fecha, r.local, r.sku, r.producto, r.promo, r.cargo,
            r.cargo_dp if r.cargo_dp is not None else "",
            r.cargo_p if r.cargo_p is not None else "",
            pct_prov,
            r.descuento_bruto,
            r.iva,
            r.descuento_neto,
            r.comision_pct,
            r.recupero,
        ]
        for c, v in enumerate(valores, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = BORDER
        ws.cell(row=i, column=1).number_format = DATE_FMT
        for c in (7, 8, 9, 13):
            ws.cell(row=i, column=c).number_format = PCT_FMT
        for c in (10, 11, 12, 14):
            ws.cell(row=i, column=c).number_format = MONEY_FMT

    total_row = len(rows_sorted) + 2
    ws.cell(row=total_row, column=1, value="TOTAL")
    ws.cell(row=total_row, column=10, value=sum(r.descuento_bruto for r in rows_sorted)).number_format = MONEY_FMT
    ws.cell(row=total_row, column=11, value=sum(r.iva for r in rows_sorted)).number_format = MONEY_FMT
    ws.cell(row=total_row, column=12, value=sum(r.descuento_neto for r in rows_sorted)).number_format = MONEY_FMT
    ws.cell(row=total_row, column=14, value=sum(r.recupero for r in rows_sorted)).number_format = MONEY_FMT
    for c in range(1, len(cols) + 1):
        ws.cell(row=total_row, column=c).fill = TOTAL_FILL
        ws.cell(row=total_row, column=c).font = TOTAL_FONT
        ws.cell(row=total_row, column=c).border = BORDER

    ws.freeze_panes = "A2"
    autofit(ws)


def escribir_ventas(ws, data: ComitenteData):
    cols = ["SKU", "Producto", "Local", "Unidades", "Importe vendido"]
    aplicar_header(ws, 1, cols)

    ventas = sorted(
        data.ventas_por_sku_local.values(),
        key=lambda v: (v.descripcion.lower(), v.sku, v.local),
    )
    i = 2
    for v in ventas:
        ws.cell(row=i, column=1, value=v.sku).border = BORDER
        ws.cell(row=i, column=2, value=v.descripcion).border = BORDER
        ws.cell(row=i, column=3, value=v.local).border = BORDER
        c4 = ws.cell(row=i, column=4, value=v.unidades)
        c4.number_format = INT_FMT
        c4.border = BORDER
        c5 = ws.cell(row=i, column=5, value=v.importe)
        c5.number_format = MONEY_FMT
        c5.border = BORDER
        i += 1

    if ventas:
        ws.cell(row=i, column=1, value="TOTAL")
        ws.cell(row=i, column=4, value=sum(v.unidades for v in ventas)).number_format = INT_FMT
        ws.cell(row=i, column=5, value=sum(v.importe for v in ventas)).number_format = MONEY_FMT
        for c in range(1, 6):
            ws.cell(row=i, column=c).fill = TOTAL_FILL
            ws.cell(row=i, column=c).font = TOTAL_FONT
            ws.cell(row=i, column=c).border = BORDER
    else:
        ws.cell(row=2, column=1, value="(No se registraron ventas de los SKUs con promocion en el periodo)").font = Font(italic=True, color="7F7F7F")
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)

    ws.freeze_panes = "A2"
    autofit(ws)


def generar_excel_consolidado(comitentes: dict[str, ComitenteData], periodo: str, out_dir: Path) -> Path:
    wb = openpyxl.Workbook()

    # ---- Hoja Resumen ADM ----
    ws = wb.active
    ws.title = "Resumen ADM"

    ws["A1"] = f"Recuperos por Liquido Producto - Consolidado {periodo}"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = "Vista general para administracion. Descuentos sin IVA; recupero neto de comision del producto."
    ws["A2"].font = Font(italic=True, color="595959")
    ws.merge_cells("A2:H2")

    cols = [
        "Codigo", "Razon social",
        "Descuento P sin IVA", "A recuperar P",
        "Descuento C sin IVA", "A recuperar C",
        "Descuento Total sin IVA", "A recuperar TOTAL",
    ]
    aplicar_header(ws, 4, cols)

    filas = []
    for clave, data in comitentes.items():
        if not data.recuperos:
            continue
        desc_p = sum(r.descuento_neto for r in data.recuperos if r.cargo == "P")
        desc_c = sum(r.descuento_neto for r in data.recuperos if r.cargo == "C")
        rec_p = sum(r.recupero for r in data.recuperos if r.cargo == "P")
        rec_c = sum(r.recupero for r in data.recuperos if r.cargo == "C")
        filas.append((clave, data.razon_social, desc_p, rec_p, desc_c, rec_c, desc_p + desc_c, rec_p + rec_c))

    filas.sort(key=lambda f: f[7], reverse=True)

    r = 5
    for fila in filas:
        for c, v in enumerate(fila, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            if c >= 3:
                cell.number_format = MONEY_FMT
        r += 1

    if filas:
        total_row = r
        ws.cell(row=total_row, column=1, value="TOTAL")
        for c in range(3, 9):
            idx = c - 1
            ws.cell(row=total_row, column=c, value=sum(f[idx] for f in filas)).number_format = MONEY_FMT
        for c in range(1, 9):
            cell = ws.cell(row=total_row, column=c)
            cell.fill = TOTAL_FILL
            cell.font = TOTAL_FONT
            cell.border = BORDER

    ws.freeze_panes = "A5"
    for col, w in zip("ABCDEFGH", (10, 50, 18, 18, 18, 18, 18, 20)):
        ws.column_dimensions[col].width = w

    # ---- Hoja Detalle completo ----
    ws2 = wb.create_sheet("Detalle completo")
    cols2 = [
        "Codigo", "Razon social", "Fecha", "Local", "SKU", "Producto",
        "Promocion", "Cargo", "% Cargo DP", "% Cargo P", "% Recupero Proveedor",
        "Descuento con IVA", "IVA", "Descuento sin IVA", "% Comision", "A recuperar",
    ]
    aplicar_header(ws2, 1, cols2)

    todas = []
    for clave, data in comitentes.items():
        for rec in data.recuperos:
            todas.append((clave, data.razon_social, rec))
    todas.sort(key=lambda t: (
        t[1].lower(),
        t[2].fecha if isinstance(t[2].fecha, datetime) else datetime.min,
        t[2].local,
        t[2].sku,
    ))

    i = 2
    for clave, razon, rec in todas:
        pct_prov = pct_recupero_p(rec.cargo, rec.cargo_dp, rec.cargo_p)
        valores = [
            clave, razon, rec.fecha, rec.local, rec.sku, rec.producto, rec.promo,
            rec.cargo,
            rec.cargo_dp if rec.cargo_dp is not None else "",
            rec.cargo_p if rec.cargo_p is not None else "",
            pct_prov,
            rec.descuento_bruto, rec.iva, rec.descuento_neto,
            rec.comision_pct, rec.recupero,
        ]
        for c, v in enumerate(valores, 1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.border = BORDER
        ws2.cell(row=i, column=3).number_format = DATE_FMT
        for c in (9, 10, 11, 15):
            ws2.cell(row=i, column=c).number_format = PCT_FMT
        for c in (12, 13, 14, 16):
            ws2.cell(row=i, column=c).number_format = MONEY_FMT
        i += 1

    if todas:
        total_row = i
        ws2.cell(row=total_row, column=1, value="TOTAL")
        ws2.cell(row=total_row, column=12, value=sum(t[2].descuento_bruto for t in todas)).number_format = MONEY_FMT
        ws2.cell(row=total_row, column=13, value=sum(t[2].iva for t in todas)).number_format = MONEY_FMT
        ws2.cell(row=total_row, column=14, value=sum(t[2].descuento_neto for t in todas)).number_format = MONEY_FMT
        ws2.cell(row=total_row, column=16, value=sum(t[2].recupero for t in todas)).number_format = MONEY_FMT
        for c in range(1, len(cols2) + 1):
            cell = ws2.cell(row=total_row, column=c)
            cell.fill = TOTAL_FILL
            cell.font = TOTAL_FONT
            cell.border = BORDER

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(cols2))}{max(i, 2)}"
    autofit(ws2)

    nombre = sanitize_filename(f"Recuperos Consolidado ADM - {periodo}") + ".xlsx"
    path = out_dir / nombre
    wb.save(path)
    return path


def generar_excel_proveedor(data: ComitenteData, periodo: str, out_dir: Path) -> Path | None:
    # ADM pide que al proveedor solo se le informe el cargo C; el P se comunica por otra via.
    recuperos_c = [r for r in data.recuperos if r.cargo == "C"]
    if not recuperos_c:
        return None
    skus_c = {r.sku for r in recuperos_c}
    ventas_c = {k: v for k, v in data.ventas_por_sku_local.items() if k[0] in skus_c}
    data_filtrada = ComitenteData(
        codigo=data.codigo,
        razon_social=data.razon_social,
        recuperos=recuperos_c,
        ventas_por_sku_local=ventas_c,
    )

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Resumen"
    escribir_resumen(ws1, data_filtrada, periodo)

    ws2 = wb.create_sheet("Detalle Recuperos")
    escribir_detalle_recuperos(ws2, data_filtrada)

    ws3 = wb.create_sheet("Ventas SKU x Local")
    escribir_ventas(ws3, data_filtrada)

    nombre = sanitize_filename(f"Recupero {data_filtrada.razon_social} - {periodo}") + ".xlsx"
    path = out_dir / nombre
    wb.save(path)
    return path


def encontrar_archivo_input() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    here = Path(__file__).resolve().parent
    candidatos = [p for p in here.glob("*.xlsx") if not p.name.startswith("~$")]
    if len(candidatos) != 1:
        raise SystemExit(
            f"No se pudo elegir archivo automaticamente. Pasalo como argumento. Candidatos: {candidatos}"
        )
    return candidatos[0]


def main():
    archivo = encontrar_archivo_input()
    print(f"Leyendo: {archivo}")
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)

    if "Control Promos" not in wb.sheetnames:
        raise SystemExit("No se encontro la hoja 'Control Promos'.")
    if "Ventas" not in wb.sheetnames:
        raise SystemExit("No se encontro la hoja 'Ventas'.")

    print("Procesando 'Control Promos'...")
    comitentes = leer_control_promos(wb["Control Promos"])
    print(f"  Comitentes con cargos P/C: {len(comitentes)}")

    skus_necesarios = {r.sku for data in comitentes.values() for r in data.recuperos}
    print(f"  SKUs distintos con promo P/C: {len(skus_necesarios)}")

    print("Procesando 'Ventas' (filtrando por SKUs con promo)...")
    cargar_ventas_relevantes(wb["Ventas"], comitentes)

    print("Cruzando descripciones y comisiones (Ventas + Maestro)...")
    wb_ro = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    mapa_desc, mapa_com = construir_mapas_maestro(wb_ro["Ventas"], wb_ro["Maestro 11-11"], skus_necesarios)
    aplicar_metadatos(comitentes, mapa_desc, mapa_com)
    print(f"  Descripciones resueltas: {len(mapa_desc)}/{len(skus_necesarios)}")
    print(f"  Comisiones resueltas:    {len(mapa_com)}/{len(skus_necesarios)}")
    sin_com = [s for s in skus_necesarios if s not in mapa_com]
    if sin_com:
        print(f"  ATENCION: {len(sin_com)} SKU sin comision en Maestro (recupero sin ajuste): {sin_com[:10]}...")

    periodo = detectar_periodo(comitentes)
    print(f"Periodo detectado: {periodo}")

    out_dir = archivo.parent / f"Informes {periodo}"
    out_dir.mkdir(exist_ok=True)

    generados = 0
    omitidos_sin_c = 0
    for clave, data in sorted(comitentes.items(), key=lambda kv: kv[1].razon_social.lower()):
        if not data.recuperos:
            continue
        path = generar_excel_proveedor(data, periodo, out_dir)
        if path is None:
            omitidos_sin_c += 1
            print(f"  [{clave}] {data.razon_social}: sin cargos C, no se genera informe.")
            continue
        total = sum(r.recupero for r in data.recuperos if r.cargo == "C")
        print(f"  [{clave}] {data.razon_social}: $ {total:,.2f}  -> {path.name}")
        generados += 1

    path_cons = generar_excel_consolidado(comitentes, periodo, out_dir)
    print(f"\nConsolidado ADM: {path_cons.name}")

    print(f"Listo. {generados} informes por proveedor + 1 consolidado en: {out_dir}")


if __name__ == "__main__":
    main()
