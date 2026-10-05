#!/usr/bin/env python3
"""Genera cuadro de cálculo de pensión alimenticia (2007-2026)."""

from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_MENSUAL = Decimal("1414.80")
INCREMENTO_ANUAL = Decimal("1.10")
TASA_MENSUAL = Decimal("0.0075")  # 9% anual / 12

INICIO = date(2007, 4, 1)
FIN = date(2026, 12, 1)  # mensualidad de diciembre 2026


def mensualidad_para_mes(d: date) -> Decimal:
    """Pensión mensual vigente; incremento del 10% cada enero (año calendario)."""
    if d.year == 2007:
        return BASE_MENSUAL
    exponente = d.year - 2007
    return (BASE_MENSUAL * (INCREMENTO_ANUAL ** exponente)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )


def iter_meses_pension():
    y, m = INICIO.year, INICIO.month
    while (y, m) <= (FIN.year, FIN.month):
        yield date(y, m, 1)
        m += 1
        if m > 12:
            m = 1
            y += 1


@dataclass
class Cuota:
    periodo: date
    monto: Decimal


def cuotas_por_anio() -> dict[int, list[Cuota]]:
    por_anio: dict[int, list[Cuota]] = {}
    for d in iter_meses_pension():
        por_anio.setdefault(d.year, []).append(
            Cuota(d, mensualidad_para_mes(d))
        )
    return por_anio


def meses_interes_en_anio(cuota: Cuota, anio: int) -> int:
    """Meses del año `anio` en que corre interés (desde mes siguiente al devengo)."""
    inicio_interes = date(cuota.periodo.year, cuota.periodo.month, 1)
    # primer mes con interés: mes siguiente al devengo
    if inicio_interes.month == 12:
        inicio_interes = date(inicio_interes.year + 1, 1, 1)
    else:
        inicio_interes = date(inicio_interes.year, inicio_interes.month + 1, 1)

    fin_anio = date(anio, 12, 1)
    inicio_anio = date(anio, 1, 1)
    if inicio_interes > fin_anio:
        return 0
    if FIN < inicio_anio:
        return 0

    # interés hasta fin del periodo de cálculo (diciembre 2026 inclusive como mes de devengo;
    # el interés del mes de diciembre 2026 sobre cuotas anteriores sí corre en dic 2026)
    fin_interes = FIN
    if fin_interes.month == 12:
        pass  # diciembre 2026 cuenta como mes de interés para cuotas ya devengadas

    start = max(inicio_interes, inicio_anio)
    end = min(fin_interes, fin_anio)
    if start > end:
        return 0

    return (end.year - start.year) * 12 + (end.month - start.month) + 1


def interes_anio(anio: int, cuotas_previas_y_actuales: list[Cuota]) -> Decimal:
    total = Decimal("0")
    for c in cuotas_previas_y_actuales:
        meses = meses_interes_en_anio(c, anio)
        if meses:
            total += c.monto * TASA_MENSUAL * meses
    return total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def main():
    por_anio = cuotas_por_anio()
    todas_cuotas: list[Cuota] = []
    filas = []

    for anio in range(2007, 2027):
        cuotas_anio = por_anio.get(anio, [])
        if not cuotas_anio:
            continue
        meses = len(cuotas_anio)
        pago_mensual = cuotas_anio[0].monto
        todas_cuotas.extend(cuotas_anio)
        principal = sum((c.monto for c in cuotas_anio), Decimal("0"))
        interes = interes_anio(anio, list(todas_cuotas))
        subtotal = (principal + interes).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        filas.append(
            {
                "anio": anio,
                "meses": meses,
                "pago_mensual": pago_mensual,
                "principal": principal,
                "interes": interes,
                "subtotal": subtotal,
            }
        )

    wb = Workbook()
    ws = wb.active
    ws.title = "Pensión alimenticia"

    headers = [
        "Año",
        "Cantidad de meses",
        "Pago mensual",
        "Interés moratorio a razón de 9% anual (0.75% mensual) sobre mensualidades insolutas",
        "Subtotal del año",
    ]
    ws.append(headers)

    for r in filas:
        ws.append(
            [
                r["anio"],
                r["meses"],
                float(r["pago_mensual"]),
                float(r["interes"]),
                float(r["subtotal"]),
            ]
        )

    total_meses = sum(r["meses"] for r in filas)
    total_principal = sum(r["principal"] for r in filas)
    total_interes = sum(r["interes"] for r in filas)
    total_general = sum(r["subtotal"] for r in filas)
    ws.append(["TOTAL", total_meses, "", float(total_interes), float(total_general)])

    ws.insert_rows(1, 4)
    ws["A1"] = "CUADRO DE CÁLCULO DE PENSIÓN ALIMENTICIA"
    ws["A2"] = (
        "Pensión mensual líquida inicial: $1,414.80 (vigencia desde abril de 2007). "
        "Incremento anual: 10%. Interés moratorio: 9% anual (0.75% mensual) "
        "sobre mensualidades vencidas y no pagadas hasta diciembre de 2026."
    )
    ws.merge_cells("A1:E1")
    ws.merge_cells("A2:E2")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"].alignment = Alignment(wrap_text=True)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    total_fill = PatternFill("solid", fgColor="FFF2CC")

    header_row = 5
    for col in range(1, 6):
        cell = ws.cell(row=header_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    money_fmt = '"$"#,##0.00'
    for row_idx in range(header_row + 1, ws.max_row + 1):
        for col in (3, 4, 5):
            ws.cell(row=row_idx, column=col).number_format = money_fmt
        for col in range(1, 6):
            ws.cell(row=row_idx, column=col).border = border
        if ws.cell(row=row_idx, column=1).value == "TOTAL":
            for col in range(1, 6):
                c = ws.cell(row=row_idx, column=col)
                c.font = Font(bold=True)
                c.fill = total_fill

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 42
    ws.column_dimensions["E"].width = 18

    out = "/workspace/Cuadro_Pension_Alimenticia_2007_2026.xlsx"
    wb.save(out)

    # Resumen en consola
    print(f"Archivo: {out}")
    print(f"Total principal (pensiones): ${total_principal:,.2f}")
    print(f"Total intereses moratorios: ${total_interes:,.2f}")
    print(f"Total general: ${total_general:,.2f}")


if __name__ == "__main__":
    main()
