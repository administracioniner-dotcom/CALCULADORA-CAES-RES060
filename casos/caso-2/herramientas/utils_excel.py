#!/usr/bin/env python3
"""Libro final del caso 1 (RES060 V1.1): Resumen · Cálculo · Trazabilidad · Referencias.
Fórmulas vivas; los datos de entrada están en azul sobre fondo crema, en la hoja Cálculo."""
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SALIDA = '/home/user/casos/caso-2/Caso2_RES060_FINAL.xlsx'
FONT = 'Calibri'

# --- paleta ---------------------------------------------------------------
NAVY, BLUE, LIGHT, PALE = '1F3864', '2E75B6', 'DDEBF7', 'F3F7FB'
CREAM, GREEN, GREYTXT, LINE = 'FFF8E1', 'E2EFDA', '595959', 'BFBFBF'
INPUT_BLUE = '0000FF'

thin = Side(style='thin', color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=Side(style='medium', color=NAVY))


def fill(c):
    return PatternFill('solid', fgColor=c)


def font(size=10, bold=False, color='000000', italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def put(ws, ref, value, *, size=10, bold=False, color='000000', italic=False, bg=None,
        fmt=None, h='left', v='center', wrap=True, border=True):
    c = ws[ref]
    c.value = value
    c.font = font(size, bold, color, italic)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    if bg:
        c.fill = fill(bg)
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BOX
    return c


def widths(ws, d):
    for k, w in d.items():
        ws.column_dimensions[k].width = w


def lines(text, width):
    if text is None:
        return 1
    t = str(text)
    if t.startswith('='):
        return 1
    return sum(max(1, math.ceil(len(p) / max(1.0, width * 1.12))) for p in t.split('\n'))


def autoheight(ws, row, cols, minh=18):
    n = 1
    for col in cols:
        w = ws.column_dimensions[col].width or 10
        n = max(n, lines(ws[f'{col}{row}'].value, w))
    ws.row_dimensions[row].height = max(minh, 13.2 * n + 5)


def banner(ws, title, subtitle, last_col):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(f'B1:{last_col}1')
    put(ws, 'B1', title, size=16, bold=True, color='FFFFFF', bg=NAVY, border=False)
    ws.row_dimensions[1].height = 34
    ws.merge_cells(f'B2:{last_col}2')
    put(ws, 'B2', subtitle, size=10, italic=True, color=GREYTXT, bg=LIGHT, border=False)
    ws.row_dimensions[2].height = 20
    ws.column_dimensions['A'].width = 2


def section(ws, row, text, first, last):
    ws.merge_cells(f'{first}{row}:{last}{row}')
    put(ws, f'{first}{row}', text, size=11, bold=True, color='FFFFFF', bg=BLUE, border=False)
    ws.row_dimensions[row].height = 22


def header(ws, row, labels, cols):
    for col, lab in zip(cols, labels):
        put(ws, f'{col}{row}', lab, bold=True, color=NAVY, bg=LIGHT, h='center')
    ws.row_dimensions[row].height = 24


def page(ws, landscape=True):
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = 'Caso 2 · RES060 V1.1 · Hoja &A · Página &P de &N'


