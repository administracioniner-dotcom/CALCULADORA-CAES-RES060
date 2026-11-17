#!/usr/bin/env python3
"""Libro final del caso 1 (RES060 V1.1): Resumen · Cálculo · Trazabilidad · Referencias.
Fórmulas vivas; los datos de entrada están en azul sobre fondo crema, en la hoja Cálculo."""
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SALIDA = '/home/user/casos/caso-1/Caso1_RES060_FINAL.xlsx'
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
    ws.oddFooter.center.text = 'Caso 1 · RES060 V1.1 · Hoja &A · Página &P de &N'


wb = Workbook()
rs = wb.active
rs.title = 'Resumen'
cl = wb.create_sheet('Cálculo')
tz = wb.create_sheet('Trazabilidad')
rf = wb.create_sheet('Referencias')

# ===========================================================================
# CÁLCULO
# ===========================================================================
banner(cl, 'Cálculo del ahorro energético · Ficha RES060 V1.1',
       'Cada resultado es una fórmula enlazada a los datos de entrada. Azul sobre crema = dato introducido; negro = resultado calculado.', 'H')
widths(cl, {'B': 5, 'C': 46, 'D': 11, 'E': 16, 'F': 13, 'G': 46, 'H': 50})

section(cl, 4, '1 · DATOS DE ENTRADA', 'B', 'H')
header(cl, 5, ['Nº', 'Dato', 'Símbolo', 'Valor', 'Unidad', 'Documento de origen', 'Ubicación en el documento'],
       'BCDEFGH')

inputs = [
    # clave, dato, símbolo, valor, unidad, formato, documento, ubicación
    ('S', 'Superficie útil habitable', 'S', 151.6, 'm²', '#,##0.00',
     'D1 · Certificado de eficiencia energética previo', 'p. 2 · Superficie útil (RD 390/2021)'),
    ('DCAL', 'Demanda de energía en calefacción', 'Dcal', 100.7, 'kWh/m²·año', '#,##0.00',
     'D1 · Certificado de eficiencia energética previo', 'p. 5 y 7 · Etiqueta de demanda de calefacción (100,7)'),
    ('SCAL', 'Superficie de calefacción cubierta por la caldera', 'S_cal', 126.45, 'm²', '#,##0.00',
     'D2 · Archivo CE3X del certificado previo', 'Instalaciones › Sólo calefacción › Demanda cubierta: 126,45 m² (83,41 %)'),
    ('ETA', 'Rendimiento de la caldera retirada (referido a PCS)', 'ηi', 0.83, '—', '0.00',
     'D5 · Criterios de verificación v14.0 + D6 · Ficha RES210, Anexo II',
     'Criterio 24/11.03 · Tabla B.3: gas, desde 1998, condensación, encendido automático = 83 %'),
    ('ZONA', 'Zona climática del edificio (clima de referencia: cálido)', 'Zona', 'B4', '—', '@',
     'D1 · Certificado previo + D4 · Ficha RES060, Anexo III', 'CEE p. 1 · Anexo III: B4 → clima cálido'),
    ('TIMP', 'Temperatura de impulsión (instalación con radiadores)', 'T', 55, '°C', '0',
     'Dato de la instalación', 'Emisores: radiadores → temperatura media (55 °C) en EPREL'),
    ('ETAS', 'Eficiencia energética estacional en calefacción (55 °C, clima cálido)', 'ηs,h', 160, '%', '0',
     'D3 · Registro EPREL 345915 (Panasonic WH-MDC09J3E5)', 'Eficiencia estacional a media temperatura, clima cálido: 160 %'),
    ('FP', 'Factor de ponderación', 'FP', 1, '—', '0',
     'D4 · Ficha RES060 V1.1', 'Apartado 3, tabla de variables'),
    ('CC', 'Coeficiente de conversión a energía primaria de la electricidad', 'CC', 2.5, '—', '0.0',
     'D4 · Ficha RES060 V1.1', 'Anexo IV, nota 1'),
    ('F1', 'Factor de corrección F(1) de la bomba aerotérmica', 'F(1)', 0.03, '—', '0%',
     'D4 · Ficha RES060 V1.1', 'Anexo IV, nota 2 (F(2) = 0 en aerotermia)'),
]
R = {}
for i, (key, dato, sim, val, uni, fmt, doc, ubi) in enumerate(inputs):
    r = 6 + i
    R[key] = f'$E${r}'
    bgc = CREAM
    put(cl, f'B{r}', i + 1, h='center', color=GREYTXT)
    put(cl, f'C{r}', dato)
    put(cl, f'D{r}', sim, bold=True, h='center')
    put(cl, f'E{r}', val, color=INPUT_BLUE, bg=bgc, fmt=fmt, h='right', bold=True)
    put(cl, f'F{r}', uni, h='center')
    put(cl, f'G{r}', doc)
    put(cl, f'H{r}', ubi)
    autoheight(cl, r, 'CGH')

r_sec2 = 6 + len(inputs) + 1  # 17
section(cl, r_sec2, '2 · CÁLCULO PASO A PASO', 'B', 'H')
header(cl, r_sec2 + 1, ['Nº', 'Concepto', 'Símbolo', 'Resultado', 'Unidad', 'Fórmula', 'Sustitución numérica'], 'BCDEFGH')

S, DC, SC, ETA, ETAS, FP, CC, F1 = (R[k] for k in ('S', 'DCAL', 'SCAL', 'ETA', 'ETAS', 'FP', 'CC', 'F1'))
p0 = r_sec2 + 2  # 19
row = {n: p0 + n for n in range(9)}  # filas de los pasos 1..9 -> row[0..8]
steps = [
    ('Demanda anual de calefacción del edificio', 'D', f'={DC}*{S}', 'kWh/año', '#,##0.0',
     'Dcal × S', f'=ROUND({DC},2)&" × "&ROUND({S},2)&" = "&ROUND(E{row[0]},1)'),
    ('Parte de la calefacción que cubría la caldera', 'f', f'={SC}/{S}', '—', '0.00%',
     'S_cal ÷ S', f'=ROUND({SC},2)&" ÷ "&ROUND({S},2)&" = "&ROUND(E{row[1]}*100,2)&" %"'),
    ('Demanda de calefacción cubierta por la caldera', 'Dc', f'=E{row[0]}*E{row[1]}', 'kWh/año', '#,##0.0',
     'D × f', f'=ROUND(E{row[0]},1)&" × "&ROUND(E{row[1]},4)&" = "&ROUND(E{row[2]},1)'),
    ('SCOP de la bomba de calor según Anexo IV', 'SCOP', f'={CC}*({ETAS}/100+{F1})', '—', '0.000',
     'CC × (ηs,h + F(1))', f'={CC}&" × ("&{ETAS}/100&" + "&{F1}&") = "&ROUND(E{row[3]},4)'),
    ('SCOP adoptado (dos decimales)', 'SCOP', f'=ROUNDDOWN(E{row[3]},2)', '—', '0.00',
     'SCOP con 2 decimales', f'="Se adopta "&E{row[4]}&" (error de redondeo de "&ROUND((E{row[3]}-E{row[4]})/E{row[3]}*100,2)&" %, dentro del 1 % admitido)"'),
    ('Consumo de energía final con la caldera retirada', 'Ei', f'=E{row[2]}/{ETA}', 'kWh/año', '#,##0.0',
     'Dc ÷ ηi', f'=ROUND(E{row[2]},1)&" ÷ "&{ETA}&" = "&ROUND(E{row[5]},1)'),
    ('Consumo de energía final equivalente con la bomba de calor', 'Ef', f'=E{row[2]}/E{row[4]}', 'kWh/año', '#,##0.0',
     'Dc ÷ SCOP', f'=ROUND(E{row[2]},1)&" ÷ "&E{row[4]}&" = "&ROUND(E{row[6]},1)'),
    ('Ahorro anual de energía final', 'AE', f'={FP}*(E{row[5]}-E{row[6]})', 'kWh/año', '#,##0.0',
     'FP × (Ei − Ef)', f'={FP}&" × ("&ROUND(E{row[5]},1)&" − "&ROUND(E{row[6]},1)&") = "&ROUND(E{row[7]},2)'),
    ('Certificados de ahorro energético (1 CAE = 1 kWh)', 'CAE', f'=ROUND(E{row[7]},0)', 'CAE', '#,##0',
     'CAE = AE (redondeado)', f'="AE = "&ROUND(E{row[7]},2)&" kWh/año → "&TEXT(E{row[8]},"0")&" CAE"'),
]
for n, (concepto, sim, formula, uni, fmt, fx, sub) in enumerate(steps):
    r = row[n]
    last = n == 8
    bg_res = 'FFF2CC' if last else (GREEN if n == 7 else None)
    put(cl, f'B{r}', n + 1, h='center', color=GREYTXT)
    put(cl, f'C{r}', concepto, bold=(n >= 7))
    put(cl, f'D{r}', sim, bold=True, h='center')
    put(cl, f'E{r}', formula, fmt=fmt, h='right', bold=(n >= 7), bg=bg_res, size=(12 if last else 10))
    put(cl, f'F{r}', uni, h='center')
    put(cl, f'G{r}', fx)
    put(cl, f'H{r}', sub, color=GREYTXT)
    autoheight(cl, r, 'CGH', minh=(24 if last else (32 if n == 4 else 18)))

r_note = row[8] + 1
cl.merge_cells(f'B{r_note}:H{r_note}')
put(cl, f'B{r_note}',
    'Fórmula de la ficha: AE = FP · (Dcal · S) · (1/ηi − 1/SCOP) · f. Los pasos anteriores la desarrollan aplicada a la parte de calefacción que cubría la caldera. '
    'Solo calefacción: la caldera no producía ACS (D1, p. 3), por lo que el término de ACS de la ficha no interviene. Sin multiplicadores.',
    italic=True, color=GREYTXT, border=False, v='top')
cl.row_dimensions[r_note].height = 42

r_sec3 = r_note + 2
section(cl, r_sec3, '3 · CONTROLES', 'B', 'H')
header(cl, r_sec3 + 1, ['Nº', 'Control', '', 'Valor', 'Unidad', 'Qué comprueba', 'Resultado'], 'BCDEFGH')
c1, c2 = r_sec3 + 2, r_sec3 + 3
# 1. ahorro frente al consumo previo
put(cl, f'B{c1}', 1, h='center', color=GREYTXT)
put(cl, f'C{c1}', 'Ahorro sobre el consumo previo de la caldera')
put(cl, f'D{c1}', '')
put(cl, f'E{c1}', f'=E{row[7]}/E{row[5]}', fmt='0.0%', h='right')
put(cl, f'F{c1}', '%', h='center')
put(cl, f'G{c1}', 'El ahorro no puede superar el consumo que se evita (criterio 25-07.03)')
put(cl, f'H{c1}', f'=IF(E{row[7]}<E{row[5]},"Correcto: el ahorro es menor que el consumo previo","Revisar: el ahorro supera el consumo previo")')
# 2. sin redondeos
put(cl, f'B{c2}', 2, h='center', color=GREYTXT)
put(cl, f'C{c2}', 'Mismo cálculo sin redondeos (Dcal 100,73 · SCOP 4,075)')
put(cl, f'D{c2}', '')
put(cl, f'E{c2}', f'={FP}*100.73*{SC}*(1/{ETA}-1/E{row[3]})', fmt='#,##0.0', h='right')
put(cl, f'F{c2}', 'kWh/año', h='center')
put(cl, f'G{c2}', 'Efecto de redondear Dcal (tabla: 100,73) y el SCOP (4,075)')
put(cl, f'H{c2}', f'="Diferencia: "&ROUND(E{c2}-E{row[7]},1)&" kWh ("&ROUND((E{c2}-E{row[7]})/E{row[7]}*100,2)&" %)"')
for r in (c1, c2):
    autoheight(cl, r, 'CGH')
cl.freeze_panes = 'A4'
cl.sheet_properties.tabColor = BLUE
page(cl)

# referencias de celdas del cálculo para el resumen
CELL = {
    'AE': f"Cálculo!$E${row[7]}", 'CAE': f"Cálculo!$E${row[8]}", 'D': f"Cálculo!$E${row[0]}",
    'F': f"Cálculo!$E${row[1]}", 'DC': f"Cálculo!$E${row[2]}", 'SCOP': f"Cálculo!$E${row[4]}",
    'EI': f"Cálculo!$E${row[5]}", 'EF': f"Cálculo!$E${row[6]}",
}

# ===========================================================================
# RESUMEN
# ===========================================================================
banner(rs, 'Certificados de Ahorro Energético (CAE) · Ficha RES060 V1.1',
       'Sustitución de caldera de combustión por bomba de calor aerotérmica · Caso 1', 'G')
widths(rs, {'B': 30, 'C': 20, 'D': 20, 'E': 20, 'F': 20, 'G': 20})

section(rs, 4, 'IDENTIFICACIÓN DE LA ACTUACIÓN', 'B', 'G')
ident = [
    ('Emplazamiento', 'Vivienda caso 1 (Córdoba) · vivienda unifamiliar · zona climática B4'),
    ('Referencia catastral', '[ref. catastral omitida]'),
    ('Actuación', 'Retirada de caldera de GLP (solo calefacción) e instalación de bomba de calor aerotérmica aire-agua '
                  'Panasonic WH-MDC09J3E5-1 (nº de serie [nº serie omitido]), emisores: radiadores'),
    ('Certificado de eficiencia energética previo', 'Nº de registro [nº registro omitido] · Registro Andaluz [nº registro omitido] · registrado el 30/07/2026'),
    ('Servicio y método', 'Solo calefacción · ficha RES060 V1.1 · sin efecto multiplicador'),
]
for i, (k, v) in enumerate(ident):
    r = 5 + i
    put(rs, f'B{r}', k, bold=True, color=NAVY, bg=PALE)
    rs.merge_cells(f'C{r}:G{r}')
    put(rs, f'C{r}', v)
    for col in 'DEFG':
        rs[f'{col}{r}'].border = BOX
    rs.row_dimensions[r].height = max(20, 13.2 * lines(v, 100) + 6)

section(rs, 11, 'RESULTADO', 'B', 'G')
rs.merge_cells('B12:C13')
put(rs, 'B12', 'CAE A RECLAMAR', size=14, bold=True, color='FFFFFF', bg=NAVY, h='center')
for ref in ('C12', 'B13', 'C13'):
    rs[ref].fill = fill(NAVY)
rs.merge_cells('D12:E13')
put(rs, 'D12', f'={CELL["CAE"]}', size=26, bold=True, color=NAVY, bg='FFF2CC', fmt='#,##0', h='center')
for ref in ('E12', 'D13', 'E13'):
    rs[ref].fill = fill('FFF2CC')
rs.merge_cells('F12:G13')
put(rs, 'F12', '1 CAE = 1 kWh de ahorro anual de energía final', size=10, italic=True, color=GREYTXT, h='center')
for ref in ('G12', 'F13', 'G13'):
    rs[ref].border = BOX
rs.row_dimensions[12].height = 24
rs.row_dimensions[13].height = 24

section(rs, 15, 'CÓMO SE OBTIENE', 'B', 'G')
header(rs, 16, ['Concepto', 'Valor', 'Unidad', 'Se calcula como', '', ''], 'BCDEFG')
rs.merge_cells('E16:G16')
flow = [
    ('Demanda de calefacción cubierta por la caldera', CELL['DC'], 'kWh/año', '#,##0.0', 'Dcal × S × 83,41 % (parte cubierta por la caldera)'),
    ('Consumo con la caldera retirada (ηi = 0,83)', CELL['EI'], 'kWh/año', '#,##0.0', 'Demanda cubierta ÷ rendimiento de la caldera'),
    ('Consumo con la bomba de calor (SCOP = 4,07)', CELL['EF'], 'kWh/año', '#,##0.0', 'Demanda cubierta ÷ SCOP de la bomba de calor'),
    ('Ahorro anual de energía final', CELL['AE'], 'kWh/año', '#,##0.0', 'Consumo antes − consumo después'),
]
for i, (a, ref, uni, fmt, how) in enumerate(flow):
    r = 17 + i
    last = i == 3
    put(rs, f'B{r}', a, bold=last, bg=(GREEN if last else None))
    put(rs, f'C{r}', f'={ref}', fmt=fmt, h='right', bold=last, bg=(GREEN if last else None))
    put(rs, f'D{r}', uni, h='center', bg=(GREEN if last else None))
    rs.merge_cells(f'E{r}:G{r}')
    put(rs, f'E{r}', how, color=GREYTXT, bg=(GREEN if last else None))
    for col in 'FG':
        rs[f'{col}{r}'].border = BOX
    rs.row_dimensions[r].height = 30 if len(a) > 34 else 22

section(rs, 22, 'DATOS DE PARTIDA', 'B', 'G')
header(rs, 23, ['Dato', 'Valor', 'Unidad', 'Documento', '', ''], 'BCDEFG')
rs.merge_cells('E23:G23')
base = [
    ('Superficie útil', 'S', 'm²', '#,##0.00', 'D1 · Certificado previo'),
    ('Demanda de calefacción', 'DCAL', 'kWh/m²·año', '#,##0.00', 'D1 · Certificado previo'),
    ('Superficie cubierta por la caldera', 'SCAL', 'm²', '#,##0.00', 'D2 · Archivo CE3X del certificado previo'),
    ('Rendimiento de la caldera', 'ETA', '—', '0.00', 'D5 · Criterio 24/11.03 · D6 · Ficha RES210, Anexo II'),
    ('Eficiencia estacional de la bomba (ηs,h)', 'ETAS', '%', '0', 'D3 · Registro EPREL 345915'),
]
for i, (a, key, uni, fmt, doc) in enumerate(base):
    r = 24 + i
    put(rs, f'B{r}', a)
    put(rs, f'C{r}', f'=Cálculo!{R[key]}', fmt=fmt, h='right')
    put(rs, f'D{r}', uni, h='center')
    rs.merge_cells(f'E{r}:G{r}')
    put(rs, f'E{r}', doc)
    for col in 'FG':
        rs[f'{col}{r}'].border = BOX
    rs.row_dimensions[r].height = 20

section(rs, 30, 'CONTENIDO DEL LIBRO', 'B', 'G')
guia = [
    ('Cálculo', 'Datos de entrada con su documento de origen, cálculo paso a paso con fórmulas y controles.'),
    ('Trazabilidad', 'De qué documento y de qué página sale cada dato, y cómo se ha comprobado.'),
    ('Referencias', 'Documentos consultados (D1 a D9), normativa y criterios aplicados, datos de EPREL.'),
]
for i, (a, b) in enumerate(guia):
    r = 31 + i
    put(rs, f'B{r}', a, bold=True, color=NAVY, bg=PALE)
    rs.merge_cells(f'C{r}:G{r}')
    put(rs, f'C{r}', b)
    for col in 'DEFG':
        rs[f'{col}{r}'].border = BOX
    rs.row_dimensions[r].height = 20
rs.merge_cells('B35:G35')
put(rs, 'B35', 'Leyenda: azul sobre fondo crema = dato introducido · negro = resultado calculado con fórmula · fondo verde / amarillo = resultado.',
    italic=True, color=GREYTXT, border=False, size=9)
rs.sheet_properties.tabColor = NAVY
page(rs, landscape=False)

# ===========================================================================
# TRAZABILIDAD
# ===========================================================================
banner(tz, 'Trazabilidad de los datos',
       'Para cada dato del cálculo: valor usado, documento (código D1–D9, ver hoja Referencias), ubicación exacta y comprobación realizada.', 'G')
widths(tz, {'B': 30, 'C': 18, 'D': 9, 'E': 44, 'F': 44, 'G': 58})
section(tz, 4, 'A · DATOS DEL CÁLCULO Y SU DOCUMENTO DE ORIGEN', 'B', 'G')
header(tz, 5, ['Dato', 'Valor usado', 'Doc.', 'Ubicación en el documento', 'Comprobación', 'Observaciones'], 'BCDEFG')
tr = [
    ('Superficie útil habitable (S)', '151,6 m²', 'D1', 'p. 2 · Superficie útil (RD 390/2021)',
     'Coincide con el certificado posterior (D9, 151,6 m²).', 'Es la superficie de la vivienda completa.'),
    ('Demanda de calefacción (Dcal)', '100,7 kWh/m²·año', 'D1', 'p. 5 y 7 · Etiqueta de demanda de calefacción',
     'La tabla detallada del mismo certificado da 100,73 (diferencia del 0,03 %).',
     'Ficha RES060, apartado 3 y nota 5: demanda según el certificado antes de la actuación. El control 2 del cálculo muestra el efecto de usar 100,73.'),
    ('Superficie cubierta por la caldera (S_cal)', '126,45 m² (83,41 %)', 'D2',
     'Instalaciones › Sólo calefacción › Demanda cubierta · Superficie y porcentaje',
     'Equivale al 83,41 % de 151,6 m². La bomba de calor de la buhardilla cubre 25,0 m² (16,49 %). '
     'Contraste con D1 p. 3: caldera 61,8 % y 3 bombas de calor al 153 %, consumo final de calefacción 146,92 kWh/m²·año → la caldera cubría el 83,46 %.',
     'El resto de la demanda ya la cubrían bombas de calor eléctricas existentes, por lo que el ahorro se calcula solo sobre la parte de la caldera.'),
    ('Rendimiento de la caldera (ηi)', '0,83', 'D5, D6',
     'D5: criterio 24/11.03 · D6: Anexo II, tabla B.3, fila «Gas, de 1998 en adelante, condensación, con encendido automático» (83 %)',
     'El criterio 24/11.03 cita expresamente la RES060 y permite usar los valores de los anexos de RES210 y TER210. Las dos tablas descargadas del MITECO son idénticas.',
     'Valor de la tabla B.3 para caldera de gas de condensación con encendido automático. El certificado previo modela la caldera como «Caldera Estándar» de GLP (rendimiento estacional 61,8 %) y la placa de la caldera no es legible en las fotografías.'),
    ('Zona climática y clima de referencia', 'B4 · cálido', 'D1, D4',
     'D1: p. 1 · D4: Anexo III, tabla de equivalencia', 'La zona B4 corresponde a clima «cálido» en el Anexo III de la ficha.', 'Determina qué eficiencia estacional de EPREL se utiliza.'),
    ('Temperatura de impulsión', '55 °C', '—', 'Dato de la instalación (radiadores)',
     'Ficha RES060, Anexo III: SCOP a la temperatura de impulsión necesaria.', 'Radiadores → media temperatura (55 °C) en EPREL.'),
    ('Eficiencia estacional en calefacción (ηs,h)', '160 %', 'D3', 'Registro EPREL 345915 · media temperatura (55 °C) · clima cálido',
     'EPREL coincide con la ficha de producto Panasonic (reglamento 811/2013).', 'Modelo del registro: WH-MDC09J3E5. La placa instalada lleva el sufijo «-1».'),
    ('SCOP de la bomba de calor', '4,07', 'D3, D4',
     'D4: Anexo IV, SCOP = CC · (ηs,h + F(1) + F(2)) con CC = 2,5 y F(1) = 3 %', 'SCOP = 2,5 × (1,60 + 0,03) = 4,075; se adopta con dos decimales (4,07).',
     'Criterio 24/11.01: el error de redondeo no puede superar el 1 %. Aquí es del 0,12 %.'),
    ('Servicio de ACS', 'No interviene', 'D1', 'p. 3 · Equipo de ACS: efecto Joule eléctrico (100 %)',
     'La caldera retirada era solo de calefacción.', 'El término de ACS de la fórmula de la ficha vale cero.'),
    ('Equipo instalado', 'Panasonic WH-MDC09J3E5-1', 'D7', 'Fotografías de la placa de características',
     'Nº de serie [nº serie omitido] · fecha de producción 2026.07 · 9 kW · R32.', 'La placa no indica SCOP ni datos a 55 °C: se usa EPREL.'),
]
for i, fila in enumerate(tr):
    r = 6 + i
    for col, val in zip('BCDEFG', fila):
        put(tz, f'{col}{r}', val, bold=(col == 'B'), h=('center' if col == 'D' else 'left'),
            bg=(PALE if col == 'B' else None), v='top')
    autoheight(tz, r, 'BCEFG', minh=30)

r2 = 6 + len(tr) + 1
section(tz, r2, 'B · CONTRASTE CON OTROS CRITERIOS (INFORMATIVO)', 'B', 'G')
header(tz, r2 + 1, ['Caso', 'CAE (kWh/año)', '', 'Datos', 'Diferencia con el cálculo', 'Comentario'], 'BCDEFG')
X = lambda k: 'Cálculo!' + R[k]
SCOPX = CELL['SCOP']
cmp_rows = [
    ('Cálculo de este libro', f'={CELL["AE"]}', 'ηi 0,83 · parte de la caldera 83,41 % · SCOP 4,07', '', 'Criterio 24/11.03 y demanda cubierta por la caldera.'),
    ('Sin repartir la demanda (ηi 0,83)', f'={X("FP")}*{X("DCAL")}*{X("S")}*(1/{X("ETA")}-1/{SCOPX})',
     'Caldera cubriendo el 100 % de la calefacción', '', 'Lectura literal de la ficha con ηi 0,83.'),
    ('ηi 0,92 (valor por defecto de la ficha) con reparto', f'={X("FP")}*{X("DCAL")}*{X("SCAL")}*(1/0.92-1/{SCOPX})',
     'ηi 0,92 · parte de la caldera 83,41 % · SCOP 4,07', '', 'Rendimiento por defecto del apartado 3 de la ficha.'),
    ('ηi 0,92 sin reparto', f'={X("FP")}*{X("DCAL")}*{X("S")}*(1/0.92-1/{SCOPX})',
     'ηi 0,92 · caldera 100 % · SCOP 4,07', '', 'Lectura literal de la ficha con el rendimiento por defecto.'),
]
for i, (a, f, d, dif, com) in enumerate(cmp_rows):
    r = r2 + 2 + i
    put(tz, f'B{r}', a, bold=(i == 0), bg=(PALE if i else GREEN))
    put(tz, f'C{r}', f, fmt='#,##0.0', h='right', bold=(i == 0), bg=(GREEN if i == 0 else None))
    put(tz, f'D{r}', '')
    put(tz, f'E{r}', d)
    put(tz, f'F{r}', (f'=C{r}-$C${r2+2}' if i else '—'), fmt='+#,##0.0;-#,##0.0;0.0', h='right')
    put(tz, f'G{r}', com)
    autoheight(tz, r, 'BEG', minh=22)

r3 = r2 + 2 + len(cmp_rows) + 1
section(tz, r3, 'C · EVIDENCIAS EN EL EXPEDIENTE', 'B', 'G')
header(tz, r3 + 1, ['Evidencia', 'Tipo', 'Doc.', 'Qué muestra', '', ''], 'BCDEFG')
tz.merge_cells(f'E{r3+1}:G{r3+1}')
evid = [
    ('Certificado de eficiencia energética previo', 'PDF', 'D1', 'Superficie, demanda, zona climática, generadores y ACS (CE3X v2.3).'),
    ('Captura CE3X · Sólo calefacción', 'Imagen', 'D2', 'Caldera estándar, GLP, 24 kW · demanda cubierta 126,45 m² (83,41 %) · rendimiento estacional 61,8 %.'),
    ('Captura CE3X · Calefacción y refrigeración Buhardilla', 'Imagen', 'D2', 'Bomba de calor · demanda cubierta 25,0 m² (16,49 %) · rendimiento estacional 153 %.'),
    ('Fotografías de la caldera retirada', 'Imagen', '—', 'Caldera mural; la placa de características no es legible.'),
    ('Fotografías de la placa de la bomba instalada', 'Imagen', 'D7', 'Modelo, nº de serie, potencia, COP A7/W35 y A2/W35.'),
]
for i, (a, b, c, d) in enumerate(evid):
    r = r3 + 2 + i
    put(tz, f'B{r}', a, bold=True, bg=PALE)
    put(tz, f'C{r}', b, h='center')
    put(tz, f'D{r}', c, h='center')
    tz.merge_cells(f'E{r}:G{r}')
    put(tz, f'E{r}', d)
    for col in 'FG':
        tz[f'{col}{r}'].border = BOX
    tz.row_dimensions[r].height = max(20, 13.2 * lines(d, 140) + 6, 13.2 * lines(a, 30) + 6)
tz.freeze_panes = 'A4'
tz.sheet_properties.tabColor = 'C55A11'
page(tz)

# ===========================================================================
# REFERENCIAS
# ===========================================================================
banner(rf, 'Referencias', 'Documentos consultados, normativa y criterios aplicados, y datos del registro EPREL.', 'F')
widths(rf, {'B': 7, 'C': 46, 'D': 30, 'E': 62, 'F': 44})
section(rf, 4, 'A · DOCUMENTOS', 'B', 'F')
header(rf, 5, ['Cód.', 'Documento', 'Emisor · fecha', 'Identificación / enlace', 'Datos que se toman'], 'BCDEF')
docs = [
    ('D1', 'Certificado de eficiencia energética previo a la actuación (CE3X v2.3)', 'Registro Andaluz · 30/07/2026',
     'Nº de registro [nº registro omitido] · [nº registro omitido] · ref. catastral [ref. catastral omitida]', 'S, Dcal, zona climática, generadores, ACS'),
    ('D2', 'Archivo CE3X del certificado previo (CEE_previo.cex)', 'Técnico certificador · capturas de pantalla',
     'Instalaciones › Sólo calefacción / Calefacción y refrigeración Buhardilla › Demanda cubierta', 'Superficie de calefacción cubierta por la caldera'),
    ('D3', 'Registro EPREL 345915 · Panasonic WH-MDC09J3E5', 'Comisión Europea · versión 2, 11/03/2024',
     'https://eprel.ec.europa.eu/screen/product/spaceheaters/345915', 'Eficiencia estacional ηs,h por clima y temperatura'),
    ('D4', 'Ficha RES060 V1.1 y Anexos II a VI', 'MITECO · catálogo vigente de fichas (residencial)',
     'https://www.miteco.gob.es/es/energia/eficiencia/cae/catalogo-de-fichas/catalogo-vigente-de-fichas/residencial.html', 'Fórmula, variables, SCOP, zonas y clima'),
    ('D5', 'Criterios de verificación del Sistema CAE v14.0', 'MITECO · 16/10/2026', 'Criterios 24/11.01, 24/11.03, 25-07.03', 'Redondeos, rendimiento de la caldera, límite físico'),
    ('D6', 'Ficha RES210 · Anexo II «Rendimiento de caldera sustituida» (idéntico en la ficha TER210)', 'MITECO · catálogo vigente',
     'https://www.miteco.gob.es/content/dam/miteco/es/energia/files-1/Eficiencia/CAE/catalogo-vigente/fichas-autocompletables/residencial/Ficha%20RES210_Anexo%20II.pdf',
     'Tabla B.3: rendimiento estacional por defecto de la caldera'),
    ('D7', 'Placa de características de la bomba de calor (fotografías)', 'Fabricante · Panasonic', 'WH-MDC09J3E5-1 · nº de serie [nº serie omitido] · producción 2026.07', 'Identificación del equipo instalado'),
    ('D8', 'Preguntas frecuentes del Sistema CAE v1.12', 'MITECO · RD 36/2023', 'Valor del CAE y duración del ahorro', '1 CAE = 1 kWh de ahorro anual'),
    ('D9', 'Certificado de eficiencia energética posterior a la actuación (CE3X 2026.08.20)', 'Técnico certificador · visita 14/10/2026',
     'Firmado el 20/10/2026', 'Contraste de la superficie útil'),
]
for i, fila in enumerate(docs):
    r = 6 + i
    for col, val in zip('BCDEF', fila):
        put(rf, f'{col}{r}', val, bold=(col == 'B'), h=('center' if col == 'B' else 'left'), bg=(PALE if col == 'B' else None), v='top')
    autoheight(rf, r, 'CDEF', minh=26)

r2 = 6 + len(docs) + 1
section(rf, r2, 'B · NORMATIVA Y CRITERIOS APLICADOS', 'B', 'F')
header(rf, r2 + 1, ['Ref.', 'Documento', 'Ubicación', 'Qué establece', 'Dónde se aplica'], 'BCDEF')
norm = [
    ('N1', 'Ficha RES060 V1.1', 'Apartado 3', 'AE = FP · [(Dcal · S) · (1/ηi − 1/SCOP) + Dacs · (1/ηi − 1/SCOPdhw)], ahorro anual de energía final en kWh/año.', 'Pasos 1 a 8 del cálculo'),
    ('N2', 'Ficha RES060 V1.1', 'Apartado 3 · tabla de variables', 'FP = 1; ηi referido a PCS; Dcal y S según el certificado de eficiencia energética antes de la actuación.', 'Datos de entrada 1, 2, 8'),
    ('N3', 'Ficha RES060 V1.1', 'Anexo III', 'Se usa el SCOP del fabricante a la temperatura de impulsión; el clima de referencia depende de la zona (B4 → cálido).', 'Datos de entrada 5, 6, 7'),
    ('N4', 'Ficha RES060 V1.1', 'Anexo IV', 'SCOP = CC · (ηs,h + F(1) + F(2)); CC = 2,5; F(1) = 3 % en aerotermia; F(2) = 0.', 'Pasos 4 y 5'),
    ('N5', 'Ficha RES060 V1.1', 'Anexo V', 'Hasta 70 kW, los rendimientos son los de EPREL o de la ficha técnica del fabricante.', 'Dato ηs,h (EPREL)'),
    ('N6', 'Ficha RES060 V1.1', 'Anexo II', 'Cuando el equipo existente sigue cubriendo parte de la demanda, el ahorro se calcula sobre el porcentaje de demanda que cubre la bomba de calor. '
                                          'Se aplica por analogía: el ahorro solo existe en la parte de la demanda que cubría la caldera.', 'Pasos 2 y 3'),
    ('N7', 'Criterios de verificación v14.0', 'Criterio 24/11.01', 'Utilizar todas las cifras significativas; el error de redondeo no puede superar el 1 %.', 'Pasos 5 y 8'),
    ('N8', 'Criterios de verificación v14.0', 'Criterio 24/11.03', 'En las fichas de sustitución de caldera por bomba de calor (incluida RES060) se pueden usar los rendimientos de los anexos de las fichas RES210 y TER210.', 'Dato ηi'),
    ('N9', 'Ficha RES210 (y TER210)', 'Anexo II · tabla B.3', 'Rendimiento estacional por defecto: gas de 1998 en adelante, condensación, encendido automático = 83 %; anterior a 1998 = 85 %.', 'Dato ηi'),
    ('N10', 'Criterios de verificación v14.0', 'Criterio 25-07.03', 'Límite físico: el ahorro no puede superar una proporción razonable del consumo que se evita.', 'Control 1'),
    ('N11', 'FAQ del Sistema CAE v1.12', 'Valor del CAE', '1 CAE equivale a 1 kWh de ahorro anual de energía final.', 'Paso 9'),
]
for i, fila in enumerate(norm):
    r = r2 + 2 + i
    for col, val in zip('BCDEF', fila):
        put(rf, f'{col}{r}', val, bold=(col == 'B'), h=('center' if col == 'B' else 'left'), bg=(PALE if col == 'B' else None), v='top')
    autoheight(rf, r, 'CDEF', minh=24)

r3 = r2 + 2 + len(norm) + 1
section(rf, r3, 'C · REGISTRO EPREL 345915 · EFICIENCIA ESTACIONAL ηs EN CALEFACCIÓN (%)', 'B', 'F')
header(rf, r3 + 1, ['', 'Clima de referencia', 'Baja temperatura (35 °C)', 'Media temperatura (55 °C)', 'Se utiliza'], 'BCDEF')
for i, (clima, bt, mt) in enumerate([('Medio', 193, 130), ('Cálido', 227, 160), ('Frío', 164, 116)]):
    r = r3 + 2 + i
    sel = clima == 'Cálido'
    put(rf, f'B{r}', '')
    put(rf, f'C{r}', clima, bold=True, bg=(GREEN if sel else PALE))
    put(rf, f'D{r}', bt, color=INPUT_BLUE, h='center', fmt='0')
    put(rf, f'E{r}', mt, color=INPUT_BLUE, h='center', fmt='0', bold=sel, bg=(GREEN if sel else None))
    put(rf, f'F{r}', 'Zona B4 (clima cálido) · radiadores (55 °C) → 160 %' if sel else '', bold=sel)
    rf.row_dimensions[r].height = 20
rf.freeze_panes = 'A4'
rf.sheet_properties.tabColor = '548235'
page(rf)

# la hoja de cálculo se abre en el Resumen
wb.active = 0
wb.save(SALIDA)
print('guardado', SALIDA)
