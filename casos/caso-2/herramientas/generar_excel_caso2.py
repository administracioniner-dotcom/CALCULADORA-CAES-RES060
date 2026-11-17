#!/usr/bin/env python3
"""Libro final del caso 2 (RES060 V1.1, calefacción + ACS): Resumen · Cálculo · Trazabilidad · Referencias."""
import sys
sys.path.insert(0, '/home/user/casos/caso-2/herramientas')
from utils_excel import *  # noqa  (paleta, put, banner, section, header, autoheight, page, lines, ...)

wb = Workbook()
rs = wb.active
rs.title = 'Resumen'
cl = wb.create_sheet('Cálculo')
tz = wb.create_sheet('Trazabilidad')
rf = wb.create_sheet('Referencias')

# ===========================================================================
# CÁLCULO
# ===========================================================================
banner(cl, 'Cálculo del ahorro energético · Ficha RES060 V1.1 (calefacción y ACS)',
       'Cada resultado es una fórmula enlazada a los datos de entrada. Azul sobre crema = dato introducido; negro = resultado calculado.', 'H')
widths(cl, {'B': 5, 'C': 48, 'D': 11, 'E': 16, 'F': 13, 'G': 46, 'H': 52})
section(cl, 4, '1 · DATOS DE ENTRADA', 'B', 'H')
header(cl, 5, ['Nº', 'Dato', 'Símbolo', 'Valor', 'Unidad', 'Documento de origen', 'Ubicación en el documento'], 'BCDEFGH')

inputs = [
    ('S', 'Superficie útil habitable', 'S', 90.69, 'm²', '#,##0.00',
     'D1 · Certificado de eficiencia energética previo', 'p. 2 · Superficie habitable'),
    ('DCAL', 'Demanda de energía en calefacción', 'Dcal', 76.57, 'kWh/m²·año', '#,##0.00',
     'D1 · Certificado de eficiencia energética previo', 'Anexo III · Demanda de calefacción de la situación original (la etiqueta del Anexo II la redondea a 76,6)'),
    ('EPACS', 'Energía primaria no renovable de ACS (situación previa)', 'EP_ACS', 42.38, 'kWh/m²·año', '#,##0.00',
     'D1 · Certificado de eficiencia energética previo', 'Anexo II, apartado 2 · Energía primaria ACS'),
    ('FEP', 'Factor de paso a energía primaria no renovable del gasóleo-C', 'f_EP', 1.179, 'kWhEP/kWh', '0.000',
     'D9 · Factores de emisión y paso a energía primaria (documento reconocido)', 'Gasóleo calefacción · energía primaria no renovable'),
    ('ETACEE', 'Rendimiento estacional de la caldera en el certificado', 'ηcert', 0.724, '—', '0.000',
     'D1 · Certificado de eficiencia energética previo', 'p. 3 · Generador «Instalación Calefacción y ACS – Gasoil»: 72,4 %'),
    ('SCAL', 'Parte de la demanda (calefacción y ACS) cubierta por la caldera', 'f', 1.0, '—', '0%',
     'D1 · Certificado previo · archivo CE3X', 'p. 3 · Único generador de calefacción y ACS · demanda cubierta 100 %'),
    ('ETA', 'Rendimiento de la caldera retirada (referido a PCS)', 'ηi', 0.79, '—', '0.00',
     'D6 · Criterios de verificación v14.0 + D7 · Ficha RES210, Anexo II', 'Criterio 24/11.03 · Tabla B.3: gasóleo, de 1998 en adelante, sin condensación = 79 %'),
    ('ZONA', 'Zona climática (clima de referencia: medio)', 'Zona', 'E1', '—', '@',
     'D1 · Certificado previo + D5 · Ficha RES060, Anexo III', 'CEE p. 1 · Anexo III: E1 → clima medio'),
    ('TIMP', 'Temperatura de impulsión (instalación con radiadores)', 'T', 55, '°C', '0',
     'Dato de la instalación', 'Emisores: radiadores → temperatura media (55 °C)'),
    ('ETAS', 'Eficiencia energética estacional en calefacción (55 °C, clima medio)', 'ηs,h', 128, '%', '0',
     'D3 · Registro EPREL 305017 · THERMOR 526683', 'Eficiencia energética estacional de calefacción, clima medio, 55 °C: 128 %'),
    ('ETAWH', 'Eficiencia energética del calentamiento de agua (perfil L, clima medio)', 'ηwh', 130, '%', '0',
     'D3 · Registro EPREL 305017 · THERMOR 526683', 'Calentamiento de agua · perfil de extracción L · clima medio: 130 % · clase A+'),
    ('FP', 'Factor de ponderación', 'FP', 1, '—', '0', 'D5 · Ficha RES060 V1.1', 'Apartado 3, tabla de variables'),
    ('CC', 'Coeficiente de conversión a energía primaria de la electricidad', 'CC', 2.5, '—', '0.0', 'D5 · Ficha RES060 V1.1', 'Anexo IV, nota 1'),
    ('F1', 'Factor de corrección F(1) de la bomba aerotérmica', 'F(1)', 0.03, '—', '0%', 'D5 · Ficha RES060 V1.1', 'Anexo IV, nota 2 (F(2) = 0 en aerotermia)'),
]
R = {}
for i, (key, dato, sim, val, uni, fmt, doc, ubi) in enumerate(inputs):
    r = 6 + i
    R[key] = f'$E${r}'
    put(cl, f'B{r}', i + 1, h='center', color=GREYTXT)
    put(cl, f'C{r}', dato)
    put(cl, f'D{r}', sim, bold=True, h='center')
    put(cl, f'E{r}', val, color=INPUT_BLUE, bg=CREAM, fmt=fmt, h='right', bold=True)
    put(cl, f'F{r}', uni, h='center')
    put(cl, f'G{r}', doc)
    put(cl, f'H{r}', ubi)
    autoheight(cl, r, 'CGH')

r_sec2 = 6 + len(inputs) + 1
section(cl, r_sec2, '2 · CÁLCULO PASO A PASO', 'B', 'H')
header(cl, r_sec2 + 1, ['Nº', 'Concepto', 'Símbolo', 'Resultado', 'Unidad', 'Fórmula', 'Sustitución numérica'], 'BCDEFGH')
S, DC, EPA, FEP, ECE, SC, ETA, ETAS, ETAWH, FP, CC, F1 = (R[k] for k in
    ('S', 'DCAL', 'EPACS', 'FEP', 'ETACEE', 'SCAL', 'ETA', 'ETAS', 'ETAWH', 'FP', 'CC', 'F1'))
p0 = r_sec2 + 2
row = {n: p0 + n for n in range(14)}
E = lambda n: f'E{row[n]}'
steps = [
    ('Demanda anual de calefacción del edificio', 'Dcal·S', f'={DC}*{S}', 'kWh/año', '#,##0.0', 'Dcal × S',
     f'=ROUND({DC},2)&" × "&ROUND({S},2)&" = "&ROUND({E(0)},1)'),
    ('Demanda anual de ACS del edificio', 'Dacs', f'={EPA}/{FEP}*{ECE}*{S}', 'kWh/año', '#,##0.0',
     'EP_ACS ÷ f_EP × ηcert × S  (consumo final de ACS × rendimiento del certificado)',
     f'=ROUND({EPA},2)&" ÷ "&{FEP}&" × "&{ECE}&" × "&ROUND({S},2)&" = "&ROUND({E(1)},1)'),
    ('Demanda de calefacción cubierta por la caldera', 'Dc', f'={E(0)}*{SC}', 'kWh/año', '#,##0.0', 'Dcal·S × f',
     f'=ROUND({E(0)},1)&" × "&ROUND({SC},2)&" = "&ROUND({E(2)},1)'),
    ('Demanda de ACS cubierta por la caldera', 'Da', f'={E(1)}*{SC}', 'kWh/año', '#,##0.0', 'Dacs × f',
     f'=ROUND({E(1)},1)&" × "&ROUND({SC},2)&" = "&ROUND({E(3)},1)'),
    ('SCOP de calefacción según Anexo IV', 'SCOP', f'={CC}*({ETAS}/100+{F1})', '—', '0.000', 'CC × (ηs,h + F(1))',
     f'={CC}&" × ("&{ETAS}/100&" + "&{F1}&") = "&ROUND({E(4)},4)'),
    ('SCOP de ACS según Anexo IV', 'SCOPdhw', f'={CC}*{ETAWH}/100', '—', '0.000', 'CC × ηwh',
     f'={CC}&" × "&{ETAWH}/100&" = "&ROUND({E(5)},4)'),
    ('Consumo de calefacción con la caldera retirada', 'Ec,i', f'={E(2)}/{ETA}', 'kWh/año', '#,##0.0', 'Dc ÷ ηi',
     f'=ROUND({E(2)},1)&" ÷ "&{ETA}&" = "&ROUND({E(6)},1)'),
    ('Consumo de calefacción con la bomba de calor', 'Ec,f', f'={E(2)}/{E(4)}', 'kWh/año', '#,##0.0', 'Dc ÷ SCOP',
     f'=ROUND({E(2)},1)&" ÷ "&ROUND({E(4)},3)&" = "&ROUND({E(7)},1)'),
    ('Ahorro anual en calefacción', 'AEcal', f'={E(6)}-{E(7)}', 'kWh/año', '#,##0.0', 'Ec,i − Ec,f',
     f'=ROUND({E(6)},1)&" − "&ROUND({E(7)},1)&" = "&ROUND({E(8)},1)'),
    ('Consumo de ACS con la caldera retirada', 'Ea,i', f'={E(3)}/{ETA}', 'kWh/año', '#,##0.0', 'Da ÷ ηi',
     f'=ROUND({E(3)},1)&" ÷ "&{ETA}&" = "&ROUND({E(9)},1)'),
    ('Consumo de ACS con la bomba de calor', 'Ea,f', f'={E(3)}/{E(5)}', 'kWh/año', '#,##0.0', 'Da ÷ SCOPdhw',
     f'=ROUND({E(3)},1)&" ÷ "&ROUND({E(5)},3)&" = "&ROUND({E(10)},1)'),
    ('Ahorro anual en ACS', 'AEacs', f'={E(9)}-{E(10)}', 'kWh/año', '#,##0.0', 'Ea,i − Ea,f',
     f'=ROUND({E(9)},1)&" − "&ROUND({E(10)},1)&" = "&ROUND({E(11)},1)'),
    ('Ahorro anual de energía final', 'AE', f'={FP}*({E(8)}+{E(11)})', 'kWh/año', '#,##0.0', 'FP × (AEcal + AEacs)',
     f'={FP}&" × ("&ROUND({E(8)},1)&" + "&ROUND({E(11)},1)&") = "&ROUND({E(12)},2)'),
    ('Certificados de ahorro energético (1 CAE = 1 kWh)', 'CAE', f'=ROUND({E(12)},0)', 'CAE', '#,##0', 'CAE = AE (redondeado)',
     f'="AE = "&ROUND({E(12)},2)&" kWh/año → "&TEXT({E(13)},"0")&" CAE"'),
]
for n, (concepto, sim, formula, uni, fmt, fx, sub) in enumerate(steps):
    r = row[n]
    last = n == 13
    bg_res = 'FFF2CC' if last else (GREEN if n in (8, 11, 12) else None)
    put(cl, f'B{r}', n + 1, h='center', color=GREYTXT)
    put(cl, f'C{r}', concepto, bold=(n >= 12))
    put(cl, f'D{r}', sim, bold=True, h='center')
    put(cl, f'E{r}', formula, fmt=fmt, h='right', bold=(n >= 12), bg=bg_res, size=(12 if last else 10))
    put(cl, f'F{r}', uni, h='center')
    put(cl, f'G{r}', fx)
    put(cl, f'H{r}', sub, color=GREYTXT)
    autoheight(cl, r, 'CGH', minh=(24 if last else (32 if n == 1 else 18)))

r_note = row[13] + 1
cl.merge_cells(f'B{r_note}:H{r_note}')
put(cl, f'B{r_note}',
    'Fórmula de la ficha: AE = FP · [ (Dcal · S) · (1/ηi − 1/SCOP) + Dacs · (1/ηi − 1/SCOPdhw) ]. Calefacción y ACS: la caldera de gasóleo producía los dos servicios y la bomba '
    '(con depósito de 190 L suministrado como conjunto) los sustituye (Anexo VI, caso 1). Sin multiplicadores.',
    italic=True, color=GREYTXT, border=False, v='top')
cl.row_dimensions[r_note].height = 42

r_sec3 = r_note + 2
section(cl, r_sec3, '3 · CONTROLES', 'B', 'H')
header(cl, r_sec3 + 1, ['Nº', 'Control', '', 'Valor', 'Unidad', 'Qué comprueba', 'Resultado'], 'BCDEFGH')
c1, c2 = r_sec3 + 2, r_sec3 + 3
put(cl, f'B{c1}', 1, h='center', color=GREYTXT)
put(cl, f'C{c1}', 'Ahorro sobre el consumo previo de la caldera'); put(cl, f'D{c1}', '')
put(cl, f'E{c1}', f'={E(12)}/({E(6)}+{E(9)})', fmt='0.0%', h='right'); put(cl, f'F{c1}', '%', h='center')
put(cl, f'G{c1}', 'El ahorro no puede superar el consumo que se evita (criterio 25-07.03)')
put(cl, f'H{c1}', f'=IF({E(12)}<{E(6)}+{E(9)},"Correcto: el ahorro es menor que el consumo previo","Revisar: el ahorro supera el consumo previo")')
put(cl, f'B{c2}', 2, h='center', color=GREYTXT)
put(cl, f'C{c2}', 'Demanda de ACS calculada con el Anejo G del DB HE (Burgos, 112 l/día)'); put(cl, f'D{c2}', '')
put(cl, f'E{c2}', 2372.0, color=INPUT_BLUE, bg=CREAM, fmt='#,##0.0', h='right', bold=True); put(cl, f'F{c2}', 'kWh/año', h='center')
put(cl, f'G{c2}', 'Contraste de la demanda de ACS con el método del CTE (temperaturas de red de la capital, 60 °C)')
put(cl, f'H{c2}', f'="Diferencia con el paso 2: "&ROUND({E(1)}-E{c2},1)&" kWh ("&ROUND(({E(1)}-E{c2})/E{c2}*100,2)&" %)"')
for r in (c1, c2):
    autoheight(cl, r, 'CGH')
cl.freeze_panes = 'A4'
cl.sheet_properties.tabColor = BLUE
page(cl)
CELL = {k: f"Cálculo!$E${row[n]}" for k, n in (('D', 0), ('DA', 1), ('DC', 2), ('DAC', 3), ('SCOP', 4), ('SCOPDHW', 5),
                                                ('ECI', 6), ('ECF', 7), ('AECAL', 8), ('EAI', 9), ('EAF', 10), ('AEACS', 11), ('AE', 12), ('CAE', 13))}

# ===========================================================================
# RESUMEN
# ===========================================================================
banner(rs, 'Certificados de Ahorro Energético (CAE) · Ficha RES060 V1.1',
       'Sustitución de caldera de gasóleo por bomba de calor aerotérmica (calefacción y ACS) · Caso 2', 'G')
widths(rs, {'B': 30, 'C': 20, 'D': 20, 'E': 20, 'F': 20, 'G': 20})
section(rs, 4, 'IDENTIFICACIÓN DE LA ACTUACIÓN', 'B', 'G')
ident = [
    ('Emplazamiento', 'Vivienda caso 2 (Burgos) · vivienda unifamiliar construida en 2007 · zona climática E1'),
    ('Referencia catastral', '[ref. catastral omitida]'),
    ('Actuación', 'Retirada de la caldera de gasóleo-C Tifell Eurofell Blue 30 TI (2007, 25,66 kW; calefacción y ACS) e instalación de bomba de calor '
                  'aerotérmica aire-agua Thermor Alféa Extensa Duo A.I. 8 con depósito de ACS de 190 L (unidad exterior Fujitsu WOYA080KLT, nº de serie [nº serie omitido]); emisores: radiadores'),
    ('Certificado de eficiencia energética previo', 'CE3X v2.3 · visita 01/06/2026 · firmado el 03/06/2026'),
    ('Servicio y método', 'Calefacción y ACS · ficha RES060 V1.1 · sin efecto multiplicador'),
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
header(rs, 16, ['Concepto', 'Calefacción', 'ACS', 'Total', 'Unidad', 'Se calcula como'], 'BCDEFG')
flow = [
    ('Demanda cubierta por la caldera', CELL['DC'], CELL['DAC'], None, 'Dcal·S y Dacs'),
    ('Consumo con la caldera retirada (ηi = 0,79)', CELL['ECI'], CELL['EAI'], None, 'Demanda ÷ rendimiento de la caldera'),
    ('Consumo con la bomba de calor', CELL['ECF'], CELL['EAF'], None, 'Demanda ÷ SCOP (3,275) y ÷ SCOPdhw (3,25)'),
    ('Ahorro anual de energía final', CELL['AECAL'], CELL['AEACS'], CELL['AE'], 'Consumo antes − consumo después'),
]
for i, (a, c1_, c2_, tot, how) in enumerate(flow):
    r = 17 + i
    last = i == 3
    bgc = GREEN if last else None
    put(rs, f'B{r}', a, bold=last, bg=bgc)
    put(rs, f'C{r}', f'={c1_}', fmt='#,##0.0', h='right', bold=last, bg=bgc)
    put(rs, f'D{r}', f'={c2_}', fmt='#,##0.0', h='right', bold=last, bg=bgc)
    put(rs, f'E{r}', (f'={tot}' if tot else f'=C{r}+D{r}'), fmt='#,##0.0', h='right', bold=last, bg=bgc)
    put(rs, f'F{r}', 'kWh/año', h='center', bg=bgc)
    put(rs, f'G{r}', how, color=GREYTXT, bg=bgc)
    rs.row_dimensions[r].height = 32

section(rs, 22, 'DATOS DE PARTIDA', 'B', 'G')
header(rs, 23, ['Dato', 'Valor', 'Unidad', 'Documento', '', ''], 'BCDEFG')
rs.merge_cells('E23:G23')
base = [
    ('Superficie útil', 'S', 'm²', '#,##0.00', 'D1 · Certificado previo'),
    ('Demanda de calefacción', 'DCAL', 'kWh/m²·año', '#,##0.00', 'D1 · Certificado previo'),
    ('Rendimiento de la caldera', 'ETA', '—', '0.00', 'D6 · Criterio 24/11.03 · D7 · Ficha RES210, Anexo II'),
    ('Eficiencia estacional de calefacción (ηs,h)', 'ETAS', '%', '0', 'D3 · Registro EPREL 305017'),
    ('Eficiencia energética de ACS (ηwh)', 'ETAWH', '%', '0', 'D3 · Registro EPREL 305017'),
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
    ('Referencias', 'Documentos consultados (D1 a D12), normativa y criterios aplicados, y datos del fabricante.'),
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
       'Para cada dato del cálculo: valor usado, documento (códigos D1–D12, ver hoja Referencias), ubicación exacta y comprobación realizada.', 'G')
widths(tz, {'B': 30, 'C': 18, 'D': 9, 'E': 44, 'F': 46, 'G': 58})
section(tz, 4, 'A · DATOS DEL CÁLCULO Y SU DOCUMENTO DE ORIGEN', 'B', 'G')
header(tz, 5, ['Dato', 'Valor usado', 'Doc.', 'Ubicación en el documento', 'Comprobación', 'Observaciones'], 'BCDEFG')
tr = [
    ('Superficie útil habitable (S)', '90,69 m²', 'D1', 'p. 2 · Superficie habitable', 'Coincide con el certificado posterior (D2, 90,69 m²).', 'Superficie de la vivienda completa.'),
    ('Demanda de calefacción (Dcal)', '76,57 kWh/m²·año', 'D1', 'Anexo III · Demanda de la situación original (0,0 % de ahorro)',
     'Coincide con el certificado posterior (76,57). La etiqueta del Anexo II muestra el mismo valor redondeado (76,6).',
     'Ficha RES060, apartado 3 y nota 5: demanda según el certificado antes de la actuación. Criterio 24/11.01: todas las cifras significativas.'),
    ('Demanda de ACS (Dacs)', '2.360 kWh/año', 'D1, D9',
     'D1: Anexo II · energía primaria de ACS 42,38 kWh/m²·año (p. 4) y rendimiento 72,4 % (p. 3) · D9: factor de paso del gasóleo-C',
     'Misma cifra por emisiones: 11,18 kgCO₂/m² ÷ 0,311 × 0,724 × 90,69 m² = 2.360 kWh/año. El factor 1,179 se comprueba con el propio certificado: 124,69 kWh/m² de energía primaria de calefacción ÷ (76,57 ÷ 0,724) = 1,179. Contraste con el Anejo G del DB HE (D10): 2.372 kWh/año.',
     'El certificado no indica la demanda de ACS en kWh/año (solo 112 l/día a 60 °C); se obtiene de su consumo de ACS.'),
    ('Parte cubierta por la caldera (f)', '100 %', 'D1', 'p. 3 · Un único generador «Instalación Calefacción y ACS – Gasoil»',
     'No existen otros generadores de calefacción ni de ACS en el certificado previo.', 'La caldera cubría toda la calefacción y todo el ACS.'),
    ('Rendimiento de la caldera (ηi)', '0,79', 'D6, D7',
     'D6: criterio 24/11.03 · D7: Anexo II, tabla B.3, fila «Gasóleo, de 1998 en adelante, sin condensación» (79 %)',
     'El criterio 24/11.03 cita expresamente la RES060 y permite usar los valores de los anexos de RES210 y TER210. La fila corresponde a gasóleo, año ≥ 1998 y sin condensación.',
     'Caldera estándar (no de condensación) según el certificado previo, p. 3; fabricada en 2007 y de gasóleo-C (placa y certificado).'),
    ('Zona climática y clima de referencia', 'E1 · medio', 'D1, D5', 'D1: p. 1 · D5: Anexo III, tabla de equivalencia', 'La zona E1 corresponde a clima «medio» en el Anexo III de la ficha.', 'Determina la eficiencia estacional que se usa.'),
    ('Temperatura de impulsión', '55 °C', '—', 'Dato de la instalación (radiadores)',
     'Ficha RES060, Anexo III: se usa el SCOP a la temperatura de impulsión necesaria. Radiadores → media temperatura (55 °C).',
     'El certificado previo (Anexo III, descripción de la medida) menciona «suelo radiante» como sistema de distribución; la instalación ejecutada es con radiadores.'),
    ('Eficiencia estacional en calefacción (ηs,h)', '128 %', 'D3', 'Registro EPREL 305017 · THERMOR · modelo 526683 · ηs clima medio a 55 °C = 128 %',
     'Coincide con el catálogo del fabricante (D12): ηs 128 % y SCOP 3,28. El Anexo IV da 2,5 × (1,28 + 0,03) = 3,275.',
     'Registro publicado (versión 2, 26/05/2026). Modelo 526683 = Alféa Extensa Duo Ai 8 de la placa, con unidad exterior WOYA080KLT (D4, D8).'),
    ('Eficiencia energética de ACS (ηwh)', '130 %', 'D3', 'Registro EPREL 305017 · calentamiento de agua · perfil L · clima medio: 130 % · clase A+',
     'Coincide con D12 (130 %). Anexo VI, caso 1: depósito y bomba suministrados como conjunto; SCOPdhw = CC · ηwh (Anexo IV) en clima medio.',
     'La ficha comercial de Thermor da además un SCOP de ACS de 3,4 a 7 °C (EN 16147); se usa el dato que exige el Anexo IV.'),
    ('Equipo instalado', 'Thermor Alféa Extensa Duo A.I. 8', 'D8', 'Fotografías de las placas de la unidad exterior y del módulo hidráulico',
     'Unidad exterior Fujitsu WOYA080KLT (serie [nº serie omitido], Mar. 2023, R32); módulo hidráulico código 024 328; depósito de ACS de 190 L.', 'Código comercial Thermor 526683 (D4).'),
    ('Caldera retirada', 'Tifell Eurofell Blue 30 TI', 'D8', 'Fotografía de la placa de la caldera',
     'Año 2007 · consumo calorífico nominal 27,30 kW · potencia útil 25,66 kW (coincide con los 25,66 kW del certificado, p. 3).', 'Caldera de gasóleo de llama azul.'),
]
for i, fila in enumerate(tr):
    r = 6 + i
    for col, val in zip('BCDEFG', fila):
        put(tz, f'{col}{r}', val, bold=(col == 'B'), h=('center' if col == 'D' else 'left'), bg=(PALE if col == 'B' else None), v='top')
    autoheight(tz, r, 'BCEFG', minh=30)

r3 = 6 + len(tr) + 1
section(tz, r3, 'B · EVIDENCIAS EN EL EXPEDIENTE', 'B', 'G')
header(tz, r3 + 1, ['Evidencia', 'Tipo', 'Doc.', 'Qué muestra', '', ''], 'BCDEFG')
tz.merge_cells(f'E{r3+1}:G{r3+1}')
evid = [
    ('Certificado de eficiencia energética previo', 'PDF', 'D1', 'Superficie, demanda de calefacción, generador único de calefacción y ACS, energía primaria y emisiones (CE3X v2.3).'),
    ('Certificado de eficiencia energética posterior', 'PDF', 'D2', 'Misma superficie y demanda; bomba de calor para calefacción y ACS.'),
    ('Fotografía de la placa de la caldera', 'Imagen', 'D8', 'Tifell Eurofell Blue 30 TI · año 2007 · 27,30 kW / 25,66 kW.'),
    ('Fotografía de la unidad exterior', 'Imagen', 'D8', 'Fujitsu General · WOYA080KLT · serie [nº serie omitido] · R32 · Mar. 2023.'),
    ('Fotografía de la placa del módulo hidráulico', 'Imagen', 'D8', 'Thermor Alféa Extensa Duo A.I. R 6-8 kW · códigos 526682 / 526683 · módulo 024 328 · depósito ACS 190 L.'),
]
for i, (a, b, c, d) in enumerate(evid):
    r = r3 + 2 + i
    put(tz, f'B{r}', a, bold=True, bg=PALE); put(tz, f'C{r}', b, h='center'); put(tz, f'D{r}', c, h='center')
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
banner(rf, 'Referencias', 'Documentos consultados, normativa y criterios aplicados, y datos del fabricante.', 'F')
widths(rf, {'B': 7, 'C': 46, 'D': 30, 'E': 62, 'F': 44})
section(rf, 4, 'A · DOCUMENTOS', 'B', 'F')
header(rf, 5, ['Cód.', 'Documento', 'Emisor · fecha', 'Identificación / enlace', 'Datos que se toman'], 'BCDEF')
docs = [
    ('D1', 'Certificado de eficiencia energética previo a la actuación (CE3X v2.3)', 'Carrasco Requejo Arquitectos · 03/06/2026', 'Ref. catastral [ref. catastral omitida] · visita 01/06/2026', 'S, Dcal, generador, ACS, energía primaria y emisiones'),
    ('D2', 'Certificado de eficiencia energética posterior a la actuación (CE3X)', 'Carrasco Requejo Arquitectos · 25/08/2026', 'Ref. catastral [ref. catastral omitida] · visita 25/08/2026', 'Contraste de la superficie y de la demanda'),
    ('D3', 'Registro EPREL 305017 · THERMOR · modelo 526683 (calefactor combinado con bomba de calor)', 'Comisión Europea · versión 2, publicado el 26/05/2026',
     'https://eprel.ec.europa.eu/screen/product/spaceheaters/305017', 'ηs de calefacción por clima y temperatura, ηwh de ACS, perfil L'),
    ('D4', 'Ficha de producto Thermor «Alféa Extensa Duo Ai» (código 526683)', 'Thermor España', 'https://thermor.es/gaerp/download/ficha-poducto-alfea-extensa-duo-aipdf', 'Identificación del modelo y datos del equipo'),
    ('D5', 'Ficha RES060 V1.1 y Anexos II a VI', 'MITECO · catálogo vigente de fichas (residencial)',
     'https://www.miteco.gob.es/es/energia/eficiencia/cae/catalogo-de-fichas/catalogo-vigente-de-fichas/residencial.html', 'Fórmula, variables, SCOP, zonas y clima'),
    ('D6', 'Criterios de verificación del Sistema CAE v14.0', 'MITECO · 16/10/2026', 'Criterios 24/11.01, 24/11.03, 25-07.03', 'Redondeos, rendimiento de la caldera, límite físico'),
    ('D7', 'Ficha RES210 · Anexo II «Rendimiento de caldera sustituida» (idéntico en la ficha TER210)', 'MITECO · catálogo vigente',
     'https://www.miteco.gob.es/content/dam/miteco/es/energia/files-1/Eficiencia/CAE/catalogo-vigente/fichas-autocompletables/residencial/Ficha%20RES210_Anexo%20II.pdf', 'Tabla B.3: rendimiento estacional por defecto de la caldera'),
    ('D8', 'Fotografías de las placas de características (caldera, unidad exterior y módulo hidráulico)', 'Fabricantes · Tifell, Fujitsu General, Thermor', 'Fotos 1 a 3 del expediente', 'Identificación de los equipos retirado e instalado'),
    ('D9', 'Factores de emisión de CO₂ y coeficientes de paso a energía primaria de diferentes fuentes de energía final (documento reconocido)', 'MITECO · IDAE', 'Documento reconocido para la certificación energética', 'Factor del gasóleo-C: 1,179 kWh EP no renovable / kWh (y 0,311 kgCO₂/kWh)'),
    ('D10', 'CTE DB-HE · Anejo G · temperaturas de agua de red', 'Ministerio de Vivienda', 'Anejo G · capital de provincia: Burgos', 'Contraste de la demanda de ACS'),
    ('D11', 'Preguntas frecuentes del Sistema CAE v1.12', 'MITECO · RD 36/2023', 'Valor del CAE y duración del ahorro', '1 CAE = 1 kWh de ahorro anual'),
    ('D12', 'Catálogo del fabricante «Alféa Extensa – Fiche Produit» · gama Duo A.I. R32 (ref. 526161)', 'Atlantic · grupo al que pertenece Thermor',
     'https://d39twxq38hdb3s.cloudfront.net/Alf%C3%A9a%20Extensa%20-%20Fiche%20Produit.pdf', 'Contraste de ηs (177 / 128 %), SCOP (4,5 / 3,28) y ηwh (130 %)'),
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
    ('N1', 'Ficha RES060 V1.1', 'Apartado 3', 'AE = FP · [(Dcal · S) · (1/ηi − 1/SCOP) + Dacs · (1/ηi − 1/SCOPdhw)], ahorro anual de energía final en kWh/año.', 'Pasos 7 a 13'),
    ('N2', 'Ficha RES060 V1.1', 'Apartado 3 · tabla de variables', 'FP = 1; ηi referido a PCS; Dcal, S y Dacs según el certificado de eficiencia energética antes de la actuación.', 'Datos de entrada 1, 2, 3, 12'),
    ('N3', 'Ficha RES060 V1.1', 'Anexo III', 'Se usa el SCOP del fabricante a la temperatura de impulsión; el clima de referencia depende de la zona (E1 → medio).', 'Datos de entrada 8, 9, 10'),
    ('N4', 'Ficha RES060 V1.1', 'Anexo IV', 'SCOP = CC · (ηs,h + F(1) + F(2)); SCOPdhw = CC · ηhw; CC = 2,5; F(1) = 3 % en aerotermia; F(2) = 0.', 'Pasos 5 y 6'),
    ('N5', 'Ficha RES060 V1.1', 'Anexo V', 'Hasta 70 kW, los rendimientos son los de EPREL o de la ficha técnica del fabricante.', 'Datos ηs,h y ηwh'),
    ('N6', 'Ficha RES060 V1.1', 'Anexo VI · caso 1', 'Bomba aerotérmica y depósito suministrados como conjunto: el fabricante aporta el SCOPdhw en condiciones de clima medio (E1 → medio).', 'Paso 6'),
    ('N7', 'Criterios de verificación v14.0', 'Criterio 24/11.01', 'Utilizar todas las cifras significativas; el error de redondeo no puede superar el 1 %.', 'Todo el cálculo (sin redondeos intermedios)'),
    ('N8', 'Criterios de verificación v14.0', 'Criterio 24/11.03', 'En las fichas de sustitución de caldera por bomba de calor (incluida RES060) se pueden usar los rendimientos de los anexos de las fichas RES210 y TER210.', 'Dato ηi'),
    ('N9', 'Ficha RES210 (y TER210)', 'Anexo II · tabla B.3', 'Rendimiento estacional por defecto: gasóleo de 1998 en adelante sin condensación = 79 %; con condensación = 83 %.', 'Dato ηi'),
    ('N10', 'Criterios de verificación v14.0', 'Criterio 25-07.03', 'Límite físico: el ahorro no puede superar una proporción razonable del consumo que se evita.', 'Control 1'),
    ('N11', 'FAQ del Sistema CAE v1.12', 'Valor del CAE', '1 CAE equivale a 1 kWh de ahorro anual de energía final.', 'Paso 14'),
]
for i, fila in enumerate(norm):
    r = r2 + 2 + i
    for col, val in zip('BCDEF', fila):
        put(rf, f'{col}{r}', val, bold=(col == 'B'), h=('center' if col == 'B' else 'left'), bg=(PALE if col == 'B' else None), v='top')
    autoheight(rf, r, 'CDEF', minh=24)
r3 = r2 + 2 + len(norm) + 1
section(rf, r3, 'C · REGISTRO EPREL 305017 · THERMOR 526683 (D3)', 'B', 'F')
header(rf, r3 + 1, ['', 'Magnitud', 'Baja temperatura (35 °C)', 'Media temperatura (55 °C)', 'Se utiliza'], 'BCDEF')
mag = [
    ('Eficiencia energética estacional de calefacción ηs · clima medio (%)', 177, 128, 'Zona E1 (clima medio) y radiadores (55 °C) → 128 %'),
    ('Eficiencia energética estacional de calefacción ηs · clima cálido (%)', 228, 159, 'No se utiliza (E1 no es clima cálido)'),
    ('Eficiencia energética del calentamiento de agua ηwh · perfil L · clima medio (%)', 130, 130, '130 % → SCOPdhw = 2,5 × 1,30 = 3,25'),
]
for i, (m, a, b, u) in enumerate(mag):
    r = r3 + 2 + i
    put(rf, f'B{r}', '')
    put(rf, f'C{r}', m, bold=True, bg=PALE)
    put(rf, f'D{r}', a, color=INPUT_BLUE, h='center', fmt='0.0#' if isinstance(a, float) else '0', bold=(i == 2), bg=(GREEN if i == 2 else None))
    put(rf, f'E{r}', b, color=INPUT_BLUE, h='center', fmt='0.0#' if isinstance(b, float) else '0', bold=(i == 0), bg=(GREEN if i == 0 else None))
    put(rf, f'F{r}', u)
    autoheight(rf, r, 'CF', minh=22)
rf.freeze_panes = 'A4'
rf.sheet_properties.tabColor = '548235'
page(rf)
wb.active = 0
wb.save(SALIDA)
print('guardado', SALIDA)
