#!/usr/bin/env python3
"""Libro final del caso 4 (RES060 V1.1, calefacción + ACS, ACS con depósito independiente): Resumen · Cálculo · Trazabilidad · Referencias."""
import sys
sys.path.insert(0, '/home/user/casos/caso-4/herramientas')
from utils_excel import *  # noqa
SALIDA = '/home/user/casos/caso-4/Caso4_RES060_FINAL.xlsx'

wb = Workbook()
rs = wb.active; rs.title = 'Resumen'
cl = wb.create_sheet('Cálculo'); tz = wb.create_sheet('Trazabilidad'); rf = wb.create_sheet('Referencias')

# ============================== CÁLCULO ==============================
banner(cl, 'Cálculo del ahorro energético · Ficha RES060 V1.1 (calefacción y ACS)',
       'Cada resultado es una fórmula enlazada a los datos de entrada. Azul sobre crema = dato introducido; negro = resultado calculado.', 'H')
widths(cl, {'B': 5, 'C': 48, 'D': 11, 'E': 16, 'F': 13, 'G': 46, 'H': 52})
section(cl, 4, '1 · DATOS DE ENTRADA', 'B', 'H')
header(cl, 5, ['Nº', 'Dato', 'Símbolo', 'Valor', 'Unidad', 'Documento de origen', 'Ubicación en el documento'], 'BCDEFGH')
inputs = [
    ('S', 'Superficie útil habitable', 'S', 99.0, 'm²', '#,##0.00', 'D1 · Certificado de eficiencia energética previo', 'p. 2 · Superficie habitable'),
    ('DCAL', 'Demanda de energía en calefacción', 'Dcal', 216.0, 'kWh/m²·año', '#,##0.00', 'D1 · Certificado de eficiencia energética previo', 'Anexo II, apartado 3 · Etiqueta de demanda de calefacción de la situación original (216,0)'),
    ('LD', 'Demanda diaria de ACS a 60 °C', 'DL/D', 100, 'litros/día', '0', 'D1 · Certificado de eficiencia energética previo', 'p. 3 · Instalaciones de ACS · Demanda diaria de ACS a 60° (litros/día)'),
    ('CE', 'Calor específico del agua', 'Ce', 0.001162, 'kWh/kg·°C', '0.000000', 'D11 · Catálogo de medidas estandarizadas · fichas IND040, TER040 y TER100', 'Anexo de cálculo de la demanda de ACS: DACS = DL/D · NP · Ce · 365 · ΔT · Ce = 0,001162 kWh/kg·°C'),
    ('TACS', 'Temperatura de acumulación de ACS', 'T_acs', 60, '°C', '0', 'D11 · Catálogo de medidas estandarizadas · fichas IND040, TER040 y TER100', 'Salto térmico con instalaciones a 60 °C de acumulación'),
    ('TFRIA', 'Temperatura unificada del agua fría de red', 'T_fría', 14, '°C', '0', 'D11 · Catálogo de medidas estandarizadas · fichas IND040, TER040 y TER100', 'ΔT = 60 °C − 14 °C = 46 °C («se unifica la temperatura de agua fría a 14 °C»)'),
    ('SCAL', 'Parte de la demanda (calefacción y ACS) cubierta por la caldera', 'f', 1.0, '—', '0%', 'D1 · Certificado de eficiencia energética previo', 'p. 3 · Único generador de calefacción y ACS · demanda cubierta 100 %'),
    ('ETA', 'Rendimiento de la caldera retirada (referido a PCS)', 'ηi', 0.79, '—', '0.00', 'D6 · Criterios de verificación v14.0 + D7 · Ficha RES210, Anexo II', 'Criterio 24/11.03 · Tabla B.3: gas, de 1998 en adelante = 79 % (valor prudente, superior al 73 % de su tipo exacto)'),
    ('ZONA', 'Zona climática (clima de referencia: cálido)', 'Zona', 'D3', '—', '@', 'D1 · Certificado previo + D5 · Ficha RES060, Anexo III', 'CEE p. 1 · Anexo III, tabla de equivalencia: D3 → clima cálido'),
    ('TIMP', 'Temperatura de impulsión (instalación con radiadores)', 'T', 55, '°C', '0', 'Dato de la instalación', 'Emisores: radiadores → media temperatura (55 °C)'),
    ('ETAS', 'Eficiencia energética estacional en calefacción (55 °C, clima cálido)', 'ηs,h', 178, '%', '0', 'D3 · Registro EPREL 1791526 · Panasonic WH-SDC0509L3E5 / WH-WDG07LE5', 'Eficiencia energética estacional de calefacción, clima cálido, 55 °C: 178 %'),
    ('COP', 'COP a A7/W55 en condiciones UNE-EN 14511', 'COP', 2.98, '—', '0.00', 'D4 · Documentación técnica Panasonic · Aquarea L R290 · kit de 7 kW (unidad exterior WH-WDG07LE5)', 'Heating capacity / COP (A +7 °C, W 55 °C): 7,00 kW / 2,98'),
    ('FC', 'Factor de corrección FC de la zona D3 (Anexo VI, caso 3)', 'FC', 1.113, '—', '0.000', 'D5 · Ficha RES060 V1.1', 'Anexo VI · caso 3 · tabla de factores FC: D3 = 1,113'),
    ('FP', 'Factor de ponderación', 'FP', 1, '—', '0', 'D5 · Ficha RES060 V1.1', 'Apartado 3, tabla de variables'),
    ('CC', 'Coeficiente de conversión a energía primaria de la electricidad', 'CC', 2.5, '—', '0.0', 'D5 · Ficha RES060 V1.1', 'Anexo IV, nota 1'),
    ('F1', 'Factor de corrección F(1) de la bomba aerotérmica', 'F(1)', 0.03, '—', '0%', 'D5 · Ficha RES060 V1.1', 'Anexo IV, nota 2 (F(2) = 0 en aerotermia)'),
]
R = {}
for i, (key, dato, sim, val, uni, fmt, doc, ubi) in enumerate(inputs):
    r = 6 + i
    R[key] = f'$E${r}'
    put(cl, f'B{r}', i + 1, h='center', color=GREYTXT)
    put(cl, f'C{r}', dato); put(cl, f'D{r}', sim, bold=True, h='center')
    put(cl, f'E{r}', val, color=INPUT_BLUE, bg=CREAM, fmt=fmt, h='right', bold=True)
    put(cl, f'F{r}', uni, h='center'); put(cl, f'G{r}', doc); put(cl, f'H{r}', ubi)
    autoheight(cl, r, 'CGH')

r_sec2 = 6 + len(inputs) + 1
section(cl, r_sec2, '2 · CÁLCULO PASO A PASO', 'B', 'H')
header(cl, r_sec2 + 1, ['Nº', 'Concepto', 'Símbolo', 'Resultado', 'Unidad', 'Fórmula', 'Sustitución numérica'], 'BCDEFGH')
S, DC, LD, CE, TACS, TFRIA, SC, ETA, ETAS, COP, FC, FP, CC, F1 = (R[k] for k in
    ('S', 'DCAL', 'LD', 'CE', 'TACS', 'TFRIA', 'SCAL', 'ETA', 'ETAS', 'COP', 'FC', 'FP', 'CC', 'F1'))
p0 = r_sec2 + 2
row = {n: p0 + n for n in range(14)}
E = lambda n: f'E{row[n]}'
steps = [
    ('Demanda anual de calefacción del edificio', 'Dcal·S', f'={DC}*{S}', 'kWh/año', '#,##0.0', 'Dcal × S', f'=ROUND({DC},2)&" × "&ROUND({S},2)&" = "&ROUND({E(0)},1)'),
    ('Demanda anual de ACS del edificio', 'Dacs', f'={LD}*{CE}*365*({TACS}-{TFRIA})', 'kWh/año', '#,##0.0',
     'DL/D × Ce × 365 × (T_acs − T_fría)', f'={LD}&" × "&{CE}&" × 365 × ("&{TACS}&" − "&{TFRIA}&") = "&ROUND({E(1)},2)'),
    ('Demanda de calefacción cubierta por la caldera', 'Dc', f'={E(0)}*{SC}', 'kWh/año', '#,##0.0', 'Dcal·S × f', f'=ROUND({E(0)},1)&" × "&ROUND({SC},2)&" = "&ROUND({E(2)},1)'),
    ('Demanda de ACS cubierta por la caldera', 'Da', f'={E(1)}*{SC}', 'kWh/año', '#,##0.0', 'Dacs × f', f'=ROUND({E(1)},1)&" × "&ROUND({SC},2)&" = "&ROUND({E(3)},1)'),
    ('SCOP de calefacción según Anexo IV', 'SCOP', f'={CC}*({ETAS}/100+{F1})', '—', '0.000', 'CC × (ηs,h + F(1))', f'={CC}&" × ("&{ETAS}/100&" + "&{F1}&") = "&ROUND({E(4)},4)'),
    ('SCOP de ACS según Anexo VI, caso 3', 'SCOPdhw', f'={COP}*{FC}', '—', '0.000', 'COP A7/W55 × FC', f'={COP}&" × "&{FC}&" = "&ROUND({E(5)},4)'),
    ('Consumo de calefacción con la caldera retirada', 'Ec,i', f'={E(2)}/{ETA}', 'kWh/año', '#,##0.0', 'Dc ÷ ηi', f'=ROUND({E(2)},1)&" ÷ "&{ETA}&" = "&ROUND({E(6)},1)'),
    ('Consumo de calefacción con la bomba de calor', 'Ec,f', f'={E(2)}/{E(4)}', 'kWh/año', '#,##0.0', 'Dc ÷ SCOP', f'=ROUND({E(2)},1)&" ÷ "&ROUND({E(4)},3)&" = "&ROUND({E(7)},1)'),
    ('Ahorro anual en calefacción', 'AEcal', f'={E(6)}-{E(7)}', 'kWh/año', '#,##0.0', 'Ec,i − Ec,f', f'=ROUND({E(6)},1)&" − "&ROUND({E(7)},1)&" = "&ROUND({E(8)},1)'),
    ('Consumo de ACS con la caldera retirada', 'Ea,i', f'={E(3)}/{ETA}', 'kWh/año', '#,##0.0', 'Da ÷ ηi', f'=ROUND({E(3)},1)&" ÷ "&{ETA}&" = "&ROUND({E(9)},1)'),
    ('Consumo de ACS con la bomba de calor', 'Ea,f', f'={E(3)}/{E(5)}', 'kWh/año', '#,##0.0', 'Da ÷ SCOPdhw', f'=ROUND({E(3)},1)&" ÷ "&ROUND({E(5)},3)&" = "&ROUND({E(10)},1)'),
    ('Ahorro anual en ACS', 'AEacs', f'={E(9)}-{E(10)}', 'kWh/año', '#,##0.0', 'Ea,i − Ea,f', f'=ROUND({E(9)},1)&" − "&ROUND({E(10)},1)&" = "&ROUND({E(11)},1)'),
    ('Ahorro anual de energía final', 'AE', f'={FP}*({E(8)}+{E(11)})', 'kWh/año', '#,##0.0', 'FP × (AEcal + AEacs)', f'={FP}&" × ("&ROUND({E(8)},1)&" + "&ROUND({E(11)},1)&") = "&ROUND({E(12)},2)'),
    ('Certificados de ahorro energético (1 CAE = 1 kWh)', 'CAE', f'=ROUND({E(12)},0)', 'CAE', '#,##0', 'CAE = AE (redondeado)', f'="AE = "&ROUND({E(12)},2)&" kWh/año → "&TEXT({E(13)},"0")&" CAE"'),
]
for n, (concepto, sim, formula, uni, fmt, fx, sub) in enumerate(steps):
    r = row[n]; last = n == 13
    bg_res = 'FFF2CC' if last else (GREEN if n in (8, 11, 12) else None)
    put(cl, f'B{r}', n + 1, h='center', color=GREYTXT)
    put(cl, f'C{r}', concepto, bold=(n >= 12)); put(cl, f'D{r}', sim, bold=True, h='center')
    put(cl, f'E{r}', formula, fmt=fmt, h='right', bold=(n >= 12), bg=bg_res, size=(12 if last else 10))
    put(cl, f'F{r}', uni, h='center'); put(cl, f'G{r}', fx); put(cl, f'H{r}', sub, color=GREYTXT)
    autoheight(cl, r, 'CGH', minh=(24 if last else (32 if n == 1 else 18)))
r_note = row[13] + 1
cl.merge_cells(f'B{r_note}:H{r_note}')
put(cl, f'B{r_note}',
    'Fórmula de la ficha: AE = FP · [ (Dcal · S) · (1/ηi − 1/SCOP) + Dacs · (1/ηi − 1/SCOPdhw) ]. Calefacción y ACS: la caldera de gas producía los dos servicios. '
    'El interacumulador de ACS de 100 l de la factura no consta como conjunto suministrado por el fabricante de la bomba, por lo que el SCOPdhw se obtiene con el caso 3 del Anexo VI. Sin multiplicadores.',
    italic=True, color=GREYTXT, border=False, v='top')
cl.row_dimensions[r_note].height = 46
r_sec3 = r_note + 2
section(cl, r_sec3, '3 · CONTROLES', 'B', 'H')
header(cl, r_sec3 + 1, ['Nº', 'Control', '', 'Valor', 'Unidad', 'Qué comprueba', 'Resultado'], 'BCDEFGH')
c1, c2 = r_sec3 + 2, r_sec3 + 3
put(cl, f'B{c1}', 1, h='center', color=GREYTXT); put(cl, f'C{c1}', 'Ahorro sobre el consumo previo de la caldera'); put(cl, f'D{c1}', '')
put(cl, f'E{c1}', f'={E(12)}/({E(6)}+{E(9)})', fmt='0.0%', h='right'); put(cl, f'F{c1}', '%', h='center')
put(cl, f'G{c1}', 'El ahorro no puede superar el consumo que se evita (criterio 25-07.03)')
put(cl, f'H{c1}', f'=IF({E(12)}<{E(6)}+{E(9)},"Correcto: el ahorro es menor que el consumo previo","Revisar: el ahorro supera el consumo previo")')
put(cl, f'B{c2}', 2, h='center', color=GREYTXT); put(cl, f'C{c2}', 'Demanda de ACS calculada con el Anejo G del DB HE (Albacete, 100 l/día)'); put(cl, f'D{c2}', '')
put(cl, f'E{c2}', 2015.2, color=INPUT_BLUE, bg=CREAM, fmt='#,##0.0', h='right', bold=True); put(cl, f'F{c2}', 'kWh/año', h='center')
put(cl, f'G{c2}', 'Contraste de la demanda de ACS con el método del CTE (temperaturas de red de la capital, 60 °C)')
put(cl, f'H{c2}', f'="Diferencia con el paso 2: "&ROUND({E(1)}-E{c2},1)&" kWh ("&ROUND(({E(1)}-E{c2})/E{c2}*100,2)&" %)"')
for r in (c1, c2):
    autoheight(cl, r, 'CGH')
cl.freeze_panes = 'A4'; cl.sheet_properties.tabColor = BLUE; page(cl)
CELL = {k: f"Cálculo!$E${row[n]}" for k, n in (('D', 0), ('DA', 1), ('DC', 2), ('DAC', 3), ('SCOP', 4), ('SCOPDHW', 5), ('ECI', 6), ('ECF', 7), ('AECAL', 8), ('EAI', 9), ('EAF', 10), ('AEACS', 11), ('AE', 12), ('CAE', 13))}

# ============================== RESUMEN ==============================
banner(rs, 'Certificados de Ahorro Energético (CAE) · Ficha RES060 V1.1',
       'Sustitución de caldera de gas por bomba de calor aerotérmica (calefacción y ACS) · Caso 4', 'G')
widths(rs, {'B': 30, 'C': 20, 'D': 20, 'E': 20, 'F': 20, 'G': 20})
section(rs, 4, 'IDENTIFICACIÓN DE LA ACTUACIÓN', 'B', 'G')
ident = [
    ('Emplazamiento', 'Vivienda caso 4 (Albacete) · vivienda en bloque construida en 2000 · zona climática D3'),
    ('Referencia catastral', '[ref. catastral omitida]'),
    ('Actuación', 'Sustitución de la caldera mural de gas propano Ferroli Domina F 24 E (24 kW; calefacción y ACS) por bomba de calor aerotérmica Panasonic Aquarea L de 7 kW con R290 '
                  '(unidad exterior WH-WDG07LE5, nº de serie [nº serie omitido]; unidad interior WH-SDC0509L3E5) e interacumulador de ACS de 100 l; emisores: radiadores'),
    ('Certificado de eficiencia energética previo', 'CE3X v2.3 · visita 27/05/2024 · firmado el 27/05/2024'),
    ('Servicio y método', 'Calefacción y ACS · ficha RES060 V1.1 · sin efecto multiplicador'),
]
for i, (k, v) in enumerate(ident):
    r = 5 + i
    put(rs, f'B{r}', k, bold=True, color=NAVY, bg=PALE)
    rs.merge_cells(f'C{r}:G{r}'); put(rs, f'C{r}', v)
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
rs.row_dimensions[12].height = 24; rs.row_dimensions[13].height = 24
section(rs, 15, 'CÓMO SE OBTIENE', 'B', 'G')
header(rs, 16, ['Concepto', 'Calefacción', 'ACS', 'Total', 'Unidad', 'Se calcula como'], 'BCDEFG')
flow = [
    ('Demanda cubierta por la caldera', CELL['DC'], CELL['DAC'], None, 'Dcal·S y Dacs'),
    ('Consumo con la caldera retirada (ηi = 0,79)', CELL['ECI'], CELL['EAI'], None, 'Demanda ÷ rendimiento de la caldera'),
    ('Consumo con la bomba de calor', CELL['ECF'], CELL['EAF'], None, 'Demanda ÷ SCOP (4,525) y ÷ SCOPdhw (3,317)'),
    ('Ahorro anual de energía final', CELL['AECAL'], CELL['AEACS'], CELL['AE'], 'Consumo antes − consumo después'),
]
for i, (a, c1_, c2_, tot, how) in enumerate(flow):
    r = 17 + i; last = i == 3; bgc = GREEN if last else None
    put(rs, f'B{r}', a, bold=last, bg=bgc)
    put(rs, f'C{r}', f'={c1_}', fmt='#,##0.0', h='right', bold=last, bg=bgc)
    put(rs, f'D{r}', f'={c2_}', fmt='#,##0.0', h='right', bold=last, bg=bgc)
    put(rs, f'E{r}', (f'={tot}' if tot else f'=C{r}+D{r}'), fmt='#,##0.0', h='right', bold=last, bg=bgc)
    put(rs, f'F{r}', 'kWh/año', h='center', bg=bgc); put(rs, f'G{r}', how, color=GREYTXT, bg=bgc)
    rs.row_dimensions[r].height = 32
section(rs, 22, 'DATOS DE PARTIDA', 'B', 'G')
header(rs, 23, ['Dato', 'Valor', 'Unidad', 'Documento', '', ''], 'BCDEFG')
rs.merge_cells('E23:G23')
base = [
    ('Superficie útil', 'S', 'm²', '#,##0.00', 'D1 · Certificado previo'),
    ('Demanda de calefacción', 'DCAL', 'kWh/m²·año', '#,##0.00', 'D1 · Certificado previo'),
    ('Rendimiento de la caldera', 'ETA', '—', '0.00', 'D6 · Criterio 24/11.03 · D7 · Ficha RES210, Anexo II'),
    ('Eficiencia estacional de calefacción (ηs,h)', 'ETAS', '%', '0', 'D3 · Registro EPREL 1791526'),
    ('COP a A7/W55', 'COP', '—', '0.00', 'D4 · Documentación técnica Panasonic'),
]
for i, (a, key, uni, fmt, doc) in enumerate(base):
    r = 24 + i
    put(rs, f'B{r}', a); put(rs, f'C{r}', f'=Cálculo!{R[key]}', fmt=fmt, h='right'); put(rs, f'D{r}', uni, h='center')
    rs.merge_cells(f'E{r}:G{r}'); put(rs, f'E{r}', doc)
    for col in 'FG':
        rs[f'{col}{r}'].border = BOX
    rs.row_dimensions[r].height = 20
section(rs, 30, 'CONTENIDO DEL LIBRO', 'B', 'G')
guia = [('Cálculo', 'Datos de entrada con su documento de origen, cálculo paso a paso con fórmulas y controles.'),
        ('Trazabilidad', 'De qué documento y de qué página sale cada dato, y cómo se ha comprobado.'),
        ('Referencias', 'Documentos consultados (D1 a D14), normativa y criterios aplicados, y datos del fabricante.')]
for i, (a, b) in enumerate(guia):
    r = 31 + i
    put(rs, f'B{r}', a, bold=True, color=NAVY, bg=PALE); rs.merge_cells(f'C{r}:G{r}'); put(rs, f'C{r}', b)
    for col in 'DEFG':
        rs[f'{col}{r}'].border = BOX
    rs.row_dimensions[r].height = 20
rs.merge_cells('B35:G35')
put(rs, 'B35', 'Leyenda: azul sobre fondo crema = dato introducido · negro = resultado calculado con fórmula · fondo verde / amarillo = resultado.', italic=True, color=GREYTXT, border=False, size=9)
rs.sheet_properties.tabColor = NAVY; page(rs, landscape=False)

# ============================== TRAZABILIDAD ==============================
banner(tz, 'Trazabilidad de los datos',
       'Para cada dato del cálculo: valor usado, documento (códigos D1–D14, ver hoja Referencias), ubicación exacta y comprobación realizada.', 'G')
widths(tz, {'B': 30, 'C': 18, 'D': 9, 'E': 44, 'F': 46, 'G': 58})
section(tz, 4, 'A · DATOS DEL CÁLCULO Y SU DOCUMENTO DE ORIGEN', 'B', 'G')
header(tz, 5, ['Dato', 'Valor usado', 'Doc.', 'Ubicación en el documento', 'Comprobación', 'Observaciones'], 'BCDEFG')
tr = [
    ('Superficie útil habitable (S)', '99,0 m²', 'D1', 'p. 2 · Superficie habitable', 'Coincide con el certificado posterior (D2, 99,0 m²).', 'Superficie de la vivienda.'),
    ('Demanda de calefacción (Dcal)', '216,0 kWh/m²·año', 'D1, D12', 'D1: Anexo II, apartado 3 · etiqueta de demanda de calefacción (216,0)',
     'Comprobación: EP de calefacción 336,03 ÷ 1,201 × 0,772 = 216,0 kWh/m²·año (factor de paso del GLP, D12; rendimiento del certificado, p. 3).',
     'Ficha RES060, apartado 3 y nota 5: demanda según el certificado de eficiencia energética antes de la actuación.'),
    ('Demanda de ACS (Dacs)', '1.951 kWh/año', 'D1, D11',
     'D1: p. 3 · demanda diaria de ACS a 60 °C = 100 l/día · D11: DACS = DL/D · NP · Ce · 365 · ΔT, con Ce = 0,001162 y ΔT = 60 − 14 = 46 °C (fichas IND040 Anexo III, TER040 Anexo VI, TER100 Anexo V)',
     '100 × 0,001162 × 365 × 46 = 1.951,0 kWh/año (19,51 kWh por litro/día). Contraste con el Anejo G del DB HE (D13, temperaturas de red de la capital): 2.015 kWh/año.',
     'El certificado no indica la demanda de ACS en kWh/año. Se aplica el método de cálculo del catálogo de medidas estandarizadas, con temperatura de agua fría unificada a 14 °C, por analogía con las fichas citadas.'),
    ('Parte cubierta por la caldera (f)', '100 %', 'D1', 'p. 3 · Un único generador «Calefacción y ACS»',
     'La máquina frigorífica del certificado es solo de refrigeración; no hay otros generadores de calefacción ni de ACS.', 'La caldera cubría toda la calefacción y todo el ACS.'),
    ('Rendimiento de la caldera (ηi)', '0,79', 'D6, D7, D10',
     'D6: criterio 24/11.03 · D7: Anexo II, tabla B.3, filas del gas de 1998 en adelante (69 %, 73 %, 79 % y 83 %)',
     'El criterio 24/11.03 cita expresamente la RES060 y permite usar los valores de los anexos de RES210 y TER210. La Ferroli Domina F 24 E es una caldera mural de gas (GLP, quemador regulado a G31), de flujo forzado y cámara estanca, con encendido y control de llama electrónicos (D10, D8); el certificado previo la define como caldera estándar (D1, p. 3).',
     'El nº de serie 9946… y el impreso de instrucciones (10/99) sitúan la fabricación en 1999 (D8). Se adopta el 79 %, superior al 73 % que correspondería a su tipo exacto (sin condensación, con encendido automático): criterio prudente, porque da un ahorro menor.'),
    ('Zona climática y clima de referencia', 'D3 · clima cálido', 'D1, D5', 'D1: p. 1 · D5: Anexo III, tabla de equivalencia (D3 → cálidas)',
     'El SCOP debe ser, al menos, el de clima medio «o el indicado para la zona climática equivalente»; para D3 la tabla indica clima cálido.', 'Determina la eficiencia estacional del fabricante que se utiliza.'),
    ('Temperatura de impulsión', '55 °C', '—', 'Dato de la instalación (radiadores)',
     'Ficha RES060, Anexo III: se usa el SCOP a la temperatura de impulsión necesaria. Radiadores → media temperatura (55 °C).', 'La factura describe la instalación de calefacción con tubería multicapa (D9).'),
    ('Eficiencia estacional en calefacción (ηs,h)', '178 %', 'D3, D4', 'D3: Registro EPREL 1791526 · Panasonic WH-SDC0509L3E5 / WH-WDG07LE5 · ηs clima cálido a 55 °C = 178 %',
     'Coincide con la documentación técnica de Panasonic (D4, manual de servicio, apartado 3.2: SCOP 4,52 / ηs 178 %). El Anexo IV da 2,5 × (1,78 + 0,03) = 4,525.',
     'Registro publicado (versión 3, 24/07/2026) del kit de 7 kW con la unidad exterior de la placa (D8). Temperatura de impulsión de 55 °C (radiadores).'),
    ('COP a A7/W55 (EN 14511)', '2,98', 'D4', 'Folleto Aquarea L · KIT-ADC07L3E5 (unidad exterior WH-WDG07LE5) · Heating capacity / COP (A +7 °C, W 55 °C) = 7,00 kW / 2,98',
     'Para el kit con unidad interior WH-SDC0509L3E5 no se publica el COP a A7/W55; se toma el del kit ADC de 7 kW, que usa la misma unidad exterior y declara las mismas ηs en EPREL (195 % / 142 %).',
     'Necesario para el SCOPdhw (Anexo VI, caso 3). La placa de la unidad exterior da COP 4,93 a A7/W35 (D8), igual que D4.'),
    ('SCOP de ACS (SCOPdhw)', '3,317', 'D5', 'Anexo VI · caso 3 · SCOPdhw = COP(A7/W55) × FC · FC(D3) = 1,113',
     'El ACS se calienta en un interacumulador de 100 l (D9) del que no consta que se suministre como conjunto con la bomba por su fabricante; el EPREL del kit es de calefacción y no declara eficiencia de ACS.',
     'Caso 3 del Anexo VI (depósitos no suministrados como conjunto).'),
    ('Equipo instalado', 'Panasonic Aquarea L 7 kW R290', 'D8, D9', 'Placa de la unidad exterior WH-WDG07LE5 (serie [nº serie omitido], producción 2023.05) · factura: kit aerotermia bibloc Panasonic Aquarea 7 kW R290',
     'La factura anota la unidad exterior WH-WDG07LE5 (SN [nº serie omitido]) y la unidad interior WH-SDC0509L3E5 (SN [nº serie omitido]).', 'Bomba aire-agua bibloc con R290: unidad exterior y unidad interior mural.'),
    ('Caldera sustituida', 'Ferroli Domina F 24 E', 'D8, D10', 'Etiqueta de características y tapa de la caldera (nº de serie [nº serie omitido])',
     'Potencia nominal PCS 28,0 kW (PCI 25,8 kW), potencia útil 23,8 kW; GLP (quemador regulado a G31). El certificado previo indica 24 kW (D1, p. 3).', 'Caldera mural de gas propano.'),
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
    ('Certificado de eficiencia energética posterior', 'PDF', 'D2', 'Misma superficie y zona; bomba de calor para calefacción y ACS (modelada).'),
    ('Factura 24PV 000073 (22/08/2024)', 'PDF', 'D9', 'Kit de aerotermia bibloc Panasonic Aquarea 7 kW R290, interacumulador de ACS de 100 l, vaso de expansión, silentblocks e instalación de calefacción con multicapa.'),
    ('Fotografías de la caldera sustituida', 'Imagen', 'D8', 'Ferroli Domina F 24 E: etiqueta de características (potencias, gas, categoría II2H3P, quemador regulado G31) y tapa con el modelo y el nº de serie.'),
    ('Fotografías de la nueva instalación', 'Imagen', 'D8', 'Placa de la unidad exterior Panasonic WH-WDG07LE5 (R290, 7,00 kW, COP 4,93 a A7/W35) y unidad interior mural Aquarea.'),
    ('Manual de la caldera Ferroli Domina F 24 E', 'PDF', 'D10', 'Caldera mural a gas, de cobre, flujo forzado, cámara estanca, con encendido y control de llama electrónicos.'),
]
for i, (a, b, c, d) in enumerate(evid):
    r = r3 + 2 + i
    put(tz, f'B{r}', a, bold=True, bg=PALE); put(tz, f'C{r}', b, h='center'); put(tz, f'D{r}', c, h='center')
    tz.merge_cells(f'E{r}:G{r}'); put(tz, f'E{r}', d)
    for col in 'FG':
        tz[f'{col}{r}'].border = BOX
    tz.row_dimensions[r].height = max(20, 13.2 * lines(d, 140) + 6, 13.2 * lines(a, 30) + 6)
tz.freeze_panes = 'A4'; tz.sheet_properties.tabColor = 'C55A11'; page(tz)

# ============================== REFERENCIAS ==============================
banner(rf, 'Referencias', 'Documentos consultados, normativa y criterios aplicados, y datos del fabricante.', 'F')
widths(rf, {'B': 7, 'C': 46, 'D': 30, 'E': 62, 'F': 44})
section(rf, 4, 'A · DOCUMENTOS', 'B', 'F')
header(rf, 5, ['Cód.', 'Documento', 'Emisor · fecha', 'Identificación / enlace', 'Datos que se toman'], 'BCDEF')
docs = [
    ('D1', 'Certificado de eficiencia energética previo a la actuación (CE3X v2.3)', 'INER Smart Energy Systems SL · 27/05/2024', 'Ref. catastral [ref. catastral omitida] · visita 27/05/2024', 'S, Dcal, zona climática, generador, ACS, energía primaria y emisiones'),
    ('D2', 'Certificado de eficiencia energética posterior a la actuación (CE3X)', 'INER Smart Energy Systems SL · 30/12/2024', 'Ref. catastral [ref. catastral omitida] · visita 30/12/2024', 'Contraste de la superficie y la zona'),
    ('D3', 'Registro EPREL 1791526 · Panasonic · WH-SDC0509L3E5 / WH-WDG07LE5 (calefactor con bomba de calor)', 'Comisión Europea · versión 3, publicado el 24/07/2026', 'https://eprel.ec.europa.eu/screen/product/spaceheaters/1791526', 'ηs de calefacción por clima y temperatura'),
    ('D4', 'Documentación técnica Panasonic Aquarea L (R290): folleto de la gama y manual de servicio de la serie Hydrobox, apartado 3.2 (WH-SDC0509L3E5 + WH-WDG07LE5)', 'Panasonic',
     'https://www.panasonicproclub.com/uploads/NL/catalogues/Aquarea%20L%20Brochure%20NL%20LR.pdf', 'COP a A7/W35 y A7/W55, SCOP y ηs (contraste con EPREL)'),
    ('D5', 'Ficha RES060 V1.1 y Anexos II a VI', 'MITECO · catálogo vigente de fichas (residencial)', 'https://www.miteco.gob.es/es/energia/eficiencia/cae/catalogo-de-fichas/catalogo-vigente-de-fichas/residencial.html', 'Fórmula, variables, SCOP, SCOPdhw, zonas y clima'),
    ('D6', 'Criterios de verificación del Sistema CAE v14.0', 'MITECO · 16/10/2026', 'Criterios 24/11.01, 24/11.03, 25-07.03', 'Redondeos, rendimiento de la caldera, límite físico'),
    ('D7', 'Ficha RES210 · Anexo II «Rendimiento de caldera sustituida» (idéntico en la ficha TER210)', 'MITECO · catálogo vigente',
     'https://www.miteco.gob.es/content/dam/miteco/es/energia/files-1/Eficiencia/CAE/catalogo-vigente/fichas-autocompletables/residencial/Ficha%20RES210_Anexo%20II.pdf', 'Tabla B.3: rendimiento estacional por defecto de la caldera'),
    ('D8', 'Fotografías de la caldera sustituida y de los equipos instalados', 'INER Smart Energy Systems SL', 'Fotos 1 a 4 del expediente', 'Identificación de los equipos'),
    ('D9', 'Factura 24PV 000073', 'INER Smart Energy Systems SL · 22/08/2024', 'Instalación de aerotermia · Albacete', 'Equipos instalados e interacumulador de ACS'),
    ('D10', 'Manual de instalación y uso de la caldera Ferroli Domina F 24 E', 'Ferroli', 'Caldera mural a gas, de cobre, flujo forzado, cámara estanca, con encendido y control de llama electrónicos', 'Tipo de caldera'),
    ('D11', 'Catálogo de medidas estandarizadas de eficiencia energética · fichas IND040 (Anexo III), TER040 (Anexo VI) y TER100 (Anexo V)', 'MITECO · Orden TED/845/2023 y actualizaciones',
     'Cálculo de la demanda anual de ACS, según el Anejo F del DB HE 2022', 'DACS = DL/D · NP · Ce · 365 · ΔT; Ce = 0,001162 kWh/kg·°C; ΔT = 60 − 14 = 46 °C'),
    ('D12', 'Factores de emisión de CO₂ y coeficientes de paso a energía primaria de diferentes fuentes de energía final (documento reconocido)', 'MITECO · IDAE', 'Documento reconocido para la certificación energética', 'Factor del GLP: 1,201 kWh EP no renovable / kWh'),
    ('D13', 'CTE DB-HE · Anejo G · temperaturas de agua de red', 'Ministerio de Vivienda', 'Anejo G · capital de provincia: Albacete', 'Contraste de la demanda de ACS'),
    ('D14', 'Preguntas frecuentes del Sistema CAE v1.12', 'MITECO · RD 36/2023', 'Valor del CAE y duración del ahorro', '1 CAE = 1 kWh de ahorro anual'),
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
    ('N2', 'Ficha RES060 V1.1', 'Apartado 3 · tabla de variables', 'FP = 1; ηi referido a PCS; Dcal, S y Dacs según el certificado de eficiencia energética antes de la actuación.', 'Datos de entrada 1, 2, 3, 15'),
    ('N3', 'Ficha RES060 V1.1', 'Anexo III', 'Se usa el SCOP del fabricante a la temperatura de impulsión; el SCOP debe ser, al menos, el de las condiciones de clima medio, o el de la zona equivalente (D3 → cálido).', 'Datos de entrada 9, 10, 11'),
    ('N4', 'Ficha RES060 V1.1', 'Anexo IV', 'SCOP = CC · (ηs,h + F(1) + F(2)); CC = 2,5; F(1) = 3 % en aerotermia; F(2) = 0.', 'Paso 5'),
    ('N5', 'Ficha RES060 V1.1', 'Anexo V', 'Hasta 70 kW, los rendimientos son los de EPREL o de la ficha técnica del fabricante.', 'Datos ηs,h y COP'),
    ('N6', 'Ficha RES060 V1.1', 'Anexo VI · caso 3', 'Bomba aerotérmica y depósito no suministrados como conjunto: SCOPdhw = COP(A7/W55) × FC, con FC por zona climática del CTE (D3 = 1,113).', 'Paso 6'),
    ('N7', 'Criterios de verificación v14.0', 'Criterio 24/11.01', 'Utilizar todas las cifras significativas; el error de redondeo no puede superar el 1 %.', 'Todo el cálculo (sin redondeos intermedios)'),
    ('N8', 'Criterios de verificación v14.0', 'Criterio 24/11.03', 'En las fichas de sustitución de caldera por bomba de calor (incluida RES060) se pueden usar los rendimientos de los anexos de las fichas RES210 y TER210.', 'Dato ηi'),
    ('N9', 'Ficha RES210 (y TER210)', 'Anexo II · tabla B.3', 'Rendimiento estacional por defecto del gas: de 1998 en adelante sin condensación, con piloto permanente = 69 %, con encendido automático = 73 %; con condensación, con piloto = 79 %, con encendido automático = 83 %.', 'Dato ηi'),
    ('N10', 'Criterios de verificación v14.0', 'Criterio 25-07.03', 'Límite físico: el ahorro no puede superar una proporción razonable del consumo que se evita.', 'Control 1'),
    ('N11', 'FAQ del Sistema CAE v1.12', 'Valor del CAE', '1 CAE equivale a 1 kWh de ahorro anual de energía final.', 'Paso 14'),
    ('N12', 'Catálogo de medidas estandarizadas (IND040, TER040, TER100)', 'Anexo de demanda de ACS', 'DACS = DL/D · NP · Ce · 365 · ΔT, con Ce = 0,001162 kWh/kg·°C y ΔT = 60 °C − 14 °C = 46 °C; el técnico responsable puede proponer cálculos alternativos. Se aplica a la RES060 por analogía.', 'Paso 2'),
]
for i, fila in enumerate(norm):
    r = r2 + 2 + i
    for col, val in zip('BCDEF', fila):
        put(rf, f'{col}{r}', val, bold=(col == 'B'), h=('center' if col == 'B' else 'left'), bg=(PALE if col == 'B' else None), v='top')
    autoheight(rf, r, 'CDEF', minh=24)
r3 = r2 + 2 + len(norm) + 1
section(rf, r3, 'C · KIT PANASONIC AQUAREA L 7 kW · EPREL 1791526 (D3) Y DOCUMENTACIÓN TÉCNICA (D4)', 'B', 'F')
header(rf, r3 + 1, ['', 'Magnitud', 'A7/W35 (35 °C)', 'A7/W55 (55 °C)', 'Se utiliza'], 'BCDEF')
mag = [
    ('ηs,h clima medio (%)', 195, 142, 'No se utiliza (D3 → clima cálido)'),
    ('ηs,h clima cálido (%)', 249, 178, 'Radiadores (55 °C), clima cálido → 178 %'),
    ('COP (EN 14511)', 4.93, 2.98, 'COP A7/W55 = 2,98 → SCOPdhw = 2,98 × 1,113'),
]
for i, (m, a, b, u) in enumerate(mag):
    r = r3 + 2 + i
    put(rf, f'B{r}', ''); put(rf, f'C{r}', m, bold=True, bg=PALE)
    put(rf, f'D{r}', a, color=INPUT_BLUE, h='center', fmt='0.00' if isinstance(a, float) else '0')
    put(rf, f'E{r}', b, color=INPUT_BLUE, h='center', fmt='0.00' if isinstance(b, float) else '0', bold=(i in (1, 2)), bg=(GREEN if i in (1, 2) else None))
    put(rf, f'F{r}', u); autoheight(rf, r, 'CF', minh=22)
rf.freeze_panes = 'A4'; rf.sheet_properties.tabColor = '548235'; page(rf)
wb.active = 0
wb.save(SALIDA)
print('guardado', SALIDA)
