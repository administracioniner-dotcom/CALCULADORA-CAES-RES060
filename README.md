# Calculadora CAE · Ficha RES060 V1.1

Cálculo de Certificados de Ahorro Energético (CAE) para la ficha MITECO **RES060 V1.1**: sustitución de caldera por bomba de calor aerotérmica aire-agua.

```
AE = FP · [ (Dcal · S) · (1/ηi − 1/SCOP) + Dacs · (1/ηi − 1/SCOPdhw) ] · f        1 CAE = 1 kWh/año
```

## Contenido

Cada caso incluye:

- `CasoN_RES060_FINAL.xlsx`: libro de 4 hojas (Resumen · Cálculo · Trazabilidad · Referencias) con fórmulas vivas y la referencia documental de cada dato.
- `herramientas/`: scripts Python (openpyxl) que generan el libro.

| Caso | Provincia | CAE |
|---|---|---|
| 1 | Córdoba | 12.213 |
| 2 | Burgos | 8.931 |
| 3 | Albacete | 43.780 |
| 4 | Albacete | 24.224 |

Los documentos originales (certificados, facturas, fotografías) no se incluyen; direcciones, referencias catastrales y números de serie se han omitido.
