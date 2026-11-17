# Generador de Informes de Recuperos por Líquido Producto

Script para generar informes Excel mensuales de recuperos a proveedores (comitentes) a partir del archivo mensual de ventas de consignación.

## Archivos

- `generar_informes.py` — script principal
- `Ventas <MES> Consignación.xlsx` — archivo de entrada (uno por mes)
- `Informes <Mes Año>/` — carpeta de salida (se crea automáticamente)

## Uso

Desde la carpeta `RecuperosLP`:

```powershell
python generar_informes.py
```

Si hay un solo `.xlsx` en la carpeta, lo toma automáticamente. Si hay más de uno, pasarlo como argumento:

```powershell
python generar_informes.py "Ventas NOV Consignación.xlsx"
```

El período (ej. "Octubre 2026") se detecta a partir de las fechas en `Control Promos`.

**Importante**: cerrar cualquier Excel abierto de la carpeta de salida antes de correr, si no, Windows da `PermissionError`.

## Qué lee

Del archivo de entrada usa tres hojas:

- **Control Promos** — fuente principal. Cada fila es un descuento aplicado (fecha, local, SKU, promo, importe, cargo).
- **Ventas** — ventas por ticket/caja. Se filtra sólo por los SKUs que aparecen en Control Promos.
- **Maestro 11-11** — fallback para descripciones de productos que no aparecieron en Ventas.

## Lógica de cálculo

Columna `Cargo` en Control Promos:

| Cargo | Significado | Recupero del proveedor |
|-------|-------------|------------------------|
| `P`   | Proveedor   | 100% del descuento (`-Importe SUM`) |
| `C`   | Compartido  | `-Importe SUM × (Cargo P / (Cargo DP + Cargo P))` |
| `D`   | Duty paid (DP) | 0% — no aparece en los informes |

Ejemplo: promo 40% off con `Cargo DP = 0.15` y `Cargo P = 0.25` → proveedor absorbe `0.25 / 0.40 = 62.5%` del descuento total.

Importes **sin IVA** (la columna `Iva SUM` se ignora).

## Qué genera

Dentro de `Informes <Mes Año>/`:

### 1. Un Excel por proveedor

Archivo: `Recupero <Razón Social> - <Mes Año>.xlsx`

Tres hojas:

- **Resumen** — totales separados por Cargo P y C, con descuento aplicado al cliente y recupero al proveedor.
- **Detalle Recuperos** — línea por línea: fecha, local, SKU, producto, promoción, cargo, % Cargo DP, % Cargo P, % recupero proveedor, descuento total, a recuperar.
- **Ventas SKU x Local** — unidades e importe vendido por SKU y local, sólo de los SKUs que tuvieron promo en el período.

### 2. Un Excel consolidado para ADM

Archivo: `Recuperos Consolidado ADM - <Mes Año>.xlsx`

Dos hojas:

- **Resumen ADM** — una fila por proveedor (ordenado de mayor a menor recupero), con descuento y recupero separados por P / C y total.
- **Detalle completo** — todas las líneas de todos los proveedores, con autofiltro y paneles congelados.

## Mapeo de locales

Los códigos de local se traducen a nombre legible en toda la salida:

| Código | Nombre     |
|--------|------------|
| D07    | Aeroparque |
| D10    | Mendoza    |
| D14    | Tucumán    |
| D20    | Bariloche  |
| D25    | Iguazú     |
| D30    | Córdoba    |
| D35    | Salta      |
| WEB    | WEB        |

Si aparece un local nuevo, agregarlo al diccionario `LOCALES` arriba en `generar_informes.py`.

## Chequeos / spot-checks realizados en octubre 2026

- 71 proveedores con cargo P o C.
- Total descuento aplicado: **$119.004.306,31**
- Total a recuperar: **$111.462.583,26** (la diferencia es la porción DP de las promos compartidas).
- Verificado en GRAPPOLO SA: promo 40% Off `C` con DP=0.15 / P=0.25 → descuento $15.412 × 0.625 = $9.632,50 de recupero. ✓

## Notas / cosas a revisar manualmente

- Puede haber comitentes con **total negativo** (ej. FYN 5 SRL / BOTIK: -$5.000 en octubre). Suelen ser ajustes o devoluciones del mes; el script los deja tal cual para que ADM los revise.
- El cruce con Ventas se hace por SKU (no por nombre de comitente) para evitar problemas de matching entre el código (`QUEBR`) y la razón social (`SABATE SANTIAGO`).
