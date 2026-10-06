# Mapa de las planillas de Los Cántaros

Contenido:
1. Dónde está cada archivo
2. La hoja Cartola: columnas y encabezado
3. Fechas: los tres formatos que conviven
4. Estado de Gastos Comunes: dónde sale el saldo actual
5. Historia de la cuota mensual
6. Trampas conocidas (léelas antes de dar cifras por buenas)

---

## 1. Dónde está cada archivo

Raíz: `C:\Users\rheis\OneDrive\Trabajo\Los Cántaros - ONE DRIVE\GASTOS COMUNES\`

```
Gastos Comunes año <AAAA>\
├── Cartola Mensual Gastos Comunes Año <AAAA>.xlsx   ← el libro mayor de movimientos
├── Estados Gastos Comunes año <AAAA>.xlsx           ← las hojas de cobro mensuales
├── GCLC <MM>-<AAAA>.pdf                             ← los cobros ya emitidos
└── Cartolas\                                        ← cartolas del banco sin procesar (solo 2026)
```

Hay carpetas desde 2012. Qué esperar según el año:

| Años | Estado |
|------|--------|
| 2016 – 2026 | Misma estructura de cartola. Verificado. |
| 2022 – 2026 | Además verificadas las observaciones y los códigos de gasto. |
| 2015 | La hoja tiene estructura de **resumen**, no de cartola de movimientos. |
| 2012 – 2014 | Nombres de archivo distintos (`Cartola Gastos Comunes año 2013.xlsx`, `Cartola Mensual Gastos Comunes 2014 4.0.xlsx`) y estructura variable. Lista la carpeta y revisa la hoja antes de filtrar. |

Cuando el patrón de nombre no calce, lista la carpeta antes de rendirte.

---

## 2. La hoja Cartola: columnas y encabezado

El nombre de la hoja es `Cartola <AAAA>`, pero **no siempre es la primera hoja
del libro** — en 2023 es `sheet2.xml`. Resuelve siempre por nombre; el script
`dump-sheet.ps1` lo hace por ti.

Encabezado en la **fila 4**; los datos parten en la **fila 5** (la 5 suele ser el
saldo inicial del año, sin movimiento). Columnas, iguales de 2022 a 2026:

| Col | Campo             | Notas |
|-----|-------------------|-------|
| A   | Fecha             | tres formatos posibles, ver §3 |
| B   | Mes               | número de mes, fórmula |
| D   | Descripción       | texto del banco, suele venir truncado |
| G   | Documento         | |
| I   | Cargos            | egreso; vacío si es abono |
| J   | Abonos            | ingreso; vacío si es cargo |
| K   | Saldo             | fórmula acumulada |
| M   | Gasto / NumeroGasto | código de gasto de la comunidad |
| N   | Casa / GCCASACobrar | número de casa cuando el abono es de una casa |
| P   | Observacion       | a qué mes se imputó el pago |
| Q   | Saldo Tesorería   | |
| S   | Cuenta ToP / Fila de Tesorería | `Tesorería` o `Personal` |

En el TSV que produce `dump-sheet.ps1` la primera columna es el número de fila de
Excel, así que **la columna A queda en el campo 2, la N en el 15 y la P en el 17**.
Un movimiento de la casa 25 se filtra con `awk -F'\t' '$15=="25"'`.

Los encabezados cambian de nombre entre años (`NumeroGasto`/`Gasto`,
`GCCASACobrar`/`Casa`) pero las **posiciones son las mismas**. Confirma leyendo
la fila 4 antes de filtrar; es una línea y te ahorra un informe equivocado.

---

## 3. Fechas: los tres formatos que conviven

En la misma columna A puedes encontrar:

| Forma | Ejemplo | Cómo interpretarla |
|-------|---------|--------------------|
| Serial de Excel | `45313` | días desde 1899-12-30. `date -d "1899-12-30 + 45313 days"` |
| Texto completo | `13-04-2026` | tal cual, dd-mm-aaaa |
| Texto corto | `09/01`, `17/03` | dd/mm; el año es el del libro |

Después de convertir, **verifica que las fechas suban de forma monótona**. Las
filas están en el orden del banco, así que una fecha fuera de secuencia es un
error de digitación, no un dato. Ver §6.

---

## 4. Estado de Gastos Comunes: dónde sale el saldo actual

`Estados Gastos Comunes año <AAAA>.xlsx` tiene una hoja por mes de cobro, con
nombre `Gasto Común <Mes> <Año>` (por ejemplo `Gasto Común Agosto 2026`). Usa la
más reciente para el estado al día de hoy.

Los datos parten en la fila 5. **El layout cambia entre hojas mensuales**, así
que lee siempre la fila 4 de la hoja que vas a usar en vez de asumir las letras:

| Hoja | Columnas |
|------|----------|
| Agosto 2026 (con Asfalto) | D casa · E encargado · F Asfalto 2025 · G saldo anterior · H G/C del mes · **I Total = F+G+H** |
| Julio 2026 (sin Asfalto)  | D casa · E encargado · F saldo anterior · G G/C del mes · **H Total = F+G** |

Un informe leído con el mapa equivocado sale corrido una columna entera, y las
cifras igual parecen plausibles. Es el error más caro de esta parte.

La fila 3 dice hasta qué fecha se consideraron los depósitos — cítala, porque el
informe es válido a esa fecha de corte, no a la de hoy.

**Bug del `+D` en la columna Total.** En algunas hojas la fórmula del Total suma
también el número de casa, y la deuda sale con el número pegado al final
($11.830.049 en vez de $11.830.000). La hoja de agosto 2026 está corregida; la de
julio 2026 todavía lo arrastra en 35 casas. Si el total termina justo en los dos
dígitos de la casa, sospecha: recalcula sumando las columnas a mano y repórtalo.

---

## 5. Historia de la cuota mensual

Necesaria para leer las observaciones antiguas sin equivocarse:

| Desde | Cuota mensual |
|-------|---------------|
| ...  hasta agosto 2023 | $120.000 |
| septiembre 2023 | $170.000 |
| **julio 2024** | $200.000 |

Las dos alzas tienen un mes de transición en que conviven los dos montos: en
julio de 2024 la mayoría pagó $200.000, pero varias casas transfirieron los
$170.000 de siempre y después enteraron el saldo — de ahí los abonos de $30.000
anotados "Paga saldo 07/2024". Junio de 2024 sigue a $170.000. Si cuadras un
saldo contra los meses devengados, usar agosto en vez de julio te deja $30.000 de
diferencia por casa.

**Cuotas y aportes que no son la cuota mensual** (ninguno avanza el conteo de
meses pagados):

| Concepto | Cuándo | Cómo aparece |
|----------|--------|--------------|
| Cuota extra 2023 (Cex) | $50.000 por junio, julio y agosto 2023 | `Cex Ago`, `Paga Cex Jun Jul Ago` |
| Asfalto 2025 | cuota especial, fondo aparte | `AsfaltoDic25`, sin importar el mes real de pago |
| Aporte CSO | aporte extraordinario de seguridad, 2020–2021 | `CSO`, `Aporte CSO` |
| Cuota extraordinaria cámaras | 2016 | `Ex Cámara` |

**Sufijo identificador:** se pide a los propietarios que los dos últimos dígitos
del monto sean el número de casa (la casa 25 transfiere `$200.025`). Solo algunas
lo hacen. Es la forma más rápida de atribuir un depósito de un pagador
desconocido, pero **no es prueba concluyente**: contrasta con el historial de la
cuenta pagadora.

---

## 6. Trampas conocidas

**Años mal digitados.** En la Cartola 2025, las filas de la segunda quincena de
diciembre quedaron grabadas con serial de 2026 (46366–46374 = 10 al 18 de
diciembre de **2026**). El día y el mes están bien, el año está corrido en uno.
Detéctalo con el chequeo de monotonía de §3 y corrige al año del libro,
mencionándolo como nota al pie en vez de callarlo.

En la Cartola 2024 hay además unas siete filas fuera de secuencia (alrededor de
las filas 44, 155, 159, 296, 302, 307 y 309). Son desórdenes menores de captura,
no errores de año: revísalas, pero no te alarmes si tu chequeo las marca.

**Un mismo pagador para dos casas distintas.** "INGRID LORENA SANDOVAL" es la
casa 25; "BECKER SOTO INGRID MARLENE" es la casa 19. Buscar por nombre parcial
mezcla las dos. Filtra por la **columna N (casa)**, que es el dato autoritativo,
y usa el nombre solo para revisar si quedó algún depósito sin asignar.

**Una casa puede pagar desde varias cuentas.** La casa 25 pagó desde la cuenta
personal `0102028244` hasta noviembre 2025, desde `0760624500` (Inmobiliaria
Ghisan) entre diciembre 2025 y junio 2026, y volvió a alternar después. Un cambio
de cuenta no es un cambio de casa; vale la pena mencionarlo en el informe porque
explica saltos aparentes.

**Descripciones truncadas de distinta forma.** El mismo movimiento puede venir
como `...CHILE SPA` o `...CHILE SP` según el formato de cartola. Nunca compares
movimientos por texto: compara por fecha + monto.

**TAG y Asfalto no son cuotas.** Un abono múltiplo exacto de $10.000 y menor o
igual a $50.000 es compra de TAG ($10.000 por unidad). Ni el TAG ni el Asfalto
avanzan el conteo de meses pagados. Antes de 2025 el TAG costaba menos: hay
abonos de $4.000, $8.000 y $16.000 anotados como TAG.

**Casas con régimen especial.** La casa 49 (Cristian Sugg) no se rige por el
conteo de meses pagados: arrastra una deuda histórica a la que se suman $200.000
por cada mes que pasa sin pagar. Si te toca informar esa casa, dilo explícitamente
en vez de intentar traducir el saldo a meses.

**El archivo puede estar bloqueado.** Si Excel u OneDrive tienen la planilla
abierta, `dump-sheet.ps1` copia a temporal solo. Nunca escribas sobre los
originales: contienen gráficos, tablas dinámicas y segmentaciones que se pierden
al reescribirlos con librerías.
