---
name: los-cántaros-informe-individual-casa
description: Genera el informe de pagos de una casa de la comunidad Los Cántaros a partir de las Cartolas Mensuales de Gastos Comunes, con el detalle de movimientos (fecha, descripción, cargo, abono, casa, observación), totales por año, el estado de deuda al día de hoy y un análisis de retrasos y anticipos. Entrega un PDF y una planilla Excel. Usa esta skill siempre que se pida el historial o la evolución de pagos de una casa, un estado de cuenta, un certificado de deuda, o se consulte por el comportamiento de pago de una casa concreta ("cómo viene pagando la casa 25", "hazme el informe de la 49", "desde cuándo debe la 59", "revisa los pagos del sitio 12"), aunque no se use la palabra informe y aunque la pregunta parezca una consulta puntual.
---

# Informe de pagos por casa

Como tesorero te llegan consultas de una casa a la vez: *¿desde cuándo debe?*,
*¿me pagó agosto?*, *¿por qué me cobran dos meses?*. Esta skill arma la respuesta
completa: el detalle de cada depósito de esa casa, los totales, el estado al día
de hoy y un análisis en prosa de cómo viene pagando.

Sale en dos formatos, porque sirven para cosas distintas: un **PDF** para
adjuntar al correo que le mandas al propietario, y una **planilla Excel** con el
detalle crudo para que puedas seguir revisando o filtrando.

## Alcance por defecto

Empieza por **2024, 2025 y 2026**. Ese rango contesta casi todas las consultas y
evita barrer catorce años de planillas por gusto.

Retrocede año por año cuando el rango por defecto no alcanza para responder:

- **La pregunta es sobre el origen de algo** — "desde cuándo arrastra este
  saldo", "cuándo empezó a atrasarse", "desde cuándo tiene esa diferencia".
- **El rango no muestra el comienzo del patrón.** Si el primer movimiento del
  rango ya viene con un desfase o un saldo arrastrado, ese desfase nació antes:
  sigue hacia atrás hasta ver el mes en que la casa estaba en cero.
- **Aparecen pocos movimientos o ninguno.** Un informe de una casa dormida que
  solo dice "sin movimientos en 2024–2026" no le sirve a nadie. Retrocede hasta
  encontrar el último pago y arma el informe desde ahí — es justamente en esas
  casas donde el tesorero necesita saber cuándo fue la última vez.

Las planillas de 2016 a 2026 comparten estructura. Antes de 2016 el formato varía
y conviene mirar la hoja antes de filtrar; `references/planillas.md` dice qué
esperar de cada tramo.

## Cómo trabajar

### 1. Identifica la casa y su propietario

El nombre del encargado sale de la hoja de cobro más reciente en
`Estados Gastos Comunes año <AAAA>.xlsx` (hoja `Gasto Común <Mes> <Año>`).
De ahí saca también el saldo vigente.

**Lee la fila 4 de esa hoja antes de tomar cifras**: las columnas se corren entre
un mes y otro según si esa hoja lleva o no columna de Asfalto. Leer julio con el
mapa de agosto te da cifras corridas una columna, y todas parecen plausibles.

Anota la fecha de corte que declara la fila 3 — el informe vale a esa fecha, no
necesariamente a hoy.

### 2. Vuelca las Cartolas a TSV

```powershell
.\scripts\dump-sheet.ps1 -Path "<ruta>\Cartola Mensual Gastos Comunes Año 2026.xlsx" -Sheet "Cartola" -Out c2026.tsv
```

Una vez por año del rango. El script lee el XML interno sin abrir Excel, resuelve
la hoja por nombre (no asumas que es la primera) y copia a temporal si el archivo
está bloqueado por OneDrive.

En el TSV la primera columna es el número de fila de Excel, así que la columna A
de la planilla queda en el campo 2, la casa (col. N) en el 15 y la observación
(col. P) en el 17.

### 3. Filtra los movimientos de la casa

```bash
awk -F'\t' '$15=="25" {print $2"\t"$5"\t"$10"\t"$11"\t"$15"\t"$17}' c2026.tsv
```

Filtra por la **columna de casa**, nunca por el nombre del pagador: hay pagadores
homónimos en casas distintas y una misma casa paga desde varias cuentas.

Después haz un barrido de control: busca en la descripción los nombres y números
de cuenta que ya identificaste como de esta casa, y revisa si quedó algún
depósito **sin** número de casa asignado. Si aparece uno, no lo sumes por tu
cuenta — repórtalo como hallazgo para que el tesorero decida.

### 4. Normaliza las fechas y verifica que suban

Tres formatos conviven en la misma columna: serial de Excel (`45313`), texto
completo (`13-04-2026`) y texto corto (`09/01`). Conviértelos todos a dd-mm-aaaa.

Los movimientos están en el orden en que los registró el banco, así que **las
fechas tienen que ir subiendo**. Una que se sale de la secuencia es un error de
digitación en la planilla. Corrígela al valor que indican sus vecinas y déjala
como nota al pie del informe: el tesorero necesita saber que hay una celda que
arreglar. No la escondas y tampoco la des por buena.

### 5. Clasifica cada abono

Solo un tipo avanza el conteo de meses pagados, y por eso la clasificación
decide el resultado del informe:

- **Cuota** — el pago de gastos comunes. La observación dice a qué mes se imputó.
- **TAG** — el dispositivo de la barrera, $10.000 por unidad (antes de 2025
  costaba menos: hay abonos de $4.000, $8.000 y $16.000). Un múltiplo exacto de
  $10.000 y ≤ $50.000 es TAG por defecto.
- **Asfalto** — la cuota especial `AsfaltoDic25`, un fondo aparte.
- **Otros aportes extraordinarios** — el aporte CSO de seguridad (2020–2021), la
  cuota extra de 2023 (`Cex`), la cuota de cámaras de 2016 (`Ex Cámara`). Cada
  tanto aparece uno nuevo: si la observación nombra un concepto que no es un mes,
  trátalo como aporte extraordinario y ponle su propio chip (`.chip.otro`).

Todo lo que no sea cuota mensual va sumado aparte en los totales y **no** mueve
el último mes pagado.

### 6. Escribe el análisis de retrasos y anticipos

Es la parte que el tesorero no puede sacar de un vistazo a la tabla, y por eso es
la que más se lee. Dos párrafos, en prosa, sobre el **patrón** — no una lista de
los movimientos que ya están más arriba.

Busca:

- **El desfase normal de la casa.** Casi todas pagan el mes N durante el mes N+1;
  un mes de desfase suele ser el ritmo sano, no un atraso. Di cuál es el ritmo
  propio de esta casa antes de calificar nada.
- **Meses sin depósito, y si se regularizaron.** Un mes en blanco seguido de un
  pago doble no es un atraso: es un pago juntado. Un mes en blanco que nunca se
  compensa sí lo es.
- **Rupturas del patrón.** Que la fecha de pago se corra, que cambie la cuenta
  pagadora, que se corte una racha. Eso es lo que anticipa un problema.
- **Saldos a favor que se arrastran.** Anticipos parciales que quedan rodando de
  mes en mes. Si en algún momento dejan de aparecer en las observaciones, dilo:
  puede ser plata a favor del propietario que se perdió de vista.
- **Pagos extraordinarios al día**, como el Asfalto pagado completo.

Cierra con el veredicto: último mes pagado, cuántos meses adeuda y cuánto.
Si la casa está al día o adelantada, dilo con la misma claridad — usa
`.verdict.good` y `.tile.good` en vez de las variantes de alerta.

Para las casas de régimen especial (la 49 arrastra una deuda histórica que no se
traduce a meses) explica el régimen en vez de forzar el conteo.

### 7. Arma los entregables

**HTML → PDF.** Parte de `assets/plantilla-informe.html`, reemplaza los
marcadores `{{...}}` y duplica las filas. La plantilla ya trae el `@media print`
que hace que el PDF salga bien; no lo toques. **Borra el comentario de cabecera**
que lista los marcadores, o te queda un `{{CASA}}` suelto en el título.

La plantilla es un punto de partida, no una camisa de fuerza: si la casa no
compró TAG, saca esa tarjeta en vez de dejarla en $0; si la casa tiene régimen
especial, cambia las tarjetas por las que expliquen su deuda y agrega una tabla
de composición del saldo. Vale mucho más un informe que calce con el caso que uno
que respete la plantilla al pie de la letra.

```powershell
.\scripts\html-to-pdf.ps1 -Html informe-casa-25.html -Out "Informe Pagos Casa 25.pdf"
```

**Excel.** Escribe un TSV con encabezado `Fecha, Descripcion, Cargo, Abono, Casa,
Observacion` y una fila por movimiento, y conviértelo:

```powershell
.\scripts\write-xlsx.ps1 -Tsv casa25.tsv -Out "Detalle Pagos Casa 25.xlsx" `
  -SheetName "Casa 25" -Title "Casa 25 - Juan Ghisellini - Pagos 2024-2026"
```

`write-xlsx.ps1` crea un archivo nuevo desde cero: no abre ni reescribe ninguna
planilla existente, que es justo lo que hay que evitar con los libros de la
tesorería.

Guarda ambos en la carpeta del año en curso, salvo que el tesorero pida otra
ubicación. Si el entorno tiene `SendUserFile`, entrégaselos con esa herramienta;
si no, basta con decirle dónde quedaron.

### 8. Cierra en el chat

Un resumen corto: totales, último mes pagado, deuda, y el párrafo de retrasos y
anticipos. Si encontraste incidencias en las planillas (fechas mal digitadas,
depósitos sin asignar, saldos a favor perdidos), enuméralas al final por separado
— son trabajo pendiente del tesorero, no parte del informe al propietario.

Si además quiere una página web compartible del informe, publica el HTML como
Artifact; no lo hagas por defecto.

## Reglas que no conviene romper

**Nunca escribas sobre las planillas originales.** Contienen gráficos, tablas
dinámicas y segmentaciones que se destruyen al reescribirlas con librerías. Esta
skill solo lee; todo lo que produce son archivos nuevos.

**No inventes la atribución de un depósito.** Si un abono no tiene casa asignada
en la columna N, no lo sumes al informe aunque el monto o el pagador calcen.
Repórtalo y deja que el tesorero decida.

## Referencia

`references/planillas.md` tiene el mapa completo: rutas, columnas exactas,
nombres de hojas por año, la historia de la cuota mensual ($120.000 → $170.000 en
septiembre 2023 → $200.000 en agosto 2024, más la cuota extra de 2023) y las
trampas conocidas de los datos. Léelo antes de dar cifras por buenas, sobre todo
si vas a interpretar observaciones anteriores a 2024.
