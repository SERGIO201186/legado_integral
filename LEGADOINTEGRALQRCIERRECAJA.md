# Especificación: QR de cierre de caja NovaPOS → Legado Integral

Este documento describe, para el equipo de **Legado Integral**, exactamente qué
información manda NovaPOS en el código QR del ticket de cierre de caja, cómo
se calcula cada dato, y qué debe hacer Legado Integral con ella (tanto para
guardarla como para devolverle a NovaPOS la confirmación del empleado).

No es una propuesta: es la especificación **ya implementada** del lado de
NovaPOS. Legado Integral debe adaptar su lógica de lectura/guardado a este
contrato, no al revés.

## 1. Flujo general

```
┌─────────────┐   cierra caja    ┌──────────────────┐   escanea QR    ┌───────────────────┐
│   NovaPOS   │ ───────────────▶ │  Ticket impreso   │ ───────────────▶│  Legado Integral   │
│ (punto de   │  genera "corte"  │  con código QR    │  (app empleado)  │  (app empleado)    │
│  venta)     │  + QR + guarda   │  (JSON en texto   │                  │                    │
│             │  en hoja "cortes"│   plano)          │                  │                    │
└─────────────┘                  └──────────────────┘                  └─────────┬──────────┘
                                                                                  │
                                                          empleado revisa y       │ POST
                                                          captura monto           │ action=confirmar_turno
                                                          entregado a admin       ▼
                                                                    ┌────────────────────────────┐
                                                                    │ Backend NovaPOS (Apps       │
                                                                    │ Script) → guarda en hoja    │
                                                                    │ "legado_turnos"             │
                                                                    └────────────────────────────┘
```

Dos hojas de cálculo distintas quedan involucradas, y **no se deben confundir**:

- **`cortes`**: la llena NovaPOS solo, al cerrar caja. Es la fuente de verdad completa del turno (incluye todo lo que hay en el QR y más: efectivo contado, desglose por método de pago, recargas MP, etc.).
- **`legado_turnos`**: la llena Legado Integral (vía el backend de NovaPOS), cuando el empleado escanea y confirma su ticket. Solo trae lo que venía en el QR, más lo que el empleado captura en ese momento (`monto_entregado_admin`).

Legado Integral **nunca escribe directo en Google Sheets** — todo pasa por el
backend de NovaPOS (Apps Script, `doPost` con `action: 'confirmar_turno'`),
que valida el NIP del empleado y hace el upsert a `legado_turnos`.

## 2. Cuándo se genera el QR

Se genera cada vez que se cierra una caja en NovaPOS, en `confirmarCierreCaja()`
(`index.html`). El QR se dibuja como SVG con `buildCorteQrSvg()` a partir del
JSON que arma `buildCorteQrPayload(corte)` (`index.html`, función central de
esta especificación). El mismo QR aparece también si se reconstruye un turno
pasado desde "Configuración → Venta anterior" — el formato es idéntico.

El contenido del QR es **texto plano JSON**, sin cifrar y sin firmar (ver
sección 6). Cualquier lector de QR estándar lo decodifica como texto.

## 3. Esquema completo del JSON del QR

```json
{
  "v": 1,
  "folio": "TCK-20260910-0007",
  "id_empleado": "EMP-001",
  "fecha": "2026-09-10",
  "hora_apertura": "08:00",
  "hora_cierre": "20:15",
  "venta_turno": 1540.50,
  "recargas_telefonicas": 300.00,
  "monto_entregado_admin": "",
  "inventario_vendido": 1500.00,
  "faltante": 0.00,
  "merma": 0.00,
  "copias_bn_usadas": 42,
  "copias_color_usadas": 5,
  "impresiones_bn_usadas": 10,
  "impresiones_color_usadas": 2,
  "copias_impresiones_vendido": 68.00
}
```

| Campo | Tipo | Origen / lógica de cálculo en NovaPOS |
|---|---|---|
| `v` | entero | Versión del formato del payload. Hoy siempre `1`. Ver sección 7 (versionado). |
| `folio` | string | Folio único del corte: `TCK-{fechaApertura sin guiones}-{consecutivo global}`. El consecutivo nunca decrece ni se reinicia por día — es `db.cortes.length + 1` en el momento del cierre. |
| `id_empleado` | string | `codigoEmpleado` del vendedor que tenía la sesión abierta (capturado al dar de alta al vendedor en NovaPOS, ej. `EMP-001`). Cadena vacía si la caja se operó sin vendedor identificado. |
| `fecha` | string `YYYY-MM-DD` | Fecha de **apertura** del turno (no de cierre) — un turno nocturno que cruza medianoche conserva la fecha en que abrió. |
| `hora_apertura` | string `HH:MM` | Hora local de apertura de caja. |
| `hora_cierre` | string `HH:MM` | Hora local de confirmación del cierre (no la hora en que se abrió el modal, sino la de "Confirmar cierre"). |
| `venta_turno` | número, 2 decimales | `ventasTotal + recargasComisionTotal`. `ventasTotal` = suma de ventas de artículos **excluyendo fiado** (el fiado no mete dinero a caja). `recargasComisionTotal` = suma de la comisión de recargas telefónicas completadas — se incluye porque esa comisión sí es ganancia del negocio, igual que en "TOTAL A ENTREGAR" del ticket impreso. El monto base de la recarga (lo que repone saldo) **no** se suma aquí. |
| `recargas_telefonicas` | número, 2 decimales | Suma del **monto base** (sin comisión) de las recargas telefónicas completadas del turno. Es dinero que solo repone saldo ya gastado, no es venta — por eso va separado de `venta_turno`. |
| `monto_entregado_admin` | string vacío en el QR | NovaPOS **siempre** lo manda vacío: ese paso (entrega física de efectivo a un administrador) es manual y todavía no lo cubre NovaPOS. **Este es el único campo que Legado Integral debe completar**, no leer, y devolverlo en la confirmación (ver sección 5). |
| `inventario_vendido` | número, 2 decimales | Aproximación = `ventasTotal` (solo venta de artículos, sin la comisión de recargas — una comisión no es mercancía). No hay todavía un costo/valor de inventario distinto calculado en NovaPOS. |
| `faltante` | número, 2 decimales | `efectivoEsperado - efectivoContado`, redondeado. `efectivoEsperado` = fondo inicial + efectivo (solo ventas en efectivo) + recargas (monto+comisión) + otros ingresos manuales − egresos. `efectivoContado` es lo que el usuario cuenta físicamente en el cajón al cerrar. **Positivo = faltante real, negativo = sobrante.** |
| `merma` | número | Siempre `0.00`. Placeholder: NovaPOS todavía no calcula merma. |
| `copias_bn_usadas` | entero | Lectura del **medidor físico** de la impresora (no dinero): `contadorCierreBn − contadorAperturaBn`. Ambos contadores se capturan a mano en NovaPOS (al abrir y al cerrar caja) leyendo el contador de la propia impresora. `0` si el negocio no tiene activo "Copias e impresiones" en Configuración, o si falta alguno de los dos conteos. |
| `copias_color_usadas` | entero | Igual lógica, contador de copias a color. |
| `impresiones_bn_usadas` | entero | Igual lógica, contador de impresiones en blanco y negro. |
| `impresiones_color_usadas` | entero | Igual lógica, contador de impresiones a color. |
| `copias_impresiones_vendido` | número, 2 decimales | **Dinero cobrado** (no unidades) por venta de copias/impresiones en el turno, según el catálogo de precios de NovaPOS (`SERVICIOS_COPIAS`, ver sección 4). Se recalcula desde el historial real de ventas del turno, no desde un contador. Sirve para **cuadrar** contra `copias_*_usadas`: si el uso del medidor no corresponde razonablemente con lo vendido, hay una discrepancia (copias regaladas, mal cobradas, o vendidas sin registrar en NovaPOS). |

### Nota importante sobre `copias_*_usadas` vs `copias_impresiones_vendido`

Son dos fuentes de datos **independientes** que NovaPOS ya trae por separado
y que el QR simplemente junta:

- Los 4 campos `*_usadas` son **lecturas de equipo** (cuánto avanzó el
  medidor físico) — no dependen de que el cajero haya cobrado nada.
- `copias_impresiones_vendido` es **dinero real cobrado** en el POS por los
  productos de la categoría "Copias e impresiones".

No van a coincidir en unidades directamente (uno es "unidades usadas", el
otro es "pesos cobrados") — Legado Integral no debe intentar reconciliarlos
automáticamente 1 a 1 dentro de esta versión del contrato; son dos señales
para que el proceso de auditoría del negocio las compare manualmente si le
interesa. NovaPOS no manda el desglose por tipo de servicio (cuántas "Copia
B/N carta" vs. "Copia B/N oficio", etc.) dentro del QR — eso viaja completo
en la hoja `cortes` (columna `copiasImpresionesVendidasDetalle`, JSON), no en
el QR, para no inflar su tamaño.

## 4. Catálogo de precios vigente (contexto, no se manda en el QR)

Este es el catálogo de precios que usa NovaPOS **hoy** para calcular
`copias_impresiones_vendido` (constante `SERVICIOS_COPIAS` en `index.html`).
Se documenta aquí solo como referencia — Legado Integral no necesita
replicar este cálculo, porque el total ya le llega calculado en el QR:

| Servicio | Precio actual |
|---|---|
| Copia B/N carta | $1.00 |
| Copia B/N oficio | $2.00 |
| Copia INE B/N | $2.00 |
| Copia/impresión a color (una sola tarifa, sin distinguir carta/oficio) | $5.00 |
| Descarga e impresión B/N (CURP, IMSS, ISSSTE) | $10.00 |
| Descarga e impresión a color (CURP, IMSS, ISSSTE) | $10.00 |
| Impresión suelta en blanco y negro (hoja) | $5.00 |

> Esta tabla puede cambiar en Configuración → Inventario de NovaPOS en
> cualquier momento; Legado Integral nunca debe hardcodearla — solo debe
> consumir el total ya calculado (`copias_impresiones_vendido`).

## 5. Qué debe hacer Legado Integral al escanear el QR

1. Decodificar el QR como JSON plano (sin descifrado ni verificación de
   firma — ver sección 6).
2. Mostrarle al empleado el resumen del turno con todos los campos de la
   sección 3.
3. Pedirle al empleado que capture **`monto_entregado_admin`** (el único
   campo que el QR trae vacío) — cuánto efectivo entregó físicamente.
4. Validar identidad del empleado con su NIP (mismo NIP que usa en NovaPOS —
   Legado Integral valida contra el backend, ver sección 5.1).
5. Enviar la confirmación al backend de NovaPOS (sección 5.1). **No** debe
   escribir directo a la hoja de cálculo.

### 5.1 Contrato del webhook de confirmación

**Endpoint:** el mismo Web App de Apps Script de NovaPOS (`doPost`).

**Request** (`action: "confirmar_turno"`):

```json
{
  "secret": "<clave de sincronización compartida con NovaPOS>",
  "action": "confirmar_turno",
  "folio": "TCK-20260910-0007",
  "id_empleado": "EMP-001",
  "nip": "1234",
  "fecha": "2026-09-10",
  "hora_apertura": "08:00",
  "hora_cierre": "20:15",
  "venta_turno": 1540.50,
  "recargas_telefonicas": 300.00,
  "monto_entregado_admin": 1500.00,
  "inventario_vendido": 1500.00,
  "faltante": 0.00,
  "merma": 0.00,
  "copias_bn_usadas": 42,
  "copias_color_usadas": 5,
  "impresiones_bn_usadas": 10,
  "impresiones_color_usadas": 2,
  "copias_impresiones_vendido": 68.00
}
```

Reglas:

- **`secret`**: la misma clave de sincronización configurada en NovaPOS
  (Configuración → Google Sheets). Sin ella el backend responde
  `{"ok":false,"error":"unauthorized"}`.
- **`folio`** e **`id_empleado`** son obligatorios; sin alguno de los dos el
  backend rechaza la petición.
- **`nip`**: Legado Integral debe reenviar el NIP del empleado en cada
  confirmación — el backend **vuelve a validarlo** contra el NIP guardado en
  NovaPOS (nunca confía en una sesión ya autenticada del lado del cliente,
  porque esta acción registra dinero entregado).
- Todos los demás campos deben reenviarse **tal cual venían en el QR**,
  excepto `monto_entregado_admin`, que Legado Integral rellena con lo que
  capturó el empleado.
- **Idempotencia por folio**: si Legado Integral reenvía una confirmación
  con un `folio` que ya existe en `legado_turnos`, el backend **actualiza**
  esa misma fila en vez de duplicarla (por ejemplo, si el empleado vuelve a
  escanear un ticket ya confirmado para corregir el monto entregado).

**Response:**

```json
{ "ok": true }
```

o, en caso de error:

```json
{ "ok": false, "error": "NIP incorrecto" }
```

Errores posibles: `unauthorized` (secret inválido), `Falta el folio del
ticket`, `Falta el empleado`, `Empleado no encontrado`, `NIP incorrecto`.

### 5.2 Qué guarda NovaPOS al recibir la confirmación

El backend escribe una fila en la hoja `legado_turnos` con estas columnas
(en este orden):

```
id, folio, codigoEmpleado, nombreEmpleado, fecha, hora_apertura, hora_cierre,
venta_turno, recargas_telefonicas, monto_entregado_admin, inventario_vendido,
faltante, merma, copias_bn_usadas, copias_color_usadas,
impresiones_bn_usadas, impresiones_color_usadas, copias_impresiones_vendido,
confirmado_en
```

- `id`: generado por NovaPOS (UUID) la primera vez; se conserva en
  actualizaciones posteriores del mismo folio.
- `nombreEmpleado`: NovaPOS lo resuelve internamente a partir de
  `id_empleado` (no hace falta que Legado Integral lo mande).
- `confirmado_en`: timestamp que pone el backend al momento de guardar (no
  lo manda Legado Integral).
- Los campos numéricos que Legado Integral no mande, o mande como texto no
  numérico, se guardan como `0`.

## 6. Sobre la firma / autenticidad del QR

El payload **no trae firma criptográfica** (no hay campo `firma` ni HMAC).
Es una omisión deliberada del lado de NovaPOS: requeriría coordinar una
llave secreta compartida con Legado Integral, que todavía no existe. Por
ahora, la única verificación real de identidad ocurre en el paso de
confirmación (validación de NIP contra el backend, sección 5.1) — el
contenido del QR en sí se debe tratar como **no verificado** hasta que pasa
por ese paso.

## 7. Versionado del payload

El campo `v` identifica la versión del esquema. Hoy es `1` e incluye los
campos nuevos de contador de copias/impresiones descritos aquí (antes de
este cambio, el payload `v:1` no traía esos 5 campos). Legado Integral debe:

- Leer los campos por **nombre**, nunca por posición — el orden de las
  claves del JSON no está garantizado.
- Tratar cualquier campo nuevo que no reconozca como **ignorable**, no como
  error — así NovaPOS puede seguir agregando campos sin romper integraciones
  ya hechas.
- Si algún campo documentado aquí llega ausente (por ejemplo, un ticket
  viejo generado antes de este cambio), asumir los valores por defecto
  descritos en la sección 3 (`0` para los contadores/venta de copias).

## 8. Referencias en el código de NovaPOS

- `buildCorteQrPayload()` — arma el JSON del QR (`index.html`).
- `buildCorteQrSvg()` — dibuja el QR como SVG (`index.html`).
- `confirmarCierreCaja()` — arma el objeto `corte` completo, incluye los
  contadores de impresora y el total de copias/impresiones vendidas
  (`index.html`).
- `calcularDesgloseCopias()` / `SERVICIOS_COPIAS` — catálogo y cálculo de
  venta de copias/impresiones (`index.html`).
- `handleConfirmarTurno_()` — backend que recibe la confirmación de Legado
  Integral y la guarda en `legado_turnos` (`nova_codigo.gs`).
- `SHEET_HEADERS.legado_turnos` / `SHEET_HEADERS.cortes` — esquemas de las
  hojas de cálculo involucradas (`nova_codigo.gs`).
