# Concurso Legado Integral

Documento de referencia de las reglas del concurso mensual de bonos para el
equipo de Legado Integral. Describe qué se premia, cuánto paga cada bono, y
cómo se mide cada uno dentro de la app y la hoja de cálculo. Las reglas de
negocio aquí descritas quedaron acordadas el 2026-09-11; los montos y
tolerancias viven en la hoja `config` de `Codigo_AppsScript_Backend.gs`
(con respaldo en `BONOS_DEFAULTS`) y son editables ahí sin tocar código.

Este documento es la referencia técnica/de negocio para quien administra la
app. Para la versión pensada para repartir o pegar y que lean los
empleados, ver el documento publicado ("Concurso Legado Integral").

## 1. Tu control del mes, dentro de la app

Todo tu historial y tu progreso de bonos se lleva dentro de la misma app
**Legado Integral** — no hay que llenar nada aparte ni por separado.

1. **Instálala en tu celular.** Es instalable directo desde el navegador
   (no necesitas Play Store ni App Store): abres el link, y tu teléfono te
   ofrece "Agregar a pantalla de inicio" — queda como cualquier otra app,
   con su propio ícono.
2. **Date de alta e inicia sesión** con tu nombre y tu NIP.
3. **Todos los días, al cerrar tu turno, es obligatorio capturar tu cierre
   de caja en la app.** Es lo único que alimenta tu historial y tus bonos
   — sin esa captura diaria, no hay datos correctos de los que partir.
   Es muy fácil: **escaneas el código QR** del ticket de cierre, o si no
   puedes escanear, **escribes el número de folio a mano** y la app trae
   los datos sola.
4. Revisa tu semana y tu mes en la pestaña **"Mi progreso"**, en cualquier
   momento.

> ⚠️ Un día que te tocaba trabajar y no capturas tu cierre de caja cuenta
> como **falta** para el Bono Puntualidad (ver sección 3.2) — a menos que
> sea tu día de descanso. Por eso la captura diaria es obligatoria, sin
> excepción.

## 2. Cómo se mide

Todo se calcula automáticamente en la app **Legado Integral**, a partir de
los turnos que cada empleado confirma al escanear (o capturar a mano) su
ticket de cierre de caja. En la pestaña **"Mi progreso"**, cada empleado ve
en tiempo real:
- Su resumen de la semana (turnos confirmados, faltantes de caja).
- El detalle de los 4 bonos operativos del mes en curso y los bonos por
  referido funerario que el dueño le haya registrado.

Nadie captura manualmente si "ganó" o no un bono: se calcula solo en
`resumenMes_()` (`Codigo_AppsScript_Backend.gs`) cada vez que el empleado
entra a la app.

## 3. Los 4 bonos operativos mensuales

Se evalúan mes calendario (del día 1 al último día del mes). Si se ganan
los 4 en el mismo mes, además de los $1,500 se gana **un día de descanso
adicional / pago doble** (acordado desde el inicio del proyecto).

### 3.1 Bono Ventas — $600

Se gana si, en **20 o más días** del mes, la venta de tu turno (venta del
turno + copias e impresiones vendidas) alcanzó la meta diaria de **tu**
turno (día o noche — son metas distintas).

**¿Cuál es la meta?**
- El primer mes del que hay historial: **$700/día** (turno día) y
  **$400/día** (turno noche).
- De ahí en adelante, se recalcula sola cada mes: se toma lo que
  **realmente se vendió** el mes anterior en ese turno, se le suma un
  **5% de crecimiento** (ajustable), y ese nuevo total se reparte entre los
  días que tiene el mes que se está calculando (no los del mes anterior —
  así un mes de 31 días no hereda tal cual el promedio de uno de 30).
- La meta puede subir **o bajar** de un mes a otro, según cómo se haya
  vendido realmente (no hay piso mínimo garantizado) — esto se acordó así
  a propósito.
- El dueño revisa la meta sugerida cada mes en la hoja `metas_ventas` y
  puede aceptarla tal cual o capturar un monto distinto a mano.

### 3.2 Bono Puntualidad — $400

Se gana si en el mes hay **0 retardos** (con una tolerancia de 1 retardo
por semana) **y 0 faltas**.

- **Horarios oficiales:** apertura del turno día a las **8:00 a.m.**, cierre
  del turno noche a las **8:00 p.m. (20:00)** — ambos con **15 minutos de
  tolerancia**. El relevo de en medio del día (cierre del turno día /
  apertura del turno noche, que ocurren al mismo tiempo) no tiene un
  horario fijo propio todavía.
- Si abres tarde o cierras después de tu horario (fuera de la tolerancia),
  la app te pide capturar el motivo antes de dejarte confirmar el turno.
  El dueño puede marcar ese retardo como "justificado" a mano en la hoja
  (por ejemplo, si quien te entregó la caja se retrasó) y ese retardo no
  cuenta en tu contra.
- **Tolerancia:** 1 retardo no justificado por semana está bien; el
  segundo retardo en esa misma semana ya hace perder el bono ese mes.
- **0 faltas:** un día que te tocaba trabajar (no era tu día de descanso) y
  no tienes ningún turno confirmado cuenta como falta, y **una sola falta
  en el mes pierde el bono**, sin importar los retardos.
- Cada empleado tiene su propio **día de descanso** configurado en la hoja
  `turnos_asignados` (por ejemplo, el turno noche descansando los
  domingos) — ese día nunca cuenta como falta.

### 3.3 Bono Caja — $250

Se gana si la suma de **faltantes** (dinero que faltó, no lo que sobró) en
el mes no pasa de **$100**, y no hay más de **1 turno por semana** con
faltante o sobrante (cualquiera de los dos).

### 3.4 Bono Inventario — $250

Se mide combinando la merma de **ambos turnos** (día + noche) de cada
fecha, porque los 2 hacen su propio conteo de inventario al entrar y no se
puede saber a cuál le faltó algo.

> ⚠️ **Pendiente:** todavía no está definido cuánta merma combinada se
> tolera al mes. Mientras el dueño no capture ese número en la hoja
> `config` (`bono_inventario_tolerancia_mensual`), este bono se calcula
> pero **no se paga** — se ve el total real de merma del mes para que el
> dueño lo revise.

## 4. Bonos por referido funerario

Independientes de los 4 anteriores: **no tienen tope**, se pagan por cada
servicio, y **no cuentan** para el combo de $1,500 + día libre (ese sigue
siendo solo de los 4 bonos operativos). El dueño los registra a mano en la
hoja `servicios_funerarios` conforme se van dando (no hay forma de
verificarlos automáticamente desde la app).

### 4.1 Bono Aviso de Servicio Funerario — $200 por servicio

Cada servicio funerario **efectivamente realizado** por Funerales Huerta,
del cual dieron aviso por referencia de este empleado.

### 4.2 Bono Servicio Directo Recomendado — $700 por servicio

Cuando el cliente llega **directo** con la referencia del empleado, sin
que el asesor funerario tenga que hacer ninguna labor de venta.

## 5. Resumen de montos

| Bono | Monto | Tope |
|---|---|---|
| Ventas | $600 | 1 vez/mes |
| Puntualidad | $400 | 1 vez/mes |
| Caja | $250 | 1 vez/mes |
| Inventario | $250 (pendiente de activar) | 1 vez/mes |
| **Los 4 juntos en el mismo mes** | **+$1,500 y día libre / pago doble** | — |
| Aviso de Servicio Funerario | $200 | por servicio, sin tope |
| Servicio Directo Recomendado | $700 | por servicio, sin tope |

## 6. Cuándo y cómo se paga

- **Bonos operativos** (Ventas, Puntualidad, Caja, Inventario) y el combo
  de $1,500: se pagan en la **quincena siguiente al cierre del mes** (el
  mes de septiembre se paga en la quincena de octubre).
- **Bonos por referido funerario** (Aviso de Servicio Funerario, Servicio
  Directo Recomendado): se pagan **al día siguiente** de realizado el
  servicio — no esperan a la quincena.
- **Día de descanso del premio máximo** (ganar los 4 bonos operativos el
  mismo mes): quien lo gana elige, dentro del **mes siguiente** al cierre,
  una fecha para tomarlo.
  - No puede coincidir con su día de descanso regular (el capturado en
    `turnos_asignados`).
  - **No es acumulable**: si no se agenda dentro de ese mes siguiente, se
    pierde — no se puede guardar para más adelante.
  - En vez de descansar, puede elegir **trabajar ese día** — en ese caso se
    le paga **doble** por ese día (sigue la misma regla de la quincena
    siguiente, pero del mes en que lo trabajó).
  - El dueño registra la elección (y el pago) a mano en la hoja
    `premios_dia_libre` — no hay ningún cálculo automático para esto, es
    solo la bitácora de la decisión.

## 7. Pendientes de definir por el negocio

- Tolerancia mensual de merma del Bono Inventario — se recomienda medirla
  **en dinero** (valor de lo perdido/dañado), no en piezas, para que sea
  consistente con el Bono Caja y con cómo ya viene tipado el campo `merma`
  en el contrato del QR de NovaPOS (monto en pesos, no conteo de unidades).
- Poblar la hoja `turnos_asignados` con los 2 empleados actuales y su día
  de descanso (mientras no exista esa fila, el candado de "0 faltas" del
  Bono Puntualidad no se evalúa).
