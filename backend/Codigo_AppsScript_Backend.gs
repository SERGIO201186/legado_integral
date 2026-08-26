/**
 * BACKEND — App de Empleados (escáner QR de cierre de caja)
 * ------------------------------------------------------------------
 * Instalación:
 * 1) Abre tu Google Sheet (la versión importada del libro de Excel).
 * 2) Extensiones > Apps Script.
 * 3) Borra el contenido de Code.gs y pega este archivo completo.
 * 4) Ajusta CONFIG.LLAVE_SECRETA si tu POS ya firma los QR (ver sección 3
 *    de QR_Schema_Ticket_Cierre.md). Si no, déjalo como está.
 * 5) Implementar > Nueva implementación > Tipo: Aplicación web.
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 * 6) Copia la URL que te da (".../exec") y pégala en app_empleados_escaner.html
 *    en la constante APPS_SCRIPT_URL.
 *
 * Recargas telefónicas:
 * - Si el QR del ticket incluye el campo `recargas_telefonicas` (monto en
 *   pesos), ese monto se resta de `venta_turno` antes de comparar contra la
 *   cuota del empleado, es decir, las recargas NO cuentan para la meta de
 *   venta. El total de venta del turno se sigue guardando completo (con
 *   recargas incluidas) en Registro_Diario_Operativo para no perder el
 *   ingreso real del día.
 * - El monto de recargas se registra aparte, en su propia hoja
 *   (CONFIG.HOJA_RECARGAS), que se crea automáticamente la primera vez que
 *   se procesa un ticket con recargas si todavía no existe.
 */

const CONFIG = {
  LLAVE_SECRETA: 'CAMBIA_ESTA_LLAVE_2026', // debe coincidir con la del POS si usan `firma`
  HOJA_DIRECTORIO: 'Directorio_Alta_Empleados',
  HOJA_REGISTRO: 'Registro_Diario_Operativo',
  HOJA_INVENTARIO: 'Control_Inventarios_y_Caja',
  HOJA_BITACORA: 'Bitacora_Escaneos_QR',
  HOJA_RECARGAS: 'Recargas_Telefonicas',
  DIRECTORIO_FIRST_ROW: 5,
  DIRECTORIO_LAST_ROW: 51,
  REGISTRO_FIRST_ROW: 5,
  INVENTARIO_FIRST_ROW: 5,
  BITACORA_FIRST_ROW: 5,
};

// ============================= ENTRY POINTS =============================

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'ping') return jsonOut({ ok: true, mensaje: 'Backend activo' });
    if (action === 'empleados') return jsonOut({ ok: true, empleados: getEmpleadosActivos() });
    if (action === 'resumen') return jsonOut(getResumenEmpleado(e.parameter.id, e.parameter.nip));
    return jsonOut({ ok: false, error: 'Acción no reconocida' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return jsonOut(procesarEscaneo(body));
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================= DIRECTORIO =============================

function leerDirectorio_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_DIRECTORIO);
  const last = sh.getLastRow();
  if (last < CONFIG.DIRECTORIO_FIRST_ROW) return [];
  const rango = sh.getRange(CONFIG.DIRECTORIO_FIRST_ROW, 1, last - CONFIG.DIRECTORIO_FIRST_ROW + 1, 12).getValues();
  return rango
    .filter(r => r[0]) // ID no vacío
    .map(r => ({
      id: String(r[0]).trim(),
      nombre: r[1],
      puesto: r[2],
      estatus: r[4],
      cuotaBase: Number(r[5]) || 0,
      nip: r[11] !== undefined && r[11] !== null ? String(r[11]).trim() : '',
    }));
}

function getEmpleadosActivos() {
  return leerDirectorio_()
    .filter(emp => (emp.estatus || '').toUpperCase() === 'ACTIVO')
    .map(emp => ({ id: emp.id, nombre: emp.nombre })); // NUNCA regresar el NIP aquí
}

function validarNip_(id, nip) {
  const emp = leerDirectorio_().find(e => e.id === String(id).trim());
  if (!emp) return { ok: false, error: 'Empleado no encontrado' };
  if ((emp.estatus || '').toUpperCase() !== 'ACTIVO') return { ok: false, error: 'Empleado inactivo' };
  if (emp.nip === '' || emp.nip !== String(nip).trim()) return { ok: false, error: 'NIP incorrecto' };
  return { ok: true, empleado: emp };
}

// ============================= PROCESAR ESCANEO =============================

function procesarEscaneo(body) {
  const requeridos = ['id_empleado', 'nip', 'folio', 'fecha', 'hora_apertura', 'hora_cierre', 'venta_turno', 'faltante'];
  for (const campo of requeridos) {
    if (body[campo] === undefined || body[campo] === null || body[campo] === '') {
      return { ok: false, error: 'Falta el campo: ' + campo };
    }
  }

  const auth = validarNip_(body.id_empleado, body.nip);
  if (!auth.ok) return auth;
  const emp = auth.empleado;

  // Verificación de firma opcional (antifraude) — ver sección 3 del schema
  if (body.firma) {
    const cadena = String(body.folio) + body.id_empleado + body.fecha + body.venta_turno + body.faltante + (body.merma || 0);
    const firmaEsperada = Utilities.computeHmacSha256Signature(cadena, CONFIG.LLAVE_SECRETA)
      .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').substring(0, 12);
    if (firmaEsperada !== body.firma) {
      return { ok: false, error: 'La firma del ticket no es válida (posible QR alterado)' };
    }
  }

  if (folioYaProcesado_(body.folio)) {
    return { ok: false, error: 'Este ticket ya fue registrado antes (folio duplicado): ' + body.folio };
  }

  const faltante = Number(body.faltante) || 0;
  const merma = Number(body.merma) || 0;
  const recargas = Number(body.recargas_telefonicas) || 0;

  registrarBitacora_(body, emp, 'Procesado');
  registrarRegistroDiario_(body, emp, faltante, recargas);
  registrarControlInventario_(body, emp, faltante, merma);
  registrarRecargas_(body, emp, recargas);

  return { ok: true, mensaje: 'Turno registrado correctamente', empleado: emp.nombre };
}

function folioYaProcesado_(folio) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BITACORA);
  const last = sh.getLastRow();
  if (last < CONFIG.BITACORA_FIRST_ROW) return false;
  const folios = sh.getRange(CONFIG.BITACORA_FIRST_ROW, 4, last - CONFIG.BITACORA_FIRST_ROW + 1, 1).getValues().flat();
  return folios.some(f => String(f).trim() === String(folio).trim());
}

function registrarBitacora_(body, emp, estatus) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_BITACORA);
  sh.appendRow([
    new Date(), emp.id, emp.nombre, body.folio, body.fecha, body.hora_apertura, body.hora_cierre,
    Number(body.venta_turno) || 0, Number(body.monto_entregado_admin) || 0,
    Number(body.inventario_vendido) || 0, Number(body.faltante) || 0, Number(body.merma) || 0, estatus,
  ]);
}

function horaLimiteEntrada_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_REGISTRO);
  const val = sh.getRange('N4').getValue(); // parámetro configurado en el Excel/Sheet
  if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
  return 9 * 60; // default 9:00 si el parámetro no está configurado
}

function minutosDeHora_(hhmm) {
  const [h, m] = String(hhmm).split(':').map(Number);
  return h * 60 + (m || 0);
}

function registrarRegistroDiario_(body, emp, faltante, recargas) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_REGISTRO);
  const limite = horaLimiteEntrada_();
  const entrada = minutosDeHora_(body.hora_apertura);
  const retardo = entrada > limite ? 'SÍ' : 'NO';
  const aperturaATiempo = retardo === 'SÍ' ? 'NO' : 'SÍ';
  const venta = Number(body.venta_turno) || 0;
  const ventaParaMeta = venta - recargas; // las recargas telefónicas no cuentan para la meta de venta
  const cuota = emp.cuotaBase || 0;
  const meta = ventaParaMeta >= cuota ? 'SÍ' : 'NO';
  const cierreCuadrado = faltante === 0 ? 'SÍ' : 'NO';

  sh.appendRow([
    body.fecha, emp.id, emp.nombre, body.hora_apertura, retardo, venta, cuota, meta,
    aperturaATiempo, cierreCuadrado, faltante,
  ]);
}

function registrarControlInventario_(body, emp, faltante, merma) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_INVENTARIO);
  const arqueo = (faltante > 0 || merma > 0) ? 'Con Faltante' : 'Aprobado';
  sh.appendRow([body.fecha, emp.id, faltante, merma, arqueo, faltante]);
}

function registrarRecargas_(body, emp, recargas) {
  if (recargas <= 0) return; // nada que registrar en este ticket
  let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_RECARGAS);
  if (!sh) {
    sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.HOJA_RECARGAS);
    sh.appendRow(['Fecha', 'ID Empleado', 'Nombre', 'Folio', 'Monto Recargas ($)']);
  }
  sh.appendRow([body.fecha, emp.id, emp.nombre, body.folio, recargas]);
}

// ============================= RESUMEN SEMANAL / MENSUAL =============================

function getResumenEmpleado(id, nip) {
  const auth = validarNip_(id, nip);
  if (!auth.ok) return auth;
  const emp = auth.empleado;

  const hoy = new Date();
  const semana = resumenSemana_(emp, hoy);
  const mes = resumenMes_(emp, hoy);

  return { ok: true, empleado: { id: emp.id, nombre: emp.nombre }, semana, mes };
}

function leerRegistroDelEmpleado_(idEmpleado) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_REGISTRO);
  const last = sh.getLastRow();
  if (last < CONFIG.REGISTRO_FIRST_ROW) return [];
  const datos = sh.getRange(CONFIG.REGISTRO_FIRST_ROW, 1, last - CONFIG.REGISTRO_FIRST_ROW + 1, 11).getValues();
  return datos
    .filter(r => r[1] === idEmpleado)
    .map(r => ({
      fecha: r[0] instanceof Date ? r[0] : new Date(r[0]),
      retardo: r[4], venta: Number(r[5]) || 0, cuota: Number(r[6]) || 0,
      meta: r[7], cierreCuadrado: r[9], descuadre: Number(r[10]) || 0,
    }));
}

function leerInventarioDelEmpleado_(idEmpleado) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_INVENTARIO);
  const last = sh.getLastRow();
  if (last < CONFIG.INVENTARIO_FIRST_ROW) return [];
  const datos = sh.getRange(CONFIG.INVENTARIO_FIRST_ROW, 1, last - CONFIG.INVENTARIO_FIRST_ROW + 1, 6).getValues();
  return datos
    .filter(r => r[1] === idEmpleado)
    .map(r => ({ fecha: r[0] instanceof Date ? r[0] : new Date(r[0]), faltante: Number(r[2]) || 0, merma: Number(r[3]) || 0 }));
}

function inicioSemana_(fecha) {
  const d = new Date(fecha);
  const dia = d.getDay(); // 0=domingo
  const diff = dia === 0 ? -6 : 1 - dia; // lunes como inicio
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function resumenSemana_(emp, hoy) {
  const inicio = inicioSemana_(hoy);
  const fin = new Date(inicio); fin.setDate(fin.getDate() + 7);

  const registros = leerRegistroDelEmpleado_(emp.id).filter(r => r.fecha >= inicio && r.fecha < fin);

  const diasConMeta = registros.filter(r => r.meta === 'SÍ').length;
  const retardos = registros.filter(r => r.retardo === 'SÍ').length;
  const descuadres = registros.filter(r => r.descuadre > 0).length;
  const diasTrabajados = registros.length;

  const puntosFavor = [];
  const areasOportunidad = [];

  if (diasConMeta > 0) puntosFavor.push(`Alcanzaste tu meta de venta en ${diasConMeta} de ${diasTrabajados} turnos esta semana.`);
  if (retardos === 0 && diasTrabajados > 0) puntosFavor.push('Llegaste puntual todos tus turnos de la semana. ¡Sigue así!');
  if (descuadres === 0 && diasTrabajados > 0) puntosFavor.push('Tu caja cuadró perfecto en todos tus cierres.');

  if (retardos > 0) areasOportunidad.push(`Tuviste ${retardos} retardo(s) esta semana — recuerda que 0 retardos en el mes te da $400 de bono.`);
  if (diasTrabajados > 0 && diasConMeta < diasTrabajados) areasOportunidad.push(`No alcanzaste la meta de venta en ${diasTrabajados - diasConMeta} turno(s) — cada día cuenta para tu bono mensual de ventas.`);
  if (descuadres > 0) areasOportunidad.push(`Tuviste ${descuadres} cierre(s) con faltante — un cierre exacto te acerca al bono de caja.`);
  if (diasTrabajados === 0) areasOportunidad.push('Aún no hay turnos registrados esta semana.');

  return {
    rango: { inicio: Utilities.formatDate(inicio, 'America/Mexico_City', 'yyyy-MM-dd'), fin: Utilities.formatDate(new Date(fin - 86400000), 'America/Mexico_City', 'yyyy-MM-dd') },
    dias_trabajados: diasTrabajados,
    puntos_favor: puntosFavor,
    areas_oportunidad: areasOportunidad,
  };
}

function resumenMes_(emp, hoy) {
  const anio = hoy.getFullYear();
  const mes = hoy.getMonth(); // 0-indexado
  const inicio = new Date(anio, mes, 1);
  const fin = new Date(anio, mes + 1, 1);

  const registros = leerRegistroDelEmpleado_(emp.id).filter(r => r.fecha >= inicio && r.fecha < fin);
  const inventario = leerInventarioDelEmpleado_(emp.id).filter(r => r.fecha >= inicio && r.fecha < fin);

  const diasConMeta = registros.filter(r => r.meta === 'SÍ').length;
  const bonoVentas = diasConMeta >= 20 ? 400 : 0;

  const retardos = registros.filter(r => r.retardo === 'SÍ').length;
  const bonoPuntualidad = retardos === 0 ? 400 : 0;

  const conFaltanteCaja = inventario.some(r => r.faltante > 0);
  const arqueoOk = !conFaltanteCaja;
  const bonoCaja = (bonoPuntualidad === 400 && arqueoOk) ? 200 : 0;

  const faltanteInventarioTotal = inventario.reduce((s, r) => s + r.faltante + r.merma, 0);
  const bonoInventario = (bonoVentas === 400 && faltanteInventarioTotal === 0) ? 200 : 0;

  const total = bonoVentas + bonoPuntualidad + bonoCaja + bonoInventario;
  const elegiblePremio = total === 1200;

  let felicitacion = null;
  if (elegiblePremio) {
    felicitacion = '🎉 ¡Felicidades! Este mes ganaste los 4 bonos y eres acreedor(a) a día libre / pago doble.';
  } else if (total > 0) {
    felicitacion = `¡Buen trabajo! Llevas $${total} en bonos ganados este mes.`;
  }

  return {
    mes: Utilities.formatDate(inicio, 'America/Mexico_City', 'MMMM yyyy'),
    dias_con_meta: diasConMeta, bono_ventas: bonoVentas,
    retardos: retardos, bono_puntualidad: bonoPuntualidad,
    arqueo_ok: arqueoOk, bono_caja: bonoCaja,
    faltante_inventario: faltanteInventarioTotal, bono_inventario: bonoInventario,
    total_bonos: total, elegible_premio_maximo: elegiblePremio, mensaje: felicitacion,
  };
}
