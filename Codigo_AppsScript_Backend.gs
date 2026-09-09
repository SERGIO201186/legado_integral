// NovaPOS — Apps Script Backend
// Pega este código en script.google.com y despliega como Web App

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// 🔒 Debe ser IDÉNTICA a la "Clave secreta de sincronización" que pongas en
// NovaPOS (Configuración → Google Sheets). Sin esto, cualquiera que adivine
// o encuentre la URL /exec podría leer, modificar o borrar todos los datos
// del negocio — la URL de un Web App de Apps Script no es secreta por sí sola.
const SYNC_SECRET = 'NicolasBravo11centroVeracruzXalapa';

function checkSecret(secret) {
  return SYNC_SECRET && SYNC_SECRET !== 'CAMBIA_ESTO_por_una_clave_larga_y_unica' && secret === SYNC_SECRET;
}

function doGet(e) {
  const action = e.parameter.action || '';
  if (!checkSecret(e.parameter.secret)) return json({ok:false, error:'unauthorized'});

  // Los NIP (dueño, inventario, cada vendedor) viven en Propiedades del
  // script, no en ninguna hoja — así alguien con quien compartas la hoja de
  // cálculo (un contador, un socio) puede ver ventas/inventario sin ver los
  // NIP, que solo son visibles desde el editor de Apps Script.
  if (action === 'get_pins') return json({ok:true, data: getPins_()});

  // Login de Legado Integral (empleados + validación de NIP) — ver
  // getEmpleadosLogin_/handleResumenLogin_ más abajo.
  if (action === 'empleados') return json({ok:true, empleados: getEmpleadosLogin_()});
  if (action === 'resumen') return handleResumenLogin_(e.parameter.id, e.parameter.nip);

  const sheet = getSheet(e.parameter.sheet || 'productos');
  if (action === 'get') {
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    let data = rows.slice(1).map(r => Object.fromEntries(headers.map((h,i)=>[h,r[i]])));
    // "ventas" es lo que más pesa a la larga (crece con cada venta, para
    // siempre) — con el parámetro opcional "dias" el cliente puede pedir
    // solo las recientes en vez de bajar TODO el historial cada vez que
    // sincroniza, ahorrando datos móviles. Sin "dias" (p.ej. "Cargar
    // historial completo") se sigue regresando todo, igual que antes.
    if ((e.parameter.sheet||'productos') === 'ventas' && e.parameter.dias) {
      const corte = new Date(Date.now() - parseInt(e.parameter.dias, 10) * 24*3600*1000);
      data = data.filter(r => new Date(r.fecha) >= corte);
    }
    return json({ok:true, data});
  }
  return json({ok:false, error:'Unknown action'});
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!checkSecret(body.secret)) return json({ok:false, error:'unauthorized'});
    const {action, sheet: sheetName, data, id} = body;

    // La recarga no toca hojas directamente: solo habla con el proveedor y
    // regresa el resultado. NovaPOS guarda el registro por separado con un
    // 'upsert' normal a la hoja "recargas", igual que hace con ventas/facturas.
    if (action === 'recharge') return handleRecharge_(body);

    // Acuse del empleado sobre su ticket de cierre, escaneado y revisado en
    // Legado Integral (ver buildCorteQrPayload en NovaPOS/index.html — la
    // "especificación compartida" que arma ese QR). No reemplaza el corte
    // que NovaPOS ya guardó en "cortes": es la confirmación del empleado,
    // con el monto que entregó a administración (NovaPOS lo deja en blanco
    // a propósito porque ese paso es manual y le toca completarlo aquí).
    if (action === 'confirmar_turno') return handleConfirmarTurno_(body);

    // Guarda un NIP en Propiedades del script (nunca en una hoja). scope
    // 'owner'/'inventario' son un solo valor; 'vendedor' guarda un mapa
    // {vendedorId: pin} porque puede haber varios. pin vacío = quitar/borrar.
    if (action === 'set_pin') {
      const props = PropertiesService.getScriptProperties();
      if (body.scope === 'owner') props.setProperty('OWNER_PIN', body.pin || '');
      else if (body.scope === 'inventario') props.setProperty('INV_PIN', body.pin || '');
      else if (body.scope === 'vendedor' && body.vendedorId) {
        let mapa = {};
        try { mapa = JSON.parse(props.getProperty('VENDEDOR_PINS') || '{}'); } catch(e) {}
        if (body.pin) mapa[body.vendedorId] = body.pin; else delete mapa[body.vendedorId];
        props.setProperty('VENDEDOR_PINS', JSON.stringify(mapa));
      } else {
        return json({ok:false, error:'scope inválido'});
      }
      return json({ok:true});
    }

    // Relevo simple hacia MailApp — igual que 'recharge' con el proveedor de
    // recargas, NovaPOS no necesita ninguna credencial de correo propia,
    // solo pedirle a Apps Script que mande el email con la cuenta de Google
    // que tiene desplegado el script.
    if (action === 'send_corte_email') {
      if (!body.to) return json({ok:false, error:'Falta correo destino'});
      MailApp.sendEmail({ to: body.to, subject: body.subject || 'Corte de turno', htmlBody: body.html || '' });
      return json({ok:true});
    }

    // Cobro con terminal física Mercado Pago Point. El Access Token vive en
    // Propiedades del script (MP_ACCESS_TOKEN) — nunca en NovaPOS ni en la
    // hoja de cálculo.
    //
    // Listar dispositivos usa la API de Terminales (/terminals/v1/list) y no
    // /point/integration-api/devices: en esta cuenta esa segunda devuelve
    // "At least one policy returned UNAUTHORIZED" aunque el token sea válido
    // y la terminal ya esté en modo PDV — son productos/permisos distintos
    // dentro de Mercado Pago. /terminals/v1/list sí respondió con la
    // terminal real (id, pos_id, store_id, operating_mode).
    if (action === 'mp_listar_dispositivos') {
      return mpProxy_('GET', '/terminals/v1/list?limit=50&offset=0', null, function(data) {
        // La respuesta real viene anidada: {data:{terminals:[...]}, paging:{...}}
        var lista = ((data.data && data.data.terminals) || []).map(function(t) { return { id: t.id }; });
        return json({ ok:true, data: lista });
      });
    }
    // Cobrar con la terminal usa la API de "orders" (/v1/orders), confirmada
    // directamente contra la documentación oficial (create, consultar por id
    // y cancelar) — no /point/integration-api/devices/.../payment-intents,
    // que es la que dio UNAUTHORIZED arriba. El monto va como texto decimal
    // ("24.00"), no en centavos. El estatus de "pago aprobado" es
    // transactions.payments[0].status === 'processed' — confirmado contra un
    // cobro real de $6 (no 'approved', que es lo que se había asumido antes
    // sin poder verificarlo y causó que un cobro real se hiciera en la
    // terminal pero NovaPOS nunca cerrara la venta ni descontara inventario).
    if (action === 'mp_crear_intent') {
      // "items" es obligatorio para /v1/orders — sin él, Mercado Pago
      // responde 400 "no hay artículos para cobrar" aunque el monto y la
      // terminal estén bien. Si NovaPOS no manda el detalle del ticket, se
      // manda un solo artículo genérico con el monto total, para que la
      // suma siempre cuadre exacto con transactions.payments[0].amount sin
      // importar descuentos aplicados en el ticket.
      var mpItems = (body.items && body.items.length) ? body.items : [{ title: 'Venta NovaPOS', quantity: 1, unit_price: Number(body.amount) }];
      return mpProxy_('POST', '/v1/orders', {
        type: 'point',
        external_reference: body.externalReference || '',
        expiration_time: 'PT16M',
        transactions: { payments: [ { amount: Number(body.amount).toFixed(2) } ] },
        items: mpItems.map(function(it) { return { title: String(it.title || 'Producto'), quantity: Number(it.quantity) || 1, unit_price: Number(it.unit_price).toFixed(2) }; }),
        config: { point: { terminal_id: body.deviceId, print_on_terminal: 'no_ticket' } },
        description: 'Venta NovaPOS',
      }, function(data) { return json({ ok:true, data: { paymentIntentId: data.id } }); });
    }
    if (action === 'mp_estado_intent') {
      return mpProxy_('GET', '/v1/orders/' + body.paymentIntentId, null, function(data) {
        var pago = (data.transactions && data.transactions.payments && data.transactions.payments[0]) || {};
        return json({ ok:true, data: { status: data.status, paymentStatus: pago.status || '' } });
      });
    }
    if (action === 'mp_cancelar_intent') {
      return mpProxy_('POST', '/v1/orders/' + body.paymentIntentId + '/cancel');
    }

    // Resincronización completa (botón "Sincronizar todo" y los resets de
    // "Borrar datos de prueba"/"Borrar TODO"): a diferencia de 'sync', que
    // reemplaza una sola hoja, aquí 'data' es un objeto {nombreHoja: filas[]}
    // con varias hojas a la vez.
    if (action === 'sync_all') {
      const errores = [];
      Object.keys(data || {}).forEach(sheetName2 => {
        // Cada hoja se procesa en su propio try/catch: si una hoja falla (ej.
        // filas congeladas o protegidas que impiden borrar), las demás no se
        // quedan sin procesar — antes, un error en cualquier hoja detenía el
        // forEach entero y las hojas que venían después (turnos,
        // entradas_inventario) nunca llegaban a recibir sus encabezados.
        try {
          const sh = getSheet(sheetName2);
          const shHeaders = ensureHeaders_(sh, sheetName2);
          const lastRow = sh.getLastRow();
          // clearContent() en vez de deleteRows(): borra el contenido de las
          // filas de datos sin tocar la estructura de la hoja, así que nunca
          // choca con filas congeladas/protegidas ni con el límite de "no se
          // pueden borrar todas las filas" — deja el renglón de encabezados
          // intacto en la fila 1 siempre.
          if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
          (data[sheetName2] || []).forEach(d => sh.appendRow(shHeaders.map(h=>d[h]??'')));
        } catch (e) {
          errores.push(sheetName2 + ': ' + e.message);
        }
      });
      return errores.length ? json({ok:false, error: errores.join(' | ')}) : json({ok:true});
    }

    const sheet = getSheet(sheetName || 'productos');
    const headers = ensureHeaders_(sheet, sheetName || 'productos');

    if (action === 'upsert') {
      // La mayoría de las hojas identifican cada fila por "id", pero "config"
      // es key/value puro (encabezados "key","value") — sin este fallback,
      // idCol quedaba en -1 para esa hoja, nunca encontraba la fila existente
      // (upsert = siempre insertaba una nueva) y el valor de "key" se escribía
      // en blanco porque data.key no existía (el cliente mandaba data.id).
      const idField = headers.includes('id') ? 'id' : 'key';
      const rows = sheet.getDataRange().getValues();
      const idCol = headers.indexOf(idField);
      let existing = rows.findIndex((r,i)=>i>0 && r[idCol]===data[idField]);
      // Productos duplicados: si el mismo producto se da de alta casi al
      // mismo tiempo en dos dispositivos distintos (cada uno sin saber del
      // otro todavía), cada uno genera su propio id local — sin esto, el
      // upsert de cada uno no encontraba la fila del otro (ids distintos) y
      // los subía como dos productos separados con el mismo código de
      // barras. Como respaldo al id, si no hay match por id pero sí existe
      // ya una fila con el mismo "barcode", se actualiza esa en vez de
      // agregar una nueva.
      if (existing < 0 && (sheetName || '') === 'productos' && data.barcode) {
        const bcCol = headers.indexOf('barcode');
        if (bcCol >= 0) existing = rows.findIndex((r,i)=>i>0 && String(r[bcCol])===String(data.barcode));
      }
      const row = headers.map(h => data[h] ?? '');
      if (existing > 0) sheet.getRange(existing+1,1,1,row.length).setValues([row]);
      else sheet.appendRow(row);
      return json({ok:true});
    }
    if (action === 'delete') {
      // Mismo fallback que en 'upsert' — por si algún día se borra una fila
      // de la hoja "config" u otra hoja key/value.
      const idField = headers.includes('id') ? 'id' : 'key';
      const rows = sheet.getDataRange().getValues();
      const idCol = headers.indexOf(idField);
      const idx = rows.findIndex((r,i)=>i>0 && r[idCol]===id);
      if (idx > 0) sheet.deleteRow(idx+1);
      return json({ok:true});
    }
    if (action === 'sync') {
      // Bulk sync: replace all rows. clearContent() en vez de deleteRows()
      // (ver el mismo cambio y comentario en 'sync_all' más arriba).
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      data.forEach(d => sheet.appendRow(headers.map(h=>d[h]??'')));
      return json({ok:true});
    }
    return json({ok:false, error:'Unknown action'});
  } catch(err) {
    return json({ok:false, error:err.toString()});
  }
}

// 📲 Habla con el proveedor mayorista de recargas (Seycel, Taecel, Sivetel,
// etc.) y regresa el resultado a NovaPOS. TODO: sustituye el cuerpo de
// sendRecharge_ por la llamada real a la API de tu proveedor, usando
// UrlFetchApp y credenciales guardadas en Archivo → Propiedades del proyecto
// → Propiedades del script (Configuración ⚙️ del editor) — nunca las escribas
// aquí en texto plano ni en NovaPOS, cualquiera con el HTML/JS las vería.
function sendRecharge_(compania, telefono, monto) {
  // Ejemplo de cómo quedaría (ajusta nombres de campos y endpoint a la
  // documentación real que te dé tu proveedor al darte de alta):
  //
  // const props = PropertiesService.getScriptProperties();
  // const resp = UrlFetchApp.fetch('https://api.tuproveedor.mx/recarga', {
  //   method: 'post',
  //   contentType: 'application/json',
  //   payload: JSON.stringify({
  //     compania, telefono, monto,
  //     usuario: props.getProperty('PROVEEDOR_USUARIO'),
  //     clave:   props.getProperty('PROVEEDOR_CLAVE'),
  //   }),
  //   muteHttpExceptions: true,
  // });
  // const data = JSON.parse(resp.getContentText());
  // return { ok: !!data.exito, folioProveedor: data.folio || '', mensaje: data.mensaje || '' };

  return { ok:false, mensaje:'Proveedor de recargas no configurado todavía — edita sendRecharge_ en este script.' };
}

function handleRecharge_(body) {
  const { compania, telefono, monto } = body;
  const resultado = sendRecharge_(compania, telefono, monto);
  return resultado.ok
    ? json({ ok:true, folioProveedor: resultado.folioProveedor||'', mensaje: resultado.mensaje||'' })
    : json({ ok:false, error: resultado.mensaje||'Error del proveedor' });
}

// Guarda el acuse del empleado sobre un ticket de cierre en la hoja
// "legado_turnos" (separada de "cortes", que NovaPOS ya llenó al cerrar
// caja). Vuelve a validar el NIP contra VENDEDOR_PINS por su cuenta — el
// backend nunca confía en la sesión guardada del navegador para una acción
// que registra dinero entregado. Reenviar el mismo folio (p.ej. si el
// empleado vuelve a escanear un ticket ya confirmado) actualiza esa misma
// fila en vez de duplicarla.
function handleConfirmarTurno_(body) {
  if (!body.folio) return json({ok:false, error:'Falta el folio del ticket'});
  if (!body.id_empleado) return json({ok:false, error:'Falta el empleado'});

  const empleado = getEmpleadosLogin_().find(e => String(e.id) === String(body.id_empleado));
  if (!empleado) return json({ok:false, error:'Empleado no encontrado'});
  const pinGuardado = getPins_().vendedorPins[String(empleado._vendedorId)];
  if (!pinGuardado || String(pinGuardado) !== String(body.nip)) {
    return json({ok:false, error:'NIP incorrecto'});
  }

  const sheet = getSheet('legado_turnos');
  const headers = ensureHeaders_(sheet, 'legado_turnos');
  const rows = sheet.getDataRange().getValues();
  const folioCol = headers.indexOf('folio');
  const idCol = headers.indexOf('id');
  const existing = rows.findIndex((r,i) => i>0 && r[folioCol] === body.folio);

  const registro = {
    id: existing > 0 ? rows[existing][idCol] : Utilities.getUuid(),
    folio: body.folio,
    codigoEmpleado: empleado.id,
    nombreEmpleado: empleado.nombre,
    fecha: body.fecha || '',
    hora_apertura: body.hora_apertura || '',
    hora_cierre: body.hora_cierre || '',
    venta_turno: Number(body.venta_turno) || 0,
    recargas_telefonicas: Number(body.recargas_telefonicas) || 0,
    monto_entregado_admin: Number(body.monto_entregado_admin) || 0,
    inventario_vendido: Number(body.inventario_vendido) || 0,
    faltante: Number(body.faltante) || 0,
    merma: Number(body.merma) || 0,
    confirmado_en: new Date().toISOString(),
  };
  const row = headers.map(h => registro[h] ?? '');
  if (existing > 0) sheet.getRange(existing+1,1,1,row.length).setValues([row]);
  else sheet.appendRow(row);

  return json({ok:true});
}

function getPins_() {
  const props = PropertiesService.getScriptProperties();
  let vendedorPins = {};
  try { vendedorPins = JSON.parse(props.getProperty('VENDEDOR_PINS') || '{}'); } catch(e) {}
  return {
    ownerPin: props.getProperty('OWNER_PIN') || '',
    invPin:   props.getProperty('INV_PIN') || '',
    vendedorPins,
  };
}

// Lista de empleados para el login de Legado Integral. Usa la misma hoja
// "vendedores" que administra NovaPOS — el "id" que se manda al cliente es
// el codigoEmpleado (p.ej. "EMP-001"), no el id interno de la fila: es el
// mismo valor que trae el campo "id_empleado" del QR de cierre de turno que
// genera NovaPOS (ver generación del ticket en NovaPOS/index.html), así que
// Legado Integral puede comparar sesión vs. ticket escaneado directamente.
// Vendedores sin código de empleado capturado en NovaPOS quedan fuera: no
// hay forma de identificarlos en un ticket.
function getEmpleadosLogin_() {
  const rows = getSheet('vendedores').getDataRange().getValues();
  const headers = rows[0] || [];
  const idCol = headers.indexOf('id');
  const nombreCol = headers.indexOf('nombre');
  const codigoCol = headers.indexOf('codigoEmpleado');
  if (idCol < 0 || nombreCol < 0 || codigoCol < 0) return [];
  return rows.slice(1)
    .filter(r => r[codigoCol])
    .map(r => ({ id: r[codigoCol], nombre: r[nombreCol], _vendedorId: r[idCol] }));
}

// Valida el NIP contra VENDEDOR_PINS (el mismo NIP con el que el vendedor ya
// autoriza ventas en NovaPOS, ver getPins_/set_pin arriba) y regresa su
// progreso. Los 4 bonos quedan en $0 por ahora: sus reglas (meta de venta,
// tolerancia de retardo, montos) todavía no están definidas en ningún
// lado — mostrar un número inventado aquí pagaría bonos con un criterio que
// nadie acordó. Cuando se definan las reglas, calcularlas aquí dentro de
// resumenMes_.
function handleResumenLogin_(codigoEmpleado, nip) {
  if (!codigoEmpleado || !nip) return json({ok:false, error:'Falta empleado o NIP'});
  const empleado = getEmpleadosLogin_().find(e => String(e.id) === String(codigoEmpleado));
  if (!empleado) return json({ok:false, error:'Empleado no encontrado'});

  const pinGuardado = getPins_().vendedorPins[String(empleado._vendedorId)];
  if (!pinGuardado || String(pinGuardado) !== String(nip)) {
    return json({ok:false, error:'NIP incorrecto'});
  }

  return json({
    ok: true,
    empleado: { id: empleado.id, nombre: empleado.nombre },
    semana: resumenSemana_(codigoEmpleado),
    mes: resumenMes_(),
  });
}

function cortesDelEmpleadoEntre_(codigoEmpleado, desde, hasta) {
  const rows = getSheet('cortes').getDataRange().getValues();
  const headers = rows[0] || [];
  const codCol = headers.indexOf('codigoEmpleado');
  const apCol = headers.indexOf('apertura');
  const faltCol = headers.indexOf('faltante');
  if (codCol < 0 || apCol < 0) return [];
  return rows.slice(1)
    .filter(r => String(r[codCol]) === String(codigoEmpleado) && r[apCol])
    .map(r => ({ apertura: r[apCol], faltante: Number(r[faltCol]) || 0 }))
    .filter(c => { const d = new Date(c.apertura); return d >= desde && d < hasta; });
}

function resumenSemana_(codigoEmpleado) {
  const hoy = new Date();
  const diaSemana = (hoy.getDay() + 6) % 7; // lunes=0 ... domingo=6
  const inicio = new Date(hoy); inicio.setHours(0,0,0,0); inicio.setDate(hoy.getDate() - diaSemana);
  const fin = new Date(inicio); fin.setDate(inicio.getDate() + 7);

  const turnos = cortesDelEmpleadoEntre_(codigoEmpleado, inicio, fin);
  const conFaltante = turnos.filter(t => t.faltante > 0);

  const puntos_favor = [];
  const areas_oportunidad = [];
  if (turnos.length) puntos_favor.push(turnos.length + ' turno(s) registrado(s) esta semana.');
  if (turnos.length && !conFaltante.length) puntos_favor.push('Sin faltantes de caja esta semana.');
  else conFaltante.forEach(t => areas_oportunidad.push('Faltante el ' + fmtDMY_(t.apertura) + ': $' + t.faltante.toFixed(2)));
  if (!turnos.length) areas_oportunidad.push('Todavía no registras ningún turno esta semana.');

  return {
    rango: { inicio: fmtDMY_(inicio), fin: fmtDMY_(new Date(fin - 1)) },
    puntos_favor,
    areas_oportunidad,
  };
}

// Placeholder mientras se definen las reglas de los 4 bonos — ver el
// comentario en handleResumenLogin_.
function resumenMes_() {
  const hoy = new Date();
  return {
    mes: hoy.toLocaleDateString('es-MX', { month: 'long' }),
    elegible_premio_maximo: false,
    mensaje: 'El cálculo de bonos del mes está pendiente de configurar — próximamente verás aquí tu progreso real.',
    bono_ventas: 0,
    bono_puntualidad: 0,
    bono_caja: 0,
    bono_inventario: 0,
    total_bonos: 0,
  };
}

function fmtDMY_(fecha) {
  const d = new Date(fecha);
  return ('0'+d.getDate()).slice(-2) + '/' + ('0'+(d.getMonth()+1)).slice(-2);
}

function getConfigMap_() {
  const rows = getSheet('config').getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach(r => { if (r[0]) map[String(r[0])] = r[1]; });
  return map;
}

// 🔔 Alertas de inventario (stock bajo / próximo a caducar) por correo — a
// diferencia de la campanita dentro de NovaPOS, esto llega aunque nadie
// tenga la app abierta, porque corre del lado de Google. No hace nada
// (ni manda correo vacío) si no hay destinatario configurado o no hay nada
// que avisar hoy. El destinatario se configura desde NovaPOS → Configuración
// → Alertas (se guarda en la hoja "config" como cualquier otro dato del
// negocio), o si no, cae al mismo correo del resumen de turno.
function enviarAlertasDiarias() {
  const cfgMap = getConfigMap_();
  const destino = cfgMap.alertas_email || cfgMap.corte_email;
  if (!destino) return;

  const rows = getSheet('productos').getDataRange().getValues();
  const headers = rows[0];
  const productos = rows.slice(1).map(r => Object.fromEntries(headers.map((h,i)=>[h,r[i]])));

  const stockBajo = productos.filter(p => Number(p.stockmin) > 0 && Number(p.stock) <= Number(p.stockmin));
  const limite = new Date(Date.now() + 30*24*3600*1000);
  const porVencer = productos.filter(p => p.vence && new Date(p.vence) <= limite);
  if (!stockBajo.length && !porVencer.length) return;

  let html = '<h2 style="font-family:sans-serif">🔔 Alertas NovaPOS — ' + new Date().toLocaleDateString('es-MX') + '</h2>';
  if (stockBajo.length) {
    html += '<p style="font-family:sans-serif"><b>⚠️ Stock bajo (' + stockBajo.length + ')</b></p><ul style="font-family:sans-serif">' +
      stockBajo.map(p => '<li>' + p.nombre + ' — quedan ' + p.stock + ' (mínimo ' + p.stockmin + ')</li>').join('') + '</ul>';
  }
  if (porVencer.length) {
    html += '<p style="font-family:sans-serif"><b>⏳ Por caducar en 30 días o menos (' + porVencer.length + ')</b></p><ul style="font-family:sans-serif">' +
      porVencer.map(p => '<li>' + p.nombre + ' — vence ' + p.vence + '</li>').join('') + '</ul>';
  }
  MailApp.sendEmail({ to: destino, subject: 'Alertas NovaPOS — stock bajo y caducidad', htmlBody: html });
}

// Ejecuta ESTA función UNA sola vez desde el editor de Apps Script (▶ Ejecutar,
// con "crearTriggerAlertasDiarias" seleccionado en el menú de funciones) para
// que Google mande el correo de enviarAlertasDiarias() todos los días a las
// 8am sin que nadie tenga que hacer nada más. Se puede volver a correr sin
// problema — primero borra cualquier trigger anterior de la misma función
// para no terminar con dos avisos duplicados cada día.
function crearTriggerAlertasDiarias() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarAlertasDiarias') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarAlertasDiarias').timeBased().everyDays(1).atHour(8).create();
}

// Relevo genérico hacia la API de Mercado Pago — agrega el Authorization
// Bearer con MP_ACCESS_TOKEN (Propiedades del script) y traduce cualquier
// error HTTP de MP a {ok:false, error} en vez de dejar tronar el script.
function mpProxy_(method, path, payload, onOk) {
  const token = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
  if (!token) return json({ ok:false, error:'Falta configurar MP_ACCESS_TOKEN en Propiedades del script' });
  const options = {
    method: method.toLowerCase(),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  };
  // La API de "orders" exige X-Idempotency-Key en cada POST (crear, cancelar,
  // reembolsar) para no duplicar la operación si la petición se reintenta.
  if (options.method === 'post') options.headers['X-Idempotency-Key'] = Utilities.getUuid();
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  const resp = UrlFetchApp.fetch('https://api.mercadopago.com' + path, options);
  const code = resp.getResponseCode();
  let data = {};
  try { data = JSON.parse(resp.getContentText() || '{}'); } catch(e) {}
  if (code >= 400) return json({ ok:false, error: data.message || data.error || ('Mercado Pago respondió ' + code) });
  return onOk ? onOk(data) : json({ ok:true, data: data });
}

// Encabezados esperados por hoja — usados tanto al crear una hoja nueva
// (getSheet) como para reponerlos si una hoja ya existe pero se vació por
// completo a mano (ensureHeaders_, que usa sync_all antes de escribir).
const SHEET_HEADERS = {
  productos:   ['id','sku','barcode','nombre','cat','proveedor','precio','costo','stock','stockmin','vence','unidad','desc'],
  ventas:      ['id','folio','fecha','items','subtotal','descuento','total','recibido','cambio','metodo','vendedor','clienteId','sucursal'],
  movimientos: ['id','fecha','tipo','monto','concepto','sucursal'],
  facturas:    ['id','folio','ventaId','fecha','clienteNombre','clienteRFC','clienteEmail','clienteDir','usoCFDI','metodoPago','formaPago','subtotal','iva','total','estatus','cfdiUUID'],
  recargas:    ['id','folio','fecha','compania','modalidad','telefono','monto','comisionPct','gananciaEstimada','estatus','folioProveedor','metodoPago','mensaje'],
  // ⚠️ "cortes" es la ÚNICA hoja de esta lista que no traía "id" — sin él,
  // idCol quedaba en -1 en el upsert genérico de más abajo, y la comparación
  // "r[-1] === data['key']" (ambos undefined) daba TRUE para cualquier fila,
  // así que CADA cierre de caja sobreescribía la primera fila de datos en vez
  // de agregar una nueva. Resultado: la hoja nunca acumulaba más de un corte
  // a la vez. Con "id" agregado, el corte se identifica igual que cualquier
  // otra hoja y cada cierre de caja sí agrega su propia fila.
  //
  // copiasImpresionesVendidasTotal/Count/Detalle: el "Resumen de turno" que
  // imprime NovaPOS trae una sección aparte, "Copias e impresiones
  // vendidas" (dinero cobrado por copias/impresiones, distinto del
  // "Contador de impresora" que son solo lecturas del medidor), con una fila
  // por tipo de copia (p.ej. "Copia B/N carta $18.00 (9)"). Como el número
  // de tipos varía según el catálogo de precios, el detalle se guarda como
  // JSON (mismo patrón que "items" en la hoja "ventas") y Total/Count quedan
  // aparte para poder sumarlos sin tener que parsear el JSON cada vez.
  cortes:      ['id','apertura','cierre','fondo','ingresos','egresos','saldoFinal','vendedor','ventasCount','ventasTotal','efectivo','tarjeta','transferencia','recargasCount','recargasTotal','recargasComisionTotal','folio','codigoEmpleado','efectivoEsperado','efectivoContado','faltante','sucursal','fiadoCount','fiadoTotal','contAperturaBn','contAperturaColor','contAperturaImpBn','contAperturaImpColor','contCierreBn','contCierreColor','contCierreImpBn','contCierreImpColor','copiasImpresionesVendidasTotal','copiasImpresionesVendidasCount','copiasImpresionesVendidasDetalle'],
  vendedores:  ['id','nombre','codigoEmpleado'],
  config:      ['key','value'],
  // Clientes para venta a crédito ("fiado") — "saldo" es lo que debe
  // actualmente, "limiteCredito" es informativo (el checkout no lo bloquea,
  // solo lo usa NovaPOS para avisar). "clientes_movs" son solo los abonos
  // (pagos a cuenta); las ventas a crédito ya quedan en "ventas" con su
  // propio clienteId, no hace falta duplicarlas aquí.
  clientes:      ['id','nombre','telefono','email','saldo','limiteCredito','notas'],
  clientes_movs: ['id','clienteId','fecha','monto','concepto'],
  // Recargas hechas directo en la terminal física de Mercado Pago (fuera del
  // flujo normal de "recargas" de NovaPOS) — se capturan a mano al cerrar
  // caja porque la app no tiene forma de enterarse de esas operaciones.
  recargas_mp: ['id','fecha','folioOTelefono','monto','comision','corteFolio','sucursal'],
  // Cargas de saldo para recargas telefónicas — cada vez que el negocio
  // recarga su cuenta con el proveedor/agente real, queda un renglón aquí.
  // El saldo disponible (ver calcularSaldoRecargas() en NovaPOS) se calcula
  // sumando estos montos y restando el monto (sin comisión) de cada recarga
  // ya registrada en "recargas".
  recarga_saldo_movs: ['id','fecha','monto','nota','vendedor','sucursal'],
  // Kardex de entradas de mercancía — separado de "movimientos" (que es
  // dinero, ingresos/egresos de caja) para poder comparar entradas vs.
  // ventas vs. conteo físico y saber si una merma es real o solo mal
  // registrada como si fuera venta o al revés.
  entradas_inventario: ['id','fecha','productoId','productoNombre','cantidad','stockAntes','stockDespues','usuario'],
  // Catálogo de turnos de personal (Mañana/Noche, etc.) — informativo, no
  // tiene que ver con la apertura/cierre real de cada caja.
  turnos: ['id','nombre','horaInicio','horaFin','vendedorIds'],
  // Acuse del empleado sobre su ticket de cierre, capturado en Legado
  // Integral al escanear el QR que genera NovaPOS (buildCorteQrPayload).
  // "folio" identifica el ticket (mismo folio que en "cortes"); no
  // duplica los datos de "cortes", solo agrega lo que el empleado confirma
  // desde esta app — sobre todo monto_entregado_admin, que NovaPOS deja en
  // blanco a propósito porque ese paso es manual.
  legado_turnos: ['id','folio','codigoEmpleado','nombreEmpleado','fecha','hora_apertura','hora_cierre','venta_turno','recargas_telefonicas','monto_entregado_admin','inventario_vendido','faltante','merma','confirmado_en'],
};

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    if (SHEET_HEADERS[name]) s.getRange(1,1,1,SHEET_HEADERS[name].length).setValues([SHEET_HEADERS[name]]);
  }
  return s;
}

// Si alguien borra a mano TODO el contenido de una hoja (incluyendo la fila
// de encabezados, no solo los datos), getSheet() la sigue encontrando por
// nombre y no la vuelve a preparar — solo hace eso para hojas nuevas. Sin
// esto, sync_all truena al pedir sh.getRange(1,1,1,0) sobre una hoja sin
// ninguna columna con contenido. Regresa el arreglo de encabezados vigente.
function ensureHeaders_(sh, name) {
  const esperados = SHEET_HEADERS[name];
  if (!esperados) return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0];
  const actuales = sh.getLastColumn() > 0 ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] : [];
  const hayEncabezados = actuales.some(c => c !== '' && c !== null);
  if (!hayEncabezados) {
    sh.getRange(1,1,1,esperados.length).setValues([esperados]);
    return esperados;
  }
  // Agrega al final cualquier columna que el código ya espere pero la hoja
  // todavía no tenga (p.ej. "pin" en una hoja "vendedores" creada antes de
  // que existiera esa columna) — nunca reordena ni toca las columnas que ya
  // existen, así que los datos guardados no se mueven ni se pierden.
  const faltantes = esperados.filter(h => !actuales.includes(h));
  if (faltantes.length) {
    sh.getRange(1, actuales.length+1, 1, faltantes.length).setValues([faltantes]);
    return [...actuales, ...faltantes];
  }
  return actuales;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
