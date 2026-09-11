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
  // getEmpleadosLogin_/handleResumenLogin_ más abajo. El map() quita el NIP
  // (_nip) antes de responder: getEmpleadosLogin_() lo trae para poder
  // validarlo en el propio backend, pero nunca debe salir en la lista
  // pública que arma el selector de "Empleado" en el login.
  if (action === 'empleados') return json({ok:true, empleados: getEmpleadosLogin_().map(e => ({id: e.id, nombre: e.nombre}))});
  if (action === 'resumen') return handleResumenLogin_(e.parameter.id, e.parameter.nip);

  // Respaldo de la captura manual del ticket cuando la cámara no sirve:
  // en vez de que el empleado teclee los montos a mano (con riesgo de
  // error de captura), trae el corte real consultando directo al backend
  // de NovaPOS de este negocio. Ver handleBuscarFolio_ más abajo.
  if (action === 'buscar_folio') return handleBuscarFolio_(e.parameter.folio);

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
// "legado_turnos" de ESTA hoja de cálculo (la cuenta propia del operador de
// Legado Integral — separada de "cortes" en NovaPOS, que vive en la cuenta
// de cada negocio cliente y esta hoja nunca ve). Vuelve a validar el NIP
// por su cuenta contra el directorio de empleados — el backend nunca
// confía en la sesión guardada del navegador para una acción que registra
// dinero entregado. Reenviar el mismo folio (p.ej. si el empleado vuelve a
// escanear un ticket ya confirmado) actualiza esa misma fila en vez de
// duplicarla.
function handleConfirmarTurno_(body) {
  if (!body.folio) return json({ok:false, error:'Falta el folio del ticket'});
  if (!body.id_empleado) return json({ok:false, error:'Falta el empleado'});

  const empleado = getEmpleadosLogin_().find(e => String(e.id) === String(body.id_empleado));
  if (!empleado) return json({ok:false, error:'Empleado no encontrado'});
  if (String(empleado._nip) !== String(body.nip)) {
    return json({ok:false, error:'NIP incorrecto'});
  }

  // El turno (día/noche) no lo manda el QR de NovaPOS. Primero se busca en
  // "turnos_asignados" (quién tiene asignado cada turno, ver más abajo) —
  // es más confiable porque los 2 turnos siempre son las mismas 2 personas
  // (salvo que alguien renuncie). Si ese día no hay ninguna asignación
  // vigente para este empleado (roster todavía no capturado, o suplencia
  // sin registrar), se cae al viejo criterio por hora de apertura.
  const asignacionActual = body.fecha ? asignacionVigenteEn_(asignacionesTurno_(), empleado.id, new Date(body.fecha)) : null;
  const turno = asignacionActual ? asignacionActual.turno : detectarTurno_(body.hora_apertura);
  // El bono de Puntualidad solo se checa en los 2 extremos del día que sí
  // tienen horario oficial acordado: la apertura del turno día (8:00) y el
  // cierre del turno noche (20:00) — el relevo de en medio (cierre del
  // turno día = apertura del turno noche, es el mismo momento) no tiene una
  // hora fija propia.
  const tolerancia = Number(getBonoConfig_('bono_puntualidad_tolerancia_min'));
  const retardoApertura = turno === 'dia'
    && minutosDeRetraso_(body.hora_apertura, getBonoConfig_('bono_puntualidad_apertura_hhmm')) > tolerancia;
  const retardoCierre = turno === 'noche'
    && minutosDeRetraso_(body.hora_cierre, getBonoConfig_('bono_puntualidad_cierre_hhmm')) > tolerancia;
  if ((retardoApertura || retardoCierre) && !String(body.motivo_retardo || '').trim()) {
    return json({ok:false, error:'Faltó indicar el motivo del retraso.'});
  }

  const sheet = getSheet('legado_turnos');
  const headers = ensureHeaders_(sheet, 'legado_turnos');
  const rows = sheet.getDataRange().getValues();
  const folioCol = headers.indexOf('folio');
  const idCol = headers.indexOf('id');
  const existing = rows.findIndex((r,i) => i>0 && r[folioCol] === body.folio);
  // "retardo_justificado" lo marca el dueño a mano en la hoja (excepción de
  // relevo, ver LEGADOINTEGRALQRCIERRECAJA.md) — si el empleado vuelve a
  // escanear/reenviar el mismo folio, no se debe perder ese juicio manual.
  const justificadoCol = headers.indexOf('retardo_justificado');
  const retardoJustificadoPrevio = (existing > 0 && justificadoCol >= 0) ? rows[existing][justificadoCol] : '';

  const registro = {
    id: existing > 0 ? rows[existing][idCol] : Utilities.getUuid(),
    folio: body.folio,
    codigoEmpleado: empleado.id,
    nombreEmpleado: empleado.nombre,
    fecha: body.fecha || '',
    hora_apertura: body.hora_apertura || '',
    hora_cierre: body.hora_cierre || '',
    turno,
    retardo_apertura: retardoApertura,
    retardo_cierre: retardoCierre,
    motivo_retardo: body.motivo_retardo || '',
    retardo_justificado: retardoJustificadoPrevio,
    // venta_turno ya viene de NovaPOS con la comisión de recargas incluida
    // (ventasTotal + recargasComisionTotal, ver LEGADOINTEGRALQRCIERRECAJA.md
    // sección 3) — no existe un campo "comision_recargas" separado en el
    // contrato real del QR/confirmar_turno, así que ya no se guarda aparte.
    venta_turno: Number(body.venta_turno) || 0,
    recargas_telefonicas: Number(body.recargas_telefonicas) || 0,
    monto_entregado_admin: Number(body.monto_entregado_admin) || 0,
    inventario_vendido: Number(body.inventario_vendido) || 0,
    faltante: Number(body.faltante) || 0,
    merma: Number(body.merma) || 0,
    // Lecturas del medidor físico de la impresora (unidades, no dinero) —
    // ver sección 3 de LEGADOINTEGRALQRCIERRECAJA.md. Se guardan aparte de
    // copias_impresiones_vendido (dinero cobrado) porque son dos señales
    // independientes que el negocio compara manualmente al auditar.
    copias_bn_usadas: Number(body.copias_bn_usadas) || 0,
    copias_color_usadas: Number(body.copias_color_usadas) || 0,
    impresiones_bn_usadas: Number(body.impresiones_bn_usadas) || 0,
    impresiones_color_usadas: Number(body.impresiones_color_usadas) || 0,
    copias_impresiones_vendido: Number(body.copias_impresiones_vendido) || 0,
    confirmado_en: new Date().toISOString(),
  };
  const row = headers.map(h => registro[h] ?? '');
  if (existing > 0) sheet.getRange(existing+1,1,1,row.length).setValues([row]);
  else sheet.appendRow(row);

  return json({ok:true});
}

// Trae el corte real por folio consultando directo al backend de NovaPOS
// de este negocio (su propia URL /exec y su propia SYNC_SECRET — NUNCA en
// este código fuente, igual que MP_ACCESS_TOKEN: se configuran una sola
// vez desde Propiedades del proyecto en el editor de Apps Script). Esto
// evita que el empleado tenga que volver a teclear a mano los montos del
// ticket cuando ya sabe el folio, con el riesgo de un error de captura.
// NovaPOS no ofrece una acción para buscar un solo folio, así que se baja
// toda la hoja "cortes" (action=get, igual que hace el propio NovaPOS al
// sincronizar) y se filtra aquí.
function handleBuscarFolio_(folio) {
  if (!folio) return json({ok:false, error:'Falta el folio'});
  const props = PropertiesService.getScriptProperties();
  const novaUrl = props.getProperty('NOVAPOS_URL');
  const novaSecret = props.getProperty('NOVAPOS_SECRET');
  if (!novaUrl || !novaSecret) {
    return json({ok:false, error:'Falta configurar NOVAPOS_URL/NOVAPOS_SECRET en Propiedades del script'});
  }

  const url = novaUrl + '?action=get&sheet=cortes&secret=' + encodeURIComponent(novaSecret);
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  let body;
  try { body = JSON.parse(resp.getContentText() || '{}'); } catch(e) { return json({ok:false, error:'Respuesta inválida de NovaPOS'}); }
  if (!body.ok) return json({ok:false, error:'NovaPOS: ' + (body.error || 'error desconocido')});

  const corte = (body.data || []).find(c => String(c.folio) === String(folio));
  if (!corte) return json({ok:false, error:'No se encontró ese folio en NovaPOS'});

  const hhmm = iso => { const d = new Date(iso); return ('0'+d.getHours()).slice(-2) + ':' + ('0'+d.getMinutes()).slice(-2); };
  // '' (no 0) cuando el NovaPOS de este negocio todavía no guarda este dato
  // en "cortes" (columna agregada después — versiones viejas de
  // nova_codigo.gs no la tienen) — un 0 se vería como "ya lo revisé y no
  // hubo", cuando en realidad es "este NovaPOS no lo registra". Mismo
  // criterio para los 4 contadores de copias/impresiones.
  const usados_ = (colApertura, colCierre) =>
    (corte[colApertura] === undefined || corte[colCierre] === undefined)
      ? '' : (Number(corte[colCierre]) || 0) - (Number(corte[colApertura]) || 0);

  return json({
    ok: true,
    data: {
      v: 1,
      folio: corte.folio,
      id_empleado: corte.codigoEmpleado || '',
      fecha: String(corte.apertura || '').slice(0,10),
      hora_apertura: corte.apertura ? hhmm(corte.apertura) : '',
      hora_cierre: corte.cierre ? hhmm(corte.cierre) : '',
      // venta_turno = ventasTotal + recargasComisionTotal, misma fórmula que
      // usa NovaPOS en buildCorteQrPayload (ver LEGADOINTEGRALQRCIERRECAJA.md
      // sección 3) — la comisión de recargas sí es ganancia del negocio y
      // cuenta como venta del turno, igual que en "TOTAL A ENTREGAR" del
      // ticket impreso. Antes se regresaba solo ventasTotal (sin la
      // comisión) y esta función la mandaba aparte como "comision_recargas",
      // un campo que el contrato real del QR nunca tuvo.
      venta_turno: Math.round(((Number(corte.ventasTotal) || 0) + (Number(corte.recargasComisionTotal) || 0)) * 100) / 100,
      recargas_telefonicas: Math.round((Number(corte.recargasTotal) || 0) * 100) / 100,
      monto_entregado_admin: '',
      inventario_vendido: Math.round((Number(corte.ventasTotal) || 0) * 100) / 100,
      faltante: Math.round((Number(corte.faltante) || 0) * 100) / 100,
      merma: 0,
      copias_bn_usadas: usados_('contAperturaBn', 'contCierreBn'),
      copias_color_usadas: usados_('contAperturaColor', 'contCierreColor'),
      impresiones_bn_usadas: usados_('contAperturaImpBn', 'contCierreImpBn'),
      impresiones_color_usadas: usados_('contAperturaImpColor', 'contCierreImpColor'),
      copias_impresiones_vendido: corte.copiasImpresionesVendidasTotal === undefined ? '' : Math.round((Number(corte.copiasImpresionesVendidasTotal) || 0) * 100) / 100,
    },
  });
}

// Diagnóstico manual para buscar_folio — mismo patrón que
// debugEmpleadosLogin más abajo. El navegador solo reporta "Failed to
// fetch" cuando la conexión se cae, sin decir por qué; esto corre del
// lado del servidor y muestra el detalle real: si las Propiedades están
// configuradas, cuánto tardó la llamada a NovaPOS, qué código HTTP
// respondió y el arranque de su contenido (para distinguir un JSON válido
// de, por ejemplo, una página de acceso de Google si el Apps Script de
// NovaPOS no está publicado como "Cualquier usuario"). Cambia
// FOLIO_DE_PRUEBA por un folio real antes de correrla: selecciónala en el
// desplegable del editor → Ejecutar → Ver → Registros de ejecución.
function debugBuscarFolio() {
  const FOLIO_DE_PRUEBA = 'TCK-20260908-0026'; // el que sí trae copias/impresiones en el correo

  const props = PropertiesService.getScriptProperties();
  const novaUrl = props.getProperty('NOVAPOS_URL');
  const novaSecret = props.getProperty('NOVAPOS_SECRET');
  Logger.log('NOVAPOS_URL configurada: ' + (novaUrl ? 'sí (' + novaUrl + ')' : 'NO'));
  Logger.log('NOVAPOS_SECRET configurada: ' + (novaSecret ? 'sí' : 'NO'));
  if (!novaUrl || !novaSecret) { Logger.log('Faltan Propiedades del script — configúralas antes de seguir.'); return; }

  const url = novaUrl + '?action=get&sheet=cortes&secret=' + encodeURIComponent(novaSecret);
  const inicio = Date.now();
  let resp;
  try {
    resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {
    Logger.log('UrlFetchApp.fetch lanzó una excepción: ' + e);
    return;
  }
  Logger.log('Tardó ' + (Date.now() - inicio) + ' ms en responder');
  Logger.log('Código HTTP: ' + resp.getResponseCode());
  const texto = resp.getContentText();
  Logger.log('Longitud de la respuesta: ' + texto.length + ' caracteres');
  Logger.log('Primeros 300 caracteres: ' + texto.slice(0, 300));

  Logger.log('--- Resultado de handleBuscarFolio_("' + FOLIO_DE_PRUEBA + '") ---');
  Logger.log(handleBuscarFolio_(FOLIO_DE_PRUEBA).getContent());

  // Corte completo (todas sus columnas, tal cual las tiene NovaPOS) — cada
  // negocio corre su propia copia de nova_codigo.gs, así que el nombre
  // real de la columna de copias/impresiones puede no coincidir con
  // "copiasImpresionesVendidasTotal" que usa handleBuscarFolio_ arriba.
  // Esto ayuda a encontrar el nombre correcto sin adivinar.
  try {
    const body = JSON.parse(texto);
    const corte = (body.data || []).find(c => String(c.folio) === String(FOLIO_DE_PRUEBA));
    if (corte) {
      Logger.log('--- Corte completo (todas las columnas que trae NovaPOS) ---');
      Logger.log(JSON.stringify(corte, null, 2));
      const clavesCopias = Object.keys(corte).filter(k => /copia|impres/i.test(k));
      Logger.log('Columnas que mencionan "copia" o "impres": ' + JSON.stringify(clavesCopias));
    } else {
      Logger.log('No se encontró un corte con folio ' + FOLIO_DE_PRUEBA + ' en la respuesta.');
    }
  } catch (e) {
    Logger.log('No se pudo volver a parsear la respuesta para inspeccionar el corte: ' + e);
  }
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

// Busca una columna por nombre tolerando mayúsculas/minúsculas y espacios
// sobrantes — un encabezado escrito a mano en la hoja (en vez de generado
// por el propio NovaPOS) fácilmente trae "CodigoEmpleado" o "codigoEmpleado "
// en vez de "codigoEmpleado" exacto, y headers.indexOf() no perdona eso.
function indexOfHeader_(headers, nombre) {
  // \s+ también agarra saltos de línea DENTRO de la celda (Alt+Enter al
  // capturar el encabezado, p.ej. "NIP de Acceso\na Apps de la Empresa" en
  // Directorio_Alta_Empleados) — sin esto, un encabezado envuelto a mano en
  // dos líneas nunca calzaba con el nombre de columna que este script busca,
  // aunque la columna sí existiera, y la función regresaba lista vacía.
  const norm = s => String(s).replace(/\s+/g, ' ').trim().toLowerCase();
  const target = norm(nombre);
  return headers.findIndex(h => norm(h) === target);
}

// Lista de empleados para el login de Legado Integral. Usa la hoja
// "Directorio_Alta_Empleados" de ESTA hoja de cálculo — la cuenta propia
// del operador de Legado Integral, no la de NovaPOS de cada negocio
// cliente (son cuentas y hojas de cálculo distintas: NovaPOS nunca ve
// estos datos, y este script nunca ve los de NovaPOS). Esa hoja trae 3
// filas de título/instrucciones antes de los encabezados reales, así que
// se busca la fila que contiene "ID Empleado" en vez de asumir que es la
// primera. "id" es el valor de la columna "ID Empleado" (p.ej. "EMP-001")
// — es el mismo valor que trae el campo "id_empleado" del QR de cierre de
// turno que genera NovaPOS, así que Legado Integral puede comparar sesión
// vs. ticket escaneado directamente. El NIP para entrar a Legado Integral
// vive ahí mismo, columna "NIP de Acceso a Apps de la Empresa" — es
// independiente del NIP de ventas que cada vendedor usa dentro de NovaPOS.
// Empleados marcados "INACTIVO" quedan fuera del login.
function getEmpleadosLogin_() {
  const rows = getSheet('Directorio_Alta_Empleados').getDataRange().getValues();
  const headerRowIdx = rows.findIndex(r => indexOfHeader_(r, 'ID Empleado') >= 0);
  if (headerRowIdx < 0) return [];
  const headers = rows[headerRowIdx];
  const idCol = indexOfHeader_(headers, 'ID Empleado');
  const nombreCol = indexOfHeader_(headers, 'Nombre Completo');
  const estatusCol = indexOfHeader_(headers, 'Estatus');
  const nipCol = indexOfHeader_(headers, 'NIP de Acceso a Apps de la Empresa');
  if (idCol < 0 || nombreCol < 0 || nipCol < 0) return [];
  return rows.slice(headerRowIdx + 1)
    .filter(r => String(r[idCol] || '').trim() && String(r[nipCol] || '').trim())
    .filter(r => estatusCol < 0 || String(r[estatusCol] || '').trim().toUpperCase() !== 'INACTIVO')
    .map(r => ({ id: String(r[idCol]).trim(), nombre: r[nombreCol], _nip: String(r[nipCol]).trim() }));
}

// Diagnóstico manual: selecciona esta función en el desplegable del editor
// de Apps Script y Ejecutar, luego revisa Ver → Registros de ejecución.
// Muestra los encabezados detectados en "Directorio_Alta_Empleados" y la
// lista final que action=empleados le manda a Legado Integral — útil para
// confirmar si el problema es un encabezado con nombre distinto, o
// simplemente que no hay ningún empleado activo con NIP capturado todavía.
// SIN guion bajo al final a propósito: Apps Script oculta del desplegable
// "Ejecutar" cualquier función cuyo nombre termine en "_" (la convención
// que usa este script para marcar funciones internas/privadas) — con guion
// bajo, esta función nunca aparecería como opción para correr manualmente.
function debugEmpleadosLogin() {
  const rows = getSheet('Directorio_Alta_Empleados').getDataRange().getValues();
  const headerRowIdx = rows.findIndex(r => indexOfHeader_(r, 'ID Empleado') >= 0);
  Logger.log('Fila de encabezados detectada: ' + (headerRowIdx < 0 ? 'NO ENCONTRADA' : (headerRowIdx + 1)));
  Logger.log('Encabezados: ' + JSON.stringify(headerRowIdx < 0 ? [] : rows[headerRowIdx]));
  Logger.log('Filas de datos: ' + Math.max(0, rows.length - 1 - headerRowIdx));
  Logger.log('Empleados que vería Legado Integral: ' + JSON.stringify(getEmpleadosLogin_()));
}

// Valida el NIP contra el directorio de empleados propio de Legado Integral
// (columna "NIP de Acceso a Apps de la Empresa" en Directorio_Alta_Empleados,
// ver getEmpleadosLogin_ arriba) y regresa su progreso, incluyendo el
// cálculo real de los 4 bonos (ver resumenMes_ más abajo).
function handleResumenLogin_(codigoEmpleado, nip) {
  if (!codigoEmpleado || !nip) return json({ok:false, error:'Falta empleado o NIP'});
  const empleado = getEmpleadosLogin_().find(e => String(e.id) === String(codigoEmpleado));
  if (!empleado) return json({ok:false, error:'Empleado no encontrado'});
  if (String(empleado._nip) !== String(nip)) {
    return json({ok:false, error:'NIP incorrecto'});
  }

  return json({
    ok: true,
    empleado: { id: empleado.id, nombre: empleado.nombre },
    semana: resumenSemana_(codigoEmpleado),
    mes: resumenMes_(codigoEmpleado),
  });
}

// -------------------------------------------------------------------------
// Configuración de los 4 bonos — todo editable sin tocar código: agrega o
// edita la fila correspondiente en la hoja "config" (columnas "key","value")
// para cambiar cualquiera de estos valores. Si la fila no existe todavía se
// usa el valor por defecto de BONOS_DEFAULTS. Reglas acordadas el
// 2026-09-10; ver conversación/PR para el detalle de cada una.
// -------------------------------------------------------------------------
const BONOS_DEFAULTS = {
  // Hora (0-23) que separa el turno "dia" del turno "noche" según la hora
  // de apertura del corte — el relevo real ocurre a mitad del día, entre el
  // fin del turno día (8:00) y el cierre del turno noche (20:00), y no hay
  // una hora fija acordada para ese relevo todavía.
  turno_corte_hora: 14,
  // Meta diaria de venta (venta_turno + copias_impresiones_vendido) del
  // PRIMER mes del que hay historial, por turno — de ahí en adelante la
  // meta se recalcula sola en metas_ventas (ver asegurarMetaVentas_).
  meta_ventas_dia_inicial: 700,
  meta_ventas_noche_inicial: 400,
  // % de crecimiento que se le suma al promedio diario real del mes
  // anterior para sugerir la meta del mes actual. ⚠️ Número de arranque —
  // AJÚSTALO en la hoja "config" (key "meta_ventas_crecimiento_pct") al
  // porcentaje que el negocio realmente quiera exigir mes a mes.
  meta_ventas_crecimiento_pct: 5,
  // Días del mes en que hay que alcanzar la meta diaria de tu turno.
  bono_ventas_dias_meta: 20,
  bono_ventas_monto: 600,
  bono_puntualidad_apertura_hhmm: '08:00',
  bono_puntualidad_cierre_hhmm: '20:00',
  bono_puntualidad_tolerancia_min: 15,
  // Retardos (no justificados) tolerados por semana antes de perder el
  // bono — no es "cuántos se permiten en total", es "más de esto en UNA
  // sola semana ya lo pierde".
  bono_puntualidad_retardos_max_semana: 1,
  bono_puntualidad_monto: 400,
  // Suma de faltantes (no sobrantes) tolerada en TODO el mes.
  bono_caja_tolerancia_mensual: 100,
  // Turnos con faltante O sobrante (cualquiera de los dos) tolerados por
  // semana — más de esto en una semana lo pierde aunque el mes siga dentro
  // de la tolerancia de $100.
  bono_caja_max_incidencias_semana: 1,
  bono_caja_monto: 250,
  bono_inventario_monto: 250,
  // ⚠️ "bono_inventario_tolerancia_mensual" TODAVÍA NO EXISTE aquí a
  // propósito: falta acordar cuánta merma combinada de los 2 turnos (ver
  // resumenMes_) se tolera al mes. Sin ese número no se puede decidir si se
  // gana el bono o no — se calcula el total real para que el dueño lo vea,
  // pero el bono se queda en $0 hasta que se capture ese valor en "config".
};

function getBonoConfig_(key) {
  const valor = getConfigMap_()[key];
  return (valor === undefined || valor === '') ? BONOS_DEFAULTS[key] : valor;
}

// Ticket sin hora de apertura (viejo, o captura manual incompleta) regresa
// '' — no se puede clasificar ni evaluar puntualidad para ese turno.
function detectarTurno_(horaApertura) {
  const hora = parseInt(String(horaApertura || '').split(':')[0], 10);
  if (isNaN(hora)) return '';
  return hora < Number(getBonoConfig_('turno_corte_hora')) ? 'dia' : 'noche';
}

// Minutos de diferencia entre una hora real "HH:MM" y una esperada
// "HH:MM" — positivo cuando la real es DESPUÉS de la esperada (retraso).
function minutosDeRetraso_(horaReal, horaEsperada) {
  const r = String(horaReal || '').split(':').map(Number);
  const e = String(horaEsperada || '').split(':').map(Number);
  if (r.length < 2 || e.length < 2 || r.some(isNaN) || e.some(isNaN)) return 0;
  return (r[0]*60 + r[1]) - (e[0]*60 + e[1]);
}

function filaAObjeto_(headers, row) {
  return Object.fromEntries(headers.map((h,i) => [h, row[i]]));
}

// "YYYY-MM-DD" a partir de cualquier valor de fecha — una celda de Google
// Sheets con pinta de fecha ("2026-09-10") se puede guardar como texto o
// autoconvertir a un valor de fecha real según el formato de la columna; a
// diferencia de recortar el string con .slice(0,10), esto da el mismo
// resultado sin importar cuál de los dos sea.
function fechaYMD_(valor) {
  return Utilities.formatDate(new Date(valor), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// lunes=0 ... domingo=6 — mismo criterio que ya usaba resumenSemana_/
// inicioSemana_, para que "dia_descanso" en "turnos_asignados" se
// interprete igual en todos lados.
function diaSemanaLunes0_(fecha) {
  return (new Date(fecha).getDay() + 6) % 7;
}

// Todas las filas de "turnos_asignados" (quién tiene cada turno, y de qué
// fecha a qué fecha — "hasta" en blanco significa "sigue vigente"). Se lee
// completa cada vez porque son pocas filas (2 turnos, más el historial de
// cuando alguien renuncia); no hace falta indexarla.
function asignacionesTurno_() {
  const sheet = getSheet('turnos_asignados');
  const headers = ensureHeaders_(sheet, 'turnos_asignados');
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => filaAObjeto_(headers, r));
}

// Qué asignación (turno + día de descanso) tenía vigente este empleado en
// una fecha dada — null si ese día no hay ninguna fila que lo cubra (roster
// no capturado todavía, o la fecha cae fuera de cualquier rango desde/hasta).
function asignacionVigenteEn_(asignaciones, codigoEmpleado, fecha) {
  return asignaciones.find(a => {
    if (String(a.codigoEmpleado) !== String(codigoEmpleado)) return false;
    const desde = a.desde ? new Date(a.desde) : new Date(0);
    const hasta = a.hasta ? new Date(a.hasta) : null;
    return fecha >= desde && (!hasta || fecha <= hasta);
  });
}

// Todos los turnos confirmados (de cualquier empleado) en un rango de
// fechas — usa "fecha" (la del turno) y no "confirmado_en" (cuándo se
// mandó el acuse), que es lo que se necesita para agrupar por semana/mes
// del CALENDARIO DE TRABAJO real, no por cuándo el empleado alcanzó a
// escanear su ticket.
function legadoTurnosEntre_(desde, hasta) {
  const sheet = getSheet('legado_turnos');
  const headers = ensureHeaders_(sheet, 'legado_turnos');
  const rows = sheet.getDataRange().getValues();
  const fechaCol = headers.indexOf('fecha');
  return rows.slice(1)
    .filter(r => r[fechaCol])
    .map(r => filaAObjeto_(headers, r))
    .filter(t => { const d = new Date(t.fecha); return d >= desde && d < hasta; });
}

// Turnos que el propio empleado ya confirmó en Legado Integral (hoja
// "legado_turnos", la que llena handleConfirmarTurno_ arriba) dentro de un
// rango de fechas — NO la hoja "cortes" de NovaPOS, que vive en la cuenta
// de cada negocio cliente y esta hoja de cálculo nunca llega a ver.
function turnosDelEmpleadoEntre_(codigoEmpleado, desde, hasta) {
  return legadoTurnosEntre_(desde, hasta).filter(t => String(t.codigoEmpleado) === String(codigoEmpleado));
}

// Lunes de la semana calendario (00:00) a la que pertenece una fecha, como
// string — se usa para agrupar retardos/incidencias de caja por semana.
function inicioSemana_(fecha) {
  const d = new Date(fecha); d.setHours(0,0,0,0);
  const diaSemana = (d.getDay() + 6) % 7; // lunes=0 ... domingo=6
  d.setDate(d.getDate() - diaSemana);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function resumenSemana_(codigoEmpleado) {
  const hoy = new Date();
  const diaSemana = (hoy.getDay() + 6) % 7; // lunes=0 ... domingo=6
  const inicio = new Date(hoy); inicio.setHours(0,0,0,0); inicio.setDate(hoy.getDate() - diaSemana);
  const fin = new Date(inicio); fin.setDate(inicio.getDate() + 7);

  const turnos = turnosDelEmpleadoEntre_(codigoEmpleado, inicio, fin)
    .map(t => Object.assign({}, t, { faltante: Number(t.faltante) || 0 }));
  const conFaltante = turnos.filter(t => t.faltante > 0);

  const puntos_favor = [];
  const areas_oportunidad = [];
  if (turnos.length) puntos_favor.push(turnos.length + ' turno(s) confirmado(s) esta semana.');
  if (turnos.length && !conFaltante.length) puntos_favor.push('Sin faltantes de caja esta semana.');
  else conFaltante.forEach(t => areas_oportunidad.push('Faltante el ' + fmtDMY_(t.fecha) + ': $' + t.faltante.toFixed(2)));
  if (!turnos.length) areas_oportunidad.push('Todavía no confirmas ningún turno esta semana.');

  return {
    rango: { inicio: fmtDMY_(inicio), fin: fmtDMY_(new Date(fin - 1)) },
    puntos_favor,
    areas_oportunidad,
  };
}

// Meta diaria de venta (por turno) del mes indicado — la crea la primera
// vez que alguien la necesita ese mes (ver resumenMes_) y de ahí en
// adelante ya no se recalcula sola: el dueño la revisa a mano en la hoja
// "metas_ventas". Si "aceptada" es TRUE se usa "meta_sugerida" tal cual; si
// no, se usa lo que capture en "meta_manual" — mientras el dueño no revise
// ninguna de las dos, se usa la sugerida (para no bloquear el cálculo).
function asegurarMetaVentas_(mesYYYYMM, turno) {
  // Dos empleados pueden abrir "Mi progreso" casi al mismo tiempo el primer
  // día del mes — sin candado, ambos podrían no encontrar la fila todavía y
  // cada uno insertar la suya, duplicando la meta de ese turno/mes.
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const sheet = getSheet('metas_ventas');
    const headers = ensureHeaders_(sheet, 'metas_ventas');
    const rows = sheet.getDataRange().getValues();
    const mesCol = headers.indexOf('mes');
    const turnoCol = headers.indexOf('turno');
    const existente = rows.findIndex((r,i) => i>0 && String(r[mesCol]) === mesYYYYMM && String(r[turnoCol]) === turno);
    if (existente > 0) return filaAObjeto_(headers, rows[existente]);

    const registro = {
      id: Utilities.getUuid(),
      mes: mesYYYYMM,
      turno,
      meta_sugerida: calcularMetaVentaSugerida_(mesYYYYMM, turno),
      aceptada: '',
      meta_manual: '',
      generado_en: new Date().toISOString(),
    };
    sheet.appendRow(headers.map(h => registro[h] ?? ''));
    return registro;
  } finally {
    lock.releaseLock();
  }
}

// meta_sugerida = promedio diario real de venta del turno el mes anterior
// (venta_turno + copias_impresiones_vendido, ver LEGADOINTEGRALQRCIERRECAJA.md)
// más el % de crecimiento configurado. Si no hay ningún turno confirmado el
// mes anterior (negocio nuevo en Legado Integral, o primer mes de este
// turno), se usa el monto inicial acordado ($700 día / $400 noche).
function calcularMetaVentaSugerida_(mesYYYYMM, turno) {
  const [anio, mes] = mesYYYYMM.split('-').map(Number);
  const mesAnterior = new Date(anio, mes - 2, 1); // Date usa meses 0-11; "mes-2" retrocede uno
  const desde = new Date(mesAnterior.getFullYear(), mesAnterior.getMonth(), 1);
  const hasta = new Date(mesAnterior.getFullYear(), mesAnterior.getMonth() + 1, 1);

  const turnosMesAnterior = legadoTurnosEntre_(desde, hasta).filter(t => t.turno === turno);
  if (!turnosMesAnterior.length) {
    return Number(getBonoConfig_(turno === 'dia' ? 'meta_ventas_dia_inicial' : 'meta_ventas_noche_inicial'));
  }
  const totalVentas = turnosMesAnterior.reduce((s,t) => s + (Number(t.venta_turno)||0) + (Number(t.copias_impresiones_vendido)||0), 0);
  const promedioDiario = totalVentas / turnosMesAnterior.length;
  const pct = Number(getBonoConfig_('meta_ventas_crecimiento_pct'));
  return Math.round(promedioDiario * (1 + pct/100));
}

function metaFinalDe_(meta) {
  const aceptada = meta.aceptada === true || String(meta.aceptada).trim().toUpperCase() === 'TRUE';
  const manual = meta.meta_manual;
  const tieneManual = manual !== '' && manual !== undefined && manual !== null && !isNaN(Number(manual));
  return (!aceptada && tieneManual) ? Number(manual) : Number(meta.meta_sugerida);
}

// Cálculo real de los 4 bonos del mes en curso para un empleado — ver la
// tabla de reglas en BONOS_DEFAULTS y el comentario de cada bono abajo.
function resumenMes_(codigoEmpleado) {
  const hoy = new Date();
  const anio = hoy.getFullYear(), mesNum = hoy.getMonth() + 1;
  const mesStr = anio + '-' + ('0'+mesNum).slice(-2);
  const inicio = new Date(anio, mesNum - 1, 1);
  const fin = new Date(anio, mesNum, 1);

  const misTurnos = turnosDelEmpleadoEntre_(codigoEmpleado, inicio, fin);

  // ---------- Bono Ventas ----------
  // Se cumple si, en "bono_ventas_dias_meta" días del mes (por defecto 20),
  // la venta de TU turno (venta_turno + copias_impresiones_vendido) alcanzó
  // la meta diaria de ese turno (día u noche) — la meta se recalcula cada
  // mes (ver asegurarMetaVentas_/calcularMetaVentaSugerida_).
  const metasPorTurno = {
    dia: metaFinalDe_(asegurarMetaVentas_(mesStr, 'dia')),
    noche: metaFinalDe_(asegurarMetaVentas_(mesStr, 'noche')),
  };
  const diasMetaRequeridos = Number(getBonoConfig_('bono_ventas_dias_meta'));
  const diasConMeta = misTurnos.filter(t => {
    if (!t.turno || metasPorTurno[t.turno] === undefined) return false;
    const ventaDelTurno = (Number(t.venta_turno)||0) + (Number(t.copias_impresiones_vendido)||0);
    return ventaDelTurno >= metasPorTurno[t.turno];
  }).length;
  const cumpleVentas = diasConMeta >= diasMetaRequeridos;

  // ---------- Bono Puntualidad ----------
  // "0 retardos en el mes" con tolerancia de 1 retardo (no justificado) por
  // semana calendario — más de 1 en UNA sola semana lo pierde. Los
  // retardos marcados "retardo_justificado" a mano por el dueño (excepción
  // de relevo) no cuentan.
  const toleranciaRetardosSemana = Number(getBonoConfig_('bono_puntualidad_retardos_max_semana'));
  const retardosPorSemana = {};
  misTurnos.forEach(t => {
    const justificado = t.retardo_justificado === true || String(t.retardo_justificado).trim().toUpperCase() === 'TRUE';
    const esRetardo = (t.retardo_apertura === true || t.retardo_cierre === true) && !justificado;
    if (!esRetardo) return;
    const semana = inicioSemana_(t.fecha);
    retardosPorSemana[semana] = (retardosPorSemana[semana] || 0) + 1;
  });

  // Candado de "0 faltas": un día "esperado" es un día del mes (hasta hoy,
  // nunca días futuros) en que, según "turnos_asignados", a este empleado
  // le tocaba trabajar y no era su día de descanso. Si ese día no tiene un
  // turno confirmado en legado_turnos, es una falta — y una sola falta en
  // el mes hace perder el bono, sin importar los retardos. Si no hay
  // ninguna asignación capturada para este empleado, no se puede evaluar
  // (se deja pasar, en vez de inventar una falta que no se puede probar).
  const misAsignaciones = asignacionesTurno_().filter(a => String(a.codigoEmpleado) === String(codigoEmpleado));
  const fechasConTurno = new Set(misTurnos.map(t => fechaYMD_(t.fecha)));
  let faltas = 0;
  if (misAsignaciones.length) {
    const finEvaluacion = new Date(anio, mesNum - 1, hoy.getDate() + 1); // nunca evalúa días futuros
    for (let d = new Date(inicio); d < finEvaluacion; d.setDate(d.getDate() + 1)) {
      const asignacion = asignacionVigenteEn_(misAsignaciones, codigoEmpleado, d);
      if (!asignacion) continue; // sin asignación vigente ese día — no se puede saber si le tocaba
      // "" (dia_descanso todavía no capturado) NO debe leerse como "0"
      // (lunes) — Number('') es 0 en JS, y eso pondría a todos con el
      // campo en blanco un día de descanso falso los lunes.
      const diaDescanso = asignacion.dia_descanso === '' || asignacion.dia_descanso === undefined || asignacion.dia_descanso === null
        ? null : Number(asignacion.dia_descanso);
      if (diaDescanso !== null && diaSemanaLunes0_(d) === diaDescanso) continue; // su día de descanso
      if (!fechasConTurno.has(fechaYMD_(d))) faltas++;
    }
  }

  const cumplePuntualidad = faltas === 0 && !Object.values(retardosPorSemana).some(n => n > toleranciaRetardosSemana);

  // ---------- Bono Caja ----------
  // Tolerancia de $100 de faltante (no sobrante) en TODO el mes, y no más
  // de 1 turno con faltante o sobrante por semana.
  const tolCajaMensual = Number(getBonoConfig_('bono_caja_tolerancia_mensual'));
  const maxIncidenciasSemana = Number(getBonoConfig_('bono_caja_max_incidencias_semana'));
  const totalFaltanteMes = misTurnos.reduce((s,t) => s + Math.max(0, Number(t.faltante)||0), 0);
  const incidenciasPorSemana = {};
  misTurnos.forEach(t => {
    if ((Number(t.faltante)||0) !== 0) {
      const semana = inicioSemana_(t.fecha);
      incidenciasPorSemana[semana] = (incidenciasPorSemana[semana]||0) + 1;
    }
  });
  const cumpleCaja = totalFaltanteMes <= tolCajaMensual && !Object.values(incidenciasPorSemana).some(n => n > maxIncidenciasSemana);

  // ---------- Bono Inventario ----------
  // Se mide combinando AMBOS turnos del día (día + noche): los 2 hacen su
  // propio conteo de inventario al entrar, así que no se puede saber a
  // cuál de los dos le falta algo — por eso se suma la merma de los 2
  // turnos de cada fecha. ⚠️ Todavía falta acordar cuánta merma combinada
  // se tolera al mes ("bono_inventario_tolerancia_mensual" en "config") —
  // sin ese número no se puede decidir si se gana o no, así que el bono se
  // queda en $0 aunque aquí ya se calcula y expone el total real del mes.
  const mermaPorDia = {};
  legadoTurnosEntre_(inicio, fin).forEach(t => {
    const clave = fechaYMD_(t.fecha);
    mermaPorDia[clave] = (mermaPorDia[clave] || 0) + (Number(t.merma)||0);
  });
  const mermaTotalMes = Object.values(mermaPorDia).reduce((s,n) => s+n, 0);
  const tolInventarioMensual = getBonoConfig_('bono_inventario_tolerancia_mensual');
  const tolInventarioDefinida = tolInventarioMensual !== undefined && tolInventarioMensual !== '';
  const cumpleInventario = tolInventarioDefinida && mermaTotalMes <= Number(tolInventarioMensual);

  const bono_ventas = cumpleVentas ? Number(getBonoConfig_('bono_ventas_monto')) : 0;
  const bono_puntualidad = cumplePuntualidad ? Number(getBonoConfig_('bono_puntualidad_monto')) : 0;
  const bono_caja = cumpleCaja ? Number(getBonoConfig_('bono_caja_monto')) : 0;
  const bono_inventario = cumpleInventario ? Number(getBonoConfig_('bono_inventario_monto')) : 0;
  const elegible_premio_maximo = cumpleVentas && cumplePuntualidad && cumpleCaja && cumpleInventario;

  const pendientes = [];
  if (!tolInventarioDefinida) pendientes.push('Bono Inventario: falta definir cuánta merma combinada (de los 2 turnos) se tolera al mes.');
  if (!misAsignaciones.length) pendientes.push('Bono Puntualidad: agrega tu turno y tu día de descanso en la hoja "turnos_asignados" para que se pueda revisar el candado de "0 faltas".');

  return {
    mes: hoy.toLocaleDateString('es-MX', { month: 'long' }),
    elegible_premio_maximo,
    mensaje: pendientes.join(' '),
    bono_ventas,
    bono_puntualidad,
    bono_caja,
    bono_inventario,
    total_bonos: bono_ventas + bono_puntualidad + bono_caja + bono_inventario,
    dias_falta: faltas,
    dias_con_meta: diasConMeta,
    dias_meta_requeridos: diasMetaRequeridos,
    merma_total_mes: Math.round(mermaTotalMes * 100) / 100,
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
  //
  // Columnas y nombres tal cual el contrato real del QR/confirmar_turno
  // (ver LEGADOINTEGRALQRCIERRECAJA.md, secciones 3 y 5.2) — no hay un campo
  // "comision_recargas" separado (venta_turno ya la incluye) y el nombre es
  // "copias_impresiones_vendido" (singular). copias_bn_usadas/
  // copias_color_usadas/impresiones_bn_usadas/impresiones_color_usadas son
  // lecturas del medidor físico (unidades), no dinero.
  //
  // turno/retardo_apertura/retardo_cierre: los calcula handleConfirmarTurno_
  // (ver detectarTurno_/minutosDeRetraso_) a partir de hora_apertura/
  // hora_cierre — no los manda la app. motivo_retardo lo captura el
  // empleado en Legado Integral cuando hay retraso. retardo_justificado lo
  // marca el DUEÑO a mano en esta hoja (checkbox TRUE/FALSE) para las
  // excepciones de relevo — nunca lo pisa una reconfirmación del mismo
  // folio (ver handleConfirmarTurno_). Junto con faltante, merma y
  // venta_turno son la materia prima que usa resumenMes_ para calcular los
  // 4 bonos (ver también la hoja "metas_ventas" más abajo).
  legado_turnos: ['id','folio','codigoEmpleado','nombreEmpleado','fecha','hora_apertura','hora_cierre','venta_turno','recargas_telefonicas','monto_entregado_admin','inventario_vendido','faltante','merma','copias_bn_usadas','copias_color_usadas','impresiones_bn_usadas','impresiones_color_usadas','copias_impresiones_vendido','turno','retardo_apertura','retardo_cierre','motivo_retardo','retardo_justificado','confirmado_en'],
  // Meta diaria de venta por turno (día/noche) y mes — la genera sola
  // asegurarMetaVentas_ la primera vez que se necesita ese mes
  // (meta_sugerida = promedio diario real del mes anterior + % de
  // crecimiento de "config", o el monto inicial acordado si no hay mes
  // anterior). El DUEÑO la revisa aquí a mano: "aceptada" = TRUE usa
  // meta_sugerida tal cual; si no, se usa lo que capture en "meta_manual".
  // Ver metaFinalDe_/resumenMes_ en el código.
  metas_ventas: ['id','mes','turno','meta_sugerida','aceptada','meta_manual','generado_en'],
  // Quién tiene asignado cada uno de los 2 turnos (día/noche) y su día de
  // descanso — el DUEÑO la captura y mantiene a mano. Siempre son las
  // mismas 2 personas salvo que alguien renuncie: en ese caso, en vez de
  // editar la fila existente, ciérrala poniéndole "hasta" (el último día
  // que trabajó) y agrega una fila nueva para quien la sustituye con su
  // propio "desde" — así el historial de meses anteriores no cambia.
  // "hasta" en blanco = sigue vigente. "dia_descanso" es un número
  // lunes=0 ... domingo=6 (ej. domingo=6). Ejemplo: el empleado del turno
  // noche descansa los domingos — su turno empieza domingo a las 20:00 y
  // termina lunes a las 8:00, así que ese domingo simplemente no habrá
  // ningún turno confirmado con fecha domingo, y no debe contar como falta.
  // Se usa para: (1) detectar el turno de cada confirmación con más
  // confianza que el criterio por hora (ver handleConfirmarTurno_), y (2)
  // el candado de "0 faltas" del Bono Puntualidad (ver resumenMes_): un día
  // que le tocaba trabajar (no es su dia_descanso) y no tiene turno
  // confirmado en legado_turnos cuenta como falta.
  turnos_asignados: ['id','turno','codigoEmpleado','dia_descanso','desde','hasta'],
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
