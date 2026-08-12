/**
 * ===================================================================
 *  ATELIER · API DE PEDIDOS — v4.0 (Sheet nuevo, estructura 18 tablas)
 *  Login, pedidos, estados, pagos, facturas FIGS, courier.
 * ===================================================================
 */

function testSheetAccess() {
  try {
    const ss = SpreadsheetApp.openById('1uZOHOwrAltZn6gltvhrtqeP1NhSYXagKCMQFvKdXlrI');
    Logger.log('✅ Sheet abierto OK');
    Logger.log('Nombre: ' + ss.getName());
    const sheets = ss.getSheets().map(s => s.getName());
    Logger.log('Hojas (' + sheets.length + '): ' + sheets.join(', '));
    return 'OK';
  } catch (err) {
    Logger.log('❌ ERROR: ' + err.message);
    return 'ERROR: ' + err.message;
  }
}

// ============ CONFIGURACIÓN ============
const SHEET_ID = '1uZOHOwrAltZn6gltvhrtqeP1NhSYXagKCMQFvKdXlrI';

const TABS = {
  ordenes:      'TablaOrdenes',
  items:        'TablaItems',
  costos:       'TablaCostos',
  pagos:        'TablaPagos',
  descuentos:   'TablaDescuentos',
  clientes:     'TablaClientes',
  productos:    'TablaProductos',
  sets:         'TablaSets',
  industrias:   'TablaIndustrias',
  regalos:      'TablaRegalos',
  materiales:   'TablaMateriales',
  tarifas:      'TablaTarifasCourier',
  estados:      'TablaEstados',
  usuarios:     'TablaUsuarios',
  log:          'TablaLog',
  lotes:        'TablaLotesStock',
  movimientos:  'TablaMovimientosStock',
  empaque:      'TablaEmpaqueEstandar',
  setEmpaque: 'TablaSetEmpaque',
  couriers: 'TablaCouriers',
  combos: 'TablaCombos',
  colores: 'TablaColores',
  gastosFijos:  'TablaGastosFijos',
  saldos: 'TablaSaldos',
};

const VALID_STATES = [
  'HACER PEDIDO', 'PEDIDO HECHO', 'EN TRANSITO A FL', 'CON FREDDY',
  'EN CAMINO A EC', 'EN BODEGA EC', 'LISTO PARA ENVIAR', 'ENTREGADO',
  'CANCELADO', 'ENTREGA PARCIAL'
];

const METODOS_PAGO = [
  'Transferencia', 'Efectivo', 'Tarjeta de crédito', 'Deuna', 'PayPal',
  'Crédito aplicado', 'Devolución (transferencia)', 'Devolución (crédito)'
];

const COL_ESTATUS_ENVIO = 4;

// ============ HELPERS ============

function normalize(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function getSpreadsheet() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function readSheetAsObjects(tabName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error(`Hoja no encontrada: ${tabName}`);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);
  return rows
    .filter(row => row[0] !== '' && row[0] !== null)
    .map((row, idx) => {
      const obj = { _rowNum: idx + 2 };
      headers.forEach((header, i) => {
        let value = row[i];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, 'GMT-5', 'yyyy-MM-dd');
        }
        obj[header] = value;
      });
      return obj;
    });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(message, code = 400) {
  return jsonResponse({ ok: false, error: message, code: code });
}

function findOrderRow(ordenId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.ordenes);
  const data = sheet.getRange('A:A').getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(ordenId).trim()) return i + 1;
  }
  return -1;
}

function logChange(usuario, ordenId, accion, detalle) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(TABS.log);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.log);
    sheet.appendRow(['FECHA_HORA', 'USUARIO', 'ACCION', 'ORDEN_ID', 'DETALLES']);
    sheet.getRange('A1:E1').setFontWeight('bold');
  }
  const timestamp = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([timestamp, usuario, accion, ordenId, detalle]);
}

function verifyAdmin(usuario) {
  if (!usuario) return false;
  const usuarios = readSheetAsObjects(TABS.usuarios);
  const found = usuarios.find(u =>
    String(u.USUARIO).toLowerCase().trim() === String(usuario).toLowerCase().trim()
  );
  return found && String(found.ROL).toLowerCase().trim() === 'admin';
}

// ============ HELPERS DE LOCK Y IDS ============

function withLock(fn) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('No se pudo obtener lock del documento. Intenta de nuevo.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function nextLoteId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.lotes);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'LOTE-00001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  ids.forEach(id => {
    const m = String(id).match(/LOTE-(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  return 'LOTE-' + String(maxNum + 1).padStart(5, '0');
}

function nextOrdenId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.ordenes);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '#0001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  ids.forEach(id => {
    const m = String(id).match(/#(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  return '#' + String(maxNum + 1).padStart(4, '0');
}

function nextClienteId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.clientes);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'C-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  ids.forEach(id => {
    const m = String(id).match(/C-(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  return 'C-' + String(maxNum + 1).padStart(3, '0');
}

function nextDescuentoId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.descuentos);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'DESC-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  ids.forEach(id => {
    const m = String(id).match(/DESC-(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  return 'DESC-' + String(maxNum + 1).padStart(3, '0');
}

function nextMovId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.movimientos);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'MOV-00001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  ids.forEach(id => {
    const m = String(id).match(/MOV-(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  return 'MOV-' + String(maxNum + 1).padStart(5, '0');
}

function canCambiarEstado(usuario) {
  if (!usuario) return false;
  const usuarios = readSheetAsObjects(TABS.usuarios);
  const found = usuarios.find(u =>
    String(u.USUARIO).toLowerCase().trim() === String(usuario).toLowerCase().trim()
  );
  if (!found) return false;
  const rol = String(found.ROL).toLowerCase().trim();
  return rol === 'admin' || rol === 'bodega';
}

function canRegistrarPagos(usuario) {
  return verifyAdmin(usuario);
}

// Permite admin O bodega (para acciones operativas que Sebas ya maneja)
function canOperar(usuario) {
  if (!usuario) return false;
  const usuarios = readSheetAsObjects(TABS.usuarios);
  const found = usuarios.find(u =>
    String(u.USUARIO).toLowerCase().trim() === String(usuario).toLowerCase().trim()
  );
  if (!found) return false;
  const rol = String(found.ROL).toLowerCase().trim();
  return rol === 'admin' || rol === 'bodega';
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function calcularEstadoPago(totalVenta, totalPagado) {
  if (totalPagado === 0) return 'Sin pago';
  if (totalPagado < totalVenta) return 'Parcial';
  if (totalPagado === totalVenta) return 'Pagado total';
  return 'Sobrepago';
}

// Escribe la fecha de hoy en F_ENTREGA_REAL (solo si está vacía).
// Busca la columna por su encabezado, así no importa en qué posición esté.
function marcarFechaEntrega(sheetOrdenes, rowNum) {
  const headers = sheetOrdenes.getRange(1, 1, 1, sheetOrdenes.getLastColumn()).getValues()[0];
  let col = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toUpperCase() === 'F_ENTREGA_REAL') { col = i + 1; break; }
  }
  if (col === -1) return; // si no existe la columna, no hace nada
  const actual = sheetOrdenes.getRange(rowNum, col).getValue();
  if (actual) return;     // ya tiene fecha, no la sobrescribe
  const hoy = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
  sheetOrdenes.getRange(rowNum, col).setValue(hoy);
}

// ============ ENDPOINTS ============

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'ping';

    if (action === 'ping') {
      return jsonResponse({
        ok: true,
        message: 'API v4.0 OK (Sheet nuevo)',
        timestamp: new Date().toISOString(),
        sheetId: SHEET_ID
      });
    }

    // ── DIAGNÓSTICO TEMPORAL — borrar después de confirmar ──
    if (action === 'testGetPedido') {
      const debug = { step: 'start' };
      try {
        debug.step = 'leer ordenes';
        const ordenes = readSheetAsObjects(TABS.ordenes);
        debug.countOrdenes = ordenes.length;

        debug.step = 'buscar orden';
        const orden = ordenes.find(o => String(o.ORDEN_ID).trim() === String(e.parameter.id).trim());
        debug.ordenEncontrada = !!orden;

        if (orden) {
          debug.step = 'leer items';
          const items = readSheetAsObjects(TABS.items);
          debug.countItems = items.length;

          debug.step = 'leer costos';
          const costos = readSheetAsObjects(TABS.costos);
          debug.countCostos = costos.length;

          debug.step = 'leer pagos';
          const pagos = readSheetAsObjects(TABS.pagos);
          debug.countPagos = pagos.length;

          debug.step = 'leer descuentos';
          const descuentos = readSheetAsObjects(TABS.descuentos);
          debug.countDescuentos = descuentos.length;

          debug.step = 'leer clientes';
          const clientes = readSheetAsObjects(TABS.clientes);
          debug.countClientes = clientes.length;

          debug.step = 'enrichOrden';
          const enriched = enrichOrden(orden, items, costos, pagos, descuentos, clientes);
          debug.step = 'stringify';
          const jsonStr = JSON.stringify(enriched);
          debug.jsonLength = jsonStr.length;
          debug.itemsCount = enriched.items.length;
          debug.costosCount = enriched.costos.length;
          debug.pagosCount = enriched.pagos.length;
          debug.descuentosCount = enriched.descuentos.length;
          debug.step = 'listo';

          // Variante C: reproducir EXACTO lo que hace la acción real getPedido
          if (e.parameter.modo === 'exacto') {
            return jsonResponse({ ok: true, data: enriched });
          }
        }

        return jsonResponse({ ok: true, debug: debug });
      } catch (err) {
        debug.error = err.message;
        debug.stack = String(err.stack || '').slice(0, 500);
        return jsonResponse({ ok: false, debug: debug });
      }
    }

    if (action === 'getPedidos') return jsonResponse({ ok: true, data: getPedidos() });

    if (action === 'getPedido') {
      const id = e.parameter.id;
      if (!id) return errorResponse('Falta parámetro id');
      const pedido = getPedido(id);
      if (!pedido) return errorResponse('Pedido no encontrado', 404);
      return jsonResponse({ ok: true, data: pedido });
    }

    if (action === 'getEstados') return jsonResponse({ ok: true, data: VALID_STATES });
    if (action === 'getProductos') return jsonResponse({ ok: true, data: getProductos() });
    if (action === 'getClientes') return jsonResponse({ ok: true, data: getClientes() });
    if (action === 'getSets') return jsonResponse({ ok: true, data: getSets() });
    if (action === 'getMetodosPago') return jsonResponse({ ok: true, data: METODOS_PAGO });
    if (action === 'getItemsPendientesCostos') return jsonResponse({ ok: true, data: getItemsPendientesCostos() });
    if (action === 'verificarIntegridad') return jsonResponse({ ok: true, data: verificarIntegridad() });
    if (action === 'getStockDisponiblePorItem') return jsonResponse({ ok: true, data: getStockDisponiblePorItem() });
    if (action === 'getItemsPendientesCourier') return jsonResponse({ ok: true, data: getItemsPendientesCourier() });
    

    if (action === 'asignarStock') {
      return handleAsignarStock({
        asignaciones: JSON.parse(e.parameter.asignaciones || '[]'),
        fecha: e.parameter.fecha || '',
        usuario: e.parameter.usuario
      });
    }

    if (action === 'getPedidosPendientesCourier') return jsonResponse({ ok: true, data: getPedidosPendientesCourier() });
    if (action === 'getLotesPendientesCourier') return jsonResponse({ ok: true, data: getLotesPendientesCourier() });

    if (action === 'login') {
      return handleLogin({ usuario: e.parameter.usuario, password: e.parameter.password });
    }

    if (action === 'cambiarEstadoItems') {
      return handleCambiarEstadoItems({
        ordenId: e.parameter.ordenId,
        itemRows: JSON.parse(e.parameter.itemRows || '[]'),
        nuevoEstado: e.parameter.nuevoEstado,
        usuario: e.parameter.usuario,
        forzar: e.parameter.forzar,
        tipoEmpaque: e.parameter.tipoEmpaque,
        cantidadCajas: e.parameter.cantidadCajas || '1',
        costoDelivery: e.parameter.costoDelivery,
        pines: JSON.parse(e.parameter.pines || '[]')   // ← línea nueva
      });
    }

    if (action === 'agregarPago') {
      return handleAgregarPago({
        ordenId: e.parameter.ordenId,
        cantidad: e.parameter.cantidad,
        fecha: e.parameter.fecha,
        metodo: e.parameter.metodo,
        urlComprobante: e.parameter.urlComprobante,
        notas: e.parameter.notas,
        usuario: e.parameter.usuario
      });
    }

    if (action === 'borrarPago') {
      return handleBorrarPago({
        ordenId: e.parameter.ordenId,
        fecha: e.parameter.fecha,
        cantidad: e.parameter.cantidad,
        usuario: e.parameter.usuario
      });
    }

    if (action === 'cargarFacturaFIGS') {
      return handleCargarFactura({
        items: JSON.parse(e.parameter.items || '[]'),
        subtotal: Number(e.parameter.subtotal),
        iva: Number(e.parameter.iva),
        shipping: Number(e.parameter.shipping),
        numFactura: e.parameter.numFactura || '',
        fecha: e.parameter.fecha || '',
        stockYaLlego: e.parameter.stockYaLlego,
        tracking: e.parameter.tracking || '',
        transporte: e.parameter.transporte || '',
        usuario: e.parameter.usuario
      });
    }

    if (action === 'cargarEnvioCourier') {
      return handleCargarCourier({
        pedidos: JSON.parse(e.parameter.pedidos || '[]'),
        lotes: JSON.parse(e.parameter.lotes || '[]'),
        items: JSON.parse(e.parameter.items || '[]'),
        costoCourier: Number(e.parameter.costoCourier),
        fecha: e.parameter.fecha || '',
        usuario: e.parameter.usuario
      });
    }

    if (action === 'crearPedido') {
      return handleCrearPedido({
        cliente: JSON.parse(e.parameter.cliente || '{}'),
        items: JSON.parse(e.parameter.items || '[]'),
        descuento: JSON.parse(e.parameter.descuento || '{"monto":0,"nota":""}'),
        estado: e.parameter.estado || 'HACER PEDIDO',
        notas: e.parameter.notas || '',
        fecha: e.parameter.fecha || '',
        usuario: e.parameter.usuario
      });
    }

    if (action === 'cambiarItemError') {
      return cambiarItemPorError({
        ordenId: e.parameter.ordenId,
        itemRowX: e.parameter.itemRowX,
        skuStock: e.parameter.skuStock,
        tallaStock: e.parameter.tallaStock,
        longitudStock: e.parameter.longitudStock || 'Regular',
        colorStock: e.parameter.colorStock,
        costoStock: e.parameter.costoStock,
        skuY: e.parameter.skuY,
        tallaY: e.parameter.tallaY,
        longitudY: e.parameter.longitudY,
        colorY: e.parameter.colorY,
        quitarCosto: e.parameter.quitarCosto,
        usuario: e.parameter.usuario
      });
    }

    if (action === 'crearCliente') {
      return handleCrearCliente({
        cliente: JSON.parse(e.parameter.cliente || '{}'),
        fecha: e.parameter.fecha,
        usuario: e.parameter.usuario
      });
    }

    if (action === 'avanzarLote') {
      return avanzarLote({
        loteId: e.parameter.loteId,
        accion: e.parameter.accion,
        courier: e.parameter.courier || '',
        fechaSalida: e.parameter.fechaSalida || '',
        usuario: e.parameter.usuario
      });
    }

    if (action === 'getCombos') {
      return jsonResponse({ ok: true, data: getCombos() });
    }
    if (action === 'crearCombo') {
      return crearCombo({
        loteSuperior: e.parameter.loteSuperior,
        loteInferior: e.parameter.loteInferior,
        fecha: e.parameter.fecha || '',
        usuario: e.parameter.usuario
      });
    }
    if (action === 'borrarCombo') {
      return borrarCombo({
        comboId: e.parameter.comboId,
        usuario: e.parameter.usuario
      });
    }
    if (action === 'getSetEmpaque') {
      return jsonResponse({ ok: true, data: getSetEmpaque() });
    }
    
    if (action === 'getStockDisponible') {
      return jsonResponse({ ok: true, data: getStockDisponible() });
    }

    if (action === 'getRegalos') {
      return jsonResponse({ ok: true, data: getRegalos() });
    }

    if (action === 'getCouriers') return jsonResponse({ ok: true, data: getCouriers() });

    if (action === 'calcularETA') return jsonResponse({ ok: true, data: calcularETA(e.parameter.courier, e.parameter.fecha) });

    if (action === 'getLotesEnCamino') return jsonResponse({ ok: true, data: getLotesEnCamino() });

    if (action === 'probarFedexToken') {
      return probarFedexToken();
    }

    if (action === 'probarFedexRastreo') {
      return probarFedexRastreo({ tracking: e.parameter.tracking || '' });
    }

    if (action === 'probarActualizarTracking') {
      return probarActualizarTracking();
    }

    if (action === 'backfillCourierCompleto') {
      return backfillCourierCompleto({ usuario: e.parameter.usuario });
    }

    if (action === 'getReporteGanancia') {
      return jsonResponse({ ok: true, data: getReporteGanancia(e.parameter.fechaInicio || '', e.parameter.fechaFin || '') });
    }

    if (action === 'getGastosFijos') {
      return getGastosFijos(e.parameter.mes || '');
    }
    if (action === 'agregarGastoFijo') {
      return agregarGastoFijo({
        usuario:   e.parameter.usuario,
        fecha:     e.parameter.fecha,
        concepto:  e.parameter.concepto,
        categoria: e.parameter.categoria,
        monto:     e.parameter.monto,
        nota:      e.parameter.nota
      });
    }
    if (action === 'getResumenMensual') {
      return getResumenMensual(
        e.parameter.mes || '',
        e.parameter.donacion || 0,
        e.parameter.fechaDesde || '',
        e.parameter.fechaHasta || ''
      );
    }
    if (action === 'getCapitalReal') {
      return getCapitalReal();
    }
    if (action === 'getCostosPorTipo') {
      return getCostosPorTipo(e.parameter.fechaInicio || '', e.parameter.fechaFin || '');
    }
    if (action === 'getTimelineMensual') {
      return getTimelineMensual(e.parameter.meses || 6);
    }
    if (action === 'getHistogramaGanancia') {
      return getHistogramaGanancia(e.parameter.fechaInicio || '', e.parameter.fechaFin || '');
    }

    if (action === 'getMateriales') {
      return jsonResponse({ ok: true, data: getMateriales() });
    }

    if (action === 'setStockMaterial') {
      return setStockMaterial({
        tabla:    e.parameter.tabla,
        id:       e.parameter.id,
        cantidad: e.parameter.cantidad,
        usuario:  e.parameter.usuario,
      });
    }

    if (action === 'agregarPin') {
      return agregarPin({
        nombre:    e.parameter.nombre,
        categoria: e.parameter.categoria,
        cantidad:  e.parameter.cantidad,
        costo:     e.parameter.costo,
        usuario:   e.parameter.usuario,
      });
    }

    if (action === 'actualizarCliente') {
      return actualizarCliente({
        clienteId: e.parameter.clienteId,
        campos:    JSON.parse(e.parameter.campos || '{}'),
        usuario:   e.parameter.usuario,
      });
    }
    
    if (action === 'agregarDescuento') {
      return agregarDescuento({
        ordenId: e.parameter.ordenId,
        monto:   e.parameter.monto,
        nota:    e.parameter.nota,
        usuario: e.parameter.usuario,
      });
    }

    if (action === 'perdidaEnTransito') {
      return perdidaEnTransito({
        ordenId: e.parameter.ordenId,
        items:   JSON.parse(e.parameter.items || '[]'),
        usuario: e.parameter.usuario,
      });
    }

    if (action === 'perdidaEnTransito') {
      return perdidaEnTransito({
        ordenId: e.parameter.ordenId,
        items:   JSON.parse(e.parameter.items || '[]'),
        usuario: e.parameter.usuario,
      });
    }

    if (action === 'getColores') return jsonResponse({ ok: true, data: getColores() });

    if (action === 'setColor') {
      return setColor({
        color:   e.parameter.color,
        hex:     e.parameter.hex,
        usuario: e.parameter.usuario
      });
    }

    return errorResponse('Acción desconocida: ' + action);
  } catch (err) {
    return errorResponse('Error: ' + err.message, 500);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'login') return handleLogin(body);
    if (action === 'cambiarEstadoItems') return handleCambiarEstadoItems(body);
    if (action === 'agregarPago') return handleAgregarPago(body);
    if (action === 'borrarPago') return handleBorrarPago(body);
    if (action === 'cargarFacturaFIGS') return handleCargarFactura(body);
    if (action === 'cargarEnvioCourier') return handleCargarCourier(body);
    if (action === 'asignarStock') return handleAsignarStock(body);
    if (action === 'crearPedido') return handleCrearPedido(body);
    if (action === 'crearCliente') return handleCrearCliente(body);
    // --- Gastos Fijos (Pieza 1) ---
    if (action === 'agregarGastoFijo') return agregarGastoFijo(body);
    if (action === 'borrarGastoFijo')  return borrarGastoFijo(body);
    if (action === 'getGastosFijos')   return getGastosFijos(body.mes);  

    return errorResponse('Acción desconocida: ' + action);
  } catch (err) {
    return errorResponse('Error: ' + err.message, 500);
  }
}

// ============ LOGIN ============

function handleLogin(body) {
  const usuario = String(body.usuario || '').toLowerCase().trim();
  const password = String(body.password || '');

  if (!usuario || !password) return errorResponse('Usuario y contraseña requeridos');

  const usuarios = readSheetAsObjects(TABS.usuarios);
  const found = usuarios.find(u => String(u.USUARIO).toLowerCase().trim() === usuario);

  if (!found) return errorResponse('Usuario o contraseña incorrectos', 401);
  if (String(found.PASSWORD) !== password) return errorResponse('Usuario o contraseña incorrectos', 401);

  const token = Utilities.base64Encode(usuario + ':' + new Date().getTime());
  logChange(usuario, '-', 'LOGIN', 'Inicio de sesión');

  return jsonResponse({
    ok: true,
    data: { token: token, usuario: found.USUARIO, nombre: found.NOMBRE, rol: found.ROL }
  });
}

// ============ VER STOCK ==========

function getStockDisponible() {
  const lotes     = readSheetAsObjects(TABS.lotes);
  const productos = readSheetAsObjects(TABS.productos);
  const costos    = readSheetAsObjects(TABS.costos);
  const coloresMap = getColores();

  const prodMap = {};
  productos.forEach(p => {
    prodMap[String(p.SKU).trim()] = { nombre: p.NOMBRE, tipo: p.TIPO_PRENDA };
  });

  // Estados que cuentan como "en camino" (preventa)
  const EN_CAMINO = ['EN TRANSITO A FL', 'EN BODEGA FL', 'EN CAMINO A EC'];

  return lotes
    .filter(l => {
      if ((Number(l.CANT_DISPONIBLE) || 0) <= 0) return false;
      const ev = String(l.ESTADO_VIAJE || '').toUpperCase().trim();
      // Llegados: vacío, DISPONIBLE, EN BODEGA EC  ·  En camino: los de EN_CAMINO
      return ev === '' || ev === 'DISPONIBLE' || ev === 'EN BODEGA EC' || EN_CAMINO.indexOf(ev) !== -1;
    })
    .map(l => {
      const sku = String(l.SKU).trim();
      const info = prodMap[sku] || { nombre: sku, tipo: '' };
      const loteId = String(l.LOTE_ID).trim();
      const ev = String(l.ESTADO_VIAJE || '').toUpperCase().trim();
      const enCamino = EN_CAMINO.indexOf(ev) !== -1;

      const tieneCourier = costos.some(c =>
        String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'LOTE_STOCK' &&
        String(c.REFERENCIA_ID).trim() === loteId &&
        String(c.TIPO_COSTO).toUpperCase().trim() === 'COURIER'
      );

      // ETA si está en camino
      let eta = null;
      if (enCamino) {
        const etaMin = String(l.ETA_MIN || '').trim();
        const etaMax = String(l.ETA_MAX || '').trim();
        if (etaMin || etaMax) {
          eta = { fechaMin: etaMin, fechaMax: etaMax };
        } else {
          const courier = String(l.COURIER || '').trim();
          const fSalida = String(l.FECHA_SALIDA_EC || '').trim();
          if (courier && fSalida) {
            const r = calcularETA(courier, fSalida);
            if (r && r.ok) eta = { fechaMin: r.fechaMin, fechaMax: r.fechaMax };
          }
        }
      }

      return {
        LOTE_ID:         l.LOTE_ID,
        SKU:             l.SKU,
        NOMBRE_PRODUCTO: info.nombre,
        TIPO_PRENDA:     info.tipo,
        TALLA:           l.TALLA,
        LONGITUD:        l.LONGITUD,
        COLOR:           l.COLOR,
        COLOR_HEX:       coloresMap[String(l.COLOR || '').toUpperCase().trim()] || '',
        CANT_DISPONIBLE: Number(l.CANT_DISPONIBLE) || 0,
        FECHA_ENTRADA:   l.FECHA_ENTRADA,
        COSTO_UNITARIO:  Number(l.COSTO_UNITARIO) || 0,
        tieneCourier:    tieneCourier,
        enCamino:        enCamino,        // ← NUEVO: bandera de preventa
        ESTADO_VIAJE:    String(l.ESTADO_VIAJE || '').trim(),  // ← NUEVO
        eta:             eta              // ← NUEVO: { fechaMin, fechaMax } o null
      };
    })
    .sort((a, b) => {
      // Primero los disponibles (llegados), luego los en camino; dentro, por fecha
      if (a.enCamino !== b.enCamino) return a.enCamino ? 1 : -1;
      return String(a.FECHA_ENTRADA).localeCompare(String(b.FECHA_ENTRADA));
    });
}
// ============ PEDIDOS ============

function getPedidos() {
  const ordenes    = readSheetAsObjects(TABS.ordenes);
  const items      = readSheetAsObjects(TABS.items);
  const costos     = readSheetAsObjects(TABS.costos);
  const pagos      = readSheetAsObjects(TABS.pagos);
  const descuentos = readSheetAsObjects(TABS.descuentos);
  const clientes   = readSheetAsObjects(TABS.clientes);
  return ordenes.map(orden => enrichOrden(orden, items, costos, pagos, descuentos, clientes));
}

function getPedido(ordenId) {
  const ordenes = readSheetAsObjects(TABS.ordenes);
  const orden = ordenes.find(o => String(o.ORDEN_ID).trim() === String(ordenId).trim());
  if (!orden) return null;
  const items      = readSheetAsObjects(TABS.items);
  const costos     = readSheetAsObjects(TABS.costos);
  const pagos      = readSheetAsObjects(TABS.pagos);
  const descuentos = readSheetAsObjects(TABS.descuentos);
  const clientes   = readSheetAsObjects(TABS.clientes);
  return enrichOrden(orden, items, costos, pagos, descuentos, clientes);
}

function getProductos() {
  const productos = readSheetAsObjects(TABS.productos);
  return productos
    .filter(p => {
      const activo = String(p.ACTIVO).toUpperCase().trim();
      return activo === 'TRUE' || activo === 'VERDADERO' || activo === 'SÍ' || activo === 'SI';
    })
    .map(p => ({ SKU: p.SKU, NOMBRE: p.NOMBRE, TIPO_PRENDA: p.TIPO_PRENDA }));
}

function getClientes() {
  const clientes = readSheetAsObjects(TABS.clientes);
  return clientes.map(c => ({
    CLIENTE_ID: c.CLIENTE_ID,
    NOMBRE: c.NOMBRE,
    TELEFONO: c.TELEFONO,
    DIRECCION: c.DIRECCION,
    CIUDAD: c.CIUDAD,
    CEDULA_RUC: c.CEDULA_RUC,
    INDUSTRIA: c.INDUSTRIA,
    EMAIL: c.EMAIL
  }));
}

function getSets() {
  const sets = readSheetAsObjects(TABS.sets);
  return sets
    .filter(s => {
      const activo = String(s.ACTIVO).toUpperCase().trim();
      return activo === 'TRUE' || activo === 'VERDADERO' || activo === 'SÍ' || activo === 'SI';
    })
    .map(s => ({
      SET_NOMBRE: s.SET_NOMBRE,
      SKU_TOP: s.SKU_TOP,
      SKU_PANTALON: s.SKU_PANTALON,
      PRECIO_SET: Number(s.PRECIO_SET) || 0
    }));
}

function getSetEmpaque() {
  const sets = readSheetAsObjects(TABS.setEmpaque);
  const materiales = readSheetAsObjects(TABS.empaque);

  return sets
    .filter(s => {
      const activo = String(s.ACTIVO).toUpperCase().trim();
      return activo === 'TRUE' || activo === 'VERDADERO';
    })
    .map(s => {
      const ids = String(s.MATERIALES).split(',').map(m => m.trim());
      const items = ids.map(id => {
        const mat = materiales.find(m => String(m.MATERIAL_ID).trim() === id);
        return mat ? { id, nombre: mat.NOMBRE, costo: Number(mat.COSTO_UNITARIO) || 0 } : null;
      }).filter(Boolean);
      const costoTotal = round2(items.reduce((sum, i) => sum + i.costo, 0));
      return {
        SET_ID: s.SET_ID,
        NOMBRE_SET: s.NOMBRE_SET,
        materiales: items,
        costoTotal
      };
    });
}

function getRegalos() {
  const regalos = readSheetAsObjects(TABS.regalos);
  return regalos
    .filter(r => (Number(r.STOCK) || 0) > 0)
    .map(r => ({
      REGALO_ID:          String(r.REGALO_ID).trim(),
      NOMBRE:             r.NOMBRE,
      INDUSTRIA_SUGERIDA: r.INDUSTRIA_SUGERIDA,
      STOCK:              Number(r.STOCK) || 0,
      STOCK_MINIMO:       Number(r.STOCK_MINIMO) || 0,
      COSTO_UNITARIO:     Number(r.COSTO_UNITARIO) || 0
    }));
}

function enrichOrden(orden, items, costos, pagos, descuentos, clientes) {
  const id = String(orden.ORDEN_ID).trim();

  const ordenItems      = items.filter(i => String(i.ORDEN_ID).trim() === id);
  const ordenCostos     = costos.filter(c =>
    String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'PEDIDO' &&
    String(c.REFERENCIA_ID).trim() === id
  );
  const ordenPagos      = pagos.filter(p => String(p.ORDEN_ID).trim() === id);
  const ordenDescuentos = descuentos.filter(d => String(d.ORDEN_ID).trim() === id);

  const nombreOrden = normalize(orden.CLIENTE_NOMBRE);
  const cliente = clientes.find(c => normalize(c.NOMBRE) === nombreOrden) || null;

  const itemsVivos = ordenItems.filter(i =>
    String(i.ESTATUS_ITEM || '').toUpperCase().trim() !== 'CANCELADO');
  const totalBruto = round2(itemsVivos.reduce((sum, i) =>
    sum + ((Number(i.CANTIDAD) || 0) * (Number(i.PRECIO_VENTA) || 0)), 0));
  const totalDescuentos = round2(ordenDescuentos.reduce((sum, d) => sum + (Number(d.MONTO) || 0), 0));
  const totalVenta  = round2(totalBruto + totalDescuentos);
  const totalPagado = round2(ordenPagos.reduce((sum, p) => sum + (Number(p.MONTO) || 0), 0));
  const totalCostos = round2(ordenCostos.reduce((sum, c) => sum + (Number(c.MONTO) || 0), 0));
  const saldo       = round2(totalVenta - totalPagado);
  const ganancia    = round2(totalVenta - totalCostos);
  const estadoPago  = calcularEstadoPago(totalVenta, totalPagado);

  return {
    ...orden,
    NOMBRE: orden.CLIENTE_NOMBRE,
    cliente,
    items: ordenItems,
    costos: ordenCostos,
    pagos: ordenPagos,
    descuentos: ordenDescuentos,
    totales: { bruto: totalBruto, descuentos: totalDescuentos, venta: totalVenta,
               costos: totalCostos, ganancia, pagado: totalPagado, saldo, estadoPago }
  };
}

// ============ CAMBIAR ESTADO ============

function handleCambiarEstado(body) {
  const ordenId     = body.ordenId;
  const nuevoEstado = body.nuevoEstado;
  const usuario     = body.usuario || 'desconocido';
  const forzar      = body.forzar === true || String(body.forzar) === 'true';
  const tipoEmpaque = String(body.tipoEmpaque || '').trim();

  if (!ordenId)     return errorResponse('Falta ordenId');
  if (!nuevoEstado) return errorResponse('Falta nuevoEstado');
  if (VALID_STATES.indexOf(nuevoEstado) === -1) {
    return errorResponse('Estado inválido: ' + nuevoEstado);
  }
  if (!canCambiarEstado(usuario)) {
    return errorResponse('No tienes permiso para cambiar estados', 403);
  }

  //if (nuevoEstado === 'LISTO PARA ENVIAR' && !tipoEmpaque) {
  //  return errorResponse('Falta seleccionar el tipo de empaque', 400);
  //}

  const rowNum = findOrderRow(ordenId);
  if (rowNum === -1) return errorResponse('Pedido no encontrado: ' + ordenId, 404);

  // ── Bloqueo: no avanzar sin factura FIGS (ítems PEDIDO sin costo) ─────────
  if (requiereFactura(nuevoEstado)) {
    const sinCosto = itemsPedidoSinCosto(ordenId);
    if (sinCosto.length > 0) {
      if (!forzar) {
        return errorResponse(
          `El pedido ${ordenId} tiene ${sinCosto.length} ítem(s) sin la factura FIGS cargada (sin costo). ` +
          `Carga la factura antes de pasar a "${nuevoEstado}".`,
          422
        );
      }
      if (!verifyAdmin(usuario)) {
        return errorResponse(
          `El pedido ${ordenId} tiene ítems sin factura. Solo un admin puede forzar.`,
          403
        );
      }
    }
  }

  // ── Bloqueo: no marcar ENTREGADO con saldo pendiente ─────────────────────
  if (nuevoEstado === 'ENTREGADO') {
    const saldo = calcularSaldoPedido(ordenId);
    if (saldo > 0.005) {
      if (!forzar) {
        return errorResponse(
          `El pedido ${ordenId} tiene un saldo pendiente de $${saldo.toFixed(2)}. ` +
          `No puede marcarse como ENTREGADO hasta que esté pagado.`,
          422
        );
      }
      if (!verifyAdmin(usuario)) {
        return errorResponse(
          `El pedido ${ordenId} debe $${saldo.toFixed(2)}. ` +
          `Solo un admin puede forzar la entrega con saldo pendiente.`,
          403
        );
      }
    }
  }

  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.ordenes);

  const estadoAnterior = sheet.getRange(rowNum, COL_ESTATUS_ENVIO).getValue();
  sheet.getRange(rowNum, COL_ESTATUS_ENVIO).setValue(nuevoEstado);

  // ── Fecha de entrega real al pasar a ENTREGADO ───────────────────────────
  if (nuevoEstado === 'ENTREGADO') {
    marcarFechaEntrega(sheet, rowNum);
  }
  // ── Costo de empaque al pasar a LISTO PARA ENVIAR ────────────────────────
  let costoEmpaqueFinal = 0;
  if (nuevoEstado === 'LISTO PARA ENVIAR') {
    const costos = readSheetAsObjects(TABS.costos);
    const yaEmpaque = costos.some(c =>
      String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'PEDIDO' &&
      String(c.REFERENCIA_ID).trim() === String(ordenId).trim() &&
      String(c.TIPO_COSTO).toUpperCase().trim() === 'EMPAQUE'
    );

    if (!yaEmpaque) {
      const sets      = readSheetAsObjects(TABS.setEmpaque);
      const materiales = readSheetAsObjects(TABS.empaque);
      const setElegido = sets.find(s => String(s.SET_ID).trim() === tipoEmpaque);

      if (setElegido) {
        const ids = String(setElegido.MATERIALES).split(',').map(m => m.trim());
        costoEmpaqueFinal = round2(ids.reduce((sum, id) => {
          const mat = materiales.find(m => String(m.MATERIAL_ID).trim() === id);
          return sum + (mat ? Number(mat.COSTO_UNITARIO) || 0 : 0);
        }, 0));

        const fechaHoy    = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
        const sheetCostos = ss.getSheetByName(TABS.costos);
        sheetCostos.appendRow([
          'PEDIDO', ordenId, fechaHoy, 'EMPAQUE',
          `${setElegido.NOMBRE_SET}`, costoEmpaqueFinal, 'Auto'
        ]);

        const COL_TIPO_EMPAQUE = 25;
        sheet.getRange(rowNum, COL_TIPO_EMPAQUE).setValue(setElegido.NOMBRE_SET);

        logChange(usuario, ordenId, 'COSTO_EMPAQUE',
          `${setElegido.NOMBRE_SET} · $${costoEmpaqueFinal.toFixed(2)}`);
      }
    }
  }

  // ── Log ───────────────────────────────────────────────────────────────────
  const saldoFinal = (nuevoEstado === 'ENTREGADO') ? calcularSaldoPedido(ordenId) : 0;
  if (nuevoEstado === 'ENTREGADO' && saldoFinal > 0.005) {
    logChange(usuario, ordenId, 'CAMBIO_ESTADO_FORZADO',
      `${estadoAnterior} → ${nuevoEstado} · FORZADO con saldo pendiente $${saldoFinal.toFixed(2)}`);
  } else {
    logChange(usuario, ordenId, 'CAMBIO_ESTADO', `${estadoAnterior} → ${nuevoEstado}`);
  }

  return jsonResponse({
    ok: true,
    data: {
      ordenId, estadoAnterior, nuevoEstado,
      tipoEmpaque, costoEmpaque: costoEmpaqueFinal,
      forzado: (nuevoEstado === 'ENTREGADO' && saldoFinal > 0.005),
      timestamp: new Date().toISOString()
    }
  });
}

/**
 * Calcula el saldo pendiente de un pedido (venta − pagado).
 */
function calcularSaldoPedido(ordenId) {
  const id = String(ordenId).trim();
  const items      = readSheetAsObjects(TABS.items);
  const descuentos = readSheetAsObjects(TABS.descuentos);
  const pagos      = readSheetAsObjects(TABS.pagos);

  const ordenItems = items.filter(i => String(i.ORDEN_ID).trim() === id &&
    String(i.ESTATUS_ITEM || '').toUpperCase().trim() !== 'CANCELADO');
  const ordenDesc  = descuentos.filter(d => String(d.ORDEN_ID).trim() === id);
  const ordenPagos = pagos.filter(p => String(p.ORDEN_ID).trim() === id);

  const totalBruto      = round2(ordenItems.reduce((sum, i) =>
    sum + ((Number(i.CANTIDAD) || 0) * (Number(i.PRECIO_VENTA) || 0)), 0));
  const totalDescuentos = round2(ordenDesc.reduce((sum, d) => sum + (Number(d.MONTO) || 0), 0));
  const totalVenta      = round2(totalBruto + totalDescuentos);
  const totalPagado     = round2(ordenPagos.reduce((sum, p) => sum + (Number(p.MONTO) || 0), 0));

  return round2(totalVenta - totalPagado);
}

// ============ PAGOS ============

function handleAgregarPago(body) {
  const ordenId        = String(body.ordenId || '').trim();
  const cantidad       = Number(body.cantidad);
  const fecha          = String(body.fecha || '').trim();
  const metodo         = String(body.metodo || '').trim();
  const urlComprobante = String(body.urlComprobante || '').trim();
  const notas          = String(body.notas || '').trim();
  const usuario        = body.usuario || 'desconocido';

  if (!ordenId) return errorResponse('Falta ordenId');
  if (isNaN(cantidad) || cantidad === 0) return errorResponse('Monto inválido (no puede ser 0)');
  if (!fecha)   return errorResponse('Falta fecha');
  if (!metodo)  return errorResponse('Falta método');
  if (METODOS_PAGO.indexOf(metodo) === -1) return errorResponse('Método inválido: ' + metodo);
  if (!canOperar(usuario)) return errorResponse('No tienes permiso para registrar pagos', 403);
  if (findOrderRow(ordenId) === -1) return errorResponse('Pedido no encontrado: ' + ordenId, 404);

  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.pagos);
  sheet.appendRow([ordenId, fecha, cantidad, metodo, urlComprobante, notas]);

  const tipo = cantidad < 0 ? 'DEVOLUCION' : 'PAGO';
  logChange(usuario, ordenId, tipo + '_AGREGADO',
    `$${cantidad.toFixed(2)} · ${metodo}${notas ? ' · ' + notas : ''}`);

  return jsonResponse({ ok: true, data: { ordenId, cantidad, fecha, metodo, urlComprobante, notas } });
}

function handleBorrarPago(body) {
  const ordenId  = String(body.ordenId || '').trim();
  const fecha    = String(body.fecha || '').trim();
  const cantidad = Number(body.cantidad);
  const usuario  = body.usuario || 'desconocido';

  if (!canOperar(usuario)) return errorResponse('No tienes permiso para borrar pagos', 403);
  if (!ordenId || !fecha || isNaN(cantidad)) return errorResponse('Faltan datos');

  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.pagos);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowOrdenId = String(data[i][0]).trim();
    let rowFecha = data[i][1];
    if (rowFecha instanceof Date) {
      rowFecha = Utilities.formatDate(rowFecha, 'GMT-5', 'yyyy-MM-dd');
    }
    const rowMonto = Number(data[i][2]);
    if (rowOrdenId === ordenId && String(rowFecha).trim() === fecha && rowMonto === cantidad) {
      sheet.deleteRow(i + 1);
      logChange(usuario, ordenId, 'PAGO_BORRADO', `$${cantidad.toFixed(2)} · ${fecha}`);
      return jsonResponse({ ok: true, data: { deleted: true } });
    }
  }

  return errorResponse('Pago no encontrado', 404);
}

// ============ ITEMS PENDIENTES DE COSTO ============

function getItemsPendientesCostos() {
  const ordenes = readSheetAsObjects(TABS.ordenes);
  const items   = readSheetAsObjects(TABS.items);
  const estadosValidos = ['HACER PEDIDO', 'PEDIDO HECHO'];
  const result = [];

  ordenes.forEach(orden => {
    const estado = String(orden.ESTATUS_ENVIO).trim();
    if (estadosValidos.indexOf(estado) === -1) return;

    const id = String(orden.ORDEN_ID).trim();
    const ordenItems = items.filter(i => String(i.ORDEN_ID).trim() === id);
    if (ordenItems.length === 0) return;

    const itemsPendientes = ordenItems
      .filter(item => {
        const costo = item.COSTO_UNITARIO;
        return costo === '' || costo === null || costo === undefined ||
               (typeof costo === 'number' && isNaN(costo));
      })
      .map(item => ({
        _rowNum: item._rowNum, SKU: item.SKU, NOMBRE_PRODUCTO: item.NOMBRE_PRODUCTO,
        TIPO_PRENDA: item.TIPO_PRENDA, TALLA: item.TALLA, LONGITUD: item.LONGITUD,
        COLOR: item.COLOR, CANTIDAD: item.CANTIDAD, PRECIO_VENTA: item.PRECIO_VENTA,
        PARTE_DE_SET: item.PARTE_DE_SET
      }));

    if (itemsPendientes.length === 0) return;

    result.push({
      ORDEN_ID: orden.ORDEN_ID, CLIENTE_NOMBRE: orden.CLIENTE_NOMBRE,
      ESTATUS_ENVIO: orden.ESTATUS_ENVIO, F_ORDEN: orden.F_ORDEN,
      items: itemsPendientes, totalItems: itemsPendientes.length
    });
  });

  return result;
}

function getStockDisponiblePorItem() {
  const pendientes = getItemsPendientesCostos();
  const lotes  = readSheetAsObjects(TABS.lotes);
  const costos = readSheetAsObjects(TABS.costos);

  function norm(v) { return String(v || '').trim().toLowerCase(); }

  return pendientes.map(pedido => {
    const itemsConStock = pedido.items.map(item => {
      const lotesMatch = lotes.filter(l =>
        norm(l.SKU) === norm(item.SKU) && norm(l.TALLA) === norm(item.TALLA) &&
        norm(l.LONGITUD) === norm(item.LONGITUD) && norm(l.COLOR) === norm(item.COLOR) &&
        (Number(l.CANT_DISPONIBLE) || 0) > 0
      );
      lotesMatch.sort((a, b) => String(a.FECHA_ENTRADA).localeCompare(String(b.FECHA_ENTRADA)));
      const stockTotal = lotesMatch.reduce((s, l) => s + (Number(l.CANT_DISPONIBLE) || 0), 0);

      let loteSugerido = null;
      if (lotesMatch.length > 0) {
        const lote = lotesMatch[0];
        const tieneCourier = costos.some(c =>
          String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'LOTE_STOCK' &&
          String(c.REFERENCIA_ID).trim() === String(lote.LOTE_ID).trim() &&
          String(c.TIPO_COSTO).toUpperCase().trim() === 'COURIER'
        );
        loteSugerido = { LOTE_ID: lote.LOTE_ID, COSTO_UNITARIO: Number(lote.COSTO_UNITARIO) || 0, tieneCourier };
      }

      return {
        _rowNum: item._rowNum, SKU: item.SKU, NOMBRE_PRODUCTO: item.NOMBRE_PRODUCTO,
        TALLA: item.TALLA, LONGITUD: item.LONGITUD, COLOR: item.COLOR,
        CANTIDAD: item.CANTIDAD, stockTotal, hayStock: stockTotal > 0, loteSugerido
      };
    });

    return {
      ORDEN_ID: pedido.ORDEN_ID, CLIENTE_NOMBRE: pedido.CLIENTE_NOMBRE,
      ESTATUS_ENVIO: pedido.ESTATUS_ENVIO, items: itemsConStock
    };
  });
}

// ============ COURIER: PENDIENTES ============

function getPedidosPendientesCourier() {
  const ordenes = readSheetAsObjects(TABS.ordenes);
  const items   = readSheetAsObjects(TABS.items);
  const costos  = readSheetAsObjects(TABS.costos);

  return ordenes
    .filter(orden => String(orden.ESTATUS_ENVIO).trim() !== 'CANCELADO')
    .filter(orden => {
      const id = String(orden.ORDEN_ID).trim();
      const ordenCostos = costos.filter(c =>
        String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'PEDIDO' &&
        String(c.REFERENCIA_ID).trim() === id
      );
      const tieneBruto   = ordenCostos.some(c => String(c.TIPO_COSTO).toUpperCase().trim() === 'BRUTO');
      const tieneCourier = ordenCostos.some(c => String(c.TIPO_COSTO).toUpperCase().trim() === 'COURIER');
      return tieneBruto && !tieneCourier;
    })
    .map(orden => {
      const id = String(orden.ORDEN_ID).trim();
      const ordenItems = items.filter(i => String(i.ORDEN_ID).trim() === id);
      const numItems = ordenItems.reduce((s, i) => s + (Number(i.CANTIDAD) || 0), 0);
      return {
        ORDEN_ID: orden.ORDEN_ID, CLIENTE_NOMBRE: orden.CLIENTE_NOMBRE,
        ESTATUS_ENVIO: orden.ESTATUS_ENVIO, F_ORDEN: orden.F_ORDEN, numItems
      };
    });
}

function getLotesPendientesCourier() {
  const lotes  = readSheetAsObjects(TABS.lotes);
  const costos = readSheetAsObjects(TABS.costos);

  return lotes
    .filter(lote => {
      const id = String(lote.LOTE_ID).trim();
      return !costos.some(c =>
        String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'LOTE_STOCK' &&
        String(c.REFERENCIA_ID).trim() === id &&
        String(c.TIPO_COSTO).toUpperCase().trim() === 'COURIER'
      );
    })
    .map(lote => ({
      LOTE_ID: lote.LOTE_ID, SKU: lote.SKU, TALLA: lote.TALLA,
      LONGITUD: lote.LONGITUD, COLOR: lote.COLOR,
      CANT_INICIAL: Number(lote.CANT_INICIAL) || 0,
      COSTO_UNITARIO: Number(lote.COSTO_UNITARIO) || 0
    }));
}

// ============ CARGAR FACTURA FIGS ============

function handleCargarFactura(body) {
  const itemsFactura = body.items || [];
  const subtotal  = Number(body.subtotal) || 0;
  const iva       = Number(body.iva) || 0;
  const shipping  = Number(body.shipping) || 0;
  const numFactura = String(body.numFactura || '').trim();
  const fecha     = String(body.fecha || '').trim() ||
                    Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
  const usuario   = body.usuario || 'desconocido';

  // ── NUEVO (Capa 2): viaje del stock ──
  // Si el frontend NO manda el campo (versión vieja), asumimos "ya llegó" para no romper nada.
  const stockYaLlego = (body.stockYaLlego === undefined || body.stockYaLlego === null || body.stockYaLlego === '')
    ? true
    : (body.stockYaLlego === true || String(body.stockYaLlego).toLowerCase() === 'true');
  const tracking   = String(body.tracking || '').trim();
  const transporte = String(body.transporte || '').trim();

  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede cargar facturas', 403);
  if (itemsFactura.length === 0) return errorResponse('No hay items en la factura');
  if (subtotal <= 0) return errorResponse('Subtotal inválido');
  if (iva < 0 || shipping < 0) return errorResponse('IVA o shipping inválidos');
  if (!numFactura) return errorResponse('Falta número de factura');

  const costosExistentes = readSheetAsObjects(TABS.costos);
  const origenBuscado = `Factura FIGS ${numFactura}`;
  if (costosExistentes.some(c => String(c.ORIGEN).trim() === origenBuscado)) {
    return errorResponse(
      `La factura ${numFactura} ya fue cargada anteriormente. ` +
      `Si necesitas corregirla, primero borra sus filas en TablaCostos.`, 409);
  }

  for (let i = 0; i < itemsFactura.length; i++) {
    const it = itemsFactura[i];
    const tipo = String(it.tipoReferencia || '').toUpperCase();
    if (tipo !== 'PEDIDO' && tipo !== 'STOCK') return errorResponse(`Item ${i}: tipoReferencia debe ser PEDIDO o STOCK`);
    if (tipo === 'PEDIDO' && (!it.ordenId || !it.itemRowNum)) return errorResponse(`Item ${i} (PEDIDO): falta ordenId o itemRowNum`);
    if (tipo === 'STOCK' && (!it.cantidad || Number(it.cantidad) <= 0)) return errorResponse(`Item ${i} (STOCK): cantidad debe ser > 0`);
    if (!it.sku || Number(it.precioFIGS) <= 0) return errorResponse(`Item ${i}: falta sku o precioFIGS`);
  }

  const sumaItems = itemsFactura.reduce((s, it) => {
    const tipo = String(it.tipoReferencia).toUpperCase();
    const cant = tipo === 'STOCK' ? Number(it.cantidad) : 1;
    return s + (Number(it.precioFIGS) * cant);
  }, 0);
  if (Math.abs(sumaItems - subtotal) > 0.05) {
    return errorResponse(`Suma de items ($${sumaItems.toFixed(2)}) no coincide con subtotal ($${subtotal.toFixed(2)})`);
  }

  return withLock(function() {
    const ss = getSpreadsheet();
    const sheetCostos  = ss.getSheetByName(TABS.costos);
    const sheetLotes   = ss.getSheetByName(TABS.lotes);
    const sheetMovs    = ss.getSheetByName(TABS.movimientos);
    const sheetItems   = ss.getSheetByName(TABS.items);
    const sheetOrdenes = ss.getSheetByName(TABS.ordenes);

    const itemsHeaders   = sheetItems.getRange(1, 1, 1, sheetItems.getLastColumn()).getValues()[0];
    const ordenesHeaders = sheetOrdenes.getRange(1, 1, 1, sheetOrdenes.getLastColumn()).getValues()[0];
    const colCostoUnitario = itemsHeaders.indexOf('COSTO_UNITARIO') + 1;
    const colNumOrdenProv  = ordenesHeaders.indexOf('NUM_ORDEN_PROV') + 1;

    if (colCostoUnitario === 0) throw new Error('No se encontró columna COSTO_UNITARIO en TablaItems');
    if (colNumOrdenProv  === 0) throw new Error('No se encontró columna NUM_ORDEN_PROV en TablaOrdenes');

    const origen = `Factura FIGS ${numFactura}`;
    const filasAgregadas = [];
    const lotesCreados = [];
    const ordenesActualizadas = new Set();

    // Estado de viaje del lote según el toggle
    const estadoViajeLote = stockYaLlego ? 'DISPONIBLE' : 'EN TRANSITO A FL';

    itemsFactura.forEach(item => {
      const tipo       = String(item.tipoReferencia).toUpperCase();
      const precioFinal = Number(item.precioFIGS);
      const sku        = String(item.sku);
      const talla      = String(item.talla || '');
      const longitud   = String(item.longitud || '');
      const color      = String(item.color || '');

      const pct          = precioFinal / subtotal;
      const ivaUnit      = round2(iva * pct);
      const shippingUnit = round2(shipping * pct);
      const costoUnitario = round2(precioFinal + ivaUnit + shippingUnit);
      const descripcionItem = `${sku} ${talla}-${longitud} ${color}`.trim();

      if (tipo === 'PEDIDO') {
        const ordenId = String(item.ordenId).trim();
        const rowNum  = Number(item.itemRowNum);
        sheetCostos.appendRow(['PEDIDO', ordenId, fecha, 'BRUTO', descripcionItem, costoUnitario, origen]);
        if (rowNum > 1) sheetItems.getRange(rowNum, colCostoUnitario).setValue(costoUnitario);
        ordenesActualizadas.add(ordenId);
        filasAgregadas.push({ tipo: 'PEDIDO', ordenId, sku, precioFinal, costoUnitario });
      } else {
        const cantidad = Number(item.cantidad);
        const loteId   = nextLoteId();
        const ahora    = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd HH:mm:ss');
        // appendRow con las 7 columnas nuevas de viaje al final (N..T)
        sheetLotes.appendRow([
          loteId, sku, talla, longitud, color, costoUnitario,
          cantidad, cantidad, 'FACTURA_FIGS', numFactura, fecha, usuario, '',
          estadoViajeLote,
          stockYaLlego ? '' : tracking,
          stockYaLlego ? '' : transporte,
          '', '', '', ''
        ]);
        sheetMovs.appendRow([nextMovId(), ahora, loteId, 'ENTRADA', cantidad,
          'FACTURA_FIGS', numFactura, usuario, '']);
        const montoTotalLote = round2(costoUnitario * cantidad);
        sheetCostos.appendRow(['LOTE_STOCK', loteId, fecha, 'BRUTO',
          descripcionItem + ` (×${cantidad})`, montoTotalLote, origen]);
        lotesCreados.push({ loteId, sku, talla, longitud, color, cantidad, costoUnitario, montoTotal: montoTotalLote, estadoViaje: estadoViajeLote });
        filasAgregadas.push({ tipo: 'STOCK', loteId, sku, precioFinal, cantidad, costoUnitario, montoTotal: montoTotalLote });
      }
    });

    ordenesActualizadas.forEach(ordenId => {
      const orderRow = findOrderRow(ordenId);
      if (orderRow !== -1) {
        const cellActual = sheetOrdenes.getRange(orderRow, colNumOrdenProv).getValue();
        if (!cellActual || String(cellActual).trim() === '' ||
            String(cellActual).toUpperCase().indexOf('POR DEFINIR') !== -1) {
          sheetOrdenes.getRange(orderRow, colNumOrdenProv).setValue(numFactura);
        }
      }
    });

    // ── #4: avanzar a PEDIDO HECHO si la factura completó los costos ──
    SpreadsheetApp.flush();
    const itemsAll = readSheetAsObjects(TABS.items);
    const tieneStock = {};
    itemsAll.forEach(it => {
      if (String(it.ORIGEN || '').toUpperCase().trim() === 'STOCK') {
        tieneStock[String(it.ORDEN_ID).trim()] = true;
      }
    });
    const pedidosConfirmados = [];
    ordenesActualizadas.forEach(ordenId => {
      const orderRow = findOrderRow(ordenId);
      if (orderRow === -1) return;
      const estadoActual = String(sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).getValue()).trim();
      if (estadoActual !== 'HACER PEDIDO') return;
      if (tieneStock[ordenId]) return;
      if (itemsPedidoSinCosto(ordenId).length > 0) return;
      sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).setValue('PEDIDO HECHO');
      logChange(usuario, ordenId, 'CAMBIO_ESTADO', 'HACER PEDIDO → PEDIDO HECHO (auto por factura)');
      pedidosConfirmados.push(ordenId);
    });

    const totalFactura = round2(subtotal + iva + shipping);
    logChange(usuario, '-', 'FACTURA_FIGS',
      `${numFactura}: ${itemsFactura.length} items, ${ordenesActualizadas.size} pedido(s), ${lotesCreados.length} lote(s), total $${totalFactura.toFixed(2)}` +
      (stockYaLlego ? '' : ' · stock EN CAMINO'));

    return jsonResponse({
      ok: true,
      data: { itemsCargados: filasAgregadas.length, pedidosAfectados: Array.from(ordenesActualizadas),
              pedidosConfirmados, lotesCreados, totalFactura, detalle: filasAgregadas }
    });
  });
}

// ============ CARGAR ENVÍO COURIER ============

// ============ CARGAR ENVÍO COURIER ============

function handleCargarCourier(body) {
  const pedidos      = body.pedidos || [];
  const lotes        = body.lotes || [];
  const itemsSueltos = body.items || [];
  const costoCourier = Number(body.costoCourier) || 0;
  const fecha        = String(body.fecha || '').trim() ||
                       Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
  const usuario      = body.usuario || 'desconocido';

  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede cargar courier', 403);
  if (pedidos.length === 0 && lotes.length === 0 && itemsSueltos.length === 0)
    return errorResponse('Selecciona al menos un ítem, pedido o lote');
  if (costoCourier <= 0) return errorResponse('Costo de courier inválido');

  const itemsPedidos = pedidos.reduce((s, p) => s + (Number(p.numItems) || 0), 0);
  const itemsLotes   = lotes.reduce((s, l) => s + (Number(l.cantInicial) || 0), 0);
  const itemsItems   = itemsSueltos.reduce((s, i) => s + (Number(i.cantidad) || 0), 0);
  const totalItems   = itemsPedidos + itemsLotes + itemsItems;
  if (totalItems <= 0) return errorResponse('El total de items debe ser mayor a 0');

  const costoPorItem = costoCourier / totalItems;

  return withLock(function() {
    const ss = getSpreadsheet();
    const sheetCostos  = ss.getSheetByName(TABS.costos);
    const sheetLotes   = ss.getSheetByName(TABS.lotes);
    const sheetItems   = ss.getSheetByName(TABS.items);
    const sheetOrdenes = ss.getSheetByName(TABS.ordenes);

    const lotesHeaders = sheetLotes.getRange(1, 1, 1, sheetLotes.getLastColumn()).getValues()[0];
    const colCostoUnit = lotesHeaders.indexOf('COSTO_UNITARIO') + 1;
    const colLoteId    = lotesHeaders.indexOf('LOTE_ID') + 1;
    const colLoteSku   = lotesHeaders.indexOf('SKU');   // 0-based

    // Leemos TablaItems una vez en memoria
    const itemsData = sheetItems.getDataRange().getValues();
    const ih = itemsData[0].map(h => String(h).trim());
    const cOrden       = ih.indexOf('ORDEN_ID');
    const cSku         = ih.indexOf('SKU');
    const cOrigen      = ih.indexOf('ORIGEN');
    const cEstItem     = ih.indexOf('ESTATUS_ITEM');
    const cCourierItem = ih.indexOf('COURIER_ITEM');

    // Estados de ítem que NO se deben pisar
    const ITEM_PROTEGIDO = ['ENTREGADO', 'CANCELADO'];

    // Marca un ítem (fila de hoja, 1-based): COURIER_ITEM=TRUE y ESTATUS_ITEM=EN BODEGA EC
    function marcarItemFila(sheetRow) {
      const idx = sheetRow - 1;
      if (cCourierItem >= 0) {
        sheetItems.getRange(sheetRow, cCourierItem + 1).setValue('TRUE');
        itemsData[idx][cCourierItem] = 'TRUE';
      }
      if (cEstItem >= 0) {
        const cur = String(itemsData[idx][cEstItem]).toUpperCase().trim();
        if (ITEM_PROTEGIDO.indexOf(cur) === -1) {
          sheetItems.getRange(sheetRow, cEstItem + 1).setValue('EN BODEGA EC');
          itemsData[idx][cEstItem] = 'EN BODEGA EC';
        }
      }
    }

    const detalle = [];
    const ordenesAfectadas = {};
    const movimientos = readSheetAsObjects(TABS.movimientos);
    const costosExistentes = readSheetAsObjects(TABS.costos);

    // ── ÍTEMS SUELTOS de pedido (courier por ítem) ──────────────────────────
    itemsSueltos.forEach(it => {
      const ordenId = String(it.ordenId).trim();
      const fila    = Number(it.itemRowNum);
      const n       = Number(it.cantidad) || 0;
      const courier = round2(costoPorItem * n);
      sheetCostos.appendRow(['PEDIDO', ordenId, fecha, 'COURIER', `Courier (item fila ${fila})`, courier, 'Courier']);
      if (fila > 1) marcarItemFila(fila);
      ordenesAfectadas[ordenId] = true;
      detalle.push({ tipo: 'ITEM', id: ordenId, fila, items: n, courier });
    });

    // ── PEDIDOS completos (modo viejo) → marca TODOS los ítems del pedido ────
    pedidos.forEach(p => {
      const ordenId = String(p.ordenId).trim();
      const n       = Number(p.numItems) || 0;
      const courier = round2(costoPorItem * n);
      sheetCostos.appendRow(['PEDIDO', ordenId, fecha, 'COURIER', `Courier (${n} items)`, courier, 'Courier']);
      for (let r = 1; r < itemsData.length; r++) {
        if (String(itemsData[r][cOrden]).trim() === ordenId) marcarItemFila(r + 1);
      }
      ordenesAfectadas[ordenId] = true;
      detalle.push({ tipo: 'PEDIDO', id: ordenId, items: n, courier });
    });

    // ── LOTES: courier del lote + retroactivo a pedidos que lo usaron ───────
    lotes.forEach(l => {
      const loteId           = String(l.loteId).trim();
      const n                = Number(l.cantInicial) || 0;
      const courierTotal     = round2(costoPorItem * n);
      const courierPorUnidad = n > 0 ? round2(courierTotal / n) : 0;

      sheetCostos.appendRow(['LOTE_STOCK', loteId, fecha, 'COURIER', `Courier (${n} items)`, courierTotal, 'Courier']);

      // Actualiza costo del lote y obtiene su SKU
      let loteSku = '';
      const dataLotes = sheetLotes.getDataRange().getValues();
      for (let i = 1; i < dataLotes.length; i++) {
        if (String(dataLotes[i][colLoteId - 1]).trim() === loteId) {
          const costoActual = Number(dataLotes[i][colCostoUnit - 1]) || 0;
          sheetLotes.getRange(i + 1, colCostoUnit).setValue(round2(costoActual + courierPorUnidad));
          if (colLoteSku >= 0) loteSku = String(dataLotes[i][colLoteSku]).trim().toUpperCase();
          break;
        }
      }
      detalle.push({ tipo: 'LOTE', id: loteId, items: n, courier: courierTotal });

      const salidas = movimientos.filter(m =>
        String(m.LOTE_ID).trim() === loteId &&
        String(m.TIPO_MOVIMIENTO).toUpperCase().trim() === 'SALIDA' &&
        String(m.TIPO_REFERENCIA).toUpperCase().trim() === 'PEDIDO'
      );
      const porPedido = {};
      salidas.forEach(m => {
        const ordenId = String(m.REFERENCIA_ID).trim();
        porPedido[ordenId] = (porPedido[ordenId] || 0) + (Number(m.CANTIDAD) || 0);
      });
      Object.keys(porPedido).forEach(ordenId => {
        const cant = porPedido[ordenId];
        const courierPedido = round2(courierPorUnidad * cant);
        const yaTiene = costosExistentes.some(c =>
          String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'PEDIDO' &&
          String(c.REFERENCIA_ID).trim() === ordenId &&
          String(c.TIPO_COSTO).toUpperCase().trim() === 'COURIER' &&
          String(c.DESCRIPCION).indexOf(loteId) !== -1
        );
        if (!yaTiene) {
          sheetCostos.appendRow([
            'PEDIDO', ordenId, fecha, 'COURIER',
            `Courier retroactivo (${cant} de ${loteId})`, courierPedido, 'Courier'
          ]);
          detalle.push({ tipo: 'PEDIDO_RETROACTIVO', id: ordenId, lote: loteId, items: cant, courier: courierPedido });
        }

        // Marca los ítems de stock de ese pedido que correspondan a este lote
        let porMarcar = cant;
        for (let r = 1; r < itemsData.length && porMarcar > 0; r++) {
          if (String(itemsData[r][cOrden]).trim() !== ordenId) continue;
          if (cOrigen >= 0 && String(itemsData[r][cOrigen]).toUpperCase().trim() !== 'STOCK') continue;
          if (cCourierItem >= 0 && String(itemsData[r][cCourierItem]).toUpperCase().trim() === 'TRUE') continue;
          if (loteSku && String(itemsData[r][cSku]).trim().toUpperCase() !== loteSku) continue;
          marcarItemFila(r + 1);
          porMarcar--;
        }
        ordenesAfectadas[ordenId] = true;
      });
    });

    SpreadsheetApp.flush();

    // ── AVANCE DE PEDIDO: si TODOS sus ítems tienen courier → EN BODEGA EC ──
    const PEDIDO_PROTEGIDO = ['ENTREGA PARCIAL', 'ENTREGADO', 'CANCELADO', 'EN BODEGA EC', 'LISTO PARA ENVIAR'];
    const pedidosAvanzados = [];

    Object.keys(ordenesAfectadas).forEach(ordenId => {
      let hay = false, todos = true;
      for (let r = 1; r < itemsData.length; r++) {
        if (String(itemsData[r][cOrden]).trim() !== ordenId) continue;
        hay = true;
        if (cCourierItem < 0 || String(itemsData[r][cCourierItem]).toUpperCase().trim() !== 'TRUE') {
          todos = false; break;
        }
      }
      if (!hay || !todos) return;

      const orderRow = findOrderRow(ordenId);
      if (orderRow === -1) return;
      const estadoActual = String(sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).getValue()).trim().toUpperCase();
      if (PEDIDO_PROTEGIDO.indexOf(estadoActual) !== -1) return;

      sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).setValue('EN BODEGA EC');
      logChange(usuario, ordenId, 'CAMBIO_ESTADO', `${estadoActual} → EN BODEGA EC (auto por courier completo)`);
      pedidosAvanzados.push(ordenId);
    });

    logChange(usuario, '-', 'COURIER_ENVIO',
      `Courier $${costoCourier.toFixed(2)} · ${itemsSueltos.length} ítem(s), ${pedidos.length} pedido(s), ${lotes.length} lote(s), ${totalItems} piezas` +
      (pedidosAvanzados.length ? ` · ${pedidosAvanzados.length} pedido(s) → EN BODEGA EC` : ''));

    return jsonResponse({
      ok: true,
      data: {
        costoCourier: round2(costoCourier),
        totalItems,
        costoPorItem: round2(costoPorItem),
        pedidosAvanzados,
        detalle
      }
    });
  });
}
// ============ CREAR PEDIDO ============

// ============ CREAR PEDIDO ============

function handleCrearPedido(body) {
  const clienteData = body.cliente || {};
  const items       = body.items || [];
  const descuento   = body.descuento || { monto: 0, nota: '' };
  const estado      = String(body.estado || 'HACER PEDIDO').trim();
  const notas       = String(body.notas || '').trim();
  const fecha       = String(body.fecha || '').trim() ||
                      Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
  const usuario     = body.usuario || 'desconocido';

  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede crear pedidos', 403);
  if (items.length === 0) return errorResponse('El pedido debe tener al menos un item');
  if (VALID_STATES.indexOf(estado) === -1) return errorResponse('Estado inválido: ' + estado);

  if (clienteData.esNuevo) {
    if (!clienteData.nombre || !String(clienteData.nombre).trim())
      return errorResponse('Falta el nombre del cliente nuevo');
  } else {
    if (!clienteData.clienteId) return errorResponse('Falta seleccionar el cliente');
  }

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!it.sku || !String(it.sku).trim()) return errorResponse(`Item ${i + 1}: falta SKU`);
    if (!it.talla || !String(it.talla).trim()) return errorResponse(`Item ${i + 1}: falta talla`);
    if (Number(it.cantidad) <= 0) return errorResponse(`Item ${i + 1}: cantidad inválida`);
    if (Number(it.precioVenta) < 0) return errorResponse(`Item ${i + 1}: precio inválido`);
  }

  return withLock(function () {
    const ss = getSpreadsheet();
    const sheetOrdenes    = ss.getSheetByName(TABS.ordenes);
    const sheetItems      = ss.getSheetByName(TABS.items);
    const sheetClientes   = ss.getSheetByName(TABS.clientes);
    const sheetDescuentos = ss.getSheetByName(TABS.descuentos);
    const sheetLotes      = ss.getSheetByName(TABS.lotes);
    const sheetMovs       = ss.getSheetByName(TABS.movimientos);
    const sheetCostos     = ss.getSheetByName(TABS.costos);

    // ──────────────────────────────────────────────────────────────
    // PLAN DE VENTA DE STOCK (validar ANTES de escribir nada)
    // ──────────────────────────────────────────────────────────────
    const lotesData = sheetLotes.getDataRange().getValues();
    const lh    = lotesData[0].map(h => String(h).trim());
    const cId    = lh.indexOf('LOTE_ID');
    const cSku   = lh.indexOf('SKU');
    const cTalla = lh.indexOf('TALLA');
    const cLong  = lh.indexOf('LONGITUD');
    const cColor = lh.indexOf('COLOR');
    const cCostoU= lh.indexOf('COSTO_UNITARIO');
    const cDisp  = lh.indexOf('CANT_DISPONIBLE');
    const cFecha = lh.indexOf('FECHA_ENTRADA');
    if ([cId, cSku, cCostoU, cDisp].some(x => x === -1))
      throw new Error('Faltan columnas clave en TablaLotesStock');

    function norm(v) { return String(v || '').trim().toLowerCase(); }

    const dispWork   = {};   // dataIdx -> disponible (copia de trabajo)
    const stockPlan  = [];   // { idx, sheetRow, dataIdx, loteId, costoUnitario, cantidad }
    const stockErrors = [];

    items.forEach((it, idx) => {
      const origen = String(it.origen || 'PEDIDO').toUpperCase();
      if (origen !== 'STOCK') return;

      const cantidad  = Number(it.cantidad) || 0;
      const loteIdReq = it.loteId ? String(it.loteId).trim() : '';

      let candIdx = [];
      for (let r = 1; r < lotesData.length; r++) {
        const row = lotesData[r];
        if (loteIdReq) {
          if (String(row[cId]).trim() === loteIdReq) candIdx.push(r);
        } else if (
          norm(row[cSku])   === norm(it.sku) &&
          norm(row[cTalla]) === norm(it.talla) &&
          norm(row[cLong])  === norm(it.longitud) &&
          norm(row[cColor]) === norm(it.color)
        ) {
          candIdx.push(r);
        }
      }

      if (cFecha !== -1) {
        candIdx.sort((a, b) =>
          String(lotesData[a][cFecha]).localeCompare(String(lotesData[b][cFecha])));
      }

      let chosen = -1;
      for (const r of candIdx) {
        const disp = (dispWork[r] !== undefined) ? dispWork[r] : (Number(lotesData[r][cDisp]) || 0);
        if (disp >= cantidad) { chosen = r; break; }
      }

      if (chosen === -1) {
        stockErrors.push(
          `Ítem ${idx + 1} (${it.sku} ${it.talla} ${it.color}): sin stock disponible` +
          (loteIdReq ? ` en ${loteIdReq}` : ''));
        return;
      }

      const dispNow = (dispWork[chosen] !== undefined) ? dispWork[chosen] : (Number(lotesData[chosen][cDisp]) || 0);
      dispWork[chosen] = round2(dispNow - cantidad);

      stockPlan.push({
        idx,
        sheetRow: chosen + 1,
        dataIdx: chosen,
        loteId: String(lotesData[chosen][cId]).trim(),
        costoUnitario: Number(lotesData[chosen][cCostoU]) || 0,
        cantidad
      });
    });

    if (stockErrors.length > 0) {
      return errorResponse('No se pudo crear el pedido por falta de stock:\n' + stockErrors.join('\n'), 422);
    }

    // ──────────────────────────────────────────────────────────────
    // CLIENTE
    // ──────────────────────────────────────────────────────────────
    let nombreCliente, clienteIdFinal;

    if (clienteData.esNuevo) {
      clienteIdFinal = nextClienteId();
      nombreCliente  = String(clienteData.nombre).trim();
      const cRow     = sheetClientes.getLastRow() + 1;
      sheetClientes.appendRow([
        clienteIdFinal, nombreCliente,
        "'" + String(clienteData.telefono || ''),
        String(clienteData.direccion || ''),
        String(clienteData.ciudad    || ''),
        String(clienteData.cedulaRuc || ''),
        String(clienteData.industria || ''),
        "'" + String(clienteData.email || ''),
        fecha, '', '', '', 0, '', ''
      ]);
      sheetClientes.getRange(cRow, 10).setFormula(`=COUNTIFS(TablaOrdenes!C:C,B${cRow},TablaOrdenes!D:D,"<>CANCELADO")`);
      sheetClientes.getRange(cRow, 11).setFormula(`=SUMIFS(TablaOrdenes!R:R,TablaOrdenes!C:C,B${cRow},TablaOrdenes!D:D,"<>CANCELADO")`);
      sheetClientes.getRange(cRow, 12).setFormula(`=IFERROR(MAXIFS(TablaOrdenes!B:B,TablaOrdenes!C:C,B${cRow}),"")`);
      sheetClientes.getRange(cRow, 14).setFormula(`=IF(J${cRow}>=3,TRUE,FALSE)`);
      if (clienteData.industria) agregarIndustria(clienteData.industria);
    } else {
      const clientes = readSheetAsObjects(TABS.clientes);
      const cli = clientes.find(c => String(c.CLIENTE_ID).trim() === String(clienteData.clienteId).trim());
      if (!cli) throw new Error('Cliente no encontrado: ' + clienteData.clienteId);
      clienteIdFinal = cli.CLIENTE_ID;
      nombreCliente  = cli.NOMBRE;
    }

    // ──────────────────────────────────────────────────────────────
    // ORDEN
    // ──────────────────────────────────────────────────────────────
    const ordenId = nextOrdenId();
    const oRow    = sheetOrdenes.getLastRow() + 1;

    sheetOrdenes.appendRow([
      ordenId, fecha, nombreCliente, estado,
      'POR DEFINIR', 'PEDIDO', 'FIGS',
      '', '', '', '', '', '', 0, notas
    ]);

    sheetOrdenes.getRange(oRow, 10).setFormula(`=IFERROR(IF(VLOOKUP(C${oRow},TablaClientes!B:E,4,FALSE)="Guayaquil","Delivery local","Servientrega"),"")`);
    sheetOrdenes.getRange(oRow, 16).setFormula(`=SUMIFS(TablaItems!L:L,TablaItems!A:A,A${oRow})`);
    sheetOrdenes.getRange(oRow, 17).setFormula(`=SUMIFS(TablaDescuentos!E:E,TablaDescuentos!B:B,A${oRow})`);
    sheetOrdenes.getRange(oRow, 18).setFormula(`=P${oRow}+Q${oRow}`);
    sheetOrdenes.getRange(oRow, 19).setFormula(`=SUMIFS(TablaPagos!C:C,TablaPagos!A:A,A${oRow})`);
    sheetOrdenes.getRange(oRow, 20).setFormula(`=R${oRow}-S${oRow}`);
    sheetOrdenes.getRange(oRow, 21).setFormula(`=IF(S${oRow}=0,"Sin pago",IF(S${oRow}<R${oRow},"Parcial",IF(S${oRow}=R${oRow},"Pagado total","Sobrepago")))`);
    sheetOrdenes.getRange(oRow, 22).setFormula(`=SUMIFS(TablaCostos!F:F,TablaCostos!A:A,"PEDIDO",TablaCostos!B:B,A${oRow})`);
    sheetOrdenes.getRange(oRow, 23).setFormula(`=R${oRow}-V${oRow}`);
    sheetOrdenes.getRange(oRow, 24).setFormula(`=IF(OR(AND(D${oRow}="ENTREGADO",T${oRow}<=0),D${oRow}="CANCELADO"),FALSE,TRUE)`);

    // ──────────────────────────────────────────────────────────────
    // ITEMS (guardando la fila de cada uno)
    // ──────────────────────────────────────────────────────────────
    const itemRows = [];
    items.forEach((it, idx) => {
      const iRow = sheetItems.getLastRow() + 1;
      sheetItems.appendRow([
        ordenId, String(it.sku).trim(), '', '',
        String(it.talla).trim(), String(it.longitud || '').trim(),
        String(it.color || '').trim(), Number(it.cantidad), Number(it.precioVenta),
        String(it.parteDeSet || '').trim(), '', ''
      ]);
      sheetItems.getRange(iRow, 3).setFormula(`=IFERROR(VLOOKUP(B${iRow},TablaProductos!A:B,2,FALSE),"")`);
      sheetItems.getRange(iRow, 4).setFormula(`=IFERROR(VLOOKUP(B${iRow},TablaProductos!A:C,3,FALSE),"")`);
      sheetItems.getRange(iRow, 12).setFormula(`=H${iRow}*I${iRow}`);
      itemRows[idx] = iRow;
    });

    // ──────────────────────────────────────────────────────────────
    // APLICAR VENTA DE STOCK (descontar + costear + movimiento)
    // ──────────────────────────────────────────────────────────────
    const stockAplicado = [];
    // Leer costos existentes una vez para verificar courier por lote
    const costosActuales = readSheetAsObjects(TABS.costos);

    const itemsHeadersFull = sheetItems.getRange(1, 1, 1, sheetItems.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const cCourierItem = itemsHeadersFull.indexOf('COURIER_ITEM') + 1;
    const cEstatusIt   = itemsHeadersFull.indexOf('ESTATUS_ITEM') + 1;
    const cOrigenIt    = itemsHeadersFull.indexOf('ORIGEN') + 1;

    stockPlan.forEach(p => {
      const it   = items[p.idx];
      const iRow = itemRows[p.idx];
      const descripcion = `${String(it.sku).trim()} ${String(it.talla).trim()}-${String(it.longitud || '').trim()} ${String(it.color || '').trim()}`.trim();

      // 1. Copiar costo del lote al item (columna 11 = COSTO_UNITARIO)
      sheetItems.getRange(iRow, 11).setValue(p.costoUnitario);

      // 2. Descontar disponibilidad del lote
      sheetLotes.getRange(p.sheetRow, cDisp + 1).setValue(dispWork[p.dataIdx]);

      // 3. Movimiento de salida
      const ahora = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd HH:mm:ss');
      sheetMovs.appendRow([
        nextMovId(), ahora, p.loteId, 'SALIDA', p.cantidad,
        'PEDIDO', ordenId, usuario, 'Venta de stock al crear pedido'
      ]);

      // 4. Fila de costo PEDIDO en TablaCostos
      sheetCostos.appendRow([
        'PEDIDO', ordenId, fecha, 'BRUTO', descripcion,
        round2(p.costoUnitario * p.cantidad), `Stock ${p.loteId}`
      ]);

      // 5. Si el lote ya tiene courier pagado → marcar ítem como EN BODEGA EC
      const loteYaTieneCourier = costosActuales.some(c =>
        String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'LOTE_STOCK' &&
        String(c.REFERENCIA_ID).trim() === p.loteId &&
        String(c.TIPO_COSTO).toUpperCase().trim() === 'COURIER'
      );
      if (loteYaTieneCourier) {
        if (cCourierItem > 0) sheetItems.getRange(iRow, cCourierItem).setValue('TRUE');
        if (cEstatusIt  > 0) sheetItems.getRange(iRow, cEstatusIt).setValue('EN BODEGA EC');
      }

      // 6. Marcar origen como STOCK
      if (cOrigenIt > 0) sheetItems.getRange(iRow, cOrigenIt).setValue('STOCK');

      stockAplicado.push({ ordenId, loteId: p.loteId, cantidad: p.cantidad, costoUnitario: p.costoUnitario });
    });

    // ──────────────────────────────────────────────────────────────
    // DESCUENTO
    // ──────────────────────────────────────────────────────────────
    const montoDesc = Number(descuento.monto) || 0;
    if (montoDesc > 0) {
      sheetDescuentos.appendRow([
        nextDescuentoId(), ordenId, fecha,
        'Precio negociado', -Math.abs(montoDesc), String(descuento.nota || '')
      ]);
    }

    logChange(usuario, ordenId, 'PEDIDO_CREADO',
      `${nombreCliente} · ${items.length} item(s)` +
      (stockAplicado.length > 0 ? ` · ${stockAplicado.length} de stock` : '') +
      (clienteData.esNuevo ? ` · cliente nuevo ${clienteIdFinal}` : '') +
      (montoDesc > 0 ? ` · descuento $${montoDesc.toFixed(2)}` : ''));

    return jsonResponse({
      ok: true,
      data: {
        ordenId, clienteId: clienteIdFinal, clienteNuevo: !!clienteData.esNuevo,
        itemsCreados: items.length, descuentoAplicado: montoDesc,
        stockAplicado
      }
    });
  });
}

// ============ ASIGNAR STOCK A PEDIDOS ============

function handleAsignarStock(body) {
  const asignaciones = body.asignaciones || [];
  const fecha    = String(body.fecha || '').trim() ||
                   Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
  const usuario  = body.usuario || 'desconocido';

  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede asignar stock', 403);
  if (asignaciones.length === 0) return errorResponse('No hay items para asignar');

  for (let i = 0; i < asignaciones.length; i++) {
    const a = asignaciones[i];
    if (!a.ordenId || !a.itemRowNum) return errorResponse(`Asignación ${i}: falta ordenId o itemRowNum`);
    if (!a.sku || Number(a.cantidad) <= 0) return errorResponse(`Asignación ${i}: falta sku o cantidad inválida`);
  }

  return withLock(function () {
    const ss = getSpreadsheet();
    const sheetItems  = ss.getSheetByName(TABS.items);
    const sheetLotes  = ss.getSheetByName(TABS.lotes);
    const sheetCostos = ss.getSheetByName(TABS.costos);
    const sheetMovs   = ss.getSheetByName(TABS.movimientos);

    const itemsHeaders = sheetItems.getRange(1, 1, 1, sheetItems.getLastColumn()).getValues()[0];
    const colCostoUnitario = itemsHeaders.indexOf('COSTO_UNITARIO') + 1;
    if (colCostoUnitario === 0) throw new Error('No se encontró columna COSTO_UNITARIO en TablaItems');

    const lotesHeaders     = sheetLotes.getRange(1, 1, 1, sheetLotes.getLastColumn()).getValues()[0];
    const colLoteId        = lotesHeaders.indexOf('LOTE_ID') + 1;
    const colLoteSku       = lotesHeaders.indexOf('SKU') + 1;
    const colLoteTalla     = lotesHeaders.indexOf('TALLA') + 1;
    const colLoteLongitud  = lotesHeaders.indexOf('LONGITUD') + 1;
    const colLoteColor     = lotesHeaders.indexOf('COLOR') + 1;
    const colLoteCostoUnit = lotesHeaders.indexOf('COSTO_UNITARIO') + 1;
    const colLoteDisp      = lotesHeaders.indexOf('CANT_DISPONIBLE') + 1;
    const colLoteFecha     = lotesHeaders.indexOf('FECHA_ENTRADA') + 1;
    if (colLoteId === 0 || colLoteDisp === 0) throw new Error('No se encontraron columnas clave en TablaLotesStock');

    function norm(v) { return String(v || '').trim().toLowerCase(); }

    const lotesData = sheetLotes.getDataRange().getValues();
    const asignados = [];
    const sinStock  = [];

    asignaciones.forEach(a => {
      const ordenId           = String(a.ordenId).trim();
      const itemRow           = Number(a.itemRowNum);
      const cantidadNecesaria = Number(a.cantidad);

      const candidatos = [];
      for (let r = 1; r < lotesData.length; r++) {
        const row = lotesData[r];
        if (norm(row[colLoteSku - 1])      !== norm(a.sku))      continue;
        if (norm(row[colLoteTalla - 1])    !== norm(a.talla))    continue;
        if (norm(row[colLoteLongitud - 1]) !== norm(a.longitud)) continue;
        if (norm(row[colLoteColor - 1])    !== norm(a.color))    continue;
        const disp = Number(row[colLoteDisp - 1]) || 0;
        if (disp < cantidadNecesaria) continue;
        candidatos.push({ sheetRow: r + 1, dataIdx: r });
      }

      candidatos.sort((x, y) =>
        String(lotesData[x.dataIdx][colLoteFecha - 1])
          .localeCompare(String(lotesData[y.dataIdx][colLoteFecha - 1]))
      );

      if (candidatos.length === 0) {
        sinStock.push({ ordenId, itemRowNum: itemRow, sku: a.sku, motivo: 'Ningún lote cubre la cantidad necesaria' });
        return;
      }

      const elegido = candidatos[0];
      const loteRow = elegido.sheetRow;
      const loteId  = String(lotesData[elegido.dataIdx][colLoteId - 1]).trim();

      // Nota: si el lote no tiene courier, se asigna igual.
      // El courier retroactivo se completa al cargar el courier del lote.

      const dispActual = Number(sheetLotes.getRange(loteRow, colLoteDisp).getValue()) || 0;
      if (dispActual < cantidadNecesaria) {
        sinStock.push({ ordenId, itemRowNum: itemRow, sku: a.sku, loteId, motivo: 'Stock agotado durante la asignación' });
        return;
      }

      const costoUnitario = Number(lotesData[elegido.dataIdx][colLoteCostoUnit - 1]) || 0;
      const descripcion   = `${a.sku} ${a.talla}-${a.longitud} ${a.color}`.trim();

      sheetItems.getRange(itemRow, colCostoUnitario).setValue(costoUnitario);
      sheetCostos.appendRow(['PEDIDO', ordenId, fecha, 'BRUTO', descripcion, costoUnitario, `Stock ${loteId}`]);

      const nuevaDisp = round2(dispActual - cantidadNecesaria);
      sheetLotes.getRange(loteRow, colLoteDisp).setValue(nuevaDisp);

      const ahora = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd HH:mm:ss');
      sheetMovs.appendRow([nextMovId(), ahora, loteId, 'SALIDA', cantidadNecesaria,
        'PEDIDO', ordenId, usuario, `Asignado a item fila ${itemRow}`]);

      lotesData[elegido.dataIdx][colLoteDisp - 1] = nuevaDisp;
      asignados.push({ ordenId, itemRowNum: itemRow, sku: a.sku, loteId, costoUnitario });
    });

    logChange(usuario, '-', 'ASIGNAR_STOCK',
      `${asignados.length} item(s) asignado(s) desde stock` +
      (sinStock.length > 0 ? ` · ${sinStock.length} sin stock` : ''));

    return jsonResponse({ ok: true, data: { asignados, sinStock } });
  });
}
// ============ VERIFICAR INTEGRIDAD ============

function verificarIntegridad() {
  const ordenes   = readSheetAsObjects(TABS.ordenes);
  const items     = readSheetAsObjects(TABS.items);
  const costos    = readSheetAsObjects(TABS.costos);
  const pagos     = readSheetAsObjects(TABS.pagos);
  const productos = readSheetAsObjects(TABS.productos);
  const lotes     = readSheetAsObjects(TABS.lotes);

  const hallazgos = [];
  function add(severidad, tipo, ubicacion, detalle) {
    hallazgos.push({ severidad, tipo, ubicacion, detalle });
  }

  const ordenIds = {};
  ordenes.forEach(o => { ordenIds[String(o.ORDEN_ID).trim()] = String(o.ESTATUS_ENVIO).trim(); });

  const skusValidos = {};
  productos.forEach(p => { skusValidos[String(p.SKU).trim()] = true; });

  ordenes.forEach(orden => {
    const id = String(orden.ORDEN_ID).trim();
    if (String(orden.ESTATUS_ENVIO).trim() === 'CANCELADO') return;
    if (!items.some(i => String(i.ORDEN_ID).trim() === id)) {
      add('ALTA', 'PEDIDO_SIN_ITEMS', `TablaOrdenes fila ${orden._rowNum} (${id})`,
        `El pedido ${id} no tiene ningún item en TablaItems.`);
    }
  });

  items.forEach(item => {
    const ordenId = String(item.ORDEN_ID).trim();
    if (!ordenIds.hasOwnProperty(ordenId)) {
      add('ALTA', 'ITEM_HUERFANO', `TablaItems fila ${item._rowNum}`,
        `El item referencia el pedido ${ordenId}, que no existe en TablaOrdenes.`);
    }
  });

  items.forEach(item => {
    const ordenId = String(item.ORDEN_ID).trim();
    const estado  = ordenIds[ordenId];
    if (estado === undefined || estado === 'CANCELADO') return;
    const costo = item.COSTO_UNITARIO;
    const vacio = costo === '' || costo === null || costo === undefined ||
                  (typeof costo === 'number' && isNaN(costo));
    if (vacio) {
      add('MEDIA', 'ITEM_SIN_COSTO', `TablaItems fila ${item._rowNum}`,
        `Item ${item.SKU} del pedido ${ordenId} no tiene COSTO_UNITARIO. La ganancia sale inflada.`);
    }
  });

  pagos.forEach(pago => {
    const ordenId = String(pago.ORDEN_ID).trim();
    if (!ordenIds.hasOwnProperty(ordenId)) {
      add('ALTA', 'PAGO_HUERFANO', `TablaPagos fila ${pago._rowNum}`,
        `El pago referencia el pedido ${ordenId}, que no existe en TablaOrdenes.`);
    }
  });

  costos.forEach(costo => {
    const tipoRef = String(costo.TIPO_REFERENCIA).toUpperCase().trim();
    const refId   = String(costo.REFERENCIA_ID).trim();
    if (tipoRef === 'PEDIDO' && !ordenIds.hasOwnProperty(refId)) {
      add('MEDIA', 'COSTO_HUERFANO', `TablaCostos fila ${costo._rowNum}`,
        `Costo tipo PEDIDO referencia ${refId}, que no existe en TablaOrdenes.`);
    }
  });

  items.forEach(item => {
    const sku = String(item.SKU).trim();
    if (sku && !skusValidos.hasOwnProperty(sku)) {
      add('MEDIA', 'SKU_INVALIDO_ITEM', `TablaItems fila ${item._rowNum}`,
        `El item usa el SKU "${sku}", que no existe en TablaProductos.`);
    }
  });

  lotes.forEach(lote => {
    const sku = String(lote.SKU).trim();
    if (sku && !skusValidos.hasOwnProperty(sku)) {
      add('MEDIA', 'SKU_INVALIDO_LOTE', `TablaLotesStock fila ${lote._rowNum} (${lote.LOTE_ID})`,
        `El lote usa el SKU "${sku}", que no existe en TablaProductos.`);
    }
  });

  const resumen = {
    totalHallazgos: hallazgos.length,
    alta:  hallazgos.filter(h => h.severidad === 'ALTA').length,
    media: hallazgos.filter(h => h.severidad === 'MEDIA').length,
    baja:  hallazgos.filter(h => h.severidad === 'BAJA').length,
    sano:  hallazgos.length === 0
  };

  return { resumen, hallazgos };
}

function probarIntegridad() {
  Logger.log(JSON.stringify(verificarIntegridad(), null, 2));
}

// ============ CAMBIAR ESTADO DE ÍTEMS (ENTREGA PARCIAL) ============

function handleCambiarEstadoItems(body) {
  const ordenId       = String(body.ordenId || '').trim();
  const itemRows      = (body.itemRows || []).map(Number);
  const nuevoEstado   = String(body.nuevoEstado || '').trim();
  const usuario       = body.usuario || 'desconocido';
  const forzar        = body.forzar === true || String(body.forzar) === 'true';
  const tipoEmpaque   = String(body.tipoEmpaque || '').trim();
  const costoDelivery = Number(body.costoDelivery) || 0;

  if (!ordenId) return errorResponse('Falta ordenId');
  if (!nuevoEstado) return errorResponse('Falta nuevoEstado');
  if (VALID_STATES.indexOf(nuevoEstado) === -1) return errorResponse('Estado inválido: ' + nuevoEstado);
  if (nuevoEstado === 'ENTREGA PARCIAL') return errorResponse('ENTREGA PARCIAL es automático, no se asigna a mano');
  if (!canCambiarEstado(usuario)) return errorResponse('No tienes permiso para cambiar estados', 403);
  if (itemRows.length === 0) return errorResponse('No se seleccionaron ítems');
  //if (nuevoEstado === 'LISTO PARA ENVIAR' && !tipoEmpaque) {
  //  return errorResponse('Falta seleccionar el tipo de empaque', 400);
  //}

  const orderRow = findOrderRow(ordenId);
  if (orderRow === -1) return errorResponse('Pedido no encontrado: ' + ordenId, 404);

  return withLock(function () {
    const ss = getSpreadsheet();
    const sheetItems   = ss.getSheetByName(TABS.items);
    const sheetOrdenes = ss.getSheetByName(TABS.ordenes);
    const sheetCostos  = ss.getSheetByName(TABS.costos);

    const ih = sheetItems.getRange(1, 1, 1, sheetItems.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const cOrdenId   = ih.indexOf('ORDEN_ID') + 1;
    const cEstatusIt = ih.indexOf('ESTATUS_ITEM') + 1;
    const cEntregado = ih.indexOf('ENTREGADO_ITEM') + 1;
    const cCosto     = ih.indexOf('COSTO_UNITARIO') + 1;
    const cOrigen    = ih.indexOf('ORIGEN') + 1;
    if (cEstatusIt === 0 || cEntregado === 0) {
      throw new Error('Faltan columnas ESTATUS_ITEM / ENTREGADO_ITEM en TablaItems');
    }

    const estadoHeaderActual = String(sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).getValue()).trim();

    const allData = sheetItems.getDataRange().getValues();
    const filasPedido = [];
    for (let r = 1; r < allData.length; r++) {
      if (String(allData[r][cOrdenId - 1]).trim() === ordenId) {
        filasPedido.push(r + 1);
      }
    }
    if (filasPedido.length === 0) return errorResponse('El pedido no tiene ítems', 404);

    // ── VALIDACIÓN: las filas deben pertenecer a este pedido ──
    const setPedido = new Set(filasPedido);
    const fuera = itemRows.filter(f => !setPedido.has(f));
    if (fuera.length > 0) {
      return errorResponse(
        `Las filas ${fuera.join(', ')} no pertenecen al pedido ${ordenId}. ` +
        `Filas válidas: ${filasPedido.join(', ')}`, 400);
    }

    // ── Bloqueo: no avanzar ítems PEDIDO sin factura FIGS (sin costo) ──
    if (requiereFactura(nuevoEstado)) {
      const sinCosto = [];
      itemRows.forEach(fila => {
        const origen = cOrigen > 0 ? String(sheetItems.getRange(fila, cOrigen).getValue()).toUpperCase().trim() : '';
        if (origen === 'STOCK') return;
        const costo = cCosto > 0 ? Number(sheetItems.getRange(fila, cCosto).getValue()) || 0 : 0;
        if (costo <= 0) sinCosto.push(fila);
      });
      if (sinCosto.length > 0) {
        if (!forzar) {
          return errorResponse(
            `${sinCosto.length} ítem(s) no tienen la factura FIGS cargada (sin costo). ` +
            `Carga la factura antes de pasarlos a "${nuevoEstado}".`, 422);
        }
        if (!verifyAdmin(usuario)) {
          return errorResponse('Hay ítems sin factura. Solo un admin puede forzar.', 403);
        }
      }
    }

    // ── Migración suave de ítems sin estado propio ──
    filasPedido.forEach(fila => {
      const est = String(sheetItems.getRange(fila, cEstatusIt).getValue()).trim();
      if (!est) {
        sheetItems.getRange(fila, cEstatusIt).setValue(estadoHeaderActual);
        sheetItems.getRange(fila, cEntregado).setValue(estadoHeaderActual === 'ENTREGADO' ? 'TRUE' : 'FALSE');
      }
    });

    // ── Validar saldo SOLO si esta entrega completa todo el pedido ──
    const seleccion = new Set(itemRows);
    let todosEntregadosDespues = true;
    filasPedido.forEach(fila => {
      let entregado;
      if (seleccion.has(fila)) entregado = (nuevoEstado === 'ENTREGADO');
      else entregado = String(sheetItems.getRange(fila, cEntregado).getValue()).toUpperCase().trim() === 'TRUE';
      if (!entregado) todosEntregadosDespues = false;
    });

    if (nuevoEstado === 'ENTREGADO' && todosEntregadosDespues) {
      const saldo = calcularSaldoPedido(ordenId);
      if (saldo > 0.005) {
        if (!forzar) {
          return errorResponse(
            `El pedido ${ordenId} tiene saldo pendiente de $${saldo.toFixed(2)}. ` +
            `No puede completarse la entrega hasta pagarse.`, 422);
        }
        if (!verifyAdmin(usuario)) {
          return errorResponse(
            `El pedido ${ordenId} debe $${saldo.toFixed(2)}. Solo un admin puede forzar.`, 403);
        }
      }
    }

    // ── Empaque y pin (al pasar a LISTO PARA ENVIAR) ──

let costoEmpaqueFinal = 0;
let costosPinesFinal = 0;
if (nuevoEstado === 'LISTO PARA ENVIAR') {
  // Empaque
  const sets       = readSheetAsObjects(TABS.setEmpaque);
  const materiales = readSheetAsObjects(TABS.empaque);
  const cantidadCajas = Math.max(1, Number(body.cantidadCajas) || 1);
  const setElegido = sets.find(s => String(s.SET_ID).trim() === tipoEmpaque);
  if (setElegido) {
    const ids = String(setElegido.MATERIALES).split(',').map(m => m.trim());
    const costoUnaCaja = round2(ids.reduce((sum, id) => {
      const mat = materiales.find(m => String(m.MATERIAL_ID).trim() === id);
      return sum + (mat ? Number(mat.COSTO_UNITARIO) || 0 : 0);
    }, 0));
    costoEmpaqueFinal = round2(costoUnaCaja * cantidadCajas);
    const fechaHoy = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
    const desc = cantidadCajas > 1 ? `${setElegido.NOMBRE_SET} ×${cantidadCajas}` : `${setElegido.NOMBRE_SET}`;
    sheetCostos.appendRow(['PEDIDO', ordenId, fechaHoy, 'EMPAQUE',
      desc, costoEmpaqueFinal, 'Auto']);
    sheetOrdenes.getRange(orderRow, 25).setValue(desc);
    logChange(usuario, ordenId, 'COSTO_EMPAQUE',
      `${desc} · $${costoEmpaqueFinal.toFixed(2)} · ${itemRows.length} ítem(s)`);
  }

  // Pines
  const pinesSeleccionados = body.pines || []; // [{ regaloid, cantidad }]
  if (pinesSeleccionados.length > 0) {
    const sheetRegalos = ss.getSheetByName(TABS.regalos);
    const regalosData  = sheetRegalos.getDataRange().getValues();
    const rh = regalosData[0].map(h => String(h).trim());
    const cRId    = rh.indexOf('REGALO_ID');
    const cRStock = rh.indexOf('STOCK');
    const cRCosto = rh.indexOf('COSTO_UNITARIO');
    const cRNombre= rh.indexOf('NOMBRE');
    const fechaHoy = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');

    pinesSeleccionados.forEach(p => {
      const regaloid = String(p.regaloid).trim();
      const cantidad = Number(p.cantidad) || 1;

      // Buscar fila del pin
      for (let r = 1; r < regalosData.length; r++) {
        if (String(regalosData[r][cRId]).trim() !== regaloid) continue;

        const stockActual = Number(regalosData[r][cRStock]) || 0;
        const costo       = Number(regalosData[r][cRCosto]) || 0;
        const nombre      = String(regalosData[r][cRNombre]).trim();

        if (stockActual < cantidad) {
          throw new Error(`Stock insuficiente para el pin "${nombre}" (disponible: ${stockActual})`);
        }

        // Descontar stock
        sheetRegalos.getRange(r + 1, cRStock + 1).setValue(stockActual - cantidad);

        // Registrar costo
        const costoTotal = round2(costo * cantidad);
        sheetCostos.appendRow(['PEDIDO', ordenId, fechaHoy, 'REGALO',
          `${nombre} ×${cantidad}`, costoTotal, 'Auto']);
        costosPinesFinal = round2(costosPinesFinal + costoTotal);

        logChange(usuario, ordenId, 'PIN_ASIGNADO',
          `${nombre} ×${cantidad} · $${costoTotal.toFixed(2)}`);

        // Guardar en REGALO_ENVIADO de TablaOrdenes (columna M = 13)
        const regalosExistentes = String(sheetOrdenes.getRange(orderRow, 13).getValue()).trim();
        const nuevoRegalo = regalosExistentes
          ? regalosExistentes + ', ' + nombre + (cantidad > 1 ? ` ×${cantidad}` : '')
          : nombre + (cantidad > 1 ? ` ×${cantidad}` : '');
        sheetOrdenes.getRange(orderRow, 13).setValue(nuevoRegalo);
        break;
      }
    });
  }
}

    // ── Costo de delivery (al entregar) ──
    if (nuevoEstado === 'ENTREGADO' && costoDelivery > 0) {
      const fechaHoy = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
      sheetCostos.appendRow(['PEDIDO', ordenId, fechaHoy, 'DELIVERY',
        'Delivery local', round2(costoDelivery), 'Manual']);
      logChange(usuario, ordenId, 'COSTO_DELIVERY', `$${costoDelivery.toFixed(2)}`);
    }

    // ── Si se están CANCELANDO ítems de stock: devolver las piezas al lote ──
    let stockDevuelto = [];
    if (nuevoEstado === 'CANCELADO') {
      stockDevuelto = devolverStockPorCancelacion(ordenId, itemRows, usuario);
    }
    // ── Aplicar el nuevo estado a los ítems seleccionados ──
    itemRows.forEach(fila => {
      sheetItems.getRange(fila, cEstatusIt).setValue(nuevoEstado);
      sheetItems.getRange(fila, cEntregado).setValue(nuevoEstado === 'ENTREGADO' ? 'TRUE' : 'FALSE');
    });

    // ── Recalcular cabecera (resumen) ──
    let algunoEntregado = false, todosEntregados = true, minIdx = 999;
    filasPedido.forEach(fila => {
      const ent = String(sheetItems.getRange(fila, cEntregado).getValue()).toUpperCase().trim() === 'TRUE';
      const est = String(sheetItems.getRange(fila, cEstatusIt).getValue()).trim();
      if (ent) algunoEntregado = true; else todosEntregados = false;
      const idx = VALID_STATES.indexOf(est);
      if (idx !== -1 && idx < minIdx) minIdx = idx;
    });

    let headerNuevo;
    if (todosEntregados) headerNuevo = 'ENTREGADO';
    else if (algunoEntregado) headerNuevo = 'ENTREGA PARCIAL';
    else headerNuevo = (minIdx >= 0 && minIdx < 999) ? VALID_STATES[minIdx] : estadoHeaderActual;

    sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).setValue(headerNuevo);

    // ── Fecha de entrega real cuando el pedido completo queda ENTREGADO ──
    if (headerNuevo === 'ENTREGADO') {
      marcarFechaEntrega(sheetOrdenes, orderRow);
    }

    logChange(usuario, ordenId, 'CAMBIO_ESTADO_ITEMS',
      `${itemRows.length} ítem(s) → ${nuevoEstado} · cabecera: ${estadoHeaderActual} → ${headerNuevo}`);

    return jsonResponse({
      ok: true,
      data: {
        ordenId, itemsCambiados: itemRows.length, nuevoEstadoItems: nuevoEstado,
        estadoCabecera: headerNuevo, costoEmpaque: costoEmpaqueFinal, costoDelivery: round2(costoDelivery),
        stockDevuelto
      }
    });
  });
}

// ============ DEVOLVER STOCK AL CANCELAR ÍTEMS ============
// Devuelve al lote las piezas de ORIGEN=STOCK de los ítems que se cancelan,
// borra su costo BRUTO del pedido y deja rastro en TablaMovimientosStock.
function devolverStockPorCancelacion(ordenId, filas, usuario) {
  const ss = getSpreadsheet();
  const sheetItems  = ss.getSheetByName(TABS.items);
  const sheetLotes  = ss.getSheetByName(TABS.lotes);
  const sheetMovs   = ss.getSheetByName(TABS.movimientos);
  const sheetCostos = ss.getSheetByName(TABS.costos);

  const ih = sheetItems.getRange(1, 1, 1, sheetItems.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const cSku = ih.indexOf('SKU') + 1, cTalla = ih.indexOf('TALLA') + 1;
  const cColor = ih.indexOf('COLOR') + 1, cCant = ih.indexOf('CANTIDAD') + 1;
  const cOrigen = ih.indexOf('ORIGEN') + 1, cEstIt = ih.indexOf('ESTATUS_ITEM') + 1;
  const cSub = ih.indexOf('SUBTOTAL') + 1;

  const movimientos = readSheetAsObjects(TABS.movimientos);
  const lotesObj    = readSheetAsObjects(TABS.lotes);

  // Lotes que YA se devolvieron por cancelación en este pedido (evita duplicar)
  const yaDevueltos = {};
  movimientos.forEach(m => {
    if (String(m.TIPO_MOVIMIENTO).toUpperCase().trim() !== 'ENTRADA') return;
    if (String(m.TIPO_REFERENCIA).toUpperCase().trim() !== 'CANCELACION') return;
    if (String(m.REFERENCIA_ID).trim() !== ordenId) return;
    const l = String(m.LOTE_ID).trim();
    yaDevueltos[l] = (yaDevueltos[l] || 0) + (Number(m.CANTIDAD) || 0);
  });

  // Salidas de stock de este pedido, agrupadas por lote
  const salidas = movimientos.filter(m =>
    String(m.TIPO_MOVIMIENTO).toUpperCase().trim() === 'SALIDA' &&
    String(m.TIPO_REFERENCIA).toUpperCase().trim() === 'PEDIDO' &&
    String(m.REFERENCIA_ID).trim() === ordenId
  );

  const lotesHeaders = sheetLotes.getRange(1, 1, 1, sheetLotes.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const colLoteId = lotesHeaders.indexOf('LOTE_ID') + 1;
  const colDisp   = lotesHeaders.indexOf('CANT_DISPONIBLE') + 1;

  function norm(v) { return String(v || '').trim().toLowerCase(); }
  const devueltos = [];
  const usados = {};

  filas.forEach(fila => {
    const origen = cOrigen > 0 ? String(sheetItems.getRange(fila, cOrigen).getValue()).toUpperCase().trim() : '';
    if (origen !== 'STOCK') return;
    const est = String(sheetItems.getRange(fila, cEstIt).getValue()).toUpperCase().trim();
    if (est === 'CANCELADO') return; // ya estaba cancelado, no devolver dos veces

    const sku   = String(sheetItems.getRange(fila, cSku).getValue()).trim();
    const talla = String(sheetItems.getRange(fila, cTalla).getValue()).trim();
    const color = String(sheetItems.getRange(fila, cColor).getValue()).trim();
    const cant  = Number(sheetItems.getRange(fila, cCant).getValue()) || 1;

    // Buscar la salida de este pedido cuyo lote coincida con el ítem
    let loteId = '';
    for (const m of salidas) {
      const lid = String(m.LOTE_ID).trim();
      const lote = lotesObj.find(l => String(l.LOTE_ID).trim() === lid);
      if (!lote) continue;
      if (norm(lote.SKU) !== norm(sku)) continue;
      if (norm(lote.TALLA) !== norm(talla)) continue;
      if (norm(lote.COLOR) !== norm(color)) continue;
      const cupo = (Number(m.CANTIDAD) || 0) - (yaDevueltos[lid] || 0) - (usados[lid] || 0);
      if (cupo < cant) continue;
      loteId = lid;
      break;
    }
    if (!loteId) return; // no se halló el lote de origen; no se toca nada

    usados[loteId] = (usados[loteId] || 0) + cant;

    // 1. Sumar de vuelta al lote
    const dataLotes = sheetLotes.getDataRange().getValues();
    for (let r = 1; r < dataLotes.length; r++) {
      if (String(dataLotes[r][colLoteId - 1]).trim() !== loteId) continue;
      const disp = Number(dataLotes[r][colDisp - 1]) || 0;
      sheetLotes.getRange(r + 1, colDisp).setValue(round2(disp + cant));
      break;
    }

    // 2. Movimiento de entrada (rastro)
    const ahora = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd HH:mm:ss');
    sheetMovs.appendRow([nextMovId(), ahora, loteId, 'ENTRADA', cant,
      'CANCELACION', ordenId, usuario, `Devuelto a stock por cancelación de ítem (fila ${fila})`]);

    // 3. Borrar el costo BRUTO de ese lote en este pedido
    const dataCostos = sheetCostos.getDataRange().getValues();
    const chh = dataCostos[0].map(h => String(h).trim());
    const kTipoRef = chh.indexOf('TIPO_REFERENCIA'), kRefId = chh.indexOf('REFERENCIA_ID');
    const kTipoC = chh.indexOf('TIPO_COSTO'), kOrigen = chh.indexOf('ORIGEN');
    for (let r = 1; r < dataCostos.length; r++) {
      if (String(dataCostos[r][kTipoRef]).toUpperCase().trim() !== 'PEDIDO') continue;
      if (String(dataCostos[r][kRefId]).trim() !== ordenId) continue;
      if (String(dataCostos[r][kTipoC]).toUpperCase().trim() !== 'BRUTO') continue;
      if (String(dataCostos[r][kOrigen]).trim() !== `Stock ${loteId}`) continue;
      sheetCostos.deleteRow(r + 1);
      break;
    }

    // 4. Subtotal a 0
    if (cSub > 0) sheetItems.getRange(fila, cSub).setValue(0);

    devueltos.push({ fila, loteId, sku, talla, color, cantidad: cant });
  });

  if (devueltos.length > 0) {
    logChange(usuario, ordenId, 'STOCK_DEVUELTO',
      devueltos.map(d => `${d.loteId} (${d.sku} ${d.talla} ${d.color}) ×${d.cantidad}`).join(' · '));
  }
  return devueltos;
}

// ============ AYUDANTES DEL LOCK DE FACTURA ============

// Estados que NO requieren factura cargada
function requiereFactura(estado) {
  const libres = ['HACER PEDIDO', 'CANCELADO', 'ENTREGA PARCIAL'];
  return libres.indexOf(estado) === -1;
}

// Devuelve los SKUs de ítems PEDIDO (no stock) de un pedido que aún no tienen costo
function itemsPedidoSinCosto(ordenId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.items);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const cOrden  = headers.indexOf('ORDEN_ID');
  const cCosto  = headers.indexOf('COSTO_UNITARIO');
  const cOrigen = headers.indexOf('ORIGEN');
  const cSku    = headers.indexOf('SKU');
  const faltantes = [];
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][cOrden]).trim() !== String(ordenId).trim()) continue;
    const origen = cOrigen !== -1 ? String(data[r][cOrigen]).toUpperCase().trim() : '';
    if (origen === 'STOCK') continue; // el stock no requiere factura
    const costo = cCosto !== -1 ? Number(data[r][cCosto]) || 0 : 0;
    if (costo <= 0) faltantes.push(cSku !== -1 ? String(data[r][cSku]) : 'item');
  }
  return faltantes;
}

// ============ AUTO-AGREGAR INDUSTRIA ============
// Agrega una industria a TablaIndustrias si no existe (normalizada en MAYÚSCULAS)
function agregarIndustria(industria) {
  const norm = String(industria || '').trim().toUpperCase();
  if (!norm) return;
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('TablaIndustrias');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const cInd = headers.indexOf('INDUSTRIA');
  if (cInd === -1) return;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][cInd]).trim().toUpperCase() === norm) return; // ya existe
  }
  sheet.appendRow([norm, 'TRUE']);
}

// ============ CREAR CLIENTE (por separado) ============
function handleCrearCliente(body) {
  const c = body.cliente || {};
  const usuario = body.usuario || 'desconocido';
  const fecha = String(body.fecha || '').trim() ||
                Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');

  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede crear clientes', 403);
  if (!c.nombre || !String(c.nombre).trim()) return errorResponse('Falta el nombre del cliente');

  return withLock(function () {
    const ss = getSpreadsheet();
    const sheetClientes = ss.getSheetByName(TABS.clientes);

    const clienteId = nextClienteId();
    const nombre = String(c.nombre).trim();
    const cRow = sheetClientes.getLastRow() + 1;
    sheetClientes.appendRow([
      clienteId, nombre,
      "'" + String(c.telefono || ''),
      String(c.direccion || ''),
      String(c.ciudad    || ''),
      String(c.cedulaRuc || ''),
      String(c.industria || ''),
      "'" + String(c.email || ''),
      fecha, '', '', '', 0, '', ''
    ]);
    sheetClientes.getRange(cRow, 10).setFormula(`=COUNTIFS(TablaOrdenes!C:C,B${cRow},TablaOrdenes!D:D,"<>CANCELADO")`);
    sheetClientes.getRange(cRow, 11).setFormula(`=SUMIFS(TablaOrdenes!R:R,TablaOrdenes!C:C,B${cRow},TablaOrdenes!D:D,"<>CANCELADO")`);
    sheetClientes.getRange(cRow, 12).setFormula(`=IFERROR(MAXIFS(TablaOrdenes!B:B,TablaOrdenes!C:C,B${cRow}),"")`);
    sheetClientes.getRange(cRow, 14).setFormula(`=IF(J${cRow}>=3,TRUE,FALSE)`);

    // Auto-agregar la industria a TablaIndustrias
    if (c.industria) agregarIndustria(c.industria);

    logChange(usuario, clienteId, 'CLIENTE_CREADO',
      `${nombre}${c.industria ? ' · ' + c.industria : ''}`);

    return jsonResponse({
      ok: true,
      data: {
        CLIENTE_ID: clienteId, NOMBRE: nombre,
        TELEFONO: String(c.telefono || ''), DIRECCION: String(c.direccion || ''),
        CIUDAD: String(c.ciudad || ''), CEDULA_RUC: String(c.cedulaRuc || ''),
        INDUSTRIA: String(c.industria || ''), EMAIL: String(c.email || '')
      }
    });
  });
}

// ============ MIGRACIÓN: marcar COURIER_ITEM en pedidos que ya tienen courier ============
// Correr UNA sola vez desde el editor (botón Run).
function migrarCourierItems() {
  const ss = getSpreadsheet();
  const costos = readSheetAsObjects(TABS.costos);

  // Pedidos que ya tienen una fila COURIER
  const conCourier = new Set();
  costos.forEach(c => {
    if (String(c.TIPO_REFERENCIA).toUpperCase().trim() === 'PEDIDO' &&
        String(c.TIPO_COSTO).toUpperCase().trim() === 'COURIER') {
      conCourier.add(String(c.REFERENCIA_ID).trim());
    }
  });

  const sheet = ss.getSheetByName(TABS.items);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const cOrden   = headers.indexOf('ORDEN_ID');
  const cCourier = headers.indexOf('COURIER_ITEM');
  if (cCourier === -1) { Logger.log('❌ Falta la columna COURIER_ITEM en TablaItems'); return; }

  let marcados = 0;
  for (let r = 1; r < data.length; r++) {
    const ordenId = String(data[r][cOrden]).trim();
    if (conCourier.has(ordenId)) {
      sheet.getRange(r + 1, cCourier + 1).setValue('TRUE');
      marcados++;
    }
  }
  Logger.log(`✅ COURIER_ITEM = TRUE en ${marcados} ítem(s), de ${conCourier.size} pedido(s) con courier`);
}

// ============ ÍTEMS PENDIENTES DE COURIER (por ítem) ============
function getItemsPendientesCourier() {
  const ordenes = readSheetAsObjects(TABS.ordenes);
  const items   = readSheetAsObjects(TABS.items);

  // Estados en los que tiene sentido cargar courier (tramo del viaje en adelante)
  const ESTADOS_COURIER = [
    'EN TRANSITO A FL', 'EN BODEGA FL', 'EN CAMINO A EC', 'EN BODEGA EC', 'LISTO PARA ENVIAR'
  ];

  const infoOrden = {};
  ordenes.forEach(o => {
    infoOrden[String(o.ORDEN_ID).trim()] = {
      estado: String(o.ESTATUS_ENVIO).trim(),
      cliente: o.CLIENTE_NOMBRE,
      fecha: o.F_ORDEN
    };
  });

  const porPedido = {};
  items.forEach(it => {
    const ordenId = String(it.ORDEN_ID).trim();
    const info = infoOrden[ordenId];
    if (!info) return;
    if (info.estado === 'CANCELADO' || info.estado === 'ENTREGADO') return; // pedido cerrado

    const origen = String(it.ORIGEN || '').toUpperCase().trim();
    if (origen === 'STOCK') return;                                  // el stock va por lote

    if (String(it.COURIER_ITEM).toUpperCase().trim() === 'TRUE') return;  // ya tiene courier
    if (String(it.ENTREGADO_ITEM).toUpperCase().trim() === 'TRUE') return; // ítem ya entregado

    const costo = Number(it.COSTO_UNITARIO) || 0;
    if (costo <= 0) return;                                          // sin factura aún

    const estadoItem = String(it.ESTATUS_ITEM || '').trim() || info.estado;
    if (ESTADOS_COURIER.indexOf(estadoItem) === -1) return;          // fuera del tramo de courier

    if (!porPedido[ordenId]) {
      porPedido[ordenId] = {
        ORDEN_ID: ordenId, CLIENTE_NOMBRE: info.cliente,
        ESTATUS_ENVIO: info.estado, F_ORDEN: info.fecha, items: []
      };
    }
    porPedido[ordenId].items.push({
      _rowNum: it._rowNum, SKU: it.SKU, NOMBRE_PRODUCTO: it.NOMBRE_PRODUCTO,
      TALLA: it.TALLA, LONGITUD: it.LONGITUD, COLOR: it.COLOR,
      CANTIDAD: Number(it.CANTIDAD) || 1,
      ESTATUS_ITEM: estadoItem
    });
  });

  return Object.keys(porPedido).map(k => porPedido[k]).filter(p => p.items.length > 0);
}

// ============ CÁLCULO DE ETA POR COURIER (Capa 1) ============

function getCouriers() {
  const rows = readSheetAsObjects(TABS.couriers);
  return rows
    .filter(c => String(c.ACTIVA).toUpperCase().trim() === 'TRUE')
    .map(c => ({
      COURIER: String(c.COURIER).trim(),
      TIPO_ESTIMADO: String(c.TIPO_ESTIMADO).toUpperCase().trim(),
      DIAS_MIN: Number(c.DIAS_MIN) || 0,
      DIAS_MAX: Number(c.DIAS_MAX) || 0,
      DIA_SALIDA: String(c.DIA_SALIDA || '').trim(),
      DIA_ENTREGA: String(c.DIA_ENTREGA || '').trim(),
    }));
}

// 'yyyy-MM-dd' → Date local (a mediodía, para evitar líos de zona horaria)
function parseFechaLocal(str) {
  const p = String(str).trim().split('-');
  if (p.length !== 3) return null;
  const d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
  return isNaN(d.getTime()) ? null : d;
}

function fmtFecha(d) {
  const y  = d.getFullYear();
  const m  = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
}

// Suma N días hábiles (salta sábado y domingo)
function addBusinessDays(date, n) {
  const d = new Date(date.getTime());
  let added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) added++;
  }
  return d;
}

// Nombre de día (es) → número (Domingo=0 … Sábado=6), tolera acentos/mayúsculas
function dayNameToNum(name) {
  const n = String(name).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  const map = { domingo: 0, lunes: 1, martes: 2, miercoles: 3, jueves: 4, viernes: 5, sabado: 6 };
  return (n in map) ? map[n] : -1;
}

// Próxima fecha cuyo día de semana = target. strictlyAfter=true => empieza desde el día siguiente.
function nextDayOfWeek(date, target, strictlyAfter) {
  const d = new Date(date.getTime());
  if (strictlyAfter) d.setDate(d.getDate() + 1);
  while (d.getDay() !== target) d.setDate(d.getDate() + 1);
  return d;
}

// Devuelve el rango estimado de llegada para un courier dado, desde una fecha de salida
function calcularETA(courierNombre, fechaSalidaStr) {
  const fechaSalida = parseFechaLocal(fechaSalidaStr);
  if (!fechaSalida) return { ok: false, error: 'Fecha inválida (usa yyyy-MM-dd)' };

  const c = getCouriers().find(x => x.COURIER.toLowerCase() === String(courierNombre).toLowerCase().trim());
  if (!c) return { ok: false, error: 'Courier no encontrado o inactivo: ' + courierNombre };

  if (c.TIPO_ESTIMADO === 'DIAS') {
    const dMin = addBusinessDays(fechaSalida, c.DIAS_MIN);
    const dMax = addBusinessDays(fechaSalida, c.DIAS_MAX);
    return {
      ok: true, courier: c.COURIER, tipo: 'DIAS',
      fechaSalida: fmtFecha(fechaSalida),
      fechaMin: fmtFecha(dMin), fechaMax: fmtFecha(dMax),
      diasMin: c.DIAS_MIN, diasMax: c.DIAS_MAX
    };
  }

  if (c.TIPO_ESTIMADO === 'DIA_SEMANA') {
    const salidaDow  = dayNameToNum(c.DIA_SALIDA);
    const entregaDow = dayNameToNum(c.DIA_ENTREGA);
    if (salidaDow < 0 || entregaDow < 0) return { ok: false, error: 'Día de semana inválido en el courier' };
    const proxSalida = nextDayOfWeek(fechaSalida, salidaDow, false); // próxima salida (incluye el día mismo)
    const entrega    = nextDayOfWeek(proxSalida, entregaDow, true);  // entrega tras la salida
    return {
      ok: true, courier: c.COURIER, tipo: 'DIA_SEMANA',
      fechaSalida: fmtFecha(fechaSalida),
      proximaSalida: fmtFecha(proxSalida),
      fechaMin: fmtFecha(entrega), fechaMax: fmtFecha(entrega),
      diaSalida: c.DIA_SALIDA, diaEntrega: c.DIA_ENTREGA
    };
  }

  return { ok: false, error: 'TIPO_ESTIMADO desconocido: ' + c.TIPO_ESTIMADO };
}

function getLotesEnCamino() {
  const lotes     = readSheetAsObjects(TABS.lotes);
  const productos = readSheetAsObjects(TABS.productos);
  const prodMap = {};
  productos.forEach(p => { prodMap[String(p.SKU).trim()] = { nombre: p.NOMBRE, tipo: p.TIPO_PRENDA }; });

  const enCamino = ['EN TRANSITO A FL', 'EN BODEGA FL', 'EN CAMINO A EC'];

  return lotes
    .filter(l => {
      if ((Number(l.CANT_DISPONIBLE) || 0) <= 0) return false;
      const ev = String(l.ESTADO_VIAJE || '').toUpperCase().trim();
      return enCamino.indexOf(ev) !== -1;
    })
    .map(l => {
      const sku = String(l.SKU).trim();
      const info = prodMap[sku] || { nombre: sku, tipo: '' };
      const courier     = String(l.COURIER || '').trim();
      const fechaSalida = String(l.FECHA_SALIDA_EC || '').trim();
      const etaMinG = String(l.ETA_MIN || '').trim();
      const etaMaxG = String(l.ETA_MAX || '').trim();

      let eta = null;
      if (etaMinG || etaMaxG) {
        eta = { fechaMin: etaMinG, fechaMax: etaMaxG };
      } else if (courier && fechaSalida) {
        const r = calcularETA(courier, fechaSalida);
        if (r && r.ok) eta = { fechaMin: r.fechaMin, fechaMax: r.fechaMax };
      }

      return {
        LOTE_ID: l.LOTE_ID, SKU: l.SKU, NOMBRE_PRODUCTO: info.nombre, TIPO_PRENDA: info.tipo,
        TALLA: l.TALLA, LONGITUD: l.LONGITUD, COLOR: l.COLOR,
        CANT_DISPONIBLE: Number(l.CANT_DISPONIBLE) || 0,
        COSTO_UNITARIO: Number(l.COSTO_UNITARIO) || 0,
        ESTADO_VIAJE: String(l.ESTADO_VIAJE || '').trim(),
        TRACKING: String(l.TRACKING || '').trim(),
        TRANSPORTE: String(l.TRANSPORTE || '').trim(),
        COURIER: courier,
        FECHA_SALIDA_EC: fechaSalida,
        eta
      };
    })
    .sort((a, b) => String(a.ESTADO_VIAJE).localeCompare(String(b.ESTADO_VIAJE)));
}

// ============ AVANZAR LOTE EN SU VIAJE (Capa 3) ============
function avanzarLote(body) {
  const loteId   = String(body.loteId || '').trim();
  const accion   = String(body.accion || '').toUpperCase().trim();
  const courier  = String(body.courier || '').trim();
  const fechaSal = String(body.fechaSalida || '').trim();
  const usuario  = body.usuario || 'desconocido';

  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede mover lotes', 403);
  if (!loteId) return errorResponse('Falta loteId');
  if (['LLEGO_FL', 'DESPACHAR_EC', 'LLEGO_EC'].indexOf(accion) === -1) {
    return errorResponse('Acción inválida: ' + accion);
  }

  return withLock(function () {
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName(TABS.lotes);
    const data  = sheet.getDataRange().getValues();
    const h     = data[0].map(x => String(x).trim());

    const cId      = h.indexOf('LOTE_ID');
    const cEstado  = h.indexOf('ESTADO_VIAJE');
    const cCourier = h.indexOf('COURIER');
    const cFechaS  = h.indexOf('FECHA_SALIDA_EC');
    const cEtaMin  = h.indexOf('ETA_MIN');
    const cEtaMax  = h.indexOf('ETA_MAX');
    if ([cId, cEstado, cCourier, cFechaS, cEtaMin, cEtaMax].some(x => x === -1)) {
      throw new Error('Faltan columnas de viaje en TablaLotesStock');
    }

    let fila = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][cId]).trim() === loteId) { fila = r + 1; break; }
    }
    if (fila === -1) return errorResponse('Lote no encontrado: ' + loteId, 404);

    const estadoActual = String(sheet.getRange(fila, cEstado + 1).getValue()).trim().toUpperCase();
    let nuevoEstado = '';
    let eta = null;

    if (accion === 'LLEGO_FL') {
      nuevoEstado = 'EN BODEGA FL';
      sheet.getRange(fila, cEstado + 1).setValue(nuevoEstado);

    } else if (accion === 'DESPACHAR_EC') {
      if (!courier)  return errorResponse('Falta el courier para despachar');
      if (!fechaSal) return errorResponse('Falta la fecha de salida');
      const r = calcularETA(courier, fechaSal);
      if (!r || !r.ok) return errorResponse('No se pudo calcular ETA: ' + (r ? r.error : 'desconocido'));
      nuevoEstado = 'EN CAMINO A EC';
      sheet.getRange(fila, cEstado + 1).setValue(nuevoEstado);
      sheet.getRange(fila, cCourier + 1).setValue(courier);
      sheet.getRange(fila, cFechaS + 1).setValue(fechaSal);
      sheet.getRange(fila, cEtaMin + 1).setValue(r.fechaMin);
      sheet.getRange(fila, cEtaMax + 1).setValue(r.fechaMax);
      eta = { fechaMin: r.fechaMin, fechaMax: r.fechaMax };

    } else if (accion === 'LLEGO_EC') {
      nuevoEstado = 'DISPONIBLE';
      sheet.getRange(fila, cEstado + 1).setValue(nuevoEstado);
    }

    logChange(usuario, loteId, 'LOTE_VIAJE', `${estadoActual || '(vacío)'} → ${nuevoEstado}` +
      (accion === 'DESPACHAR_EC' ? ` · ${courier} · ETA ${eta.fechaMin}–${eta.fechaMax}` : ''));

    return jsonResponse({ ok: true, data: { loteId, estadoAnterior: estadoActual, nuevoEstado, eta } });
  });
}

function nextComboId() {
  const combos = readSheetAsObjects(TABS.combos);
  let max = 0;
  combos.forEach(c => {
    const m = String(c.COMBO_ID).match(/COMBO-(\d+)/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return 'COMBO-' + String(max + 1).padStart(5, '0');
}

// ============ COMBOS / CONJUNTOS (Capa 3) ============

function getCombos() {
  const combos = readSheetAsObjects(TABS.combos);
  const disp   = getStockDisponible();    // piezas disponibles
  const camino = getLotesEnCamino();      // piezas en camino (con eta + estado)

  // Mapa de TODOS los lotes "vivos" (disponibles + en camino) por LOTE_ID
  const map = {};
  disp.forEach(l => {
    map[String(l.LOTE_ID).trim()] = {
      LOTE_ID: l.LOTE_ID, NOMBRE_PRODUCTO: l.NOMBRE_PRODUCTO, TIPO_PRENDA: l.TIPO_PRENDA,
      TALLA: l.TALLA, LONGITUD: l.LONGITUD, COLOR: l.COLOR,
      COLOR_HEX: l.COLOR_HEX || '',
      enCamino: false, eta: null, ESTADO_VIAJE: ''
    };
  });
  const coloresMap = getColores();
  camino.forEach(l => {
    map[String(l.LOTE_ID).trim()] = {
      LOTE_ID: l.LOTE_ID, NOMBRE_PRODUCTO: l.NOMBRE_PRODUCTO, TIPO_PRENDA: l.TIPO_PRENDA,
      TALLA: l.TALLA, LONGITUD: l.LONGITUD, COLOR: l.COLOR,
      COLOR_HEX: coloresMap[String(l.COLOR || '').toUpperCase().trim()] || '',
      enCamino: true, eta: l.eta || null, ESTADO_VIAJE: String(l.ESTADO_VIAJE || '').trim()
    };
  });

  // Para elegir la ETA "más lejana" de las dos piezas
  function etaMax(a, b) {
    if (!a) return b;
    if (!b) return a;
    const fa = a.fechaMax || a.fechaMin || '';
    const fb = b.fechaMax || b.fechaMin || '';
    return fa >= fb ? a : b;
  }

  const out = [];
  combos.forEach(c => {
    const sup = String(c.LOTE_SUPERIOR).trim();
    const inf = String(c.LOTE_INFERIOR).trim();
    const ls = map[sup];
    const li = map[inf];
    // Si una pieza ya no está viva (se vendió), el combo no se muestra
    if (!ls || !li) return;

    const algunoEnCamino = ls.enCamino || li.enCamino;
    const etaCombo = algunoEnCamino ? etaMax(ls.eta, li.eta) : null;

    out.push({
      COMBO_ID: c.COMBO_ID,
      FECHA: c.FECHA,
      USUARIO: c.USUARIO,
      enCamino: algunoEnCamino,
      eta: etaCombo,
      superior: {
        LOTE_ID: ls.LOTE_ID, NOMBRE_PRODUCTO: ls.NOMBRE_PRODUCTO, TIPO_PRENDA: ls.TIPO_PRENDA,
        TALLA: ls.TALLA, LONGITUD: ls.LONGITUD, COLOR: ls.COLOR,
        COLOR_HEX: ls.COLOR_HEX || '',
        enCamino: ls.enCamino, ESTADO_VIAJE: ls.ESTADO_VIAJE, eta: ls.eta
      },
      inferior: {
        LOTE_ID: li.LOTE_ID, NOMBRE_PRODUCTO: li.NOMBRE_PRODUCTO, TIPO_PRENDA: li.TIPO_PRENDA,
        TALLA: li.TALLA, LONGITUD: li.LONGITUD, COLOR: li.COLOR,
        COLOR_HEX: li.COLOR_HEX || '',
        enCamino: li.enCamino, ESTADO_VIAJE: li.ESTADO_VIAJE, eta: li.eta
      }
    });
  });
  return out;
}

function crearCombo(body) {
  const loteSup = String(body.loteSuperior || '').trim();
  const loteInf = String(body.loteInferior || '').trim();
  const usuario = body.usuario || 'desconocido';
  const fecha = String(body.fecha || '').trim() ||
                Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');

  if (!loteSup || !loteInf) return errorResponse('Faltan los dos lotes del combo');
  if (loteSup === loteInf) return errorResponse('No puedes combinar un lote consigo mismo');

  return withLock(function () {
    const disp = getStockDisponible();
    const camino = getLotesEnCamino();
    const ids = {};
    disp.forEach(l => { ids[String(l.LOTE_ID).trim()] = true; });
    camino.forEach(l => { ids[String(l.LOTE_ID).trim()] = true; });
    if (!ids[loteSup]) return errorResponse('La pieza superior no está disponible ni en camino: ' + loteSup);
    if (!ids[loteInf]) return errorResponse('La pieza inferior no está disponible ni en camino: ' + loteInf);

    const combos = readSheetAsObjects(TABS.combos);
    const dup = combos.some(c =>
      (String(c.LOTE_SUPERIOR).trim() === loteSup && String(c.LOTE_INFERIOR).trim() === loteInf) ||
      (String(c.LOTE_SUPERIOR).trim() === loteInf && String(c.LOTE_INFERIOR).trim() === loteSup)
    );
    if (dup) return errorResponse('Ese conjunto ya existe', 409);

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TABS.combos);
    const comboId = nextComboId();
    sheet.appendRow([comboId, loteSup, loteInf, fecha, usuario, '']);
    logChange(usuario, comboId, 'COMBO_CREAR', `${loteSup} + ${loteInf}`);
    return jsonResponse({ ok: true, data: { comboId, loteSuperior: loteSup, loteInferior: loteInf } });
  });
}

function borrarCombo(body) {
  const comboId = String(body.comboId || '').trim();
  const usuario = body.usuario || 'desconocido';
  if (!comboId) return errorResponse('Falta comboId');

  return withLock(function () {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(TABS.combos);
    const data = sheet.getDataRange().getValues();
    const h = data[0].map(x => String(x).trim());
    const cId = h.indexOf('COMBO_ID');
    if (cId === -1) throw new Error('Falta columna COMBO_ID en TablaCombos');

    let fila = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][cId]).trim() === comboId) { fila = r + 1; break; }
    }
    if (fila === -1) return errorResponse('Combo no encontrado: ' + comboId, 404);

    sheet.deleteRow(fila);
    logChange(usuario, comboId, 'COMBO_BORRAR', '');
    return jsonResponse({ ok: true, data: { comboId, borrado: true } });
  });
}

// ============ FEDEX — AUTENTICACIÓN (Capa 5) ============

// Por ahora usamos TEST. Cuando pasemos a producción cambiamos a 'PROD'.
const FEDEX_ENV = 'PROD';
const FEDEX_BASE_TEST = 'https://apis-sandbox.fedex.com';
const FEDEX_BASE_PROD = 'https://apis.fedex.com';

function fedexBaseUrl() {
  return FEDEX_ENV === 'PROD' ? FEDEX_BASE_PROD : FEDEX_BASE_TEST;
}

function fedexGetToken() {
  const props = PropertiesService.getScriptProperties();
  const key    = FEDEX_ENV === 'PROD' ? props.getProperty('FEDEX_PROD_KEY')    : props.getProperty('FEDEX_TEST_KEY');
  const secret = FEDEX_ENV === 'PROD' ? props.getProperty('FEDEX_PROD_SECRET') : props.getProperty('FEDEX_TEST_SECRET');

  if (!key || !secret) {
    throw new Error('Faltan credenciales FEDEX en Script Properties (' + FEDEX_ENV + ')');
  }

  const resp = UrlFetchApp.fetch(fedexBaseUrl() + '/oauth/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'client_credentials',
      client_id: key,
      client_secret: secret
    },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) {
    throw new Error('FEDEX OAuth falló (' + code + '): ' + body);
  }

  const json = JSON.parse(body);
  return json.access_token;
}

// Función de prueba: confirma que las credenciales sirven
function probarFedexToken() {
  const token = fedexGetToken();
  return jsonResponse({
    ok: true,
    data: {
      mensaje: 'Token obtenido correctamente',
      env: FEDEX_ENV,
      tokenEmpieza: token.substring(0, 12) + '…',
      largo: token.length
    }
  });
}

// ============ FEDEX — RASTREAR UN ENVÍO (Capa 5) ============

function fedexRastrear(tracking) {
  tracking = String(tracking || '').trim();
  if (!tracking) return { ok: false, error: 'Falta el número de tracking' };

  const token = fedexGetToken();
  const resp = UrlFetchApp.fetch(fedexBaseUrl() + '/track/v1/trackingnumbers', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token,
      'X-locale': 'en_US'
    },
    payload: JSON.stringify({
      includeDetailedScans: false,
      trackingInfo: [{ trackingNumberInfo: { trackingNumber: tracking } }]
    }),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) {
    return { ok: false, error: 'FEDEX track falló (' + code + '): ' + body };
  }

  const json = JSON.parse(body);

  // Extraer el estado del primer resultado
  try {
    const result = json.output.completeTrackResults[0].trackResults[0];

    // ¿Hubo error en ese tracking? (ej. número inexistente)
    if (result.error) {
      return { ok: false, error: result.error.message || 'Tracking no encontrado', tracking: tracking };
    }

    const status = result.latestStatusDetail || {};
    return {
      ok: true,
      tracking: tracking,
      statusCode: status.code || '',          // ej. "OD", "DL", "IT"
      statusDesc: status.description || '',    // ej. "Delivered", "In transit"
      statusTexto: status.statusByLocale || '' // versión legible
    };
  } catch (err) {
    return { ok: false, error: 'No se pudo leer la respuesta de FEDEX', raw: body.substring(0, 300) };
  }
}

// Función de prueba: rastrea un tracking que le pases por URL
function probarFedexRastreo(body) {
  const tracking = String(body.tracking || '').trim();
  const r = fedexRastrear(tracking);
  return jsonResponse({ ok: true, data: r });
}

// ============ FEDEX — TRADUCCIÓN DE ESTADOS (Capa 5) ============

// Convierte el código/estado de FEDEX al estado de viaje del lote.
// Devuelve uno de: 'EN TRANSITO A FL', 'EN BODEGA FL', o '' (sin cambio aún).
function fedexEstadoALote(statusCode, statusDesc) {
  const code = String(statusCode || '').toUpperCase().trim();
  const desc = String(statusDesc || '').toUpperCase();

  // Entregado en bodega FL
  if (code === 'DL' || desc.indexOf('DELIVERED') !== -1) {
    return 'EN BODEGA FL';
  }

  // Aún no se ha movido (label creado / shipping soon) → no cambiamos nada
  if (code === 'OC' || desc.indexOf('LABEL') !== -1 || desc.indexOf('SHIPMENT INFORMATION') !== -1) {
    return '';
  }

  // Cualquier otro estado de movimiento (recogido, en tránsito, en reparto)
  if (['PU', 'IT', 'IN', 'OD', 'AR', 'DP', 'AF'].indexOf(code) !== -1) {
    return 'EN TRANSITO A FL';
  }

  // Si llega un código que no conocemos, no tocamos el lote (lo dejamos como está)
  return '';
}

// ============ FEDEX — MOTOR DE ACTUALIZACIÓN DE LOTES (Capa 5) ============

// Recorre los lotes que están viajando hacia FL (con tracking) y actualiza
// su estado según lo que diga FEDEX. Solo toca el tramo FIGS → FL.
function actualizarLotesPorTracking() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.lotes);
  const data  = sheet.getDataRange().getValues();
  const h     = data[0].map(x => String(x).trim());

  const cId       = h.indexOf('LOTE_ID');
  const cEstado   = h.indexOf('ESTADO_VIAJE');
  const cTracking = h.indexOf('TRACKING');
  const cTransp   = h.indexOf('TRANSPORTE');
  if ([cId, cEstado, cTracking].some(x => x === -1)) {
    throw new Error('Faltan columnas en TablaLotesStock (ESTADO_VIAJE / TRACKING)');
  }

  // Estados donde todavía esperamos noticias de FEDEX (tramo hacia FL)
  const enViajeAFL = ['EN TRANSITO A FL', 'EN BODEGA FL'];

  const cambios = [];
  const sinTracking = [];
  const errores = [];

  for (let r = 1; r < data.length; r++) {
    const loteId   = String(data[r][cId]).trim();
    const estado   = String(data[r][cEstado]).trim().toUpperCase();
    const tracking = String(data[r][cTracking]).trim();
    const transp   = cTransp !== -1 ? String(data[r][cTransp]).trim().toUpperCase() : '';

    if (!loteId) continue;
    // Solo nos interesan los que están en el tramo hacia FL
    if (enViajeAFL.indexOf(estado) === -1) continue;
    // Por ahora solo FEDEX (USPS lo agregamos después si quieres)
    if (transp && transp !== 'FEDEX') continue;

    if (!tracking) { sinTracking.push(loteId); continue; }

    const res = fedexRastrear(tracking);
    if (!res.ok) { errores.push({ loteId, tracking, error: res.error }); continue; }

    const nuevoEstado = fedexEstadoALote(res.statusCode, res.statusDesc);
    if (!nuevoEstado) continue;                 // FEDEX no da motivo de cambio
    if (nuevoEstado === estado) continue;       // ya está en ese estado

    // Aplicar el cambio
    sheet.getRange(r + 1, cEstado + 1).setValue(nuevoEstado);
    logChange('auto-fedex', loteId, 'LOTE_VIAJE',
      `${estado} → ${nuevoEstado} (FEDEX: ${res.statusDesc || res.statusCode})`);
    cambios.push({ loteId, tracking, de: estado, a: nuevoEstado, fedex: res.statusDesc || res.statusCode });

    Utilities.sleep(300); // respiro entre llamadas para no saturar
  }

  return { revisados: cambios.length + errores.length, cambios, sinTracking, errores };
}

// Función para probar/correr manualmente y ver el resultado por URL
function probarActualizarTracking() {
  const r = actualizarLotesPorTracking();
  return jsonResponse({ ok: true, data: r });
}

function normalizarMayusculasItems() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.items);
  const data = sheet.getDataRange().getValues();
  const h = data[0].map(x => String(x).trim());

  const cTalla = h.indexOf('TALLA');
  const cColor = h.indexOf('COLOR');
  const cLong  = h.indexOf('LONGITUD');
  if ([cTalla, cColor].some(x => x === -1)) throw new Error('Faltan columnas TALLA/COLOR');

  let cambios = 0;
  for (let r = 1; r < data.length; r++) {
    [cTalla, cColor, cLong].forEach(c => {
      if (c === -1) return;
      const val = String(data[r][c] || '');
      const up = val.toUpperCase();
      if (val !== up && up.trim() !== '') {
        sheet.getRange(r + 1, c + 1).setValue(up);
        cambios++;
      }
    });
  }
  Logger.log('Celdas normalizadas: ' + cambios);
  return cambios;
}

// Cuenta cuántos ítems cubre una fila de courier según su descripción
function _courierItemsDeDescripcion(desc) {
  const d = String(desc || '');
  if (/item fila/i.test(d)) return 1;
  let m = d.match(/retroactivo\s*\((\d+)\s*de/i);
  if (m) return parseInt(m[1], 10);
  m = d.match(/\((\d+)\s*items?\)/i);
  if (m) return parseInt(m[1], 10);
  return 1;
}

// Backfill de una sola vez: pedidos con courier completo → EN BODEGA EC
function backfillCourierCompleto(body) {
  const usuario = (body && body.usuario) || 'backfill';
  if (!verifyAdmin(usuario)) return errorResponse('Solo admin', 403);

  const ss = getSpreadsheet();
  const sheetItems   = ss.getSheetByName(TABS.items);
  const sheetOrdenes = ss.getSheetByName(TABS.ordenes);
  const costos = readSheetAsObjects(TABS.costos);

  // Courier acumulado por pedido (en # de ítems cubiertos)
  const courierPorPedido = {};
  costos.forEach(c => {
    if (String(c.TIPO_REFERENCIA).toUpperCase().trim() !== 'PEDIDO') return;
    if (String(c.TIPO_COSTO).toUpperCase().trim() !== 'COURIER') return;
    const id = String(c.REFERENCIA_ID).trim();
    courierPorPedido[id] = (courierPorPedido[id] || 0) + _courierItemsDeDescripcion(c.DESCRIPCION);
  });

  const itemsData = sheetItems.getDataRange().getValues();
  const ih = itemsData[0].map(h => String(h).trim());
  const cOrden       = ih.indexOf('ORDEN_ID');
  const cEstItem     = ih.indexOf('ESTATUS_ITEM');
  const cCourierItem = ih.indexOf('COURIER_ITEM');

  const filasPorPedido = {};
  const countPorPedido = {};
  for (let r = 1; r < itemsData.length; r++) {
    const id = String(itemsData[r][cOrden]).trim();
    if (!id) continue;
    if (!filasPorPedido[id]) { filasPorPedido[id] = []; countPorPedido[id] = 0; }
    filasPorPedido[id].push(r + 1);
    countPorPedido[id]++;
  }

  const ITEM_PROTEGIDO   = ['ENTREGADO', 'CANCELADO'];
  const PEDIDO_PROTEGIDO = ['ENTREGA PARCIAL', 'ENTREGADO', 'CANCELADO', 'EN BODEGA EC', 'LISTO PARA ENVIAR'];

  const marcados = [];
  const avanzados = [];

  Object.keys(courierPorPedido).forEach(ordenId => {
    const itemCount = countPorPedido[ordenId] || 0;
    if (itemCount === 0) return;
    if (courierPorPedido[ordenId] < itemCount) return;   // courier incompleto

    filasPorPedido[ordenId].forEach(sheetRow => {
      const idx = sheetRow - 1;
      if (cCourierItem >= 0) {
        sheetItems.getRange(sheetRow, cCourierItem + 1).setValue('TRUE');
        itemsData[idx][cCourierItem] = 'TRUE';
      }
      if (cEstItem >= 0) {
        const cur = String(itemsData[idx][cEstItem]).toUpperCase().trim();
        if (ITEM_PROTEGIDO.indexOf(cur) === -1 && cur !== 'EN BODEGA EC') {
          sheetItems.getRange(sheetRow, cEstItem + 1).setValue('EN BODEGA EC');
          itemsData[idx][cEstItem] = 'EN BODEGA EC';
        }
      }
    });
    marcados.push(ordenId);

    const orderRow = findOrderRow(ordenId);
    if (orderRow === -1) return;
    const estadoActual = String(sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).getValue()).trim().toUpperCase();
    if (PEDIDO_PROTEGIDO.indexOf(estadoActual) !== -1) return;
    sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).setValue('EN BODEGA EC');
    logChange(usuario, ordenId, 'CAMBIO_ESTADO', `${estadoActual} → EN BODEGA EC (backfill courier)`);
    avanzados.push(ordenId);
  });

  return jsonResponse({ ok: true, data: { marcados, avanzados } });
}

// ============ CAMBIO DE ÍTEM POR ERROR (devolución → stock + ítem nuevo) ============
function cambiarItemPorError(body) {
  const ordenId  = String(body.ordenId || '').trim();
  const itemRowX = Number(body.itemRowX);
  const usuario  = body.usuario || 'desconocido';

  // Lo que LLEGÓ mal (va a stock)
  const skuStock   = String(body.skuStock || '').trim();
  const tallaStock = String(body.tallaStock || '').trim().toUpperCase();
  const longStock  = String(body.longitudStock || 'Regular').trim();
  const colorStock = String(body.colorStock || '').trim().toUpperCase();
  const costoStockIn = (body.costoStock !== undefined && body.costoStock !== '') ? Number(body.costoStock) : null;

  // Lo que se RE-PIDE (lo que la clienta quería). Default: igual a X.
  const skuY   = String(body.skuY || '').trim();
  const tallaY = String(body.tallaY || '').trim().toUpperCase();
  const longY  = String(body.longitudY || '').trim();
  const colorY = String(body.colorY || '').trim().toUpperCase();

  const quitarCosto = body.quitarCosto === undefined ? true
    : (body.quitarCosto === true || String(body.quitarCosto).toLowerCase() === 'true');

  if (!canCambiarEstado(usuario)) return errorResponse('No tienes permiso', 403);
  if (!ordenId) return errorResponse('Falta ordenId');
  if (!itemRowX || itemRowX < 2) return errorResponse('Falta la fila del ítem del pedido');
  if (!skuStock || !tallaStock || !colorStock) return errorResponse('Faltan datos de lo que llegó (SKU, talla, color)');

  return withLock(function () {
    const ss = getSpreadsheet();
    const sheetItems   = ss.getSheetByName(TABS.items);
    const sheetLotes   = ss.getSheetByName(TABS.lotes);
    const sheetMovs    = ss.getSheetByName(TABS.movimientos);
    const sheetCostos  = ss.getSheetByName(TABS.costos);
    const sheetOrdenes = ss.getSheetByName(TABS.ordenes);

    const ih = sheetItems.getRange(1, 1, 1, sheetItems.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const cOrden = ih.indexOf('ORDEN_ID'), cSku = ih.indexOf('SKU'), cTalla = ih.indexOf('TALLA');
    const cLong = ih.indexOf('LONGITUD'), cColor = ih.indexOf('COLOR'), cCant = ih.indexOf('CANTIDAD');
    const cPrecio = ih.indexOf('PRECIO_VENTA'), cParte = ih.indexOf('PARTE_DE_SET');
    const cCosto = ih.indexOf('COSTO_UNITARIO'), cOrigen = ih.indexOf('ORIGEN');
    const cEstItem = ih.indexOf('ESTATUS_ITEM'), cEntrItem = ih.indexOf('ENTREGADO_ITEM'), cCourIt = ih.indexOf('COURIER_ITEM');

    const filaX = sheetItems.getRange(itemRowX, 1, 1, sheetItems.getLastColumn()).getValues()[0];
    if (String(filaX[cOrden]).trim() !== ordenId) return errorResponse(`La fila ${itemRowX} no pertenece al pedido ${ordenId}`, 400);
    const estX = String(filaX[cEstItem] || '').toUpperCase().trim();
    if (estX === 'CANCELADO') return errorResponse('Ese ítem ya está cancelado', 409);
    if (String(filaX[cEntrItem]).toUpperCase().trim() === 'TRUE') return errorResponse('Ese ítem ya fue entregado a la clienta', 409);

    const skuX = String(filaX[cSku]).trim(), tallaX = String(filaX[cTalla]).trim();
    const longX = String(filaX[cLong]).trim(), colorX = String(filaX[cColor]).trim();
    const precioX = Number(filaX[cPrecio]) || 0;
    const parteX = cParte >= 0 ? String(filaX[cParte] || '') : '';
    const costoX = Number(filaX[cCosto]) || 0;
    const costoStock = (costoStockIn !== null) ? costoStockIn : costoX;

    const yS = skuY || skuX, yT = tallaY || tallaX, yL = longY || longX, yC = colorY || colorX;

    // PASO 1: lo que llegó → lote de stock DISPONIBLE
    const loteId = nextLoteId();
    const ahora = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd HH:mm:ss');
    const hoy   = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
    sheetLotes.appendRow([
      loteId, skuStock, tallaStock, longStock, colorStock, round2(costoStock),
      1, 1, 'DEVOLUCION', ordenId, hoy, usuario,
      `Cambio por error en ${ordenId} (llegó ${skuStock} en vez de ${skuX})`,
      'DISPONIBLE', '', '', '', '', '', ''
    ]);
    sheetMovs.appendRow([nextMovId(), ahora, loteId, 'ENTRADA', 1, 'DEVOLUCION', ordenId, usuario, `Cambio por error en ${ordenId}`]);

    // PASO 2: cancelar X (trazabilidad)
    sheetItems.getRange(itemRowX, cEstItem + 1).setValue('CANCELADO');
    sheetItems.getRange(itemRowX, cEntrItem + 1).setValue('FALSE');
    const cSub = ih.indexOf('SUBTOTAL');
    if (cSub >= 0) sheetItems.getRange(itemRowX, cSub + 1).setValue(0);

    // PASO 3: crear Y (re-pedido), mismo precio, sin costo, a pedir
    const iRow = sheetItems.getLastRow() + 1;
    sheetItems.appendRow([ordenId, yS, '', '', yT, yL, yC, 1, precioX, parteX, '', '']);
    sheetItems.getRange(iRow, 3).setFormula(`=IFERROR(VLOOKUP(B${iRow},TablaProductos!A:B,2,FALSE),"")`);
    sheetItems.getRange(iRow, 4).setFormula(`=IFERROR(VLOOKUP(B${iRow},TablaProductos!A:C,3,FALSE),"")`);
    sheetItems.getRange(iRow, 12).setFormula(`=H${iRow}*I${iRow}`);
    if (cOrigen   >= 0) sheetItems.getRange(iRow, cOrigen + 1).setValue('PEDIDO');
    if (cEstItem  >= 0) sheetItems.getRange(iRow, cEstItem + 1).setValue('HACER PEDIDO');
    if (cEntrItem >= 0) sheetItems.getRange(iRow, cEntrItem + 1).setValue('FALSE');
    if (cCourIt   >= 0) sheetItems.getRange(iRow, cCourIt + 1).setValue('');

    // PASO 4: quitar el costo BRUTO de X del pedido (su dinero se fue al stock)
    let costoQuitado = 0;
    if (quitarCosto) {
      const data = sheetCostos.getDataRange().getValues();
      const chh = data[0].map(h => String(h).trim());
      const kTipoRef = chh.indexOf('TIPO_REFERENCIA'), kRefId = chh.indexOf('REFERENCIA_ID');
      const kTipoC = chh.indexOf('TIPO_COSTO'), kDesc = chh.indexOf('DESCRIPCION'), kMonto = chh.indexOf('MONTO');
      const skuXUp = skuX.toUpperCase(), colorXUp = colorX.toUpperCase();
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][kTipoRef]).toUpperCase().trim() !== 'PEDIDO') continue;
        if (String(data[r][kRefId]).trim() !== ordenId) continue;
        if (String(data[r][kTipoC]).toUpperCase().trim() !== 'BRUTO') continue;
        const d = String(data[r][kDesc]).toUpperCase();
        if (d.indexOf(skuXUp) === 0 && (colorXUp ? d.indexOf(colorXUp) !== -1 : true)) {
          costoQuitado = Number(data[r][kMonto]) || 0;
          sheetCostos.deleteRow(r + 1);
          break;
        }
      }
    }

    // PASO 5: recalcular cabecera (ignorando CANCELADO)
    SpreadsheetApp.flush();
    const allItems = sheetItems.getDataRange().getValues();
    let algunoEntregado = false, todosEntregados = true, minIdx = 999, hayVivos = false;
    for (let r = 1; r < allItems.length; r++) {
      if (String(allItems[r][cOrden]).trim() !== ordenId) continue;
      const est = String(allItems[r][cEstItem]).toUpperCase().trim();
      if (est === 'CANCELADO') continue;
      hayVivos = true;
      const ent = String(allItems[r][cEntrItem]).toUpperCase().trim() === 'TRUE';
      if (ent) algunoEntregado = true; else todosEntregados = false;
      const idx = VALID_STATES.indexOf(est);
      if (idx !== -1 && idx < minIdx) minIdx = idx;
    }
    const orderRow = findOrderRow(ordenId);
    if (orderRow !== -1 && hayVivos) {
      let header;
      if (todosEntregados) header = 'ENTREGADO';
      else if (algunoEntregado) header = 'ENTREGA PARCIAL';
      else header = (minIdx >= 0 && minIdx < 999) ? VALID_STATES[minIdx] : null;
      if (header) sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).setValue(header);
    }

    logChange(usuario, ordenId, 'CAMBIO_ITEM_ERROR',
      `Llegó ${skuStock} ${tallaStock} ${colorStock} → stock ${loteId} ($${round2(costoStock)}) · ` +
      `Cancelado ${skuX} ${tallaX} ${colorX} · Re-pide ${yS} ${yT} ${yC}` +
      (quitarCosto ? ` · costo quitado $${costoQuitado.toFixed(2)}` : ''));

    return jsonResponse({
      ok: true,
      data: {
        ordenId,
        loteStock: { loteId, sku: skuStock, talla: tallaStock, color: colorStock, costo: round2(costoStock) },
        cancelado: { sku: skuX, talla: tallaX, color: colorX, fila: itemRowX },
        rePedido:  { sku: yS, talla: yT, longitud: yL, color: yC, precio: precioX, fila: iRow },
        costoQuitado: round2(costoQuitado)
      }
    });
  });
}

// ============ REPORTE DE GANANCIA (todos los pedidos con factura, costos reales + asumidos) ============
function getReporteGanancia(fechaInicio, fechaFin) {
  const ordenes = readSheetAsObjects(TABS.ordenes);
  const items   = readSheetAsObjects(TABS.items);
  const costos  = readSheetAsObjects(TABS.costos);
  return calcularReporteGanancia(ordenes, items, costos, fechaInicio, fechaFin);
}

// Misma lógica que antes, pero recibe las tablas YA LEÍDAS en vez de releerlas.
// Así getTimelineMensual puede leer una sola vez y calcular 6 meses en memoria.
function calcularReporteGanancia(ordenes, items, costos, fechaInicio, fechaFin) {
  // Montos asumidos para costos que falten
  const COURIER_ASUMIDO  = 7.50;  // por pieza sin courier cargado
  const EMPAQUE_ASUMIDO  = 3.04;  // por pedido sin empaque
  const PIN_ASUMIDO      = 1.10;  // por pedido sin pin
  const DELIVERY_ASUMIDO = 6.10;  // por pedido sin delivery

  const ini = fechaInicio ? String(fechaInicio).trim() : '';
  const fin = fechaFin ? String(fechaFin).trim() : '';

  function toYMD(v) {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, 'GMT-5', 'yyyy-MM-dd');
    return String(v).trim().slice(0, 10);
  }

  // Fecha de pedido + nombre por orden
  const infoPedido = {};
  ordenes.forEach(o => {
    const id = String(o.ORDEN_ID).trim();
    if (!id) return;
    infoPedido[id] = {
      fOrden: toYMD(o.F_ORDEN),
      nombre: String(o.CLIENTE_NOMBRE || '').trim(),
      notas: String(o.NOTAS || '').trim()
    };
  });

  // Ítems no cancelados por pedido
  const itemsPorPedido = {};
  items.forEach(it => {
    if (String(it.ESTATUS_ITEM || '').toUpperCase().trim() === 'CANCELADO') return;
    const id = String(it.ORDEN_ID).trim();
    if (!id) return;
    if (!itemsPorPedido[id]) itemsPorPedido[id] = [];
    itemsPorPedido[id].push(it);
  });

  // Agrupar costos por pedido y tipo
  const cx = {}; // cx[ordenId] = { bruto, courierReal, nCourier, tieneEmpaque, tienePin, tieneDelivery, brutos:[] }
  costos.forEach(c => {
    if (String(c.TIPO_REFERENCIA).toUpperCase().trim() !== 'PEDIDO') return;
    const id = String(c.REFERENCIA_ID).trim();
    if (!cx[id]) cx[id] = { bruto: 0, courierReal: 0, nCourier: 0, tieneEmpaque: false, tienePin: false, tieneDelivery: false, brutos: [] };
    const tipo  = String(c.TIPO_COSTO).toUpperCase().trim();
    const monto = Number(c.MONTO) || 0;
    if (tipo === 'BRUTO') {
      cx[id].bruto += monto;
      cx[id].brutos.push(String(c.ORIGEN || '').toUpperCase());
    } else if (tipo === 'COURIER') {
      cx[id].courierReal += monto;
      cx[id].nCourier += 1; // cada fila de courier = 1 pieza cubierta
    } else if (tipo === 'EMPAQUE') {
      cx[id].tieneEmpaque = true;
    } else if (tipo === 'REGALO') {
      cx[id].tienePin = true;
    } else if (tipo === 'DELIVERY') {
      cx[id].tieneDelivery = true;
    }
  });

  let stock  = { ventas: 0, venta: 0, costo: 0 };
  let pedido = { ventas: 0, venta: 0, costo: 0 };
  const detalle = [];

  Object.keys(itemsPorPedido).forEach(ordenId => {
    const c = cx[ordenId];
    // Excluir pedidos sin factura (sin BRUTO)
    if (!c || c.bruto <= 0) return;

    const info = infoPedido[ordenId] || { fOrden: '', nombre: '' };

    // Filtro por rango de fecha de pedido
    if (ini && (!info.fOrden || info.fOrden < ini)) return;
    if (fin && (!info.fOrden || info.fOrden > fin)) return;
    // Excluir pedidos de marketing/sorteo (precio simbolico, no son ventas reales)
    const notas = String((infoPedido[ordenId] && infoPedido[ordenId].notas) || '').toUpperCase();
    if (notas.indexOf('MARKETING') !== -1 || notas.indexOf('SORTEO') !== -1) return;

    const piezas = itemsPorPedido[ordenId];
    const nPiezas = piezas.length;
    const venta = piezas.reduce((s, it) =>
      s + (Number(it.CANTIDAD) || 0) * (Number(it.PRECIO_VENTA) || 0), 0);

    // Costo = real + asumido donde falte
    let costo = c.bruto;                 // BRUTO siempre real
    costo += c.courierReal;              // courier real cargado
    const piezasSinCourier = Math.max(0, nPiezas - c.nCourier);
    costo += piezasSinCourier * COURIER_ASUMIDO;   // courier asumido por pieza faltante
    if (!c.tieneEmpaque)  costo += EMPAQUE_ASUMIDO;
    if (!c.tienePin)      costo += PIN_ASUMIDO;
    if (!c.tieneDelivery) costo += DELIVERY_ASUMIDO;

    // Clasificar pedido: STOCK si alguna pieza vino de stock; si no, PEDIDO
    const tieneStock = c.brutos.some(og => og.indexOf('STOCK') !== -1);
    const tipo = tieneStock ? 'STOCK' : 'PEDIDO';

    const ganancia = round2(venta - costo);
    if (tipo === 'STOCK') {
      stock.ventas++; stock.venta += venta; stock.costo += costo;
    } else {
      pedido.ventas++; pedido.venta += venta; pedido.costo += costo;
    }

    detalle.push({
      ORDEN_ID: ordenId,
      NOMBRE: info.nombre,
      F_ORDEN: info.fOrden,
      TIPO: tipo,
      venta: round2(venta),
      costo: round2(costo),
      ganancia: ganancia
    });
  });

  // Ordenar detalle por ID
  detalle.sort((a, b) => a.ORDEN_ID.localeCompare(b.ORDEN_ID));

  return {
    stock:  { ventas: stock.ventas,  venta: round2(stock.venta),  costo: round2(stock.costo),  ganancia: round2(stock.venta - stock.costo) },
    pedido: { ventas: pedido.ventas, venta: round2(pedido.venta), costo: round2(pedido.costo), ganancia: round2(pedido.venta - pedido.costo) },
    detalle: detalle
  };
}

// ============ MATERIALES DE EMPAQUE Y REGALOS ============

function getMateriales() {
  const empaque = readSheetAsObjects(TABS.empaque);
  const regalos = readSheetAsObjects(TABS.regalos);

  const materialesEmpaque = empaque
    .filter(m => {
      const activo = String(m.ACTIVO || '').toUpperCase().trim();
      return activo === 'TRUE' || activo === 'VERDADERO';
    })
    .map(m => ({
      ID:           String(m.MATERIAL_ID).trim(),
      NOMBRE:       String(m.NOMBRE || '').trim(),
      CATEGORIA:    'EMPAQUE',
      STOCK:        Number(m.STOCK) || 0,
      STOCK_MINIMO: Number(m.STOCK_MINIMO) || 0,
      COSTO:        Number(m.COSTO_UNITARIO) || 0,
    }));

  const materialesRegalos = regalos
    .filter(r => String(r.REGALO_ID || '').trim() !== '')
    .map(r => ({
      ID:           String(r.REGALO_ID).trim(),
      NOMBRE:       String(r.NOMBRE || '').trim(),
      CATEGORIA:    String(r.INDUSTRIA_SUGERIDA || 'General').trim(),
      STOCK:        Number(r.STOCK) || 0,
      STOCK_MINIMO: Number(r.STOCK_MINIMO) || 0,
      COSTO:        Number(r.COSTO_UNITARIO) || 0,
    }));

  return { empaque: materialesEmpaque, regalos: materialesRegalos };
}

function setStockMaterial(body) {
  const tabla    = String(body.tabla || '').trim();   // 'empaque' | 'regalos'
  const id       = String(body.id || '').trim();
  const cantidad = Number(body.cantidad);
  const usuario  = body.usuario || 'desconocido';

  if (!id) return errorResponse('Falta id');
  if (isNaN(cantidad) || cantidad < 0) return errorResponse('Cantidad inválida');
  if (tabla !== 'empaque' && tabla !== 'regalos') return errorResponse('Tabla inválida');

  const user = readSheetAsObjects(TABS.usuarios).find(u =>
    String(u.USUARIO).toLowerCase().trim() === String(usuario).toLowerCase().trim()
  );
  if (!user) return errorResponse('Usuario no encontrado', 403);
  const rol = String(user.ROL).toLowerCase().trim();
  if (rol !== 'admin' && rol !== 'bodega') return errorResponse('Sin permiso', 403);

  const ss = getSpreadsheet();
  const tabName = tabla === 'empaque' ? TABS.empaque : TABS.regalos;
  const colId   = tabla === 'empaque' ? 'MATERIAL_ID' : 'REGALO_ID';
  const sheet   = ss.getSheetByName(tabName);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const cId    = headers.indexOf(colId);
  const cStock = headers.indexOf('STOCK');
  if (cId === -1)    return errorResponse(`Columna ${colId} no encontrada en ${tabName}`);
  if (cStock === -1) return errorResponse(`Columna STOCK no encontrada en ${tabName}`);

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][cId]).trim() === id) {
      sheet.getRange(r + 1, cStock + 1).setValue(cantidad);
      logChange(usuario, id, 'STOCK_MATERIAL',
        `${tabName} · ${id} · stock → ${cantidad}`);
      return jsonResponse({ ok: true, data: { id, tabla, cantidad } });
    }
  }

  return errorResponse(`${id} no encontrado en ${tabName}`, 404);
}


function agregarPin(body) {
  const nombre   = String(body.nombre || '').trim();
  const categoria = String(body.categoria || 'General').trim();
  const cantidad = Number(body.cantidad) || 0;
  const costo    = Number(body.costo) || 1.1;
  const usuario  = body.usuario || 'desconocido';

  if (!nombre) return errorResponse('Falta el nombre del pin');
  if (cantidad < 0) return errorResponse('Cantidad inválida');

  const user = readSheetAsObjects(TABS.usuarios).find(u =>
    String(u.USUARIO).toLowerCase().trim() === String(usuario).toLowerCase().trim()
  );
  if (!user) return errorResponse('Usuario no encontrado', 403);
  const rol = String(user.ROL).toLowerCase().trim();
  if (rol !== 'admin' && rol !== 'bodega') return errorResponse('Sin permiso', 403);

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.regalos);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const cId = headers.indexOf('REGALO_ID');

  // Generar siguiente ID
  let maxNum = 0;
  for (let r = 1; r < data.length; r++) {
    const m = String(data[r][cId]).match(/PIN-(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  }
  const nuevoId = 'PIN-' + String(maxNum + 1).padStart(3, '0');

  sheet.appendRow([nuevoId, nombre, categoria, cantidad, 1, costo]);
  logChange(usuario, nuevoId, 'PIN_CREADO', `${nombre} · ${categoria} · stock ${cantidad}`);

  return jsonResponse({ ok: true, data: { id: nuevoId, nombre, categoria, cantidad, costo } });
}

function actualizarCliente(body) {
  const clienteId = String(body.clienteId || '').trim();
  const campos    = body.campos || {};
  const usuario   = body.usuario || 'desconocido';

  if (!clienteId) return errorResponse('Falta clienteId');
  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede actualizar clientes', 403);

  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.clientes);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());

  const cId = headers.indexOf('CLIENTE_ID');
  if (cId === -1) return errorResponse('Columna CLIENTE_ID no encontrada');

  let fila = -1;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][cId]).trim() === clienteId) { fila = r + 1; break; }
  }
  if (fila === -1) return errorResponse('Cliente no encontrado: ' + clienteId, 404);

  const mapaCampos = {
    direccion: 'DIRECCION',
    ciudad:    'CIUDAD',
    telefono:  'TELEFONO',
    email:     'EMAIL',
    industria: 'INDUSTRIA',
    cedulaRuc: 'CEDULA_RUC',
  };

  const cambios = [];
  Object.keys(campos).forEach(campo => {
    const col = mapaCampos[campo];
    if (!col) return;
    const idx = headers.indexOf(col);
    if (idx === -1) return;
    sheet.getRange(fila, idx + 1).setValue(String(campos[campo] || ''));
    cambios.push(`${col}=${campos[campo]}`);
  });

  logChange(usuario, clienteId, 'CLIENTE_ACTUALIZADO', cambios.join(' · '));
  return jsonResponse({ ok: true, data: { clienteId, cambios } });
}

function agregarDescuento(body) {
  const ordenId = String(body.ordenId || '').trim();
  const monto   = Number(body.monto);
  const nota    = String(body.nota || '').trim();
  const usuario = body.usuario || 'desconocido';

  if (!ordenId) return errorResponse('Falta ordenId');
  if (isNaN(monto) || monto <= 0) return errorResponse('El descuento debe ser mayor a 0');
  if (!canOperar(usuario)) return errorResponse('No tienes permiso para aplicar descuentos', 403);
  if (findOrderRow(ordenId) === -1) return errorResponse('Pedido no encontrado: ' + ordenId, 404);

  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.descuentos);
  const fecha = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');

  sheet.appendRow([
    nextDescuentoId(), ordenId, fecha,
    'Descuento último minuto', -Math.abs(monto),
    nota || 'Descuento aplicado al cobrar'
  ]);

  logChange(usuario, ordenId, 'DESCUENTO_APLICADO',
    `$${monto.toFixed(2)}${nota ? ' · ' + nota : ''}`);

  return jsonResponse({ ok: true, data: { ordenId, monto: round2(monto), nota } });
}

function perdidaEnTransito(body) {
  const ordenId = String(body.ordenId || '').trim();
  const items   = body.items || []; // [{ itemRowNum, figsRepone, sku, talla, longitud, color }]
  const usuario = body.usuario || 'desconocido';

  if (!canOperar(usuario)) return errorResponse('No tienes permiso para registrar pérdidas', 403);
  if (!ordenId) return errorResponse('Falta ordenId');
  if (items.length === 0) return errorResponse('No se seleccionaron ítems perdidos');

  return withLock(function () {
    const ss = getSpreadsheet();
    const sheetItems   = ss.getSheetByName(TABS.items);
    const sheetCostos  = ss.getSheetByName(TABS.costos);
    const sheetOrdenes = ss.getSheetByName(TABS.ordenes);

    const ih = sheetItems.getRange(1, 1, 1, sheetItems.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const cOrden = ih.indexOf('ORDEN_ID'), cSku = ih.indexOf('SKU'), cTalla = ih.indexOf('TALLA');
    const cLong = ih.indexOf('LONGITUD'), cColor = ih.indexOf('COLOR'), cCosto = ih.indexOf('COSTO_UNITARIO');
    const cEstItem = ih.indexOf('ESTATUS_ITEM'), cEntr = ih.indexOf('ENTREGADO_ITEM'), cCour = ih.indexOf('COURIER_ITEM');

    const costosData = sheetCostos.getDataRange().getValues();
    const chh = costosData[0].map(h => String(h).trim());
    const kTipoRef = chh.indexOf('TIPO_REFERENCIA'), kRefId = chh.indexOf('REFERENCIA_ID');
    const kTipoC = chh.indexOf('TIPO_COSTO'), kDesc = chh.indexOf('DESCRIPCION');

    const resultado = [];
    const costRowsToDelete = [];

    items.forEach(it => {
      const fila = Number(it.itemRowNum);
      if (!fila || fila < 2) return;
      const oldRow = sheetItems.getRange(fila, 1, 1, sheetItems.getLastColumn()).getValues()[0];
      if (String(oldRow[cOrden]).trim() !== ordenId) return;
      const oldSku = String(oldRow[cSku]).trim();
      const oldColor = String(oldRow[cColor]).trim();

      const newSku = it.sku ? String(it.sku).trim() : oldSku;
      const newTalla = it.talla ? String(it.talla).trim() : String(oldRow[cTalla]).trim();
      const newLong = it.longitud ? String(it.longitud).trim() : String(oldRow[cLong]).trim();
      const newColor = it.color ? String(it.color).trim() : oldColor;

      // Opción A: editar la fila en sitio (conserva precio de venta)
      sheetItems.getRange(fila, cSku + 1).setValue(newSku);
      sheetItems.getRange(fila, cTalla + 1).setValue(newTalla);
      sheetItems.getRange(fila, cLong + 1).setValue(newLong);
      sheetItems.getRange(fila, cColor + 1).setValue(newColor);
      if (cCosto >= 0) sheetItems.getRange(fila, cCosto + 1).setValue('');
      if (cEstItem >= 0) sheetItems.getRange(fila, cEstItem + 1).setValue('HACER PEDIDO');
      if (cEntr >= 0) sheetItems.getRange(fila, cEntr + 1).setValue('FALSE');
      if (cCour >= 0) sheetItems.getRange(fila, cCour + 1).setValue('');

      // Si FIGS repone → borrar el costo bruto viejo (no es pérdida tuya)
      if (it.figsRepone) {
        for (let r = 1; r < costosData.length; r++) {
          if (String(costosData[r][kTipoRef]).toUpperCase().trim() !== 'PEDIDO') continue;
          if (String(costosData[r][kRefId]).trim() !== ordenId) continue;
          if (String(costosData[r][kTipoC]).toUpperCase().trim() !== 'BRUTO') continue;
          const d = String(costosData[r][kDesc]).toUpperCase();
          if (d.indexOf(oldSku.toUpperCase()) === 0 && (oldColor ? d.indexOf(oldColor.toUpperCase()) !== -1 : true)) {
            costRowsToDelete.push(r + 1);
            costosData[r][kTipoRef] = '__DELETED__';
            break;
          }
        }
      }

      resultado.push({ fila, sku: newSku, talla: newTalla, color: newColor, figsRepone: !!it.figsRepone });
    });

    costRowsToDelete.sort((a, b) => b - a).forEach(r => sheetCostos.deleteRow(r));

    // Recalcular cabecera al estado mínimo de los ítems vivos
    SpreadsheetApp.flush();
    const allItems = sheetItems.getDataRange().getValues();
    let minIdx = 999, hayVivos = false;
    for (let r = 1; r < allItems.length; r++) {
      if (String(allItems[r][cOrden]).trim() !== ordenId) continue;
      const est = String(allItems[r][cEstItem]).toUpperCase().trim();
      if (est === 'CANCELADO') continue;
      hayVivos = true;
      const idx = VALID_STATES.indexOf(est);
      if (idx !== -1 && idx < minIdx) minIdx = idx;
    }
    const orderRow = findOrderRow(ordenId);
    if (orderRow !== -1 && hayVivos && minIdx < 999) {
      sheetOrdenes.getRange(orderRow, COL_ESTATUS_ENVIO).setValue(VALID_STATES[minIdx]);
    }

    logChange(usuario, ordenId, 'PERDIDA_TRANSITO',
      `${resultado.length} ítem(s) perdido(s) · ${resultado.filter(r => r.figsRepone).length} repuesto(s) por FIGS`);

    return jsonResponse({ ok: true, data: { ordenId, items: resultado } });
  });
}

// ============ COLORES (hex por color FIGS) ============

function getColores() {
  const rows = readSheetAsObjects(TABS.colores);
  const map = {};
  rows.forEach(r => {
    const color = String(r.COLOR || '').toUpperCase().trim();
    const hex   = String(r.HEX || '').trim();
    if (color) map[color] = hex;
  });
  return map;  // { "WILD IRIS": "#8E7CC3", ... }
}

function setColor(body) {
  const color = String(body.color || '').toUpperCase().trim();
  const hex   = String(body.hex || '').trim();
  const usuario = body.usuario || 'desconocido';

  if (!verifyAdmin(usuario)) return errorResponse('Solo admin puede editar colores', 403);
  if (!color) return errorResponse('Falta el color');
  if (!/^#[0-9A-Fa-f]{6}$/.test(hex)) return errorResponse('Hex inválido (usa formato #RRGGBB)');

  return withLock(function () {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(TABS.colores);
    if (!sheet) {
      sheet = ss.insertSheet(TABS.colores);
      sheet.appendRow(['COLOR', 'HEX']);
    }
    const data = sheet.getDataRange().getValues();
    const h = data[0].map(x => String(x).trim().toUpperCase());
    const cColor = h.indexOf('COLOR');
    const cHex   = h.indexOf('HEX');
    if (cColor === -1 || cHex === -1) throw new Error('Faltan columnas COLOR/HEX en TablaColores');

    // ¿Ya existe? actualizar; si no, agregar
    let fila = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][cColor]).toUpperCase().trim() === color) { fila = r + 1; break; }
    }
    if (fila === -1) {
      sheet.appendRow([color, hex]);
    } else {
      sheet.getRange(fila, cHex + 1).setValue(hex);
    }

    logChange(usuario, '-', 'COLOR_HEX', `${color} → ${hex}`);
    return jsonResponse({ ok: true, data: { color, hex } });
  });
}

 
// ---- Generador de ID, siguiendo tu patron nextDescuentoId ----
function nextGastoFijoId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.gastosFijos);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'GF-0001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;
  ids.forEach(id => {
    const m = String(id).match(/GF-(\d+)/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });
  return 'GF-' + String(maxNum + 1).padStart(4, '0');
}
 
 
// ---- Categorias validas (para mantener consistencia) ----
const CATEGORIAS_GASTO_FIJO = ['SUELDO', 'SUSCRIPCION', 'SERVICIOS', 'PUBLICIDAD', 'OTRO'];
 
 
// ---- Agregar un gasto fijo (entrada manual) ----
// body: { fecha, concepto, categoria, monto, nota, usuario }
//   fecha    : 'YYYY-MM-DD' (opcional; si falta usa hoy)
//   concepto : 'Sebas semana 1', 'Shopify agosto', etc.
//   categoria: una de CATEGORIAS_GASTO_FIJO
//   monto    : numero positivo
function agregarGastoFijo(body) {
  const usuario   = body.usuario || 'desconocido';
  const concepto  = String(body.concepto || '').trim();
  const categoria = String(body.categoria || 'OTRO').toUpperCase().trim();
  const monto     = Number(body.monto);
  const nota      = String(body.nota || '').trim();
  let   fecha     = String(body.fecha || '').trim();
 
  // Validaciones (mismo estilo que agregarDescuento)
  if (!canOperar(usuario)) return errorResponse('No tienes permiso para registrar gastos', 403);
  if (!concepto) return errorResponse('Falta el concepto del gasto');
  if (isNaN(monto) || monto <= 0) return errorResponse('El monto debe ser mayor a 0');
  if (CATEGORIAS_GASTO_FIJO.indexOf(categoria) === -1) {
    return errorResponse('Categoria invalida. Usa: ' + CATEGORIAS_GASTO_FIJO.join(', '));
  }
  if (!fecha) fecha = Utilities.formatDate(new Date(), 'GMT-5', 'yyyy-MM-dd');
 
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.gastosFijos);
  if (!sheet) return errorResponse('No existe la hoja TablaGastosFijos. Creala primero.', 500);
 
  const id = nextGastoFijoId();
  sheet.appendRow([ id, fecha, concepto, categoria, round2(monto), nota ]);
 
  logChange(usuario, id, 'GASTO_FIJO_AGREGADO',
    `${concepto} · ${categoria} · $${round2(monto)}`);
 
  return jsonResponse({ ok: true, data: {
    id, fecha, concepto, categoria, monto: round2(monto), nota
  }});
}
 
 
// ---- Borrar un gasto fijo (por si te equivocas) ----
// body: { gastoId, usuario }
function borrarGastoFijo(body) {
  const usuario = body.usuario || 'desconocido';
  const gastoId = String(body.gastoId || '').trim();
 
  if (!canOperar(usuario)) return errorResponse('No tienes permiso para borrar gastos', 403);
  if (!gastoId) return errorResponse('Falta gastoId');
 
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(TABS.gastosFijos);
  if (!sheet) return errorResponse('No existe la hoja TablaGastosFijos.', 500);
 
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return errorResponse('No hay gastos registrados', 404);
 
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(x => String(x).trim() === gastoId);
  if (idx === -1) return errorResponse('Gasto no encontrado: ' + gastoId, 404);
 
  sheet.deleteRow(idx + 2); // +2: fila 1 es encabezado, idx es 0-based
  logChange(usuario, gastoId, 'GASTO_FIJO_BORRADO', '');
 
  return jsonResponse({ ok: true, data: { gastoId, borrado: true } });
}
 
 
// ---- Leer gastos fijos de un mes (para el resumen) ----
// mes: 'YYYY-MM' (ej '2026-08'). Si se omite, devuelve todos.
// Devuelve la lista + el total, listo para restar en el resumen mensual.
function getGastosFijos(mes) {
  const filas = readSheetAsObjects(TABS.gastosFijos);
  return jsonResponse({ ok: true, data: calcularGastosFijos(filas, mes) });
}

// Misma lógica que antes, pero recibe la tabla YA LEÍDA en vez de releerla.
function calcularGastosFijos(filas, mes) {
  const filtro = mes ? String(mes).trim().slice(0, 7) : '';
 
  const lista = [];
  let total = 0;
  const porCategoria = {};
 
  filas.forEach(g => {
    const fecha = String(g.FECHA || '').trim().slice(0, 10);
    if (!fecha) return;
    if (filtro && fecha.slice(0, 7) !== filtro) return;
 
    const monto = Number(g.MONTO) || 0;
    const cat   = String(g.CATEGORIA || 'OTRO').toUpperCase().trim();
    total += monto;
    porCategoria[cat] = (porCategoria[cat] || 0) + monto;
 
    lista.push({
      id: String(g.GASTO_ID || '').trim(),
      fecha: fecha,
      concepto: String(g.CONCEPTO || '').trim(),
      categoria: cat,
      monto: round2(monto),
      nota: String(g.NOTA || '').trim()
    });
  });
 
  // Redondear los subtotales por categoria
  Object.keys(porCategoria).forEach(k => porCategoria[k] = round2(porCategoria[k]));
 
  return {
    mes: filtro || 'todos',
    gastos: lista,
    porCategoria: porCategoria,
    total: round2(total),
    cantidad: lista.length
  };
}


// ============================================================
// A) RESUMEN MENSUAL — la pieza central
// ============================================================
// mes: 'YYYY-MM'.  donacion: monto opcional que devuelven a ScrubMe.
// Devuelve las dos vistas: repartible (socios) y salud (capital).
function getResumenMensual(mes, donacion) {
  const m = String(mes || '').trim().slice(0, 7);
  if (!m) return errorResponse('Falta el mes (formato YYYY-MM)');
  const don = Number(donacion) || 0;
 
  // Rango de fechas del mes
  const ini = m + '-01';
  const fin = m + '-31';
 
  // 1) Ganancia de pedidos del mes (usa el reporte existente)
  const rep = getReporteGanancia(ini, fin);
 
  // 2) SOLO cuenta lo ENTREGADO Y PAGADO para el reparto.
  //    getReporteGanancia trae detalle con ORDEN_ID; cruzamos
  //    contra TablaOrdenes para saber cuales estan entregados+pagados.
  const ordenes = readSheetAsObjects(TABS.ordenes);
  const estadoPago = {};
  ordenes.forEach(o => {
    const id = String(o.ORDEN_ID).trim();
    estadoPago[id] = {
      estado: String(o.ESTATUS_ENVIO || '').toUpperCase().trim(),
      saldo:  Number(o.SALDO) || 0
    };
  });
 
  let gananciaRepartiblePedido = 0;  // pedidos (no stock) entregados y pagados
  let gananciaStock = 0;             // ventas de stock (van a ScrubMe)
  let gananciaProyectada = 0;        // todo lo del mes, entregado o no
  const detalleRepartible = [];
 
  rep.detalle.forEach(d => {
    const ep = estadoPago[d.ORDEN_ID] || { estado: '', saldo: 999 };
    gananciaProyectada += d.ganancia;
 
    if (d.TIPO === 'STOCK') {
      gananciaStock += d.ganancia;      // stock siempre va a ScrubMe
      return;
    }
    // PEDIDO: solo entra al reparto si esta ENTREGADO y saldo 0
    const entregadoPagado = (ep.estado === 'ENTREGADO' && ep.saldo <= 0.5);
    if (entregadoPagado) {
      gananciaRepartiblePedido += d.ganancia;
      detalleRepartible.push(d);
    }
  });
 
  // 3) Gastos fijos del mes (Pieza 1)
  const gf = JSON.parse(getGastosFijos(m).getContent());
  const gastosFijos = (gf.data && gf.data.total) || 0;
 
  // 4) Las dos ecuaciones
  const repartibleAntesGastos = round2(gananciaRepartiblePedido);
  const repartibleNeto = round2(gananciaRepartiblePedido - gastosFijos - don);
  const porSocio = round2(repartibleNeto / 2);
 
  return jsonResponse({ ok: true, data: {
    mes: m,
 
    // --- BOLSILLO 1: repartible entre socios ---
    repartible: {
      gananciaPedidosPagados: repartibleAntesGastos,
      gastosFijos: round2(gastosFijos),
      donacion: round2(don),
      neto: repartibleNeto,
      porSocio: porSocio,
      cantidadPedidos: detalleRepartible.length
    },
 
    // --- BOLSILLO 2: capital ScrubMe (no se reparte) ---
    scrubme: {
      gananciaStock: round2(gananciaStock)
    },
 
    // --- referencia: proyeccion (incluye no cobrados) ---
    proyeccion: {
      gananciaTotalMes: round2(gananciaProyectada),
      brecha: round2(gananciaProyectada - repartibleAntesGastos) // lo aun no cobrado
    },
 
    detalleRepartible: detalleRepartible
  }});
}
 
 
// ============================================================
// B) CAPITAL REAL — foto del negocio (usa TablaSaldos)
// ============================================================
// Suma liquido y deuda de la ULTIMA foto en TablaSaldos, y le
// agrega el invertido (stock vivo + pedidos activos) calculado
// en vivo desde las tablas.
function getCapitalReal() {
  // 1) Ultima foto de saldos (por fecha mas reciente)
  const saldos = readSheetAsObjects(TABS.saldos);
  let ultimaFecha = '';
  saldos.forEach(s => {
    const f = String(s.FECHA || '').trim().slice(0, 10);
    if (f > ultimaFecha) ultimaFecha = f;
  });
 
  let liquido = 0, deuda = 0;
  const cuentas = [];
  saldos.forEach(s => {
    const f = String(s.FECHA || '').trim().slice(0, 10);
    if (f !== ultimaFecha) return;
    const monto = Number(s.MONTO) || 0;
    const tipo = String(s.TIPO || '').toUpperCase().trim();
    if (tipo === 'ACTIVO') liquido += monto;
    else if (tipo === 'DEUDA') deuda += monto;
    cuentas.push({
      cuenta: String(s.CUENTA || '').trim(),
      tipo: tipo,
      monto: round2(monto)
    });
  });
 
  // 2) Stock vivo (CANT_DISPONIBLE > 0)
  const lotes = readSheetAsObjects(TABS.lotes);
  let stock = 0;
  lotes.forEach(l => {
    const disp = Number(l.CANT_DISPONIBLE) || 0;
    if (disp > 0) stock += disp * (Number(l.COSTO_UNITARIO) || 0);
  });
 
  // 3) Pedidos activos (no entregado, no cancelado)
  const ordenes = readSheetAsObjects(TABS.ordenes);
  let pedidosActivos = 0;
  ordenes.forEach(o => {
    const est = String(o.ESTATUS_ENVIO || '').toUpperCase().trim();
    if (est === 'ENTREGADO' || est === 'CANCELADO') return;
    pedidosActivos += Number(o.TOTAL_COSTOS) || 0;
  });
 
  const invertido = round2(stock + pedidosActivos);
  const capital = round2(liquido + invertido - deuda);
 
  return jsonResponse({ ok: true, data: {
    fechaFoto: ultimaFecha,
    liquido: round2(liquido),
    stock: round2(stock),
    pedidosActivos: round2(pedidosActivos),
    invertido: invertido,
    deuda: round2(deuda),
    capitalReal: capital,
    cuentas: cuentas
  }});
}
 
 
// ============================================================
// C) DATA PARA GRAFICAS
// ============================================================
 
// C1) PIE CHART — a donde va el dinero (costos por tipo) en un rango
// Devuelve [{tipo:'BRUTO', monto:...}, {tipo:'COURIER',...}, ...]
function getCostosPorTipo(fechaInicio, fechaFin) {
  const costos = readSheetAsObjects(TABS.costos);
  const ini = String(fechaInicio || '').trim();
  const fin = String(fechaFin || '').trim();
 
  const porTipo = {};
  costos.forEach(c => {
    const f = String(c.FECHA || '').trim().slice(0, 10);
    if (ini && f < ini) return;
    if (fin && f > fin) return;
    const tipo = String(c.TIPO_COSTO || 'OTRO').toUpperCase().trim();
    porTipo[tipo] = (porTipo[tipo] || 0) + (Number(c.MONTO) || 0);
  });
 
  const data = Object.keys(porTipo).map(t => ({ tipo: t, monto: round2(porTipo[t]) }));
  data.sort((a, b) => b.monto - a.monto);
  const total = round2(data.reduce((s, x) => s + x.monto, 0));
 
  return jsonResponse({ ok: true, data: { desde: ini, hasta: fin, total: total, porTipo: data } });
}
 
 
/* ============================================================
   getTimelineMensual v2 — MESES COMPLETOS + PROYECCIÓN
   ============================================================
   REEMPLAZA la función getTimelineMensual actual por esta.

   Cambios:
   - Opción A: la línea de datos REALES solo llega hasta el
     último mes COMPLETO (no incluye el mes en curso a medias).
   - Opción C: agrega un punto PROYECTADO para el mes en curso,
     estimado por el ritmo diario de lo que va del mes.

   Cada punto trae una bandera:
     completo: true  -> mes terminado, dato real
     completo: false -> mes en curso, valores PROYECTADOS
   El frontend usa esa bandera para dibujar sólido vs punteado.
   ============================================================ */
function getTimelineMensual(meses) {
  const n = Number(meses) || 6;
  const hoy = new Date();
  const anioHoy = hoy.getFullYear();
  const mesHoy = hoy.getMonth();          // 0-11
  const diaHoy = hoy.getDate();

  // Leer las tablas grandes UNA SOLA VEZ, no una vez por mes.
  const ordenes     = readSheetAsObjects(TABS.ordenes);
  const items       = readSheetAsObjects(TABS.items);
  const costos       = readSheetAsObjects(TABS.costos);
  const gastosFijos = readSheetAsObjects(TABS.gastosFijos);

  const salida = [];

  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(anioHoy, mesHoy - i, 1);
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    const mes = `${y}-${mo}`;
    const ini = `${mes}-01`;
    const fin = `${mes}-31`;

    const esMesActual = (d.getFullYear() === anioHoy && d.getMonth() === mesHoy);

    // Cálculo en memoria: nada de esto vuelve a tocar el Sheet.
    const rep = calcularReporteGanancia(ordenes, items, costos, ini, fin);
    const gastos = calcularGastosFijos(gastosFijos, mes).total || 0;

    if (!esMesActual) {
      // Mes COMPLETO -> dato real
      salida.push({
        mes: mes,
        completo: true,
        ventaPedido: rep.pedido.venta,
        gananciaPedido: rep.pedido.ganancia,
        ventaStock: rep.stock.venta,
        gananciaStock: rep.stock.ganancia,
        gastosFijos: round2(gastos),
        gananciaNeta: round2(rep.pedido.ganancia - gastos)
      });
    } else {
      // Mes EN CURSO -> proyección por ritmo diario
      const diasEnMes = new Date(y, d.getMonth() + 1, 0).getDate(); // último día del mes
      const factor = diaHoy > 0 ? (diasEnMes / diaHoy) : 1;

      salida.push({
        mes: mes,
        completo: false,
        diasTranscurridos: diaHoy,
        diasDelMes: diasEnMes,
        // valores REALES de lo que va del mes (por si el frontend los quiere)
        ventaReal: rep.pedido.venta,
        gananciaReal: rep.pedido.ganancia,
        // valores PROYECTADOS a fin de mes
        ventaPedido: round2(rep.pedido.venta * factor),
        gananciaPedido: round2(rep.pedido.ganancia * factor),
        ventaStock: round2(rep.stock.venta * factor),
        gananciaStock: round2(rep.stock.ganancia * factor),
        gastosFijos: round2(gastos), // los gastos fijos NO se proyectan (son fijos)
        gananciaNeta: round2(rep.pedido.ganancia * factor - gastos)
      });
    }
  }

  return jsonResponse({ ok: true, data: { meses: n, timeline: salida } });
}
 
 
// C3) HISTOGRAMA — distribucion de ganancia por pedido en un rango
// Para ver cuantos pedidos caen en cada rango de ganancia.
function getHistogramaGanancia(fechaInicio, fechaFin) {
  const rep = getReporteGanancia(fechaInicio, fechaFin);
  // Rangos (buckets) de ganancia en USD
  const buckets = [
    { label: '< 0 (perdida)', min: -99999, max: 0, count: 0 },
    { label: '0 - 20',   min: 0,  max: 20, count: 0 },
    { label: '20 - 40',  min: 20, max: 40, count: 0 },
    { label: '40 - 60',  min: 40, max: 60, count: 0 },
    { label: '60 - 100', min: 60, max: 100, count: 0 },
    { label: '100+',     min: 100, max: 99999, count: 0 }
  ];
  rep.detalle.forEach(d => {
    const g = d.ganancia;
    for (const b of buckets) {
      if (g >= b.min && g < b.max) { b.count++; break; }
    }
  });
  return jsonResponse({ ok: true, data: {
    desde: fechaInicio, hasta: fechaFin,
    totalPedidos: rep.detalle.length,
    buckets: buckets
  }});
}
 
 
// C4) CAPITAL EN EL TIEMPO — todas las fotos de TablaSaldos
// Para line chart de si el negocio sube o baja mes a mes.
// (Util cuando tengas 2+ fotos.)
function getCapitalTimeline() {
  const saldos = readSheetAsObjects(TABS.saldos);
  const porFecha = {};  // fecha -> {liquido, deuda}
  saldos.forEach(s => {
    const f = String(s.FECHA || '').trim().slice(0, 10);
    if (!f) return;
    if (!porFecha[f]) porFecha[f] = { liquido: 0, deuda: 0 };
    const monto = Number(s.MONTO) || 0;
    const tipo = String(s.TIPO || '').toUpperCase().trim();
    if (tipo === 'ACTIVO') porFecha[f].liquido += monto;
    else if (tipo === 'DEUDA') porFecha[f].deuda += monto;
  });
 
  const fechas = Object.keys(porFecha).sort();
  const linea = fechas.map(f => ({
    fecha: f,
    liquido: round2(porFecha[f].liquido),
    deuda: round2(porFecha[f].deuda),
    // nota: no incluye stock/pedidos historicos (no se guardan),
    // solo liquido y deuda de cada foto
    liquidoMenosDeuda: round2(porFecha[f].liquido - porFecha[f].deuda)
  }));
 
  return jsonResponse({ ok: true, data: { fotos: linea.length, timeline: linea } });
}

function TEST_resumen() {
  Logger.log(getResumenMensual('2026-07', 0).getContent());
}

function TEST_capital() {
  Logger.log(getCapitalReal().getContent());
}

function TEST_costos() {
  Logger.log(getCostosPorTipo('2026-07-01', '2026-07-31').getContent());
}

/* ============================================================
   getResumenMensual v2 — REPARTO POR FECHA DE COBRO (Opción 2a)
   ============================================================
   REEMPLAZA la función getResumenMensual actual por esta.

   Cambio principal: un pedido cuenta para el reparto en el
   período donde se registró su ÚLTIMO pago (el que lo dejó en
   saldo 0), NO por su fecha de pedido. Así el reparto refleja
   "cuánta plata entró" en el período — ideal para el corte 15-a-15.

   NUEVO: acepta rango libre de fechas (para cortes 15 a 15).
   Uso:
     getResumenMensual('2026-07')                  -> mes completo
     getResumenMensual(null, don, '2026-07-15', '2026-08-15') -> corte 15 a 15
   ============================================================ */
function getResumenMensual(mes, donacion, fechaDesde, fechaHasta) {
  const don = Number(donacion) || 0;

  // Rango: o bien un mes 'YYYY-MM', o bien un rango libre desde/hasta
  let ini, fin, etiqueta;
  if (fechaDesde && fechaHasta) {
    ini = String(fechaDesde).trim().slice(0, 10);
    fin = String(fechaHasta).trim().slice(0, 10);
    etiqueta = `${ini} a ${fin}`;
  } else {
    const m = String(mes || '').trim().slice(0, 7);
    if (!m) return errorResponse('Falta el mes (YYYY-MM) o el rango desde/hasta');
    ini = m + '-01';
    fin = m + '-31';
    etiqueta = m;
  }

  const ordenes = readSheetAsObjects(TABS.ordenes);
  const items   = readSheetAsObjects(TABS.items);
  const costos  = readSheetAsObjects(TABS.costos);
  const pagos   = readSheetAsObjects(TABS.pagos);

  function ymd(v) { return String(v || '').trim().slice(0, 10); }
  function num(v) { return Number(v) || 0; }

  // ---------------------------------------------------------
  // 1) Para cada pedido: fecha de su ULTIMO pago y total pagado.
  //    (los pagos negativos = devoluciones, tambien cuentan al neto)
  // ---------------------------------------------------------
  const pagoPorOrden = {}; // id -> { totalPagado, ultimaFecha }
  pagos.forEach(p => {
    const id = String(p.ORDEN_ID || '').trim();
    if (!id) return;
    const f = ymd(p.FECHA_PAGO);
    const monto = num(p.MONTO);
    if (!pagoPorOrden[id]) pagoPorOrden[id] = { totalPagado: 0, ultimaFecha: '' };
    pagoPorOrden[id].totalPagado += monto;
    // la ultima fecha entre pagos POSITIVOS (un reembolso no "completa" el pedido)
    if (monto > 0 && f > pagoPorOrden[id].ultimaFecha) {
      pagoPorOrden[id].ultimaFecha = f;
    }
  });

  // ---------------------------------------------------------
  // 2) Info de cada orden: venta, estado, tipo (stock/pedido), notas
  // ---------------------------------------------------------
  const infoOrden = {};
  ordenes.forEach(o => {
    const id = String(o.ORDEN_ID || '').trim();
    if (!id) return;
    infoOrden[id] = {
      estado: String(o.ESTATUS_ENVIO || '').toUpperCase().trim(),
      venta:  num(o.TOTAL_VENTA),
      saldo:  num(o.SALDO),
      notas:  String(o.NOTAS || '').toUpperCase(),
      nombre: String(o.CLIENTE_NOMBRE || '').trim()
    };
  });

  // Tipo (STOCK vs PEDIDO) desde items: si alguna prenda viva es STOCK -> STOCK
  const tipoOrden = {};
  items.forEach(it => {
    const id = String(it.ORDEN_ID || '').trim();
    if (!id) return;
    const est = String(it.ESTATUS_ITEM || '').toUpperCase().trim();
    if (est === 'CANCELADO') return;
    const origen = String(it.ORIGEN || '').toUpperCase().trim();
    if (!tipoOrden[id]) tipoOrden[id] = 'PEDIDO';
    if (origen === 'STOCK') tipoOrden[id] = 'STOCK';
  });

  // Costo real por pedido (suma de TablaCostos tipo PEDIDO)
  const costoOrden = {};
  costos.forEach(c => {
    if (String(c.TIPO_REFERENCIA || '').toUpperCase().trim() !== 'PEDIDO') return;
    const id = String(c.REFERENCIA_ID || '').trim();
    costoOrden[id] = (costoOrden[id] || 0) + num(c.MONTO);
  });

  // ---------------------------------------------------------
  // 3) Recorrer pedidos: cuentan los que se TERMINARON DE PAGAR
  //    dentro del rango (ultima fecha de pago en [ini, fin])
  //    y estan entregados + saldo 0.
  // ---------------------------------------------------------
  let gananciaRepartiblePedido = 0;
  let gananciaStock = 0;
  const detalleRepartible = [];

  Object.keys(infoOrden).forEach(id => {
    const info = infoOrden[id];
    if (info.estado === 'CANCELADO') return;
    // excluir marketing/sorteo
    if (info.notas.indexOf('MARKETING') !== -1 || info.notas.indexOf('SORTEO') !== -1) return;

    const pg = pagoPorOrden[id];
    if (!pg) return;                       // nunca se pago -> no entra
    if (info.saldo > 0.5) return;          // aun debe -> no entra
    if (info.estado !== 'ENTREGADO') return; // no entregado -> no entra

    // La fecha que decide el período es la del ULTIMO pago
    const fechaCobro = pg.ultimaFecha;
    if (!fechaCobro) return;
    if (fechaCobro < ini || fechaCobro > fin) return;  // fuera del corte

    const venta = info.venta;
    const costo = round2(costoOrden[id] || 0);
    const ganancia = round2(venta - costo);

    if (tipoOrden[id] === 'STOCK') {
      gananciaStock += ganancia;           // stock -> ScrubMe
    } else {
      gananciaRepartiblePedido += ganancia;
      detalleRepartible.push({
        ORDEN_ID: id, NOMBRE: info.nombre, FECHA_COBRO: fechaCobro,
        TIPO: 'PEDIDO', venta: round2(venta), costo: costo, ganancia: ganancia
      });
    }
  });

  // ---------------------------------------------------------
  // 4) Gastos fijos del período
  // ---------------------------------------------------------
  let gastosFijos = 0;
  const gf = readSheetAsObjects(TABS.gastosFijos);
  gf.forEach(g => {
    const f = ymd(g.FECHA);
    if (!f) return;
    if (f < ini || f > fin) return;
    gastosFijos += num(g.MONTO);
  });

  // ---------------------------------------------------------
  // 5) Las ecuaciones
  // ---------------------------------------------------------
  const repartibleAntesGastos = round2(gananciaRepartiblePedido);
  const repartibleNeto = round2(gananciaRepartiblePedido - gastosFijos - don);
  const porSocio = round2(repartibleNeto / 2);

  // Brecha: pedidos ENTREGADOS en el rango pero AUN sin cobrar del todo.
  // (plata que trabajaste en el período pero todavía no entra)
  let porCobrar = 0;
  Object.keys(infoOrden).forEach(id => {
    const info = infoOrden[id];
    if (info.estado === 'CANCELADO') return;
    if (info.notas.indexOf('MARKETING') !== -1 || info.notas.indexOf('SORTEO') !== -1) return;
    if (info.saldo <= 0.5) return;   // ya pagado, no es brecha
    // saldo pendiente de pedidos vivos
    porCobrar += info.saldo;
  });

  return jsonResponse({ ok: true, data: {
    periodo: etiqueta,
    desde: ini,
    hasta: fin,
    baseCalculo: 'fecha de cobro (último pago)',

    repartible: {
      gananciaPedidosPagados: repartibleAntesGastos,
      gastosFijos: round2(gastosFijos),
      donacion: round2(don),
      neto: repartibleNeto,
      porSocio: porSocio,
      cantidadPedidos: detalleRepartible.length
    },

    scrubme: {
      gananciaStock: round2(gananciaStock)
    },

    // se mantiene para compatibilidad con el frontend actual
    proyeccion: {
      gananciaTotalMes: round2(repartibleAntesGastos + gananciaStock),
      brecha: round2(porCobrar)
    },

    detalleRepartible: detalleRepartible
  }});
}

function TEST_corte15() {
  Logger.log(getResumenMensual(null, 0, '2026-07-15', '2026-08-15').getContent());
}
