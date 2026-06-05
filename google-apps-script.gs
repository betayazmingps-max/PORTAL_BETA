/**
 * ════════════════════════════════════════════════════════════════
 *  PORTAL BETA — Google Apps Script
 *  Guarda registros, documentos y datos de IA + responde al panel
 *
 *  CÓMO INSTALAR:
 *  1. Abre tu Google Sheet
 *  2. Extensiones → Apps Script
 *  3. Borra todo y pega este código
 *  4. Guardar → Implementar → Nueva implementación
 *  5. Tipo: Aplicación web
 *     - Ejecutar como: Yo
 *     - Quién accede: Cualquier usuario
 *  6. Copia la URL /exec y ponla en el Worker (GSHEET_URL)
 * ════════════════════════════════════════════════════════════════
 */

const HOJA_PROVEEDORES  = 'Proveedores';
const HOJA_VENCIMIENTOS = 'Vencimientos';
const HOJA_LOG          = 'Historial';

// Columnas de la hoja Proveedores
const COLS = ['fecha','ruc','razonSocial','contacto','email','telefono',
              'categoria','analista','estadoSunat','condicion','direccion',
              'obs','estado','notas','docs','ai_extraido'];

/* ════════════ POST — guardar / actualizar ════════════ */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    // ── Acciones del panel ──
    if (data.action === 'aprobar')              return accionEstado(ss, data.ruc, 'aprobado', data);
    if (data.action === 'rechazar')             return accionEstado(ss, data.ruc, 'rechazado', data);
    if (data.action === 'notas')                return accionNotas(ss, data.ruc, data.notas);
    if (data.action === 'log')                  return accionLog(ss, data);
    if (data.action === 'guardarVencimientos')  return accionVencimientos(ss, data);
    if (data.action === 'guardarAI')            return accionGuardarAI(ss, data);

    // ── Registro nuevo (sin action) ──
    return guardarRegistro(ss, data);

  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ════════════ GET — listar para el panel ════════════ */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = e.parameter.action || 'listar';

    if (action === 'listar') {
      const analista = e.parameter.analista || 'TODOS';
      const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
      if (!sheet || sheet.getLastRow() < 2) return json({ ok: true, proveedores: [] });

      const datos = sheet.getDataRange().getValues();
      const proveedores = [];

      for (let i = 1; i < datos.length; i++) {
        const fila = datos[i];
        const prov = {};
        COLS.forEach((col, idx) => { prov[col] = fila[idx] || ''; });
        prov.fila = i + 1;

        // Filtrar por analista (ADMIN ve todos)
        if (analista !== 'TODOS' && String(prov.analista) !== String(analista)) continue;

        // Parsear docs y ai_extraido (vienen como texto JSON)
        try { prov.docs        = prov.docs        ? JSON.parse(prov.docs)        : {}; } catch { prov.docs = {}; }
        try { prov.ai_extraido = prov.ai_extraido ? JSON.parse(prov.ai_extraido) : {}; } catch { prov.ai_extraido = {}; }

        proveedores.push(prov);
      }
      return json({ ok: true, proveedores });
    }

    return json({ ok: false, error: 'Acción GET desconocida' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ════════════ Guardar registro nuevo ════════════ */
function guardarRegistro(ss, data) {
  let sheet = ss.getSheetByName(HOJA_PROVEEDORES);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_PROVEEDORES);
    sheet.appendRow(['Fecha','RUC','Razón Social','Contacto','Email','Teléfono',
                     'Categoría','Analista','Estado SUNAT','Condición','Dirección',
                     'Observaciones','Estado','Notas','Docs','AI Extraído']);
  }
  // Evitar duplicado por RUC
  const existente = buscarFilaPorRuc(sheet, data.ruc);
  if (existente > 0) {
    return json({ ok: false, error: 'RUC ya registrado', duplicado: true });
  }

  sheet.appendRow([
    data.fecha || new Date().toLocaleString('es-PE'),
    data.ruc || '', data.razonSocial || '', data.contacto || '', data.email || '',
    data.telefono || '', data.categoria || '', data.analista || '',
    data.estadoSunat || '', data.condicion || '', data.direccion || '',
    data.obs || '', data.estado || 'pendiente', '',
    data.docs || '', data.ai_extraido || ''
  ]);
  return json({ ok: true });
}

/* ════════════ Aprobar / Rechazar ════════════ */
function accionEstado(ss, ruc, nuevoEstado, data) {
  const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
  const fila  = buscarFilaPorRuc(sheet, ruc);
  if (fila < 0) return json({ ok: false, error: 'RUC no encontrado' });
  sheet.getRange(fila, 13).setValue(nuevoEstado); // columna Estado
  if (data.motivo) sheet.getRange(fila, 14).setValue('RECHAZO: ' + data.motivo);
  return json({ ok: true });
}

/* ════════════ Notas ════════════ */
function accionNotas(ss, ruc, notas) {
  const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
  const fila  = buscarFilaPorRuc(sheet, ruc);
  if (fila < 0) return json({ ok: false, error: 'RUC no encontrado' });
  sheet.getRange(fila, 14).setValue(notas || ''); // columna Notas
  return json({ ok: true });
}

/* ════════════ Guardar datos IA (desde ingresar.html) ════════════ */
function accionGuardarAI(ss, data) {
  const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
  const fila  = buscarFilaPorRuc(sheet, data.ruc);
  if (fila < 0) return json({ ok: false, error: 'RUC no encontrado' });
  sheet.getRange(fila, 16).setValue(JSON.stringify(data.datos || {})); // columna AI Extraído
  return json({ ok: true });
}

/* ════════════ Vencimientos (hoja aparte) ════════════ */
function accionVencimientos(ss, data) {
  let sheet = ss.getSheetByName(HOJA_VENCIMIENTOS);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_VENCIMIENTOS);
    sheet.appendRow(['Fecha','RUC','Razón Social','Analista','Documento','Titular','Vencimiento','Número']);
  }
  (data.vencimientos || []).forEach(v => {
    sheet.appendRow([
      data.fecha || new Date().toLocaleString('es-PE'),
      data.ruc || '', data.razonSocial || '', data.analista || '',
      v.docNombre || v.id || '', v.titular || '', v.vencimiento || '', v.numero || ''
    ]);
  });
  return json({ ok: true });
}

/* ════════════ Log de acciones (hoja aparte) ════════════ */
function accionLog(ss, data) {
  let sheet = ss.getSheetByName(HOJA_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_LOG);
    sheet.appendRow(['Fecha','RUC','Analista','Acción','Detalle']);
  }
  sheet.appendRow([data.fecha || new Date().toLocaleString('es-PE'),
                   data.ruc || '', data.analista || '', data.accion || '', data.detalle || '']);
  return json({ ok: true });
}

/* ════════════ Helpers ════════════ */
function buscarFilaPorRuc(sheet, ruc) {
  if (!sheet || sheet.getLastRow() < 2) return -1;
  const rucs = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rucs.length; i++) {
    if (String(rucs[i][0]) === String(ruc)) return i + 2;
  }
  return -1;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
