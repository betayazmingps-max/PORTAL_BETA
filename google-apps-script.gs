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
const HOJA_DOCUMENTOS   = 'Documentos'; // un documento por fila/celda — ver accionGuardarDocs

// Columnas de la hoja Proveedores (orden REAL de tu Sheet)
// A=Fecha B=RUC C=Razón D=Contacto E=Email F=Tel G=Categoría
// H=Estado SUNAT I=Condición J=Dirección K=Observaciones L=Estado
// M=Notas N=Docs O=Notas Analista P=AI Extraído Q=Carpeta SharePoint
const COLS = ['fecha','ruc','razonSocial','contacto','email','telefono','categoria',
              'estadoSunat','condicion','direccion','obs','estado',
              'notas','docs','notasAnalista','ai_extraido','carpetaSP'];

// Índices de columnas (1-based para getRange)
const COL_DOCS            = 14; // N — solo metadata de documentos (JSON), contenido real en hoja "Documentos"
const COL_ESTADO          = 12; // L
const COL_NOTAS_ANALISTA  = 15; // O
const COL_AI_EXTRAIDO     = 16; // P
const COL_CARPETA_SP      = 17; // Q
const COL_ANALISTA        = 18; // R
const COL_DOCS_OPCIONALES = 19; // S — lista de IDs marcados como obligatorios por el analista (JSON)
const COL_DOCS_PERSON     = 20; // T — docs personalizados pedidos por el analista (JSON)
const COL_SUBCATEGORIA    = 21; // U — Sub categoría (solo Servicios Norte Y Sur: Materia Prima, Carga Seca, Refrigerado, Flota Vehicular)
const COL_SEDE            = 22; // V — Sede (solo Servicios Norte Y Sur: NORTE o SUR)

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
    if (data.action === 'guardarDocs')          return accionGuardarDocs(ss, data);

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
        prov.fila              = i + 1;
        prov.analista          = fila[17] || '';  // columna R
        prov.subcategoria      = fila[20] || '';  // columna U
        prov.sede              = fila[21] || '';  // columna V
        // Lista de docs condicionales marcados como obligatorios por el analista (S)
        try {
          const raw = fila[18];
          prov.docs_obligatorios = raw ? JSON.parse(raw) : [];
        } catch { prov.docs_obligatorios = []; }
        // Lista de docs personalizados (T)
        try {
          const raw = fila[19];
          prov.docs_personalizados = raw ? JSON.parse(raw) : [];
        } catch { prov.docs_personalizados = []; }

        // Filtrar por analista (ADMIN ve todos)
        if (analista !== 'TODOS' && String(prov.analista) !== String(analista)) continue;

        // Parsear docs y ai_extraido (vienen como texto JSON)
        try { prov.docs        = prov.docs        ? JSON.parse(prov.docs)        : {}; } catch { prov.docs = {}; }
        try { prov.ai_extraido = prov.ai_extraido ? JSON.parse(prov.ai_extraido) : {}; } catch { prov.ai_extraido = {}; }

        proveedores.push(prov);
      }
      return json({ ok: true, proveedores });
    }

    // Trae el contenido (base64) de UN documento a demanda — Proveedores!Docs
    // ya no guarda el contenido, solo metadata, así que preview.html pide el
    // archivo real acá, solo cuando el analista hace clic en "Ver documento"
    if (action === 'verDoc') {
      const ruc   = e.parameter.ruc || '';
      const docId = e.parameter.docId || '';
      const sheetDocs = ss.getSheetByName(HOJA_DOCUMENTOS);
      if (!sheetDocs || sheetDocs.getLastRow() < 2) return json({ ok: false, error: 'Documento no encontrado' });

      const datos = sheetDocs.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) {
        if (String(datos[i][0]) === String(ruc) && String(datos[i][1]) === String(docId)) {
          return json({ ok: true, nombre: datos[i][2], tipo: datos[i][3], contenido: datos[i][4] });
        }
      }
      return json({ ok: false, error: 'Documento no encontrado' });
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

  // Documentos del registro inicial (registro.html manda un JSON string con
  // [{id, nombre, tipo, contenido}, ...]) — se guardan uno por celda en la hoja
  // "Documentos" (ver guardarDocumentosIndividuales); acá solo queda metadata liviana.
  let docsArray = [];
  try { docsArray = data.docs ? JSON.parse(data.docs) : []; } catch(e) { docsArray = []; }
  const guardado = guardarDocumentosIndividuales(ss, data.ruc, docsArray);

  // Orden de columnas: A→Q (17 columnas)
  // OJO: El Analista NO está en tu Sheet — si lo necesitas, agrega columna R y descomenta abajo
  sheet.appendRow([
    data.fecha || new Date().toLocaleString('es-PE'),  // A Fecha
    data.ruc || '',                                     // B RUC
    data.razonSocial || '',                             // C Razón Social
    data.contacto || '',                                // D Contacto
    data.email || '',                                   // E Email
    data.telefono || '',                                // F Teléfono
    data.categoria || '',                               // G Categoría
    data.estadoSunat || '',                             // H Estado SUNAT
    data.condicion || '',                               // I Condición
    data.direccion || '',                               // J Dirección
    data.obs || '',                                     // K Observaciones
    data.estado || 'pendiente',                         // L Estado
    '',                                                  // M Notas
    JSON.stringify(guardado.resumen),                   // N Docs (solo metadata)
    '',                                                  // O Notas Analista
    data.ai_extraido || '',                             // P AI Extraído
    ''                                                   // Q Carpeta SharePoint
  ]);
  // Guardar analista en columna R si existe (opcional)
  const ultimaFila = sheet.getLastRow();
  if (data.analista) {
    sheet.getRange(ultimaFila, 18).setValue(data.analista);  // R Analista
  }
  // Sub categoría (solo aplica a Servicios Norte Y Sur: Materia Prima, Carga Seca, Refrigerado, Flota Vehicular)
  if (data.subcategoria) {
    sheet.getRange(ultimaFila, COL_SUBCATEGORIA).setValue(data.subcategoria);  // U Sub categoría
  }
  // Sede (solo aplica a Servicios Norte Y Sur: NORTE o SUR)
  if (data.sede) {
    sheet.getRange(ultimaFila, COL_SEDE).setValue(data.sede);  // V Sede
  }
  return json({ ok: true, omitidos: guardado.omitidos });
}

/* ════════════ Aprobar / Rechazar ════════════ */
function accionEstado(ss, ruc, nuevoEstado, data) {
  const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
  const fila  = buscarFilaPorRuc(sheet, ruc);
  if (fila < 0) return json({ ok: false, error: 'RUC no encontrado' });
  sheet.getRange(fila, COL_ESTADO).setValue(nuevoEstado);                                          // L Estado
  if (data.motivo) sheet.getRange(fila, COL_NOTAS_ANALISTA).setValue('RECHAZO: ' + data.motivo);   // O Notas Analista
  if (data.carpeta) sheet.getRange(fila, COL_CARPETA_SP).setValue(data.carpeta);                   // Q Carpeta SharePoint
  // Lista de docs condicionales que el analista marcó como obligatorios
  if (data.docs_obligatorios) {
    sheet.getRange(fila, COL_DOCS_OPCIONALES).setValue(JSON.stringify(data.docs_obligatorios));    // S
  }
  // Lista de docs personalizados (texto libre del analista)
  if (data.docs_personalizados) {
    sheet.getRange(fila, COL_DOCS_PERSON).setValue(JSON.stringify(data.docs_personalizados));     // T
  }
  return json({ ok: true });
}

/* ════════════ Notas ════════════ */
function accionNotas(ss, ruc, notas) {
  const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
  const fila  = buscarFilaPorRuc(sheet, ruc);
  if (fila < 0) return json({ ok: false, error: 'RUC no encontrado' });
  sheet.getRange(fila, COL_NOTAS_ANALISTA).setValue(notas || '');  // O Notas Analista
  return json({ ok: true });
}

/* ════════════ Guardar datos IA (desde ingresar.html y registro.html)
   MEZCLA con lo que ya hay (no sobrescribe), así no se pierden los del registro
══════════════════════════════════════════════════════════════════ */
function accionGuardarAI(ss, data) {
  const sheet = ss.getSheetByName(HOJA_PROVEEDORES);
  const fila  = buscarFilaPorRuc(sheet, data.ruc);
  if (fila < 0) return json({ ok: false, error: 'RUC no encontrado' });

  // Leer lo que ya hay en la columna AI Extraído
  let existente = {};
  try {
    const valor = sheet.getRange(fila, COL_AI_EXTRAIDO).getValue();
    if (valor) existente = JSON.parse(valor) || {};
  } catch(e) { existente = {}; }

  // Merge con los nuevos datos
  const merged = Object.assign({}, existente, data.datos || {});
  sheet.getRange(fila, COL_AI_EXTRAIDO).setValue(JSON.stringify(merged));  // P AI Extraído
  return json({ ok: true });
}

/* ════════════ Guardar documentos — un documento por fila/celda ════════════
   Usada por guardarRegistro (registro inicial) y accionGuardarDocs (subidas
   posteriores desde ingresar.html). Cada documento vive en su PROPIA celda
   de la hoja "Documentos" (upsert por RUC+DocId) — así un documento grande
   nunca se acumula junto a los demás en una sola celda.
   Google Sheets no permite más de 50,000 caracteres en una celda: si el
   base64 de un documento no cabe, se guarda igual su metadata (para que
   preview.html sepa que SÍ se subió) pero con muy_grande=true y sin
   contenido — ese documento se queda visible solo en SharePoint.
   docsArray: [{id, nombre, tipo, contenido}, ...]
   Devuelve { resumen, omitidos } — resumen es SOLO metadata (nombre/tipo/
   flags, nunca el contenido), pensado para guardarse en Proveedores!Docs,
   que por eso siempre cabe sin importar cuántos documentos tenga el proveedor.
══════════════════════════════════════════════════════════════════ */
const LIMITE_CELDA_DOC = 49000; // margen de seguridad bajo el límite real de 50,000

function guardarDocumentosIndividuales(ss, ruc, docsArray) {
  let sheetDocs = ss.getSheetByName(HOJA_DOCUMENTOS);
  if (!sheetDocs) {
    sheetDocs = ss.insertSheet(HOJA_DOCUMENTOS);
    sheetDocs.appendRow(['RUC', 'DocId', 'Nombre', 'Tipo', 'Contenido', 'Fecha']);
  }

  // Índice RUC+DocId → nº de fila, para actualizar (upsert) sin duplicar filas
  const indice = {};
  if (sheetDocs.getLastRow() > 1) {
    sheetDocs.getRange(2, 1, sheetDocs.getLastRow() - 1, 2).getValues()
      .forEach((f, i) => { indice[f[0] + '|' + f[1]] = i + 2; });
  }

  const resumen  = {};
  const omitidos = [];
  (docsArray || []).forEach(doc => {
    if (!doc || !doc.id) return;
    const contenido = doc.contenido || '';
    const cabe      = contenido.length <= LIMITE_CELDA_DOC;
    const key       = ruc + '|' + doc.id;
    const valores   = [ruc, doc.id, doc.nombre || '', doc.tipo || '', cabe ? contenido : '', new Date().toLocaleString('es-PE')];

    if (indice[key]) {
      sheetDocs.getRange(indice[key], 1, 1, 6).setValues([valores]);
    } else {
      sheetDocs.appendRow(valores);
      indice[key] = sheetDocs.getLastRow();
    }

    resumen[doc.id] = { nombre: doc.nombre || '', tipo: doc.tipo || '', guardado: true, muy_grande: !cabe };
    if (!cabe) omitidos.push(doc.nombre || doc.id);
  });

  return { resumen: resumen, omitidos: omitidos };
}

function accionGuardarDocs(ss, data) {
  const sheetProv = ss.getSheetByName(HOJA_PROVEEDORES);
  const fila      = buscarFilaPorRuc(sheetProv, data.ruc);
  if (fila < 0) return json({ ok: false, error: 'RUC no encontrado' });

  // Metadata previa de Proveedores!Docs, para no perder registro de documentos
  // guardados en llamadas anteriores (mezcla, no sobrescribe).
  // OJO: si la celda existe pero no se puede parsear (ej. una celda vieja de
  // antes de este fix, demasiado larga para leerse completa con getValue()),
  // NO se debe seguir — sobrescribir a ciegas borraría esa metadata para
  // siempre. Mejor cortar y avisar que se necesita revisión manual.
  let existente = {};
  const valor = sheetProv.getRange(fila, COL_DOCS).getValue();
  if (valor) {
    try {
      const parsed = JSON.parse(valor);
      existente = Array.isArray(parsed) ? {} : (parsed || {});
    } catch(e) {
      return json({ ok: false, error: 'No se pudo leer el contenido previo de Docs para RUC ' + data.ruc + ' (celda dañada o demasiado larga). No se modificó nada — revisa esa celda manualmente en el Sheet antes de reintentar.' });
    }
  }

  // Migración: filas viejas (de antes de este cambio) pueden traer el base64
  // completo embebido dentro de Proveedores!Docs. Si eso sigue ahí, se muda a la
  // hoja "Documentos" junto con los nuevos, para que la celda deje de estar
  // sobrecargada y no vuelva a fallar el próximo guardado de este proveedor.
  const docsArray = Object.keys(data.docs || {}).map(id => Object.assign({ id: id }, data.docs[id]));
  Object.keys(existente).forEach(id => {
    if (existente[id] && existente[id].contenido && !docsArray.some(d => d.id === id)) {
      docsArray.push(Object.assign({ id: id }, existente[id]));
    }
  });

  const guardado = guardarDocumentosIndividuales(ss, data.ruc, docsArray);

  // Reconstruir el resumen: metadata vieja SIN contenido embebido + lo recién guardado
  const existenteLimpio = {};
  Object.keys(existente).forEach(id => {
    const d = existente[id] || {};
    existenteLimpio[id] = { nombre: d.nombre || '', tipo: d.tipo || '', guardado: true, muy_grande: !!d.muy_grande };
  });

  const resumen = Object.assign({}, existenteLimpio, guardado.resumen);
  sheetProv.getRange(fila, COL_DOCS).setValue(JSON.stringify(resumen));  // N Docs (solo metadata)
  return json({ ok: true, omitidos: guardado.omitidos });
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

function probar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('ID: ' + ss.getId());
  Logger.log('Pestañas: ' + ss.getSheets().map(s => s.getName()).join(', '));
}
