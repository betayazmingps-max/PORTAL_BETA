/**
 * ════════════════════════════════════════════════════════════════
 *  PORTAL BETA — Cloudflare Worker
 *  URL: https://portalbeta.betayazmingps.workers.dev
 *
 *  RUTAS:
 *    POST /login_analista   → Valida credenciales del panel (preview.html)
 *    POST /analizar_doc     → Proxy de Claude AI para validar documentos
 *    POST /enviar_otp       → Envía código OTP por correo
 *    POST /verificar_otp    → Verifica el código OTP
 *    POST /sunat_ruc        → Consulta RUC en apis.net.pe
 *    POST /listar           → Lista proveedores desde Google Sheets
 *    POST /sheet            → Actualiza datos en Google Sheets
 *    POST /subir_docs       → Sube documentos a SharePoint vía Power Automate
 *    POST /crear_carpetas   → Crea carpetas en SharePoint vía Power Automate
 *    POST /notificar        → Envía notificación por correo al proveedor
 *
 *  VARIABLES DE ENTORNO (configura en Cloudflare Dashboard → Workers → Settings → Variables):
 *    ANTHROPIC_API_KEY      → sk-ant-...
 *    GSHEET_URL             → URL del Google Apps Script
 *    FLOW_SUBIR_DOCS        → URL del Flow de Power Automate para subir docs
 *    FLOW_CREAR_CARPETAS    → URL del Flow de Power Automate para crear carpetas
 *    RESEND_API_KEY         → API key de Resend (o tu servicio de email)
 *    EMAIL_FROM             → correo remitente (ej: noreply@betaagroindustrial.com)
 *    ANALISTAS_JSON         → JSON con credenciales: {"NORTE":{"pass":"xxx","nombre":"NORTE"},...}
 *    OTP_SECRET             → clave secreta para firmar OTPs (cualquier string largo)
 * ════════════════════════════════════════════════════════════════
 */

export default {
  async fetch(request, env, ctx) {
    // ── CORS ──
    const CORS = {
      'Access-Control-Allow-Origin':  '*',
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }

    if (request.method !== 'POST') {
      return json({ ok: false, error: 'Solo POST' }, 405, CORS);
    }

    const url  = new URL(request.url);
    const ruta = url.pathname;

    let body = {};
    try { body = await request.json(); } catch {}

    try {
      // ══════════════════════════════════════════
      //  🔐 LOGIN ANALISTA
      // ══════════════════════════════════════════
      if (ruta === '/login_analista') {
        const { analista, pass } = body;
        if (!analista || !pass) return json({ ok: false, error: 'Datos incompletos' }, 400, CORS);

        // Las contraseñas viven en la variable de entorno ANALISTAS_JSON
        // Formato: {"NORTE":{"pass":"beta2026","nombre":"NORTE"},"ADMIN":{"pass":"admin2026","nombre":"Administrador"}}
        let auth = {};
        try { auth = JSON.parse(env.ANALISTAS_JSON || '{}'); } catch {}

        const entrada = auth[analista];
        if (!entrada || entrada.pass !== pass) {
          return json({ ok: false, error: 'Contraseña incorrecta.' }, 401, CORS);
        }

        return json({
          ok:      true,
          nombre:  entrada.nombre,
          esAdmin: analista === 'ADMIN',
        }, 200, CORS);
      }

      // ══════════════════════════════════════════
      //  🤖 ANALIZAR DOC con Claude
      // ══════════════════════════════════════════
      if (ruta === '/analizar_doc') {
        const { docName, mediaType, isImg, data: b64, prompt, modo } = body;

        // Modo regularización: recibe un prompt de texto (sin archivo)
        if (modo === 'regularizacion' && prompt) {
          const resp = await fetch('https://api.anthropic.com/v1/messages', {
            method:  'POST',
            headers: {
              'Content-Type':      'application/json',
              'x-api-key':         env.ANTHROPIC_API_KEY,
              'anthropic-version': '2023-06-01',
            },
            body: JSON.stringify({
              model:      'claude-sonnet-4-20250514',
              max_tokens: 800,
              messages:   [{ role: 'user', content: prompt }],
            }),
          });
          const d    = await resp.json();
          const texto = d.content?.find(c => c.type === 'text')?.text || '{}';
          return json({ ok: true, texto }, 200, CORS);
        }

        // Modo simple o completo: recibe archivo en base64
        if (!b64 || !mediaType) return json({ ok: false, error: 'Faltan datos del archivo' }, 400, CORS);

        const promptCompleto = `Eres un validador de documentos de una empresa agroindustrial peruana.
El campo solicitado es: "${docName || 'Documento'}"
Analiza el documento y responde ÚNICAMENTE con este JSON (sin texto extra, sin markdown):
{
  "valido": true/false,
  "nitido": true/false,
  "motivo": "texto corto en español",
  "datos": {
    "titular": "nombre de la persona o empresa en el doc, o null",
    "numero": "número de documento/placa/licencia/RUC que aparezca, o null",
    "vencimiento": "fecha de vencimiento en formato YYYY-MM-DD si existe, o null",
    "emision": "fecha de emisión en formato YYYY-MM-DD si existe, o null",
    "observacion": "dato relevante adicional, o null"
  }
}
Reglas:
- valido: true si el documento corresponde al campo solicitado.
- nitido: true si se puede leer bien. false si está borroso o ilegible.
- motivo: si válido di "Documento verificado". Si no, explica brevemente.
- Sé flexible: acepta si tiene relación con el campo. Solo rechaza si es completamente diferente.`;

        const promptSimple = `Campo: "${docName}". Responde solo JSON: {"valido":true/false,"motivo":"texto","nitido":true/false}`;

        const contentParts = [
          isImg
            ? { type: 'image',    source: { type: 'base64', media_type: mediaType, data: b64 } }
            : { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64 } },
          { type: 'text', text: modo === 'simple' ? promptSimple : promptCompleto },
        ];

        const resp = await fetch('https://api.anthropic.com/v1/messages', {
          method:  'POST',
          headers: {
            'Content-Type':      'application/json',
            'x-api-key':         env.ANTHROPIC_API_KEY,
            'anthropic-version': '2023-06-01',
          },
          body: JSON.stringify({
            model:      'claude-sonnet-4-20250514',
            max_tokens: modo === 'simple' ? 150 : 300,
            messages:   [{ role: 'user', content: contentParts }],
          }),
        });

        const d     = await resp.json();
        const texto = d.content?.find(c => c.type === 'text')?.text || '{}';
        return json({ ok: true, texto }, 200, CORS);
      }

      // ══════════════════════════════════════════
      //  📧 ENVIAR OTP
      // ══════════════════════════════════════════
      if (ruta === '/enviar_otp') {
        const { ruc, email, contacto } = body;
        if (!ruc || !email) return json({ ok: false, error: 'RUC y email requeridos' }, 400, CORS);

        // Generar OTP de 6 dígitos y guardarlo en KV o en memoria con TTL
        const otp    = String(Math.floor(100000 + Math.random() * 900000));
        const expiry = Date.now() + 10 * 60 * 1000; // 10 minutos

        // Guardar OTP en KV (necesitas crear un KV namespace "OTP_STORE" en Cloudflare)
        if (env.OTP_STORE) {
          await env.OTP_STORE.put(`${ruc}:${email}`, JSON.stringify({ otp, expiry }), { expirationTtl: 600 });
        }

        // Enviar email con Resend
        const emailBody = {
          from:    env.EMAIL_FROM || 'Portal Beta <noreply@betaagroindustrial.com>',
          to:      [email],
          subject: `Tu código de acceso — Portal Proveedores Beta`,
          html: `
            <div style="font-family:'Outfit',Arial,sans-serif;max-width:480px;margin:0 auto;background:#F0EDF8;border-radius:16px;overflow:hidden">
              <div style="background:#5212A0;padding:1.5rem 2rem;border-bottom:3px solid #5BAD1E">
                <span style="font-size:26px;font-weight:800;color:#fff">bet<span style="color:#79CC38">a</span></span>
                <p style="color:rgba(255,255,255,.7);font-size:12px;margin-top:4px">Complejo Agroindustrial</p>
              </div>
              <div style="padding:2rem">
                <p style="font-size:15px;color:#1C0840;font-weight:600;margin-bottom:.5rem">Hola, ${contacto || 'Proveedor'} 👋</p>
                <p style="font-size:13px;color:#5C4880;margin-bottom:1.5rem">Tu código de verificación para acceder al Portal de Proveedores es:</p>
                <div style="background:#fff;border:2px solid #5212A0;border-radius:12px;padding:1.5rem;text-align:center;margin-bottom:1.5rem">
                  <span style="font-size:42px;font-weight:900;letter-spacing:10px;color:#5212A0;font-family:monospace">${otp}</span>
                </div>
                <p style="font-size:12px;color:#9B89B8">⏱ Válido por <strong>10 minutos</strong>. No compartas este código con nadie.</p>
              </div>
              <div style="background:#1C0840;padding:1rem 2rem;text-align:center">
                <p style="color:rgba(255,255,255,.4);font-size:10px">Portal Proveedores v4.0 · Complejo Agroindustrial Beta S.A.</p>
              </div>
            </div>`,
        };

        if (env.RESEND_API_KEY) {
          const res = await fetch('https://api.resend.com/emails', {
            method:  'POST',
            headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${env.RESEND_API_KEY}` },
            body:    JSON.stringify(emailBody),
          });
          if (!res.ok) {
            const err = await res.text();
            return json({ ok: false, error: 'Error enviando email: ' + err }, 500, CORS);
          }
        } else {
          // Sin servicio de email configurado — log para debug
          console.log('OTP generado (sin email configurado):', { ruc, email, otp });
        }

        return json({ ok: true }, 200, CORS);
      }

      // ══════════════════════════════════════════
      //  ✅ VERIFICAR OTP
      // ══════════════════════════════════════════
      if (ruta === '/verificar_otp') {
        const { ruc, email, otp } = body;
        if (!ruc || !email || !otp) return json({ ok: false, error: 'Datos incompletos' }, 400, CORS);

        if (!env.OTP_STORE) {
          // Sin KV configurado — modo permisivo para desarrollo
          console.warn('OTP_STORE no configurado — verificación omitida');
          return json({ ok: true, valido: true }, 200, CORS);
        }

        const stored = await env.OTP_STORE.get(`${ruc}:${email}`);
        if (!stored) return json({ ok: true, valido: false, error: 'OTP expirado o no encontrado' }, 200, CORS);

        const { otp: otpGuardado, expiry } = JSON.parse(stored);

        if (Date.now() > expiry) {
          await env.OTP_STORE.delete(`${ruc}:${email}`);
          return json({ ok: true, valido: false, error: 'OTP expirado' }, 200, CORS);
        }

        if (otp !== otpGuardado) {
          return json({ ok: true, valido: false, error: 'OTP incorrecto' }, 200, CORS);
        }

        // OTP válido — eliminar para que no se reutilice
        await env.OTP_STORE.delete(`${ruc}:${email}`);
        return json({ ok: true, valido: true }, 200, CORS);
      }

      // ══════════════════════════════════════════
      //  🔍 SUNAT RUC
      // ══════════════════════════════════════════
      if (ruta === '/sunat_ruc') {
        const { ruc, token } = body;
        if (!ruc) return json({ ok: false, error: 'RUC requerido' }, 400, CORS);

        const apiToken = env.APIS_NET_PE_TOKEN || token || '';
        const res = await fetch(`https://api.apis.net.pe/v2/sunat/ruc?numero=${ruc}`, {
          headers: { Authorization: `Bearer ${apiToken}`, Accept: 'application/json' },
        });

        if (!res.ok) return json({ ok: false, error: 'RUC no encontrado en SUNAT' }, 404, CORS);
        const data = await res.json();
        return json({ ok: true, ...data }, 200, CORS);
      }

      // ══════════════════════════════════════════
      //  📋 LISTAR PROVEEDORES (desde Google Sheets)
      // ══════════════════════════════════════════
      if (ruta === '/listar') {
        const { analista } = body;
        const gsheetUrl = env.GSHEET_URL;
        if (!gsheetUrl) return json({ ok: false, error: 'GSHEET_URL no configurado' }, 500, CORS);

        const params = new URLSearchParams({ action: 'listar', analista: analista || 'TODOS' });
        const res = await fetch(`${gsheetUrl}?${params}`);
        const data = await res.json();
        return json(data, 200, CORS);
      }

      // ══════════════════════════════════════════
      //  💾 ACTUALIZAR SHEET (aprobar, rechazar, notas, guardarAI)
      // ══════════════════════════════════════════
      if (ruta === '/sheet') {
        const gsheetUrl = env.GSHEET_URL;
        if (!gsheetUrl) return json({ ok: false, error: 'GSHEET_URL no configurado' }, 500, CORS);

        // Usamos no-cors desde el Worker no aplica — aquí podemos leer la respuesta
        const res = await fetch(gsheetUrl, {
          method:  'POST',
          headers: { 'Content-Type': 'application/json' },
          body:    JSON.stringify(body),
        });
        try {
          const data = await res.json();
          return json(data, 200, CORS);
        } catch {
          return json({ ok: true }, 200, CORS);
        }
      }

      // ══════════════════════════════════════════
      //  📤 SUBIR DOCUMENTOS a SharePoint
      // ══════════════════════════════════════════
      if (ruta === '/subir_docs') {
        const flowUrl = env.FLOW_SUBIR_DOCS;
        if (!flowUrl) return json({ ok: false, error: 'FLOW_SUBIR_DOCS no configurado' }, 500, CORS);

        const ctrl   = new AbortController();
        const tid    = setTimeout(() => ctrl.abort(), 120000); // 2 min por archivo
        try {
          const res = await fetch(flowUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify(body),
            signal:  ctrl.signal,
          });
          clearTimeout(tid);
          // Power Automate puede devolver 202 Accepted — ambos son OK
          if (res.status === 202 || res.ok) return json({ ok: true }, 200, CORS);
          return json({ ok: false, error: `Flow respondió ${res.status}` }, res.status, CORS);
        } catch(e) {
          clearTimeout(tid);
          if (e.name === 'AbortError') return json({ ok: true }, 200, CORS); // Timeout = flow aceptó
          throw e;
        }
      }

      // ══════════════════════════════════════════
      //  📁 CREAR CARPETAS en SharePoint
      // ══════════════════════════════════════════
      if (ruta === '/crear_carpetas') {
        const { flowUrl, ...payload } = body;
        const targetUrl = flowUrl || env.FLOW_CREAR_CARPETAS;
        if (!targetUrl) return json({ ok: false, error: 'URL del flow no configurada' }, 500, CORS);

        const ctrl = new AbortController();
        const tid  = setTimeout(() => ctrl.abort(), 30000);
        try {
          const res = await fetch(targetUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify(payload),
            signal:  ctrl.signal,
          });
          clearTimeout(tid);
          if (res.status === 202 || res.ok) return json({ ok: true }, 200, CORS);
          return json({ ok: false, error: `Flow respondió ${res.status}` }, res.status, CORS);
        } catch(e) {
          clearTimeout(tid);
          if (e.name === 'AbortError') return json({ ok: true }, 200, CORS);
          throw e;
        }
      }

      // ══════════════════════════════════════════
      //  🔔 NOTIFICAR PROVEEDOR por correo
      // ══════════════════════════════════════════
      if (ruta === '/notificar') {
        const { email, contacto, razonSocial, ruc, mensaje, linkIngresar, motivo, tipo } = body;
        if (!email) return json({ ok: false, error: 'Email requerido' }, 400, CORS);

        const esAprobacion = tipo === 'aprobacion';
        const color        = esAprobacion ? '#5BAD1E' : '#E02020';
        const titulo       = esAprobacion ? '¡Tu solicitud fue aprobada! 🎉' : 'Actualización sobre tu solicitud';
        const cuerpo       = esAprobacion
          ? `<p style="font-size:14px;color:#1C0840">Hola <strong>${contacto}</strong>,</p>
             <p style="font-size:13px;color:#5C4880;line-height:1.6;margin:1rem 0">
               Tu empresa <strong>${razonSocial}</strong> (RUC: ${ruc}) ha sido <strong style="color:#5BAD1E">aprobada</strong>
               como proveedor de <strong>Complejo Agroindustrial Beta</strong>.
             </p>
             ${mensaje ? `<p style="font-size:13px;color:#5C4880;font-style:italic;margin-bottom:1rem">💬 Mensaje del analista: "${mensaje}"</p>` : ''}
             <div style="text-align:center;margin:1.5rem 0">
               <a href="${linkIngresar}" style="background:#5212A0;color:#fff;padding:12px 28px;border-radius:10px;font-weight:700;font-size:14px;text-decoration:none;display:inline-block">
                 Acceder al Portal de Documentos →
               </a>
             </div>
             <p style="font-size:11.5px;color:#9B89B8">Usa tu RUC y este correo para ingresar.</p>`
          : `<p style="font-size:14px;color:#1C0840">Hola <strong>${contacto}</strong>,</p>
             <p style="font-size:13px;color:#5C4880;line-height:1.6;margin:1rem 0">
               Lamentamos informarte que la solicitud de <strong>${razonSocial}</strong> (RUC: ${ruc})
               no ha podido ser aprobada en esta oportunidad.
             </p>
             ${motivo ? `<div style="background:#FEE8E8;border:1px solid #E02020;border-radius:10px;padding:1rem;margin-bottom:1rem;font-size:13px;color:#E02020"><strong>Motivo:</strong> ${motivo}</div>` : ''}
             <p style="font-size:13px;color:#5C4880">Por favor contáctate con tu analista para más información.</p>`;

        const emailPayload = {
          from:    env.EMAIL_FROM || 'Portal Beta <noreply@betaagroindustrial.com>',
          to:      [email],
          subject: titulo,
          html: `
            <div style="font-family:'Outfit',Arial,sans-serif;max-width:520px;margin:0 auto;background:#F0EDF8;border-radius:16px;overflow:hidden">
              <div style="background:#5212A0;padding:1.5rem 2rem;border-bottom:3px solid ${color}">
                <span style="font-size:26px;font-weight:800;color:#fff">bet<span style="color:#79CC38">a</span></span>
                <p style="color:rgba(255,255,255,.7);font-size:12px;margin-top:4px">Complejo Agroindustrial</p>
              </div>
              <div style="padding:2rem">${cuerpo}</div>
              <div style="background:#1C0840;padding:1rem 2rem;text-align:center">
                <p style="color:rgba(255,255,255,.4);font-size:10px">Portal Proveedores v4.0 · Complejo Agroindustrial Beta S.A.</p>
              </div>
            </div>`,
        };

        if (env.RESEND_API_KEY) {
          const res = await fetch('https://api.resend.com/emails', {
            method:  'POST',
            headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${env.RESEND_API_KEY}` },
            body:    JSON.stringify(emailPayload),
          });
          if (!res.ok) {
            const err = await res.text();
            return json({ ok: false, error: 'Error enviando notificación: ' + err }, 500, CORS);
          }
        } else {
          console.log('NOTIFICAR (sin RESEND_API_KEY):', { email, tipo, titulo });
        }

        return json({ ok: true }, 200, CORS);
      }

      // ══════════════════════════════════════════
      //  404 — Ruta no encontrada
      // ══════════════════════════════════════════
      return json({ ok: false, error: `Ruta no encontrada: ${ruta}` }, 404, CORS);

    } catch(e) {
      console.error('Worker error:', e);
      return json({ ok: false, error: 'Error interno: ' + e.message }, 500, CORS);
    }
  }
};

/** Helper: respuesta JSON con headers CORS */
function json(data, status = 200, cors = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...cors },
  });
}
