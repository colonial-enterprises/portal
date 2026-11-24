/**
 * Colonial · Recepción de pedidos del portal de compras DHL
 * ---------------------------------------------------------
 * Recibe cada pedido enviado desde el portal, lo guarda como una fila en esta hoja
 * de Google Sheets y envía un correo con el detalle a la persona dueña de la cuenta
 * (y a los correos extra que se configuren). Además envía una confirmación al
 * comprador, solo si su correo es de un dominio permitido (por defecto dhl.com).
 *
 * Instalación (una sola vez):
 *   1. Abra la hoja de Google Sheets "Pedidos portal DHL" > Extensiones > Apps Script.
 *   2. Pegue este archivo completo en Code.gs y guarde.
 *   3. Elija la función "configurar" y presione Ejecutar. Autorice los permisos.
 *   4. Implementar > Nueva implementación > Tipo: Aplicación web.
 *        Ejecutar como: Yo.  Quién tiene acceso: Cualquier usuario.
 *   5. Copie la URL de la aplicación web (termina en /exec) y compártala con Claude.
 */

// Correos adicionales que reciben cada pedido (además de la dueña de la cuenta).
// Ejemplo: const EXTRA_RECIPIENTS = ['ventas@colonialenterprise.com'];
const EXTRA_RECIPIENTS = [];

// Solo se envía confirmación al comprador si su correo termina en uno de estos dominios.
const CONFIRM_DOMAINS = ['dhl.com'];

const SHEET_NAME = 'Pedidos';
const HEADERS = ['Fecha', 'Referencia', 'Estado', 'Cliente', 'Nombre', 'Área', 'Correo', 'Teléfono',
  'Propósito', 'Fecha deseada', 'Lugar de entrega', 'Centro de costo', 'PO', 'Productos',
  'Subtotal ref. (US$)', 'Ítems a cotizar', 'Comentarios', 'Idioma'];

/** Ejecutar una vez desde el editor: prepara la hoja y pide los permisos. */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  hoja_();
  Logger.log('Listo. Los pedidos llegarán a: ' + destinatarios_().join(', '));
}

function doGet() {
  return json_({ ok: true, service: 'Colonial pedidos DHL' });
}

/**
 * Alternativa sin "Implementar": ejecutar una vez desde el editor.
 * Crea un formulario de Google que recibe los pedidos del portal, lo deja abierto
 * sin inicio de sesión y conecta el aviso automático. Al terminar muestra en el
 * registro la dirección que se debe compartir con Claude.
 */
function crearFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  hoja_();
  const form = FormApp.create('Pedidos portal DHL (no editar)');
  form.setDescription('Formulario técnico que recibe los pedidos del portal de compras. No lo edite ni lo comparta.');
  try { form.setRequireLogin(false); } catch (err) { Logger.log('Aviso: no se pudo quitar el inicio de sesión: ' + err); }
  form.setCollectEmail(false);
  const item = form.addParagraphTextItem().setTitle('pedido');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alRecibirFormulario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alRecibirFormulario').forForm(form).onFormSubmit().create();
  PropertiesService.getScriptProperties().setProperty('FORM_ID', form.getId());
  const url = form.getPublishedUrl().replace(/\/viewform.*$/, '/formResponse');
  Logger.log('COMPARTA ESTO CON CLAUDE:  ' + url + '  |  entry.' + item.getId());
}

/** Se ejecuta solo cada vez que llega un pedido por el formulario. */
function alRecibirFormulario(e) {
  const r = e.response.getItemResponses()[0];
  const raw = r ? String(r.getResponse() || '') : '';
  const res = procesarPedido_(raw);
  if (!res.ok) console.warn('Pedido rechazado: ' + res.error);
}

function doPost(e) {
  return json_(procesarPedido_((e && e.postData && e.postData.contents) || ''));
}

function procesarPedido_(raw) {
  try {
    if (raw.length > 50000) return { ok: false, error: 'too_large' };
    const d = JSON.parse(raw);

    // Trampa anti-spam: los humanos nunca llenan este campo.
    if (d._gotcha) return { ok: true, ref: d.referencia || '' };

    const req = ['referencia', 'nombre', 'area', 'correo', 'telefono', 'proposito', 'fecha_entrega', 'lugar_entrega'];
    for (const k of req) if (!String(d[k] || '').trim()) return { ok: false, error: 'missing_' + k };
    const items = Array.isArray(d.items) ? d.items.slice(0, 200) : [];
    const custom = Array.isArray(d.custom) ? d.custom.slice(0, 50) : [];
    if (!items.length && !custom.length) return { ok: false, error: 'empty_order' };

    // Límite simple: máximo 10 pedidos por correo cada 10 minutos.
    const cache = CacheService.getScriptCache();
    const key = 'n_' + String(d.correo).toLowerCase();
    const n = Number(cache.get(key) || 0);
    if (n >= 10) return { ok: false, error: 'rate_limited' };
    cache.put(key, String(n + 1), 600);

    const s = v => String(v == null ? '' : v).slice(0, 500);
    const priced = items.filter(i => typeof i.price === 'number');
    const subtotal = priced.reduce((t, i) => t + i.price * (Number(i.qty) || 0), 0);
    const toQuote = items.length - priced.length + custom.length;
    const lines = items.map(i => `${Number(i.qty) || 0} × ${s(i.name)}` + (typeof i.price === 'number' ? ` · US$ ${i.price.toFixed(2)} ${s(i.unit)}` : ' · A cotizar'))
      .concat(custom.map(c => `${Number(c.qty) || 0} × [Especial] ${s(c.text)}`));

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      hoja_().appendRow([new Date(), s(d.referencia), 'Nuevo', s(d.cliente), s(d.nombre), s(d.area), s(d.correo),
        s(d.telefono), s(d.proposito), s(d.fecha_entrega), s(d.lugar_entrega), s(d.centro_costo), s(d.po),
        lines.join('\n'), Math.round(subtotal * 100) / 100, toQuote, s(d.comentarios), s(d.idioma)]);
    } finally { lock.releaseLock(); }

    enviarAviso_(d, items, custom, subtotal, toQuote);
    enviarConfirmacion_(d, items, custom, subtotal, toQuote);
    return { ok: true, ref: s(d.referencia) };
  } catch (err) {
    console.error(err);
    return { ok: false, error: 'server_error' };
  }
}

/* ---------------- Correos ---------------- */

function enviarAviso_(d, items, custom, subtotal, toQuote) {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  const html = plantilla_({
    titulo: 'Nuevo pedido ' + esc_(d.referencia),
    intro: `${esc_(d.nombre)} (${esc_(d.area)}) envió un pedido desde el portal de compras. Responda este correo para escribirle directamente.`,
    d, items, custom, subtotal, toQuote,
    pie: `<a href="${ss.getUrl()}" style="color:#1F656F">Ver todos los pedidos en la hoja de Google</a>`
  });
  MailApp.sendEmail({
    to: destinatarios_().join(','),
    replyTo: validEmail_(d.correo) ? d.correo : '',
    name: 'Portal de compras Colonial',
    subject: `Pedido DHL ${d.referencia}: ${d.area} · ${d.nombre}`,
    htmlBody: html
  });
}

function enviarConfirmacion_(d, items, custom, subtotal, toQuote) {
  const email = String(d.correo || '').trim().toLowerCase();
  if (!validEmail_(email)) return;
  const domain = email.split('@')[1];
  if (!CONFIRM_DOMAINS.some(dom => domain === dom || domain.endsWith('.' + dom))) return;
  const en = d.idioma === 'en';
  const html = plantilla_({
    titulo: (en ? 'We received your order ' : 'Recibimos su pedido ') + esc_(d.referencia),
    intro: en
      ? 'Thank you. We will send you the formal proposal for Basware in under 12 business hours. This is a summary of your request.'
      : 'Gracias. Le enviaremos la propuesta formal para Basware en menos de 12 horas hábiles. Este es el resumen de su solicitud.',
    d, items, custom, subtotal, toQuote, en,
    pie: en ? 'Reply to this email if you need to change anything.' : 'Si necesita cambiar algo, responda este correo.'
  });
  MailApp.sendEmail({
    to: email,
    name: 'Colonial',
    subject: (en ? 'Order received ' : 'Pedido recibido ') + d.referencia + ' · Colonial',
    htmlBody: html
  });
}

function plantilla_(o) {
  const en = !!o.en;
  const L = en
    ? { qty: 'Qty', item: 'Description', price: 'Ref. price', quote: 'Quote on request', special: 'Special request', sub: 'Reference subtotal (before ITBMS)', toq: 'item(s) to be quoted' }
    : { qty: 'Cant.', item: 'Descripción', price: 'Precio ref.', quote: 'A cotizar', special: 'Solicitud especial', sub: 'Subtotal de referencia (antes de ITBMS)', toq: 'ítem(s) a cotizar' };
  const td = 'padding:8px;border-bottom:1px solid #BFDDE1;font-size:14px;vertical-align:top';
  const rows = o.items.map(i => `<tr><td style="${td}">${Number(i.qty) || 0}</td><td style="${td}">${esc_(i.name)}</td><td style="${td};text-align:right;white-space:nowrap">${typeof i.price === 'number' ? 'US$ ' + i.price.toFixed(2) : L.quote}</td></tr>`).join('')
    + o.custom.map(c => `<tr><td style="${td}">${Number(c.qty) || 0}</td><td style="${td}">${L.special}: ${esc_(c.text)}</td><td style="${td};text-align:right">${L.quote}</td></tr>`).join('');
  const f = [['Nombre', 'Name', 'nombre'], ['Área', 'Area', 'area'], ['Correo', 'Email', 'correo'], ['Teléfono', 'Phone', 'telefono'],
    ['Propósito', 'Purpose', 'proposito'], ['Fecha deseada', 'Desired date', 'fecha_entrega'], ['Lugar de entrega', 'Delivery location', 'lugar_entrega'],
    ['Centro de costo', 'Cost center', 'centro_costo'], ['PO', 'PO', 'po'], ['Comentarios', 'Comments', 'comentarios']]
    .filter(x => String(o.d[x[2]] || '').trim())
    .map(x => `<tr><td style="padding:4px 12px 4px 0;font-weight:bold;font-size:14px;vertical-align:top">${en ? x[1] : x[0]}</td><td style="padding:4px 0;font-size:14px">${esc_(o.d[x[2]])}</td></tr>`).join('');
  return `<div style="font-family:Arial,Helvetica,sans-serif;color:#222C49;max-width:640px">
    <div style="border-top:3px solid #46999F;padding-top:16px">
      <div style="font-size:12px;letter-spacing:.1em;text-transform:uppercase;color:#1F656F;font-weight:bold">Colonial · ${esc_(o.d.cliente || 'Portal de compras')}</div>
      <h2 style="margin:8px 0 12px;font-size:22px">${o.titulo}</h2>
      <p style="font-size:14px;line-height:1.5">${o.intro}</p>
      <table style="border-collapse:collapse;margin:12px 0">${f}</table>
      <table style="border-collapse:collapse;width:100%;margin-top:8px">
        <tr><th style="text-align:left;padding:8px;font-size:12px;border-bottom:2px solid #46999F">${L.qty}</th><th style="text-align:left;padding:8px;font-size:12px;border-bottom:2px solid #46999F">${L.item}</th><th style="text-align:right;padding:8px;font-size:12px;border-bottom:2px solid #46999F">${L.price}</th></tr>
        ${rows}
      </table>
      <p style="font-size:15px;margin:12px 0 4px"><strong>${L.sub}: US$ ${o.subtotal.toFixed(2)}</strong></p>
      ${o.toQuote ? `<p style="font-size:13px;color:#1F656F;margin:0">+ ${o.toQuote} ${L.toq}</p>` : ''}
      <p style="font-size:13px;color:#1F656F;margin-top:20px">${o.pie}</p>
    </div></div>`;
}

/* ---------------- Utilidades ---------------- */

function hoja_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E0F1F3');
    sh.setColumnWidth(14, 380);
  }
  return sh;
}

function destinatarios_() {
  const owner = Session.getEffectiveUser().getEmail();
  return [owner].concat(EXTRA_RECIPIENTS).filter((v, i, a) => v && a.indexOf(v) === i);
}

function validEmail_(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e || '').trim()); }
function esc_(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
