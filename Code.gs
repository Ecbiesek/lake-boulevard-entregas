// ============================================================
// Lake Boulevard — Agendamento de Entrega
// Google Apps Script — Cole este código em script.google.com
// ============================================================

var CALENDAR_ID   = 'c_efdcb1f03e36922ef053c4a5288afbac407fa5b2b396dca4d77da9c9479a3611@group.calendar.google.com';
var NOTIFY_EMAIL  = 'engenharia@ecbiesek.com';
var LOCK_HOURS    = 24;

// ------------------------------------------------------------
// Endpoint principal
// ------------------------------------------------------------
function doGet(e) {
  var p = e.parameter;
  var result = {};

  if (p.action === 'reserve') {
    result = reserveSlot(p.iCalUID, p.start, p.name, p.unit);
  } else if (p.action === 'assign') {
    result = assignSlot(p.start, p.unit, p.token);
  } else if (p.action === 'release') {
    result = releaseSlot(p.start, p.token);
  } else if (p.action === 'clearBefore') {
    result = clearSlotsBeforeDate(p.cutoff, p.token);
  }

  var json = JSON.stringify(result);
  var out;
  if (p.callback) {
    out = ContentService.createTextOutput(p.callback + '(' + json + ')');
    out.setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    out = ContentService.createTextOutput(json);
    out.setMimeType(ContentService.MimeType.JSON);
  }
  return out;
}

// ------------------------------------------------------------
// Reserva um slot (busca por horário de início)
// ------------------------------------------------------------
function reserveSlot(iCalUID, startISO, clientName, clientUnit) {
  try {
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    var startTime = new Date(startISO);
    var endSearch = new Date(startTime.getTime() + 60000);
    var events = cal.getEvents(startTime, endSearch);
    var event = null;
    for (var i = 0; i < events.length; i++) {
      if (events[i].getTitle() === 'SLOT DISPONÍVEL') { event = events[i]; break; }
    }
    if (!event) return { success: false, error: 'Este horário já foi reservado por outro cliente. Escolha outro slot.' };

    var slotDate = Utilities.formatDate(startTime, 'America/Porto_Velho', 'dd/MM/yyyy');
    var slotTime = Utilities.formatDate(startTime, 'America/Porto_Velho', 'HH:mm');

    event.setTitle('⏳ ' + clientName + ' · ' + clientUnit);
    event.setColor(CalendarApp.EventColor.YELLOW);
    event.setDescription(
      'Cliente: '   + clientName  + '\n' +
      'Unidade: '   + clientUnit  + '\n' +
      'Reservado: ' + new Date().toLocaleString('pt-BR') + '\n' +
      'Status: AGUARDANDO\n' +
      'Timestamp: ' + Date.now()
    );

    sendEmail(
      '🏠 Nova solicitação de entrega — Lake Boulevard',
      'Cliente: '  + clientName + '\n' +
      'Unidade: '  + clientUnit + '\n' +
      'Data: '     + slotDate   + '\n' +
      'Horário: '  + slotTime   + '\n\n' +
      'Cliente manifestou interesse. Entre em contato pelo WhatsApp para confirmar. O slot será liberado automaticamente em 24h se não confirmado.',
      '<h3 style="color:#1a3a5c">🏠 Nova solicitação de entrega — Lake Boulevard</h3>' +
      '<table style="font-family:sans-serif;font-size:14px;border-collapse:collapse">' +
      '<tr><td style="padding:6px 16px 6px 0;color:#888">Cliente</td><td style="padding:6px 0"><strong>' + clientName + '</strong></td></tr>' +
      '<tr><td style="padding:6px 16px 6px 0;color:#888">Unidade</td><td style="padding:6px 0"><strong>' + clientUnit + '</strong></td></tr>' +
      '<tr><td style="padding:6px 16px 6px 0;color:#888">Data</td><td style="padding:6px 0">' + slotDate + '</td></tr>' +
      '<tr><td style="padding:6px 16px 6px 0;color:#888">Horário</td><td style="padding:6px 0">' + slotTime + '</td></tr>' +
      '</table>' +
      '<p style="margin-top:16px;font-size:13px;color:#888">Cliente manifestou interesse. Entre em contato pelo WhatsApp para confirmar. O slot será liberado automaticamente em 24h se não confirmado.</p>'
    );

    return { success: true, date: slotDate, time: slotTime };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ------------------------------------------------------------
// Admin: atribui slot diretamente a uma unidade
// ------------------------------------------------------------
var ADMIN_TOKEN = 'lcr2026';

function assignSlot(startISO, unit, token) {
  if (token !== ADMIN_TOKEN) return { success: false, error: 'Token inválido.' };
  try {
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    var startTime = new Date(startISO);
    var endSearch = new Date(startTime.getTime() + 60000);
    var events = cal.getEvents(startTime, endSearch);
    if (!events.length) return { success: false, error: 'Slot não encontrado nesse horário.' };
    var ev = events[0];
    ev.setTitle(unit);
    ev.setColor(CalendarApp.EventColor.GREEN);
    ev.setDescription('Atribuído manualmente em ' + new Date().toLocaleString('pt-BR'));
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ------------------------------------------------------------
// Admin: libera slot manualmente
// ------------------------------------------------------------
function releaseSlot(startISO, token) {
  if (token !== ADMIN_TOKEN) return { success: false, error: 'Token inválido.' };
  try {
    var cal = CalendarApp.getCalendarById(CALENDAR_ID);
    var startTime = new Date(startISO);
    var endSearch = new Date(startTime.getTime() + 60000);
    var events = cal.getEvents(startTime, endSearch);
    if (!events.length) return { success: false, error: 'Slot não encontrado.' };
    var ev = events[0];
    ev.setTitle('SLOT DISPONÍVEL');
    ev.setColor(CalendarApp.EventColor.CYAN);
    ev.setDescription('');
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ------------------------------------------------------------
// Admin: apaga slots disponíveis antes de uma data de corte
// ------------------------------------------------------------
function clearSlotsBeforeDate(cutoffISO, token) {
  if (token !== ADMIN_TOKEN) return { success: false, error: 'Token inválido.' };
  try {
    var cal     = CalendarApp.getCalendarById(CALENDAR_ID);
    var past    = new Date('2026-06-01T00:00:00-04:00');
    var cutoff  = new Date(cutoffISO);
    var events  = cal.getEvents(past, cutoff);
    var removed = 0;
    for (var i = 0; i < events.length; i++) {
      if (events[i].getTitle() === 'SLOT DISPONÍVEL') {
        events[i].deleteEvent();
        removed++;
      }
    }
    return { success: true, removed: removed };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ------------------------------------------------------------
// Envio de e-mail
// ------------------------------------------------------------
function sendEmail(subject, body, htmlBody) {
  MailApp.sendEmail({
    to:       NOTIFY_EMAIL,
    subject:  subject,
    body:     body,
    htmlBody: htmlBody
  });
}

// ------------------------------------------------------------
// Trigger: libera slots expirados (roda a cada 1 hora)
// ------------------------------------------------------------
function releaseExpiredSlots() {
  var cal    = CalendarApp.getCalendarById(CALENDAR_ID);
  var now    = new Date();
  var future = new Date(now.getTime() + 30 * 24 * 3600000);
  var events = cal.getEvents(now, future);

  for (var i = 0; i < events.length; i++) {
    var ev = events[i];
    if (ev.getTitle().indexOf('⏳') !== 0) continue;

    var desc  = ev.getDescription() || '';
    var match = desc.match(/Timestamp: (\d+)/);
    if (!match) continue;

    var hoursPassed = (Date.now() - parseInt(match[1])) / 3600000;
    if (hoursPassed < LOCK_HOURS) continue;

    var clientInfo = ev.getTitle().replace('⏳ ', '');
    var d = Utilities.formatDate(ev.getStartTime(), 'America/Porto_Velho', 'dd/MM/yyyy');
    var t = Utilities.formatDate(ev.getStartTime(), 'America/Porto_Velho', 'HH:mm');

    ev.setTitle('SLOT DISPONÍVEL');
    ev.setColor(CalendarApp.EventColor.CYAN);
    ev.setDescription('');

    sendEmail(
      '⚠️ Slot liberado automaticamente — Lake Boulevard',
      'Horário: ' + d + ' às ' + t + '\nCliente: ' + clientInfo + '\nMotivo: 24h sem confirmação pelo WhatsApp.',
      '<h3 style="color:#b45309">⚠️ Slot liberado automaticamente — Lake Boulevard</h3>' +
      '<table style="font-family:sans-serif;font-size:14px;border-collapse:collapse">' +
      '<tr><td style="padding:6px 16px 6px 0;color:#888">Horário</td><td>' + d + ' às ' + t + '</td></tr>' +
      '<tr><td style="padding:6px 16px 6px 0;color:#888">Cliente</td><td>' + clientInfo + '</td></tr>' +
      '<tr><td style="padding:6px 16px 6px 0;color:#888">Motivo</td><td>24h sem confirmação pelo WhatsApp</td></tr>' +
      '</table>'
    );
  }
}
