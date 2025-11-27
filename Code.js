/**
 * Code.gs - Backend V4 (Final Consolidado - Automatizado)
 */
const APP_VERSION = 'v4.1-Fixes';

// === RUTAS E INICIO ===

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('index');
  const user = getUserInfo_();

  // Inyectar firma
  t.signatureUrl = getSignatureDataUrl_();

  if (!user || user.estado !== 'ACTIVO') {
    const denied = HtmlService.createTemplateFromFile('Denied');
    denied.signatureUrl = t.signatureUrl;
    return denied.evaluate()
      .setTitle('Acceso Denegado')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  t.user = user;
  t.appVersion = APP_VERSION;

  return t.evaluate()
    .setTitle('Solicitud Almuerzo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// === API PÚBLICA ===

function apiGetInitData(requestedDateStr) {
  try {
    const user = getUserInfo_();
    if (!user) throw new Error("Usuario no encontrado.");

    const availableDates = getAvailableMenuDates_(true);

    if (availableDates.length === 0) {
      return { ok: true, empty: true, msg: "No hay menús disponibles." };
    }

    let targetDateStr = requestedDateStr;
    if (!targetDateStr || !availableDates.some(d => d.value === targetDateStr)) {
      targetDateStr = availableDates[0].value;
    }

    const existingOrder = getOrderByUserDate_(user.email, targetDateStr);
    const menu = getMenuByDate_(targetDateStr);
    const cutoffPassed = !isDateOpenForOrdering_(targetDateStr);

    let adminSummary = null;
    if (user.rol === 'ADMIN_GEN' || user.rol === 'ADMIN_DEP') {
      adminSummary = getDepartmentStats_(targetDateStr, (user.rol === 'ADMIN_GEN' ? null : user.departamento));
    }

    return {
      ok: true,
      user: user,
      currentDate: targetDateStr,
      dates: availableDates,
      menu: menu,
      myOrder: existingOrder,
      adminData: adminSummary,
      isCutoffPassed: cutoffPassed // Flag para el frontend
    };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSubmitOrder(payload) {
  try {
    const user = getUserInfo_();
    const dateStr = payload.date;

    if (!isDateOpenForOrdering_(dateStr)) {
      throw new Error("El tiempo límite para pedir el almuerzo de esta fecha ha expirado.");
    }

    validateOrderRules_(payload);
    saveOrderToSheet_(user, dateStr, payload);

    return { ok: true };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiCancelOrder(orderId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Pedidos');
  const rows = sh.getDataRange().getValues();
  const user = getUserInfo_();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(orderId) && String(rows[i][3]).toLowerCase() === user.email.toLowerCase()) {
      const orderDate = formatDate_(new Date(rows[i][2]));
      if (!isDateOpenForOrdering_(orderDate)) {
        return { ok: false, msg: "Ya no puedes cancelar este pedido (hora de cierre pasada)." };
      }
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, msg: "Pedido no encontrado." };
}

// === AUTOMATIZACIÓN (TRIGGERS) ===

function scheduledSendReminders() {
  const nextBusinessDay = getNextBusinessDay_(new Date());
  if (!nextBusinessDay) return;
  const dateStr = formatDate_(nextBusinessDay);

  const ss = SpreadsheetApp.getActive();
  const uSh = ss.getSheetByName('Usuarios');
  const uData = uSh.getDataRange().getValues();
  const pSh = ss.getSheetByName('Pedidos');
  const pData = pSh.getDataRange().getValues();

  const orderedEmails = new Set();
  for (let i = 1; i < pData.length; i++) {
    if (formatDate_(new Date(pData[i][2])) === dateStr && pData[i][8] !== 'CANCELADO') {
      orderedEmails.add(String(pData[i][3]).toLowerCase());
    }
  }

  uData.slice(1).forEach(row => {
    const email = String(row[0]).toLowerCase();
    const prefs = JSON.parse(row[5] || '{}');
    if (row[4] === 'ACTIVO' && !orderedEmails.has(email) && prefs.reminders !== false) {
      sendEmail_(email, "Recordatorio de Almuerzo",
        `Hola ${row[1]},<br><br>Aún no has realizado tu pedido de almuerzo para el <b>${formatDisplayDate_(dateStr)}</b>.<br>` +
        `Recuerda que tienes hasta las ${getCutoffTimeStr_()} para hacerlo.<br><br>` +
        `<a href="${ScriptApp.getService().getUrl()}">Ir a la App</a>`
      );
    }
  });
}

function scheduledLockSheet() { console.log("Bloqueo lógico activado."); }

function scheduledDailyClose() {
  const now = new Date();
  const nextBusinessDay = getNextBusinessDay_(now);
  if (!nextBusinessDay) return;
  const targetDateStr = formatDate_(nextBusinessDay);
  backupOrdersToDrive_(targetDateStr);
  sendDailyAdminSummary_(targetDateStr);
  checkMenuIntegrity_();
  updateSheetHeaders_(targetDateStr);
}

function scheduledDepartmentReports() {
  const now = new Date();
  const nextBusinessDay = getNextBusinessDay_(now);
  if (!nextBusinessDay) return;
  const dateStr = formatDate_(nextBusinessDay);
  const orders = getOrdersByDate_(dateStr);
  const byDept = {};
  orders.forEach(o => {
    const dept = o.departamento || 'Sin Depto';
    if (!byDept[dept]) byDept[dept] = [];
    byDept[dept].push(o);
  });

  let deptEmails = {};
  try { deptEmails = JSON.parse(getConfigValue_('RESPONSIBLES_EMAILS_JSON') || '{}'); } catch (e) {}

  Object.keys(byDept).forEach(dept => {
    const recipient = deptEmails[dept];
    if (recipient) {
      const listHtml = byDept[dept].map(o => `<li><b>${o.nombre}:</b> ${o.resumen}</li>`).join('');
      sendEmail_(recipient, `Reporte Almuerzo ${dept} - ${dateStr}`,
        `<h3>Pedidos para ${formatDisplayDate_(dateStr)}</h3>` +
        `<p>Total platos: ${byDept[dept].length}</p>` +
        `<ul>${listHtml}</ul>`
      );
    }
  });
}

// === CORE LÓGICA (Optimized) ===

function getAvailableMenuDates_(fetchAll) {
  const ss = SpreadsheetApp.getActive();
  const menuSh = ss.getSheetByName('Menu');
  const data = menuSh.getRange(2, 1, menuSh.getLastRow()-1, 2).getValues();
  const now = new Date();
  const todayStr = formatDate_(now);
  const datesSet = new Set();
  data.forEach(r => {
    const dStr = formatDate_(new Date(r[1]));
    if (dStr >= todayStr) datesSet.add(dStr);
  });

  const sorted = Array.from(datesSet).sort();
  const holidays = getHolidaysSet_();
  const valid = [];

  sorted.forEach(dStr => {
    if (isDateOpenForOrdering_(dStr, holidays)) {
      valid.push({ value: dStr, label: formatDisplayDate_(dStr) });
    }
  });
  return valid;
}

function isDateOpenForOrdering_(targetDateStr, holidaysSet) {
  if (!holidaysSet) holidaysSet = getHolidaysSet_();
  const now = new Date();
  const targetDate = new Date(targetDateStr + 'T12:00:00');

  if (targetDate <= now) return false;
  if ([0, 6].includes(targetDate.getDay())) return false;
  if (holidaysSet.has(targetDateStr)) return false;

  const prevBizDay = getPreviousBusinessDay_(targetDate, holidaysSet);
  const prevBizDayStr = formatDate_(prevBizDay);
  const todayStr = formatDate_(now);

  if (todayStr === prevBizDayStr) {
    const limit = getCutoffDateObj_(now);
    if (now > limit) return false;
  }

  const zeroNow = new Date(now); zeroNow.setHours(0,0,0,0);
  const zeroPrev = new Date(prevBizDay); zeroPrev.setHours(0,0,0,0);
  if (zeroNow > zeroPrev) return false;

  return true;
}

// === UTILS ===

// Helper seguro para obtener hora de cierre
function getCutoffTimeStr_() {
  const val = getConfigValue_('HORA_CIERRE');
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  return String(val || '14:30');
}

function getCutoffDateObj_(baseDate) {
  const timeStr = getCutoffTimeStr_(); // "14:30"
  const [h, m] = timeStr.split(':').map(Number);
  const d = new Date(baseDate);
  d.setHours(h, m, 0, 0);
  return d;
}

function getConfigValue_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return data[i][1];
  }
  return '';
}

// Versión optimizada con Caché para Holidays
function getHolidaysSet_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('HOLIDAYS_SET_V1');
  if (cached) {
    return new Set(JSON.parse(cached));
  }

  const set = new Set();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('DiasLibres');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues().forEach(r => {
      if (r[0]) set.add(formatDate_(new Date(r[0])));
    });
  }

  try {
    const calId = 'es.do#holiday@group.v.calendar.google.com';
    const now = new Date();
    const start = new Date(now.getTime() - 30 * 86400000);
    const end = new Date(now.getTime() + 365 * 86400000);
    // API Call lenta
    CalendarApp.getCalendarById(calId).getEvents(start, end).forEach(e => {
      set.add(formatDate_(e.getStartTime()));
    });
  } catch (e) {
    console.warn("Error fetching calendar: " + e.message);
  }

  // Guardar en caché por 6 horas
  cache.put('HOLIDAYS_SET_V1', JSON.stringify([...set]), 21600);
  return set;
}

function getNextBusinessDay_(date) {
  let d = new Date(date);
  d.setDate(d.getDate() + 1);
  const holidays = getHolidaysSet_();
  while ([0, 6].includes(d.getDay()) || holidays.has(formatDate_(d))) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

function getPreviousBusinessDay_(date, holidaysSet) {
  let d = new Date(date);
  do { d.setDate(d.getDate() - 1); }
  while ([0, 6].includes(d.getDay()) || holidaysSet.has(formatDate_(d)));
  return d;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDisplayDate_(dateStr) {
  const d = new Date(dateStr + 'T12:00:00');
  const days = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  return `${days[d.getDay()]} ${d.getDate()}/${d.getMonth() + 1}`;
}

// === EXTRAS ===

function backupOrdersToDrive_(dateStr) {
  const rootId = getConfigValue_('BACKUP_FOLDER_ID');
  if (!rootId) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const filtered = data.filter((row, i) => i === 0 || formatDate_(new Date(row[2])) === dateStr);
  if (filtered.length <= 1) return;

  try {
    const rootFolder = DriveApp.getFolderById(rootId);
    const d = new Date(dateStr + 'T12:00:00');
    const year = String(d.getFullYear());
    const month = String(d.getMonth() + 1).padStart(2, '0');

    let yFolder = rootFolder.getFoldersByName(year).hasNext() ? rootFolder.getFoldersByName(year).next() : rootFolder.createFolder(year);
    let mFolder = yFolder.getFoldersByName(month).hasNext() ? yFolder.getFoldersByName(month).next() : yFolder.createFolder(month);

    const tempSheet = SpreadsheetApp.create(`Pedidos_${dateStr}`);
    tempSheet.getSheets()[0].getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
    const tempFile = DriveApp.getFileById(tempSheet.getId());
    tempFile.moveTo(mFolder);
    mFolder.createFile(tempFile.getAs('application/pdf')).setName(`Pedidos_${dateStr}.pdf`);
  } catch (e) { console.error("Backup error: " + e.message); }
}

function sendDailyAdminSummary_(dateStr) {
  const admins = getConfigValue_('ADMIN_EMAILS');
  if (!admins) return;
  const orders = getOrdersByDate_(dateStr);
  if (orders.length > 0) {
    sendEmail_(admins, `Resumen Pedidos ${dateStr}`, `Se han registrado <b>${orders.length}</b> pedidos.`);
  }
}

function checkMenuIntegrity_() {
  const ss = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName('Menu');
  const data = mSh.getDataRange().getValues();
  const menuMap = {};
  for (let i = 1; i < data.length; i++) {
    const dStr = formatDate_(new Date(data[i][1]));
    if (dStr > formatDate_(new Date())) {
      if (!menuMap[dStr]) menuMap[dStr] = false;
      if (data[i][2] === 'Arroces' && String(data[i][3]).toLowerCase().includes('arroz blanco')) menuMap[dStr] = true;
    }
  }
  const warnings = Object.keys(menuMap).filter(d => !menuMap[d]);
  if (warnings.length > 0) {
    const admins = getConfigValue_('ADMIN_EMAILS');
    if (admins) sendEmail_(admins, "Alerta Menú", `Falta Arroz Blanco en: ${warnings.join(', ')}`);
  }
}

function updateSheetHeaders_(nextDateStr) {
  const ss = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName('Menu');
  if (mSh) mSh.getRange('A1').setNote(`Semana de pedido activo: ${nextDateStr}`);
}

function getOrdersByDate_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (formatDate_(new Date(data[i][2])) === dateStr && data[i][8] !== 'CANCELADO') {
      list.push({ nombre: data[i][4], departamento: data[i][5], resumen: data[i][6] });
    }
  }
  return list;
}

function getUserInfo_() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      return { email: data[i][0], nombre: data[i][1], departamento: data[i][2], rol: data[i][3], estado: data[i][4] };
    }
  }
  return null;
}

function getMenuByDate_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
  const data = sh.getDataRange().getValues();
  const menu = {};
  for (let i = 1; i < data.length; i++) {
    if (formatDate_(new Date(data[i][1])) === dateStr && String(data[i][5]).toUpperCase() === 'SI') {
      const cat = data[i][2];
      if (!menu[cat]) menu[cat] = [];
      menu[cat].push({ id: data[i][0], plato: data[i][3], desc: data[i][4] });
    }
  }
  return menu;
}

function getOrderByUserDate_(email, dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (formatDate_(new Date(data[i][2])) === dateStr && String(data[i][3]).toLowerCase() === email) {
      return { id: data[i][0], resumen: data[i][6], detalle: JSON.parse(data[i][7] || '{}') };
    }
  }
  return null;
}

function saveOrderToSheet_(user, dateStr, selection) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (formatDate_(new Date(data[i][2])) === dateStr && String(data[i][3]).toLowerCase() === user.email) {
      rowIdx = i + 1; break;
    }
  }
  const id = rowIdx > 0 ? data[rowIdx - 1][0] : Utilities.getUuid();
  const now = new Date();
  const rowData = [id, now, dateStr, user.email, user.nombre, user.departamento, selection.items.join(', '), JSON.stringify(selection), 'ACTIVO', now];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  else sh.appendRow(rowData);
}

function getDepartmentStats_(dateStr, deptFilter) {
  const orders = getOrdersByDate_(dateStr);
  const stats = { total: 0, byUser: [] };
  orders.forEach(o => {
    if (!deptFilter || o.departamento === deptFilter) {
      stats.total++;
      stats.byUser.push({ nombre: o.nombre, pedido: o.resumen, depto: o.departamento });
    }
  });
  return stats;
}

function validateOrderRules_(sel) {
  const cats = sel.categorias || [];
  const items = sel.items || [];
  const specialList = ['Vegetariana', 'Caldo', 'Opcion_Rapida'];
  if (cats.some(c => specialList.includes(c)) && cats.length > 1) {
    if ([...new Set(cats)].some(c => !specialList.includes(c))) throw new Error("Platos especiales no combinables.");
  }
  if (cats.includes('Granos') && !items.some(i => i.toLowerCase().includes('arroz blanco'))) throw new Error("Granos requieren Arroz Blanco.");
  if (cats.includes('Arroces') && cats.includes('Viveres')) throw new Error("No combinar Arroz y Víveres.");
}

function getSignatureDataUrl_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('SIG_V3_PROD');
  if (cached) return cached;
  const fileId = getConfigValue_('FOOTER_SIGNATURE_ID');
  if (!fileId) return '';
  try {
    const file = Drive.Files.get(fileId, { fields: 'thumbnailLink' });
    if (!file || !file.thumbnailLink) return '';
    const blob = UrlFetchApp.fetch(file.thumbnailLink.replace(/=s\d+$/, '=s300'), { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    const dataUrl = 'data:image/png;base64,' + Utilities.base64Encode(blob.getBytes());
    cache.put('SIG_V3_PROD', dataUrl, 21600);
    return dataUrl;
  } catch (e) { return ''; }
}