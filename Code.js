/**
 * Code.gs - Backend V4 (Final Consolidado - Automatizado)
 */
const APP_VERSION = 'v4.3.1-Cleanup';

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

/**
 * Obtiene datos para la app.
 * Devuelve TODAS las fechas futuras válidas con menú para permitir adelantado.
 */
function apiGetInitData(requestedDateStr) {
  try {
    const user = getUserInfo_();
    if (!user) throw new Error("Usuario no encontrado.");

    // Obtener TODAS las fechas futuras válidas con menú
    const availableDates = getAvailableMenuDates_(true); // true = fetchAll

    if (availableDates.length === 0) {
      return { ok: true, empty: true, msg: "No hay menús disponibles." };
    }

    // Fecha objetivo: la solicitada O la primera disponible (siguiente día hábil)
    let targetDateStr = requestedDateStr;
    if (!targetDateStr || !availableDates.some(d => d.value === targetDateStr)) {
      targetDateStr = availableDates[0].value;
    }

    // Optimization: Fetch ALL menus and orders
    const allMenus = getAllMenus_(availableDates);
    const allOrders = getAllUserOrders_(user.email, availableDates);

    const menu = allMenus[targetDateStr] || {};
    const existingOrder = allOrders[targetDateStr] || null;

    let adminSummary = null;
    if (user.rol === 'ADMIN_GEN' || user.rol === 'ADMIN_DEP') {
      adminSummary = getDepartmentStats_(targetDateStr, (user.rol === 'ADMIN_GEN' ? null : user.departamento));
    }

    // User preferences
    const prefs = getUserPrefs_(user.email);

    return {
      ok: true,
      user: user,
      userPrefs: prefs,
      currentDate: targetDateStr,
      dates: availableDates,
      menu: menu,
      allMenus: allMenus,
      allOrders: allOrders,
      myOrder: existingOrder,
      adminData: adminSummary
    };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSubmitOrder(payload) {
  try {
    const user = getUserInfo_();
    const dateStr = payload.date;

    // 1. Validar Bloqueo Global (Hora de Cierre o Bloqueo explícito)
    if (!isDateOpenForOrdering_(dateStr)) {
      throw new Error("El tiempo límite para pedir el almuerzo de esta fecha ha expirado.");
    }

    // 2. Validar Reglas de Menú
    validateOrderRules_(payload);

    // 3. Guardar
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

function apiSetUserPreference(key, value) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === userEmail) {
        const currentPrefs = JSON.parse(data[i][5] || '{}');
        currentPrefs[key] = value;
        sh.getRange(i + 1, 6).setValue(JSON.stringify(currentPrefs));
        return { ok: true };
      }
    }
    return { ok: false, msg: "Usuario no encontrado" };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// === AUTOMATIZACIÓN (TRIGGERS DIARIOS) ===

/**
 * 1:00 PM: Recordatorios
 */
function scheduledSendReminders() {
  const nextBusinessDay = getNextBusinessDay_(new Date());
  if (!nextBusinessDay) return; // No hay día hábil próximo configurado

  const dateStr = formatDate_(nextBusinessDay);

  // Lista de usuarios activos
  const ss = SpreadsheetApp.getActive();
  const uSh = ss.getSheetByName('Usuarios');
  const uData = uSh.getDataRange().getValues();
  const pSh = ss.getSheetByName('Pedidos');
  const pData = pSh.getDataRange().getValues();

  // Set de emails que YA pidieron para esa fecha
  const orderedEmails = new Set();
  for (let i = 1; i < pData.length; i++) {
    const rowDate = formatDate_(new Date(pData[i][2]));
    if (rowDate === dateStr && pData[i][8] !== 'CANCELADO') {
      orderedEmails.add(String(pData[i][3]).toLowerCase());
    }
  }

  // Recorrer usuarios y notificar
  uData.slice(1).forEach(row => {
    const email = String(row[0]).toLowerCase();
    const estado = row[4];
    const prefs = JSON.parse(row[5] || '{}');

    // Solo activos, que NO han pedido, y que tienen recordatorios ON (default ON)
    if (estado === 'ACTIVO' && !orderedEmails.has(email) && prefs.reminders !== false) {
      sendEmail_(email, "Recordatorio de Almuerzo",
        `Hola ${row[1]},<br><br>Aún no has realizado tu pedido de almuerzo para el <b>${formatDisplayDate_(dateStr)}</b>.<br>` +
        `Recuerda que tienes hasta las ${getConfigValue_('HORA_CIERRE')} para hacerlo.<br><br>` +
        `<a href="${ScriptApp.getService().getUrl()}">Ir a la App</a>`
      );
    }
  });
}

/**
 * 2:00 PM: Bloqueo de Hoja (Visual)
 * En realidad el bloqueo lógico está en isDateOpenForOrdering_ basado en HORA_CIERRE.
 * Esta función es simbólica o para protección de celdas si se usara Sheet UI.
 * Como es Web App, solo aseguramos que HORA_CIERRE se respete.
 */
function scheduledLockSheet() {
  console.log("Bloqueo lógico activado por horario (HORA_CIERRE).");
}

/**
 * 2:30 PM: Cierre, Respaldo y Limpieza
 */
function scheduledDailyClose() {
  const now = new Date();
  const nextBusinessDay = getNextBusinessDay_(now);
  if (!nextBusinessDay) return; // Nada que cerrar si no hay operación mañana?

  // Nota: El cierre suele ser para el día siguiente.
  // Ejemplo: Hoy Lunes 2:30 PM cierro pedidos para Martes.
  const targetDateStr = formatDate_(nextBusinessDay);

  // 1. Generar Respaldo PDF/Excel de pedidos para targetDateStr
  backupOrdersToDrive_(targetDateStr);

  // 2. Enviar Resumen General a Admins
  sendDailyAdminSummary_(targetDateStr);

  // 3. Validar Integridad de Menú (Arroz Blanco) para días futuros
  checkMenuIntegrity_();

  // 4. Actualizar Encabezados Visuales (si se usa la hoja como referencia)
  updateSheetHeaders_(targetDateStr);
}

/**
 * 3:00 PM: Reportes por Departamento
 */
function scheduledDepartmentReports() {
  const now = new Date();
  const nextBusinessDay = getNextBusinessDay_(now);
  if (!nextBusinessDay) return;

  const dateStr = formatDate_(nextBusinessDay);

  // Agrupar pedidos por departamento
  const orders = getOrdersByDate_(dateStr);
  const byDept = {};

  orders.forEach(o => {
    const dept = o.departamento || 'Sin Depto';
    if (!byDept[dept]) byDept[dept] = [];
    byDept[dept].push(o);
  });

  // Obtener mapeo de responsables
  const rawMap = getConfigValue_('RESPONSIBLES_EMAILS_JSON');
  let deptEmails = {};
  try { deptEmails = JSON.parse(rawMap); } catch (e) {}

  // Enviar correos
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

// === LÓGICA CORE (FECHAS Y MENÚ) ===

/**
 * Devuelve fechas futuras con menú disponible.
 * @param {boolean} fetchAll Si es true, trae todas. Si false, solo las inmediatas.
 */
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
    // Para "adelantar", permitimos pedir si la fecha es futura,
    // AUNQUE hoy ya haya pasado la hora de cierre de "mañana".
    // La regla de cierre es: "Para pedir para mañana, hazlo antes de hoy 2:30".
    // Para pedir para pasado mañana, la hora de cierre es mañana 2:30.
    // Por tanto, fechas > mañana siempre están abiertas hoy.

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

  if (targetDate <= now) return false; // Pasado

  const day = targetDate.getDay();
  if (day === 0 || day === 6) return false; // Finde
  if (holidaysSet.has(targetDateStr)) return false; // Feriado

  // Regla Cierre: Pedir antes de HORA_CIERRE del día hábil previo.
  const prevBizDay = getPreviousBusinessDay_(targetDate, holidaysSet);
  const prevBizDayStr = formatDate_(prevBizDay);
  const todayStr = formatDate_(now);

  // Si hoy es el día previo (ej. pido para mañana)
  if (todayStr === prevBizDayStr) {
    const configVal = getConfigValue_('HORA_CIERRE');
    let h, m;

    // Robust parsing: Handle String "14:30" or Date object
    if (configVal instanceof Date) {
      h = configVal.getHours();
      m = configVal.getMinutes();
    } else {
      const cutStr = String(configVal || '14:30');
      const parts = cutStr.split(':');
      if (parts.length >= 2) {
        h = parseInt(parts[0], 10);
        m = parseInt(parts[1], 10);
      } else {
        h = 14; m = 30; // Fallback
      }
    }

    const limit = new Date();
    limit.setHours(h, m, 0, 0);
    if (now > limit) return false;
  }

  // Si hoy es posterior al día previo (ya pasó el día de pedir)
  const zeroNow = new Date(now); zeroNow.setHours(0,0,0,0);
  const zeroPrev = new Date(prevBizDay); zeroPrev.setHours(0,0,0,0);
  if (zeroNow > zeroPrev) return false;

  return true;
}

// === RESPALDOS Y HELPERS ===

function backupOrdersToDrive_(dateStr) {
  const rootId = getConfigValue_('BACKUP_FOLDER_ID');
  if (!rootId) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();

  // Filtrar pedidos de la fecha
  const filtered = data.filter((row, i) => i === 0 || formatDate_(new Date(row[2])) === dateStr);
  if (filtered.length <= 1) return; // Solo header

  try {
    const rootFolder = DriveApp.getFolderById(rootId);
    const d = new Date(dateStr + 'T12:00:00');
    const year = String(d.getFullYear());
    const month = String(d.getMonth() + 1).padStart(2, '0');

    // Estructura Año/Mes
    let yFolder = rootFolder.getFoldersByName(year).hasNext() ? rootFolder.getFoldersByName(year).next() : rootFolder.createFolder(year);
    let mFolder = yFolder.getFoldersByName(month).hasNext() ? yFolder.getFoldersByName(month).next() : yFolder.createFolder(month);

    // Crear Spreadsheet temporal
    const tempSheet = SpreadsheetApp.create(`Pedidos_${dateStr}`);
    tempSheet.getSheets()[0].getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
    const tempFile = DriveApp.getFileById(tempSheet.getId());

    // Mover a carpeta
    tempFile.moveTo(mFolder);

    // Generar PDF (básico)
    const pdfBlob = tempFile.getAs('application/pdf');
    mFolder.createFile(pdfBlob).setName(`Pedidos_${dateStr}.pdf`);

    // Borrar temporal si se quiere solo PDF o dejar excel. La guía dice "PDF y Excel".
    // Dejamos ambos.

  } catch (e) {
    console.error("Error backup: " + e.message);
  }
}

function sendDailyAdminSummary_(dateStr) {
  const admins = getConfigValue_('ADMIN_EMAILS');
  if (!admins) return;

  const orders = getOrdersByDate_(dateStr);
  const count = orders.length;

  if (count > 0) {
    sendEmail_(admins, `Resumen Pedidos ${dateStr}`,
      `Se han registrado <b>${count}</b> pedidos para el día ${dateStr}.<br>` +
      `El respaldo ha sido generado en Drive.`
    );
  }
}

function sendEmail_(to, subject, htmlBody) {
  const testMode = getConfigValue_('TEST_EMAIL_MODE') === 'TRUE';
  const testDest = getConfigValue_('TEST_EMAIL_DEST');
  const senderName = getConfigValue_('MAIL_SENDER_NAME');

  const recipient = testMode ? testDest : to;
  if (!recipient) return;

  const finalSubject = testMode ? `[TEST] ${subject}` : subject;

  // Inyectar firma si existe
  const sigUrl = getSignatureDataUrl_();
  const signatureHtml = sigUrl ? `<br><br><img src="${sigUrl}" style="max-height:100px;">` : '';

  MailApp.sendEmail({
    to: recipient,
    subject: finalSubject,
    htmlBody: htmlBody + signatureHtml,
    name: senderName
  });
}

function checkMenuIntegrity_() {
  const ss = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName('Menu');
  const data = mSh.getDataRange().getValues();

  const datesToCheck = new Set();
  const menuMap = {}; // { date: { hasRice: bool, id: ... } }

  // Recolectar datos
  for (let i = 1; i < data.length; i++) {
    const dStr = formatDate_(new Date(data[i][1]));
    const cat = data[i][2];
    const item = String(data[i][3]).toLowerCase();

    // Solo verificar futuro
    if (dStr > formatDate_(new Date())) {
      if (!menuMap[dStr]) menuMap[dStr] = { hasRice: false };

      if (cat === 'Arroces' && item.includes('arroz blanco')) {
        menuMap[dStr].hasRice = true;
      }
    }
  }

  // Verificar
  const warnings = [];
  Object.keys(menuMap).forEach(d => {
    if (!menuMap[d].hasRice) {
      warnings.push(d);
    }
  });

  if (warnings.length > 0) {
    const msg = `Advertencia: El menú para las siguientes fechas no incluye "Arroz Blanco" en la categoría Arroces: ${warnings.join(', ')}. Esto bloqueará la selección de Granos.`;
    console.warn(msg);
    // Notificar admin
    const admins = getConfigValue_('ADMIN_EMAILS');
    if (admins) sendEmail_(admins, "Alerta: Integridad de Menú", msg);
  }
}

function updateSheetHeaders_(nextDateStr) {
  // Actualiza textos visuales en las hojas para referencia de usuarios que entran directo al Sheet
  const ss = SpreadsheetApp.getActive();

  // 1. Hoja Solicitud (Pedidos)
  // Nota: "Solicitud" en el prompt original se refiere a donde piden. Aquí es "Pedidos".
  // Si existiera una hoja UI legacy llamada "Solicitud", la actualizaríamos.
  // Como usamos "Pedidos" como DB, actualizaremos una celda de info si existe, o creamos un log.
  // Asumiremos que el usuario podría ver la hoja Menu.

  // 2. Hoja Menu
  const mSh = ss.getSheetByName('Menu');
  if (mSh) {
    // Intentar poner un texto descriptivo en la fila 1 si hay espacio o en una celda libre
    // El requerimiento dice: "Menú del [lunes] al [viernes]..."
    // Calculamos inicio y fin de esa semana
    const d = new Date(nextDateStr + 'T12:00:00');
    const day = d.getDay(); // 0-6
    const diffToMon = d.getDate() - day + (day == 0 ? -6 : 1); // adjust when day is sunday
    const monday = new Date(d.setDate(diffToMon));
    const friday = new Date(d.setDate(monday.getDate() + 4));

    const text = `Menú de la semana: ${formatDate_(monday)} al ${formatDate_(friday)}`;
    // Lo ponemos en una nota o celda A1 si es posible sin romper headers.
    // Como A1 son headers, lo ponemos como Nota en A1.
    mSh.getRange('A1').setNote(text);
  }
}

// === DATA ACCESS HELPERS ===

function getOrdersByDate_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate === dateStr && data[i][8] !== 'CANCELADO') {
      list.push({
        nombre: data[i][4],
        departamento: data[i][5],
        resumen: data[i][6]
      });
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
      return {
        email: data[i][0],
        nombre: data[i][1],
        departamento: data[i][2],
        rol: data[i][3],
        estado: data[i][4]
      };
    }
  }
  return null;
}

function getMenuByDate_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
  const data = sh.getDataRange().getValues();
  const menu = {};
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][1]));
    if (rowDate === dateStr && String(data[i][5]).toUpperCase() === 'SI') {
      const cat = data[i][2];
      if (!menu[cat]) menu[cat] = [];
      menu[cat].push({ id: data[i][0], plato: data[i][3], desc: data[i][4] });
    }
  }
  return menu;
}

function getAllMenus_(availableDates) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
  const data = sh.getDataRange().getValues();
  const menuMap = {}; // { dateStr: { cat: [items] } }

  // Initialize map
  const validDates = new Set(availableDates.map(d => d.value));
  availableDates.forEach(d => {
    menuMap[d.value] = {};
  });

  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][1]));
    if (validDates.has(rowDate) && String(data[i][5]).toUpperCase() === 'SI') {
      const cat = data[i][2];
      const item = { id: data[i][0], plato: data[i][3], desc: data[i][4] };

      if (!menuMap[rowDate][cat]) menuMap[rowDate][cat] = [];
      menuMap[rowDate][cat].push(item);
    }
  }
  return menuMap;
}

function getUserPrefs_(email) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      return JSON.parse(data[i][5] || '{}');
    }
  }
  return {};
}

function getAllUserOrders_(email, availableDates) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const ordersMap = {};

  const validDates = new Set(availableDates.map(d => d.value));

  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    // Check if it's one of the relevant dates AND belongs to user AND not canceled
    if (validDates.has(rowDate) &&
        String(data[i][3]).toLowerCase() === email &&
        data[i][8] !== 'CANCELADO') {
      ordersMap[rowDate] = { id: data[i][0], resumen: data[i][6], detalle: JSON.parse(data[i][7] || '{}') };
    }
  }
  return ordersMap;
}

function getOrderByUserDate_(email, dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate === dateStr && String(data[i][3]).toLowerCase() === email) {
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
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate === dateStr && String(data[i][3]).toLowerCase() === user.email) {
      rowIdx = i + 1;
      break;
    }
  }
  const id = rowIdx > 0 ? data[rowIdx - 1][0] : Utilities.getUuid();
  const now = new Date();
  const rowData = [
    id, now, dateStr, user.email, user.nombre, user.departamento,
    selection.items.join(', '), JSON.stringify(selection), 'ACTIVO', now
  ];
  if (rowIdx > 0) {
    sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }
}

function getDepartmentStats_(dateStr, deptFilter) {
  const departmentStats = { total: 0, byUser: [] };
  const orders = getOrdersByDate_(dateStr);
  orders.forEach(o => {
    if (!deptFilter || o.departamento === deptFilter) {
      departmentStats.total++;
      departmentStats.byUser.push({ nombre: o.nombre, pedido: o.resumen, depto: o.departamento });
    }
  });
  return departmentStats;
}

function validateOrderRules_(sel) {
  const cats = sel.categorias || [];
  const items = sel.items || [];
  const specialList = ['Vegetariana', 'Caldo', 'Opcion_Rapida'];
  const hasSpecial = cats.some(c => specialList.includes(c));
  if (hasSpecial && cats.length > 1) {
     const uniqueCats = [...new Set(cats)];
     if (uniqueCats.some(c => !specialList.includes(c))) {
       throw new Error("Platos especiales no se pueden combinar con el menú regular.");
     }
  }
  if (cats.includes('Granos')) {
    const hasWhiteRice = items.some(i => i.toLowerCase().includes('arroz blanco'));
    if (!hasWhiteRice) throw new Error("Los granos requieren seleccionar Arroz Blanco.");
  }
  if (cats.includes('Arroces') && cats.includes('Viveres')) {
    throw new Error("No puedes combinar Arroz y Víveres.");
  }
}

// === UTILS ===

let _configCache = null;

function getConfigValue_(key) {
  if (!_configCache) {
    _configCache = {};
    const sh = SpreadsheetApp.getActive().getSheetByName('Config');
    if (sh) {
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        _configCache[String(data[i][0])] = data[i][1];
      }
    }
  }
  return _configCache[key] !== undefined ? _configCache[key] : '';
}

function getHolidaysSet_() {
  const cache = CacheService.getScriptCache();
  const cachedHolidays = cache.get('HOLIDAYS_CACHE_V2');

  if (cachedHolidays) {
    return new Set(JSON.parse(cachedHolidays));
  }

  const set = new Set();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('DiasLibres');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      if (r[0]) set.add(formatDate_(new Date(r[0])));
    });
  }
  try {
    const calId = 'es.do#holiday@group.v.calendar.google.com';
    const now = new Date();
    const start = new Date(now.getTime() - 30 * 86400000);
    const end = new Date(now.getTime() + 365 * 86400000);
    CalendarApp.getCalendarById(calId).getEvents(start, end).forEach(e => set.add(formatDate_(e.getStartTime())));
  } catch (e) {
    console.error('Error fetching calendar holidays: ' + e.message);
  }

  // Cache for 6 hours
  cache.put('HOLIDAYS_CACHE_V2', JSON.stringify(Array.from(set)), 21600);

  return set;
}

function getNextBusinessDay_(date) {
  let d = new Date(date);
  d.setDate(d.getDate() + 1);
  const holidays = getHolidaysSet_();
  while (d.getDay() === 0 || d.getDay() === 6 || holidays.has(formatDate_(d))) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

function getPreviousBusinessDay_(date, holidaysSet) {
  let d = new Date(date);
  do { d.setDate(d.getDate() - 1); }
  while (d.getDay() === 0 || d.getDay() === 6 || holidaysSet.has(formatDate_(d)));
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

function getSignatureDataUrl_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('SIG_V3_PROD');
  if (cached) return cached;
  const fileId = getConfigValue_('FOOTER_SIGNATURE_ID');
  if (!fileId) return '';
  try {
    const file = Drive.Files.get(fileId, { fields: 'thumbnailLink' });
    if (!file || !file.thumbnailLink) return '';
    const url = file.thumbnailLink.replace(/=s\d+$/, '=s300');
    const blob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    const b64 = Utilities.base64Encode(blob.getBytes());
    const dataUrl = 'data:image/png;base64,' + b64;
    cache.put('SIG_V3_PROD', dataUrl, 21600);
    return dataUrl;
  } catch (e) { return ''; }
}
