/**
 * Code.gs - Backend V4 (Final Consolidado)
 */
const APP_VERSION = 'v3.2';

// === RUTAS E INICIO ===

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('index');
  const user = getUserInfo_();

  // Inyectar firma (versión optimizada)
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

// === API PÚBLICA (CLIENTE -> SERVIDOR) ===

/**
 * Obtiene datos iniciales para la app.
 * Si `requestedDateStr` es nulo, calcula automáticamente el primer día disponible.
 */
function apiGetInitData(requestedDateStr) {
  try {
    const user = getUserInfo_();
    if (!user) throw new Error("Usuario no encontrado en la base de datos.");

    // 1. Calcular días disponibles (Filtrando hora de cierre y feriados)
    const availableDates = getAvailableMenuDates_();

    // Si no hay menús futuros cargados o todos cerraron
    if (availableDates.length === 0) {
      return {
        ok: true,
        empty: true,
        msg: "No hay menús disponibles para ordenar en este momento."
      };
    }

    // 2. Determinar qué fecha mostrar (la pedida o la primera disponible)
    let targetDateStr = requestedDateStr;
    if (!targetDateStr || !availableDates.some(d => d.value === targetDateStr)) {
      targetDateStr = availableDates[0].value;
    }

    // 3. Cargar datos del contexto (Menú y Pedido de esa fecha)
    const existingOrder = getOrderByUserDate_(user.email, targetDateStr);
    const menu = getMenuByDate_(targetDateStr);

    // 4. Datos para Admin (si aplica)
    let adminSummary = null;
    if (user.rol === 'ADMIN_GEN' || user.rol === 'ADMIN_DEP') {
      adminSummary = getDepartmentStats_(targetDateStr, (user.rol === 'ADMIN_GEN' ? null : user.departamento));
    }

    return {
      ok: true,
      user: user,
      currentDate: targetDateStr, // Fecha activa (YYYY-MM-DD)
      dates: availableDates,      // Lista para las pestañas
      menu: menu,
      myOrder: existingOrder,
      adminData: adminSummary
    };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

/**
 * Recibe y guarda el pedido validando reglas de negocio y tiempos.
 */
function apiSubmitOrder(payload) {
  try {
    const user = getUserInfo_();
    const dateStr = payload.date; // Fecha objetivo

    // 1. Firewall de Tiempo: Validar que NO se haya cerrado el día mientras el usuario elegía
    if (!isDateOpenForOrdering_(dateStr)) {
      throw new Error("El tiempo límite para pedir el almuerzo de esta fecha ha expirado.");
    }

    // 2. Firewall de Reglas: Validar combinaciones prohibidas
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
    // Buscar por ID y Email (seguridad)
    if (String(rows[i][0]) === String(orderId) && String(rows[i][3]).toLowerCase() === user.email.toLowerCase()) {

      // Validar si aún es tiempo de cancelar (Regla de Cierre)
      const orderDate = Utilities.formatDate(new Date(rows[i][2]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (!isDateOpenForOrdering_(orderDate)) {
        return { ok: false, msg: "Ya no puedes cancelar este pedido porque la hora de cierre ha pasado." };
      }

      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, msg: "Pedido no encontrado." };
}

// === LÓGICA DE FECHAS Y CIERRE (CORE) ===

/**
 * Devuelve lista de fechas futuras que tienen menú Y están abiertas para pedir.
 */
function getAvailableMenuDates_() {
  const ss = SpreadsheetApp.getActive();
  const menuSh = ss.getSheetByName('Menu');
  // Obtenemos solo las fechas del menú (columna B, índice 1)
  const menuData = menuSh.getRange(2, 1, menuSh.getLastRow() - 1, 2).getValues();

  const now = new Date();
  const todayStr = formatDate_(now);

  // Set de fechas únicas que existen en la hoja Menu (Hoy o Futuro)
  const menuDatesSet = new Set();
  menuData.forEach(r => {
    const dStr = formatDate_(new Date(r[1]));
    if (dStr >= todayStr) menuDatesSet.add(dStr);
  });

  const validDates = [];
  const holidays = getHolidaysSet_();

  // Revisamos si es tarde HOY para pedir para mañana
  // (Esta lógica se delega a isDateOpenForOrdering_, pero preparamos el entorno)

  const sortedMenuDates = Array.from(menuDatesSet).sort();

  sortedMenuDates.forEach(dateStr => {
    // Verificar si la fecha cumple con todas las reglas (Feriado, Finde, Hora Cierre)
    if (isDateOpenForOrdering_(dateStr, holidays)) {
      validDates.push({
        value: dateStr,
        label: formatDisplayDate_(dateStr)
      });
    }
  });

  return validDates;
}

/**
 * Verifica si una fecha de consumo es válida para pedir AHORA MISMO.
 */
function isDateOpenForOrdering_(targetDateStr, holidaysSet) {
  if (!holidaysSet) holidaysSet = getHolidaysSet_();

  const now = new Date();
  const targetDate = new Date(targetDateStr + 'T12:00:00'); // Fijar mediodía para evitar problemas de timezone

  // 1. No se puede pedir para el pasado ni para hoy (según regla "siguiente día hábil")
  if (targetDate <= now) return false;

  // 2. Fin de semana (Sábado 6, Domingo 0) -> Cerrado
  const day = targetDate.getDay();
  if (day === 0 || day === 6) return false;

  // 3. Feriado -> Cerrado
  if (holidaysSet.has(targetDateStr)) return false;

  // 4. REGLA DEL CIERRE (2:30 PM)
  // Para comer el Día X, el pedido debe hacerse antes de la HORA_CIERRE del "Día Previo Hábil".

  const prevBusinessDay = getPreviousBusinessDay_(targetDate, holidaysSet);
  const prevDayStr = formatDate_(prevBusinessDay);
  const todayStr = formatDate_(now);

  // Si HOY es el día previo hábil (ej. Hoy martes, quiero pedir para Miércoles)
  if (todayStr === prevDayStr) {
     const cutoffHourStr = getConfigValue_('HORA_CIERRE') || '14:30'; // Por defecto 2:30 PM
     const [cutH, cutM] = cutoffHourStr.split(':').map(Number);

     const limitTime = new Date();
     limitTime.setHours(cutH, cutM, 0, 0);

     // Si ahora es más tarde que el límite, CERRAMOS el pedido para esa fecha
     if (now > limitTime) return false;
  }

  // Si HOY es posterior al día previo hábil (ya pasó el día de pedir), cerrado.
  // (Ej. Hoy es Jueves y quiero pedir para Miércoles -> Imposible, ya filtrado por fecha,
  // pero cubre casos raros de feriados intermedios).
  const zeroNow = new Date(now); zeroNow.setHours(0,0,0,0);
  const zeroPrev = new Date(prevBusinessDay); zeroPrev.setHours(0,0,0,0);

  if (zeroNow > zeroPrev) return false;

  return true;
}

// === HELPERS DE NEGOCIO (VALIDACIONES) ===

function validateOrderRules_(sel) {
  const cats = sel.categorias || [];
  const items = sel.items || []; // Nombres de los platos

  // 1. Platos Especiales Exclusivos
  // (Vegetariana, Caldo, Opción Rápida no se mezclan con buffet normal)
  const specialList = ['Vegetariana', 'Caldo', 'Opcion_Rapida'];
  const hasSpecial = cats.some(c => specialList.includes(c));

  if (hasSpecial && cats.length > 1) {
     // Si hay especial, todas las categorías deben ser especiales (o la misma)
     // Ej: Se permite 2 caldos, pero no Caldo + Arroz
     const uniqueCats = [...new Set(cats)];
     if (uniqueCats.some(c => !specialList.includes(c))) {
       throw new Error("Los platos especiales (Vegetariano, Caldo, Rápido) no se pueden combinar con el menú regular (Arroz, Carne, etc.).");
     }
  }

  // 2. Regla Granos
  // Si pides Granos, DEBES tener Arroz Blanco
  if (cats.includes('Granos')) {
    const hasWhiteRice = items.some(i => i.toLowerCase().includes('arroz blanco'));
    if (!hasWhiteRice) {
      throw new Error("Los granos requieren seleccionar Arroz Blanco.");
    }
  }

  // 3. Regla Arroz vs Víveres
  if (cats.includes('Arroces') && cats.includes('Viveres')) {
    throw new Error("No puedes seleccionar Arroz y Víveres en el mismo pedido.");
  }
}

// === HELPERS DE DATOS (BASES DE DATOS) ===

function getUserInfo_() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  const data = sh.getDataRange().getValues();
  // Headers: email, nombre, departamento, rol, estado
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
    // Filtramos por fecha y que esté habilitado ('SI')
    if (rowDate === dateStr && String(data[i][5]).toUpperCase() === 'SI') {
      const cat = data[i][2];
      if (!menu[cat]) menu[cat] = [];
      menu[cat].push({
        id: data[i][0],
        plato: data[i][3],
        desc: data[i][4]
      });
    }
  }
  return menu;
}

function getOrderByUserDate_(email, dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  // Buscamos de abajo hacia arriba (último pedido válido)
  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = formatDate_(new Date(data[i][2])); // Col C: Fecha Consumo
    if (rowDate === dateStr && String(data[i][3]).toLowerCase() === email) {
      return {
        id: data[i][0],
        resumen: data[i][6],
        detalle: JSON.parse(data[i][7] || '{}')
      };
    }
  }
  return null;
}

function saveOrderToSheet_(user, dateStr, selection) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;

  // Buscar si ya existe pedido para esa fecha/usuario para sobrescribir
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate === dateStr && String(data[i][3]).toLowerCase() === user.email) {
      rowIdx = i + 1;
      break;
    }
  }

  const id = rowIdx > 0 ? data[rowIdx - 1][0] : Utilities.getUuid();
  const now = new Date();

  // Headers: id, fecha_solicitud, fecha_consumo, email, nombre, depto, resumen, json, estado, timestamp
  const rowData = [
    id,
    now,
    dateStr,
    user.email,
    user.nombre,
    user.departamento,
    selection.items.join(', '),
    JSON.stringify(selection),
    'ACTIVO',
    now
  ];

  if (rowIdx > 0) {
    sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }
}

function getDepartmentStats_(dateStr, deptFilter) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const stats = { total: 0, byUser: [] };

  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    const rowDept = data[i][5]; // Col F

    if (rowDate === dateStr) {
      // Si es Admin Gen (filter null) o coincide depto
      if (!deptFilter || rowDept === deptFilter) {
        stats.total++;
        stats.byUser.push({
          nombre: data[i][4],
          pedido: data[i][6],
          depto: rowDept
        });
      }
    }
  }
  return stats;
}

// === UTILS (CONFIG, FECHAS, HELPERS) ===

function getConfigValue_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return data[i][1];
  }
  return '';
}

function getHolidaysSet_() {
  const set = new Set();

  // 1. Días libres Institucionales (desde Sheet "DiasLibres")
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('DiasLibres');
  if (sh) {
    const data = sh.getRange(2, 1, sh.getLastRow(), 1).getValues();
    data.forEach(r => {
      if (r[0]) set.add(formatDate_(new Date(r[0])));
    });
  }

  // 2. Feriados Oficiales República Dominicana (desde Google Calendar)
  try {
    const calId = 'es.do#holiday@group.v.calendar.google.com';
    const now = new Date();
    // Consultamos desde hace 30 días hasta un año a futuro
    const start = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const end = new Date(now.getTime() + 365 * 24 * 60 * 60 * 1000);

    const calendar = CalendarApp.getCalendarById(calId);
    if (calendar) {
      const events = calendar.getEvents(start, end);
      events.forEach(e => {
        // Los eventos de todo el día suelen ser de medianoche a medianoche
        // Tomamos el start date que es la fecha del feriado
        set.add(formatDate_(e.getStartTime()));
      });
    }
  } catch (e) {
    console.warn('No se pudo obtener el calendario de feriados: ' + e.message);
  }

  return set;
}

function getPreviousBusinessDay_(date, holidaysSet) {
  let d = new Date(date);
  // Retroceder 1 día hasta encontrar uno que NO sea finde ni feriado
  do {
    d.setDate(d.getDate() - 1);
  } while (d.getDay() === 0 || d.getDay() === 6 || holidaysSet.has(formatDate_(d)));
  return d;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDisplayDate_(dateStr) {
  // Formato amigable: "Jueves 28/11"
  // Usamos T12:00:00 para evitar saltos de fecha por zona horaria
  const d = new Date(dateStr + 'T12:00:00');
  const days = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  return `${days[d.getDay()]} ${d.getDate()}/${d.getMonth() + 1}`;
}

// === IMAGEN / FIRMA (OPTIMIZADA Y DEBUGGEADA) ===

function getSignatureDataUrl_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'SIG_V3_PROD';

  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const fileId = getConfigValue_('FOOTER_SIGNATURE_ID');
  if (!fileId) return '';

  try {
    const blob = getDriveThumbnailBlob_(fileId, 300);
    if (!blob) return ''; // Silencioso si falla

    const b64 = Utilities.base64Encode(blob.getBytes());
    const mime = blob.getContentType() || 'image/png';
    const dataUrl = 'data:' + mime + ';base64,' + b64;

    try { cache.put(cacheKey, dataUrl, 21600); } catch (e) {}

    return dataUrl;
  } catch (e) {
    console.error('Error firma: ' + e.message);
    return '';
  }
}

function getDriveThumbnailBlob_(fileId, size) {
  try {
    const file = Drive.Files.get(fileId, { fields: 'thumbnailLink' });
    let url = file && file.thumbnailLink;
    if (!url) return null;

    url = url.replace(/=s\d+$/, '=s' + Math.max(32, size || 300));

    const resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) return null;

    const blob = resp.getBlob();
    const type = blob.getContentType();
    return type.startsWith('image/') ? (type.match(/png|jpeg|jpg/) ? blob : blob.setContentType('image/png')) : null;

  } catch (e) {
    console.warn('Error getDriveThumbnailBlob_: ' + e.message);
    return null;
  }
}