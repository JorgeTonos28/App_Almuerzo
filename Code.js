/**
 * Code.gs - Backend V5 (Refactor & New Features)
 */
const APP_VERSION = 'v7.36';
const SPREADSHEET_RETRY_ATTEMPTS = 4;
const SPREADSHEET_RETRY_DELAY_MS = 1500;
const CHEF_GAME_DAILY_LIMIT_SECONDS = 15 * 60;
const CHEF_GAME_SCORE_CHEF = 1;
const CHEF_GAME_SCORE_COMBO_3 = 1;
const CHEF_GAME_SCORE_COMBO_5 = 2;
const CHEF_GAME_SCORE_PERFECT = 3;
const CHEF_GAME_SCORE_MISS = -1;
const CHEF_GAME_SCORE_ONION = -2;
const CHEF_GAME_SCORE_PAN = -4;
const CHEF_GAME_SCHEMA_CACHE_KEY = 'CHEF_GAME_SCHEMA_READY_V1';
const MENU_CATEGORIES_SHEET = 'CategoriasMenu';
const DEFAULT_MENU_CATEGORIES = [
  ['Arroces', 'Arroces', 10, 'ACTIVO', '', 'SI', 'Granos, Carnes, Ensaladas, Frituritas', 'UNICA'],
  ['Granos', 'Granos', 20, 'ACTIVO', 'Legumbres, Habichuelas', 'SI', 'Arroces, Carnes, Ensaladas, Frituritas', 'UNICA'],
  ['Carnes', 'Carnes', 30, 'ACTIVO', 'Proteinas, Carnes y Pescados', 'SI', 'Arroces, Granos, Ensaladas, Viveres, Frituritas', 'UNICA'],
  ['Ensaladas', 'Ensaladas', 40, 'ACTIVO', 'Guarnición, Guarnición / Ensalada', 'SI', 'Arroces, Granos, Carnes, Viveres, Frituritas', 'UNICA'],
  ['Viveres', 'Viveres', 50, 'ACTIVO', 'Víveres, Tuberculos', 'SI', 'Carnes, Ensaladas, Frituritas', 'UNICA'],
  ['Vegetariana', 'Vegetariana', 60, 'ACTIVO', 'Vegetariano, Veggie, Menú Vegetariano', 'NO', '', 'UNICA'],
  ['Caldo', 'Caldo', 70, 'ACTIVO', 'Caldos, Sopas, Sopa', 'NO', '', 'MULTIPLE'],
  ['Opcion_Rapida', 'Opcion Rapida', 80, 'ACTIVO', 'Opción Rápida, Rapida, Sandwich, Dieta', 'NO', '', 'UNICA'],
  ['Frituritas', 'Frituritas', 90, 'ACTIVO', 'Frituras, Fritos, Snack', 'SI', 'Arroces, Granos, Carnes, Ensaladas, Viveres', 'UNICA']
];

// === RUTAS E INICIO ===

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (isMenuDayEndpointRequest_(params)) {
    return handleMenuDayEndpointRequest_(params);
  }

  const t = HtmlService.createTemplateFromFile('index');
  const user = getUserInfo_();
  t.signatureUrl = '';
  t.initialDataJson = 'null';

  if (!user || user.estado !== 'ACTIVO') {
    const denied = HtmlService.createTemplateFromFile('Denied');
    denied.signatureUrl = getSignatureDataUrl_();
    denied.email = Session.getActiveUser().getEmail().toLowerCase();
    denied.status = user && user.estado ? user.estado.trim().toUpperCase() : null;
    return denied.evaluate()
      .setTitle('Acceso Denegado')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  t.user = user;
  t.appVersion = APP_VERSION;
  t.signatureUrl = getSignatureDataUrl_();
  t.initialDataJson = serializeForInlineScript_(apiGetInitData());

  return t.evaluate()
    .setTitle('Solicitud Almuerzo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureMenuCategoriesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(MENU_CATEGORIES_SHEET);
  const expectedHeaders = ['id', 'nombre', 'orden', 'estado', 'alias_importacion', 'es_combinable', 'combinable_con', 'tipo_seleccion'];
  if (!sheet) {
    sheet = ss.insertSheet(MENU_CATEGORIES_SHEET);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  } else {
    ensureSheetHeaders_(sheet, expectedHeaders);
  }
  ensureDefaultMenuCategories_(sheet);
  return sheet;
}

function ensureDefaultMenuCategories_(sheet) {
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    sheet.getRange(2, 1, DEFAULT_MENU_CATEGORIES.length, 8).setValues(DEFAULT_MENU_CATEGORIES);
    return;
  }

  const existingMap = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '').trim();
    if (id) existingMap[id] = { rowIdx: i + 1, row: data[i] };
  }

  const defaultDict = {};
  DEFAULT_MENU_CATEGORIES.forEach(cat => {
    defaultDict[cat[0]] = cat;
  });

  // 1. Backfill missing or empty constraint columns for existing categories
  Object.keys(existingMap).forEach(id => {
    const info = existingMap[id];
    const def = defaultDict[id];
    if (def) {
      let needsUpdate = false;
      const rowVals = [...info.row];
      while (rowVals.length < 8) rowVals.push('');

      if (!String(rowVals[5] || '').trim()) {
        rowVals[5] = def[5];
        needsUpdate = true;
      }
      if (!String(rowVals[6] || '').trim() && def[6]) {
        rowVals[6] = def[6];
        needsUpdate = true;
      }
      if (!String(rowVals[7] || '').trim()) {
        rowVals[7] = def[7];
        needsUpdate = true;
      }

      if (needsUpdate) {
        sheet.getRange(info.rowIdx, 1, 1, 8).setValues([[
          rowVals[0], rowVals[1], rowVals[2], rowVals[3], rowVals[4], rowVals[5], rowVals[6], rowVals[7]
        ]]);
      }
    }
  });

  // 2. Append missing default categories
  const missing = DEFAULT_MENU_CATEGORIES.filter(category => !existingMap[category[0]]);
  if (missing.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 8).setValues(missing);
  }
}

function parseMenuCategoryAliases_(value) {
  const values = Array.isArray(value) ? value : String(value || '').split(/[,;\n]/);
  const seen = {};
  return values.map(alias => String(alias || '').trim().replace(/\s+/g, ' ')).filter(alias => {
    const key = normalizeCategoryLookupKey_(alias);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function parseCategoryCombinableWith_(value) {
  if (!value) return [];
  const values = Array.isArray(value) ? value : String(value || '').split(/[,;\n]/);
  const seen = {};
  return values.map(item => String(item || '').trim()).filter(item => {
    if (!item || seen[item]) return false;
    seen[item] = true;
    return true;
  });
}

function getMenuCategories_(includeInactive) {
  const cacheKey = ['MENU_CATEGORIES', getRevisionValue_('APP_MENU_CATEGORIES_REVISION'), includeInactive ? 'ALL' : 'ACTIVE'].join(':');
  const cached = readJsonCache_(cacheKey);
  if (cached) return cached;

  const data = readSheetValues_(ensureMenuCategoriesSheet_(), 8);
  const specialDefaultNonCombinable = ['Vegetariana', 'Caldo', 'Opcion_Rapida'];
  const specialDefaultMulti = ['Vegetariana', 'Caldo', 'Opcion_Rapida'];

  const categories = data.slice(1)
    .map(row => {
      const id = String(row[0] || '').trim();
      const rawCombinable = String(row[5] || '').trim().toUpperCase();
      const rawTipo = String(row[7] || '').trim().toUpperCase();
      const esCombinable = rawCombinable ? rawCombinable !== 'NO' : !specialDefaultNonCombinable.includes(id);
      const tipoSeleccion = rawTipo ? (rawTipo === 'MULTIPLE' ? 'MULTIPLE' : 'UNICA') : (specialDefaultMulti.includes(id) ? 'MULTIPLE' : 'UNICA');

      return {
        id: id,
        nombre: normalizeMenuText_(row[1]),
        orden: Number.isFinite(Number(row[2])) ? Number(row[2]) : 999,
        estado: String(row[3] || '').trim().toUpperCase() === 'INACTIVO' ? 'INACTIVO' : 'ACTIVO',
        aliases: parseMenuCategoryAliases_(row[4]),
        es_combinable: esCombinable,
        combinable_con: parseCategoryCombinableWith_(row[6]),
        tipo_seleccion: tipoSeleccion
      };
    })
    .filter(category => category.id && category.nombre)
    .filter(category => includeInactive || category.estado === 'ACTIVO')
    .sort((a, b) => a.orden - b.orden || a.nombre.localeCompare(b.nombre, 'es'));

  writeJsonCache_(cacheKey, categories, 300);
  return categories;
}

function getMenuCategoryMap_() {
  const map = {};
  getMenuCategories_(true).forEach(category => { map[category.id] = category; });
  return map;
}

function getMenuCategoryById_(categoryId) {
  return getMenuCategoryMap_()[String(categoryId || '').trim()] || null;
}

function invalidateMenuCategoriesCache_() {
  bumpRevisionValue_('APP_MENU_CATEGORIES_REVISION');
  invalidateUserInitCache_();
  invalidateMenuDataCache_();
}

function isMenuDayEndpointRequest_(params) {
  const endpoint = String(params.endpoint || params.api || params.action || '').trim().toLowerCase();
  return endpoint === 'menu-dia' || endpoint === 'menu_day' || endpoint === 'menu-day';
}

function handleMenuDayEndpointRequest_(params) {
  try {
    ensureOperationalConfigKeys_();

    const configuredToken = String(getConfigValue_('MENU_DAY_ENDPOINT_TOKEN') || '').trim();
    const providedToken = String(params.token || params.apiKey || params.key || '').trim();

    if (!configuredToken) {
      return createJsonResponse_({
        ok: false,
        status: 503,
        error: 'ENDPOINT_NOT_CONFIGURED',
        msg: 'Endpoint no configurado. Define MENU_DAY_ENDPOINT_TOKEN en Config.'
      });
    }

    if (!providedToken || providedToken !== configuredToken) {
      return createJsonResponse_({
        ok: false,
        status: 401,
        error: 'UNAUTHORIZED',
        msg: 'Token invalido.'
      });
    }

    return createJsonResponse_(getMenuDayEndpointPayload_(params.fecha || params.date));
  } catch (err) {
    const msg = err && err.message ? err.message : 'Error interno.';
    const isDateError = msg.indexOf('Fecha') === 0;
    return createJsonResponse_({
      ok: false,
      status: isDateError ? 400 : 500,
      error: isDateError ? 'INVALID_DATE' : 'SERVER_ERROR',
      msg: msg
    });
  }
}

function createJsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeEndpointDate_(value) {
  const dateStr = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    throw new Error('Fecha requerida en formato YYYY-MM-DD.');
  }

  const date = new Date(dateStr + 'T12:00:00');
  if (isNaN(date.getTime()) || formatDate_(date) !== dateStr) {
    throw new Error('Fecha invalida.');
  }

  return dateStr;
}

function formatMenuRowDate_(rawDate) {
  if (rawDate instanceof Date) return formatDate_(rawDate);

  const raw = String(rawDate || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return formatDate_(new Date(raw + 'T12:00:00'));
  }

  return formatDate_(new Date(rawDate));
}

function getMenuDayEndpointPayload_(dateValue) {
  const dateStr = normalizeEndpointDate_(dateValue);
  const cacheKey = [
    'MENU_DAY_ENDPOINT',
    getRevisionValue_('APP_MENU_REVISION'),
    dateStr
  ].join(':');

  const cachedPayload = readJsonCache_(cacheKey);
  if (cachedPayload) return cachedPayload;

  const menuSheet = SpreadsheetApp.getActive().getSheetByName('Menu');
  if (!menuSheet) throw new Error('Hoja Menu no encontrada.');

  const data = readSheetValues_(menuSheet, 6);
  const menu = {};
  const items = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;

    let rowDate = '';
    try {
      rowDate = formatMenuRowDate_(row[1]);
    } catch (e) {
      continue;
    }

    if (rowDate !== dateStr || String(row[5]).trim().toUpperCase() !== 'SI') continue;

    const categoria = String(row[2] || '').trim();
    if (!categoria) continue;

    const item = {
      id: row[0],
      categoria: categoria,
      plato: normalizeMenuText_(row[3]),
      descripcion: normalizeMenuText_(row[4])
    };

    if (!menu[categoria]) menu[categoria] = [];
    menu[categoria].push({
      id: item.id,
      plato: item.plato,
      descripcion: item.descripcion
    });
    items.push(item);
  }

  const payload = {
    ok: true,
    fecha: dateStr,
    date: dateStr,
    label: formatDisplayDate_(dateStr),
    existeMenu: items.length > 0,
    exists: items.length > 0,
    menu: menu,
    items: items,
    appVersion: APP_VERSION,
    generadoEn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss")
  };

  writeJsonCache_(cacheKey, payload, 300);
  return payload;
}

function serializeForInlineScript_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
}

// === API PÚBLICA ===

function apiGetInitData(requestedDateStr, impersonateEmail) {
  try {
    const ss = SpreadsheetApp.getActive();
    const usersData = readSheetValues_(ss.getSheetByName('Usuarios'));
    const deptMap = getDepartmentMap_();

    const activeUser = getUserInfo_(null, usersData, deptMap);
    if (!activeUser) throw new Error("Usuario no encontrado.");
    ensureOperationalConfigKeys_();

    let targetUser = activeUser;
    let deptUsers = [];

    // Logic for Impersonation (ADMIN_DEP only)
    if (activeUser.rol === 'ADMIN_DEP') {
       // Filter out self (using pre-fetched usersData optimization in helper if needed, but simple filter here)
       deptUsers = getUsersByDept_(activeUser.departamentoId, usersData).filter(u => u.email.toLowerCase() !== activeUser.email.toLowerCase());

       if (impersonateEmail && impersonateEmail !== activeUser.email) {
          const checkUser = getUserInfo_(impersonateEmail, usersData, deptMap);
          if (checkUser && checkUser.departamentoId === activeUser.departamentoId) {
             targetUser = checkUser;
          }
       }
    }

    const initCacheKey = getInitCacheKey_(activeUser.email, targetUser.email, requestedDateStr || '');
    const cachedResponse = readJsonCache_(initCacheKey);
    if (cachedResponse) {
      cachedResponse.chefGame = getChefGameState_(targetUser.email);
      return cachedResponse;
    }

    const menuBundle = getMenuBundle_();
    const availableDates = menuBundle.dates || [];

    let targetDateStr = requestedDateStr;
    if (availableDates.length > 0) {
       if (!targetDateStr || !availableDates.some(d => d.value === targetDateStr)) {
         targetDateStr = availableDates[0].value;
       }
    } else {
       targetDateStr = null;
    }

    const ordersData = readSheetValues_(ss.getSheetByName('Pedidos'), 9);
    const allMenus = menuBundle.menusByDate || {};
    const menu = targetDateStr ? (allMenus[targetDateStr] || {}) : {};
    const allOrders = getAllUserOrders_(targetUser.email, null, ordersData);

    const existingOrder = allOrders[targetDateStr] || null;

    let adminSummary = null;
    if (activeUser.rol === 'ADMIN_GEN' || activeUser.rol === 'ADMIN_DEP') {
      adminSummary = getDepartmentStats_(targetDateStr, (activeUser.rol === 'ADMIN_GEN' ? null : activeUser.departamentoId), ordersData, deptMap);
    }

    const prefs = getUserPrefs_(targetUser.email, usersData);
    const mealPriceCurrent = getCurrentMealPrice_();
    const mealPriceHistory = getMealPriceHistory_();
    const announcementConfig = getAnnouncementConfig_();
    const provInfo = getProviderInfo_();
    const userMealRatings = getUserMealRatingsMap_(targetUser.email);
    const userProviderRating = getUserProviderRating_(targetUser.email, provInfo.periodId);
    const todayYmd = getTodayYmd_();
    const serverHour = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'H'));

    const nextBizDay = getNextBusinessDay_(new Date());

    const response = {
      ok: true,
      nextBusinessDay: formatDate_(nextBizDay),
      user: targetUser,
      activeUser: activeUser,
      deptUsers: deptUsers,
      userPrefs: prefs,
      currentDate: targetDateStr,
      dates: availableDates,
      menu: menu,
      allMenus: allMenus,
      allOrders: allOrders,
      myOrder: existingOrder,
      adminData: adminSummary,
      mealPricing: { current: mealPriceCurrent, history: mealPriceHistory },
      announcementConfig: announcementConfig,
      providerInfo: provInfo,
      userMealRatings: userMealRatings,
      userProviderRating: userProviderRating,
      serverHour: serverHour,
      serverTime: new Date().toISOString(),
      todayYmd: todayYmd,
      chefGame: getChefGameState_(targetUser.email),
      deptMap: deptMap,
      menuCategories: getMenuCategories_(true)
    };

    writeJsonCache_(initCacheKey, response, 45);
    return response;

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiGetDateViewData(requestedDateStr, impersonateEmail) {
  try {
    if (!requestedDateStr) throw new Error("Fecha requerida.");

    const ss = SpreadsheetApp.getActive();
    const usersData = readSheetValues_(ss.getSheetByName('Usuarios'), 7);
    const deptMap = getDepartmentMap_();
    const activeUser = getUserInfo_(null, usersData, deptMap);
    if (!activeUser) throw new Error("Usuario no encontrado.");
    ensureOperationalConfigKeys_();

    let targetUser = activeUser;
    if (activeUser.rol === 'ADMIN_DEP' && impersonateEmail && impersonateEmail !== activeUser.email) {
      const checkUser = getUserInfo_(impersonateEmail, usersData, deptMap);
      if (checkUser && checkUser.departamentoId === activeUser.departamentoId) {
        targetUser = checkUser;
      }
    }

    const cacheKey = getDateViewCacheKey_(activeUser.email, targetUser.email, requestedDateStr);
    const cachedResponse = readJsonCache_(cacheKey);
    if (cachedResponse) return cachedResponse;

    const menuBundle = getMenuBundle_();
    if (!menuBundle.dates.some(d => d.value === requestedDateStr)) {
      throw new Error("La fecha solicitada ya no está disponible.");
    }

    const ordersData = readSheetValues_(ss.getSheetByName('Pedidos'), 9);
    const myOrder = getUserOrderByDate_(targetUser.email, requestedDateStr, ordersData);

    let adminSummary = null;
    if (activeUser.rol === 'ADMIN_GEN' || activeUser.rol === 'ADMIN_DEP') {
      adminSummary = getDepartmentStats_(requestedDateStr, (activeUser.rol === 'ADMIN_GEN' ? null : activeUser.departamentoId), ordersData, deptMap);
    }

    const response = {
      ok: true,
      currentDate: requestedDateStr,
      menu: menuBundle.menusByDate[requestedDateStr] || {},
      myOrder: myOrder,
      adminData: adminSummary
    };

    writeJsonCache_(cacheKey, response, 45);
    return response;
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiCheckUserStatus() {
   const user = getUserInfo_();
   return user ? user.estado : null;
}

function apiRequestAccess(data) {
  try {
     const email = Session.getActiveUser().getEmail().toLowerCase();
     if (!email.endsWith('@infotep.gob.do')) throw new Error("Dominio no permitido.");

     const existing = getUserInfo_(email);
     if (existing) {
        if (existing.estado === 'PENDIENTE') throw new Error("Ya tienes una solicitud pendiente.");
        if (existing.estado === 'ACTIVO') throw new Error("Tu cuenta ya está activa.");
        if (existing.estado === 'INACTIVO') throw new Error("Tu cuenta está inactiva. Contacta a un administrador.");
     }

     const ss = SpreadsheetApp.getActive();
     const sh = ss.getSheetByName('Usuarios');

     // Validate Code Uniqueness
     if (!/^\d{4}$/.test(data.code)) throw new Error("Código inválido. Deben ser 4 dígitos.");
     const uData = sh.getDataRange().getValues();
     for(let i=1; i<uData.length; i++) {
        if (String(uData[i][6]) === String(data.code)) {
           throw new Error("El código de empleado " + data.code + " ya está en uso.");
        }
     }

     // Append PENDING user
     sh.appendRow([email, data.name || 'Sin Nombre', data.dept || 'Sin Depto', 'USER', 'PENDIENTE', '{}', data.code]);
     SpreadsheetApp.flush();

     // Notify Admins
     const admins = getConfigValue_('ADMIN_EMAILS');
     if (admins) {
        const html = getEmailTemplate_({
           title: 'Nueva Solicitud de Acceso',
           body: `
             <p>El usuario <strong>${data.name}</strong> ha solicitado acceso al sistema de almuerzo.</p>
             <div style="background-color: #f3f4f6; padding: 16px; border-radius: 8px; margin: 16px 0;">
                <p style="margin: 4px 0;"><strong>Correo:</strong> ${email}</p>
                <p style="margin: 4px 0;"><strong>Departamento:</strong> ${data.dept}</p>
                <p style="margin: 4px 0;"><strong>Código:</strong> ${data.code}</p>
             </div>
             <p>Ingresa al Panel de Administración para verificar y aprobar esta solicitud.</p>
           `,
           cta: { text: 'Ir al Panel de Administración', url: getAppUrl_() }
        });
        sendEmail_(admins, "Almuerzo Pre-empacado | Nueva Solicitud de Acceso", html);
     }

     // Notify User
     const userHtml = getEmailTemplate_({
        title: 'Solicitud Recibida',
        subtitle: 'Acceso en Proceso',
        body: `
          <p>Hola <strong>${data.name}</strong>,</p>
          <p>Hemos recibido tu solicitud de acceso al sistema de almuerzo.</p>
          <p>Tu solicitud está siendo procesada por el equipo administrativo. Recibirás un correo de confirmación una vez que tu acceso haya sido aprobado.</p>
        `,
        footerNote: 'Gracias por tu paciencia.'
     });
     sendEmail_(email, "Almuerzo Pre-empacado | Solicitud Recibida", userHtml);

     invalidateUserInitCache_();
     return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function apiSubmitOrder(payload) {
  try {
    const ss = SpreadsheetApp.getActive();
    const usersData = readSheetValues_(ss.getSheetByName('Usuarios'), 7);
    const activeUser = getUserAccessRecord_(null, usersData);
    if (!activeUser) throw new Error("Usuario no encontrado.");
    let targetUser = activeUser;

    if (payload.impersonateEmail && activeUser.rol === 'ADMIN_DEP') {
       const checkUser = getUserAccessRecord_(payload.impersonateEmail, usersData);
       if (checkUser && checkUser.departamentoId === activeUser.departamentoId) {
          targetUser = checkUser;
       } else {
          throw new Error("No tienes permiso para pedir por este usuario.");
       }
    }

    const dateStr = payload.date;
    if (!isDateOpenForOrdering_(dateStr)) {
      throw new Error("El tiempo límite para pedir el almuerzo de esta fecha ha expirado.");
    }

    validateOrderRules_(payload);
    const savedOrder = saveOrderToSheet_(targetUser, dateStr, payload, activeUser.email);
    invalidateUserInitCache_();
    return { ok: true, order: savedOrder };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiCancelOrder(orderId) {
  try {
    const activeEmail = Session.getActiveUser().getEmail().toLowerCase();

    const result = cancelOrderRecordById_(orderId, function(orderSnapshot) {
      return String(orderSnapshot.email).toLowerCase() === activeEmail;
    });

    if (!result.found) return { ok: false, msg: "Pedido no encontrado." };
    if (!result.allowed) return { ok: false, msg: "No tienes permiso para cancelar este pedido." };
    if (!isDateOpenForOrdering_(result.date)) {
      return { ok: false, msg: "Ya no puedes cancelar este pedido (hora de cierre pasada)." };
    }

    invalidateUserInitCache_();
    return { ok: true, date: result.date };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSetUserPreference(key, value, targetEmail) {
  try {
    const activeUser = getUserInfo_();
    let email = activeUser.email.toLowerCase();

    // Allow Admin to set prefs for others (e.g. disable reminders)
    if (targetEmail) {
       if (['ADMIN_GEN', 'ADMIN_DEP'].includes(activeUser.rol)) {
          // Verify scope for ADMIN_DEP
          if (activeUser.rol === 'ADMIN_DEP') {
             const target = getUserInfo_(targetEmail);
             if (!target || target.departamentoId !== activeUser.departamentoId) {
                throw new Error("No puedes modificar usuarios de otro departamento.");
             }
          }
          email = targetEmail.toLowerCase();
       } else {
          throw new Error("No tienes permisos para modificar otros usuarios.");
       }
    }

    const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email) {
        const currentPrefs = JSON.parse(data[i][5] || '{}');
        currentPrefs[key] = value;
        sh.getRange(i + 1, 6).setValue(JSON.stringify(currentPrefs));
        invalidateUserInitCache_();
        return { ok: true };
      }
    }
    return { ok: false, msg: "Usuario no encontrado" };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiDismissAnnouncement(announcementId) {
  try {
    const user = getUserInfo_();
    if (!user) throw new Error("Usuario no autenticado");
    const activeId = String(announcementId || getConfigValue_('ANNOUNCEMENT_ID') || 'default').trim();
    const prefs = getUserPrefs_(user.email);
    const announcements = prefs.announcements || {};
    announcements[activeId] = (Number(announcements[activeId]) || 0) + 1;
    return apiSetUserPreference('announcements', announcements);
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSaveAnnouncementConfig(payload) {
  try {
    const admin = getUserInfo_();
    if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");

    ensureOperationalConfigKeys_();
    const enabled = payload.enabled ? 'TRUE' : 'FALSE';
    const expiresOn = normalizeAnnouncementDate_(payload.expiresOn, formatDateWithOffset_(30));
    const maxDismiss = String(parsePositiveInt_(payload.maxDismiss, 3));
    let announcementId = String(payload.id || '').trim();
    if (payload.forceNewId || !announcementId) {
      announcementId = 'anuncio_' + getTodayYmd_().replace(/-/g, '_') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmm');
    }

    const slides = Array.isArray(payload.slides) ? payload.slides : [];
    const payloadJson = JSON.stringify({ slides: slides });

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Config');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]);
      if (key === 'ANNOUNCEMENT_ENABLED') sh.getRange(i + 1, 2).setValue(enabled);
      if (key === 'ANNOUNCEMENT_ID') sh.getRange(i + 1, 2).setValue(announcementId);
      if (key === 'ANNOUNCEMENT_EXPIRES_ON') sh.getRange(i + 1, 2).setValue(expiresOn);
      if (key === 'ANNOUNCEMENT_MAX_DISMISS') sh.getRange(i + 1, 2).setValue(maxDismiss);
      if (key === 'ANNOUNCEMENT_PAYLOAD_JSON') sh.getRange(i + 1, 2).setValue(payloadJson);
    }

    _configCache = null;
    invalidateUserInitCache_();
    return {
      ok: true,
      announcementConfig: {
        enabled: enabled === 'TRUE',
        id: announcementId,
        expiresOn: expiresOn,
        maxDismiss: parseInt(maxDismiss, 10),
        slides: slides
      }
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// === SISTEMA DE VALORACIONES (COMIDAS Y PROVEEDOR) ===

function ensureRatingsSheets_() {
  const ss = SpreadsheetApp.getActive();

  // 1. ValoracionesComida
  let mealSheet = ss.getSheetByName('ValoracionesComida');
  const mealHeaders = [
    'id', 'pedido_id', 'fecha_consumo', 'email_usuario', 'nombre_usuario', 'departamento',
    'puntuacion', 'comentario', 'platos_resumen', 'timestamp_creacion', 'timestamp_actualizacion'
  ];
  if (!mealSheet) {
    mealSheet = ss.insertSheet('ValoracionesComida');
    mealSheet.getRange(1, 1, 1, mealHeaders.length).setValues([mealHeaders]);
    mealSheet.getRange(1, 1, 1, mealHeaders.length).setFontWeight('bold').setBackground('#f3f4f6');
    mealSheet.setFrozenRows(1);
  } else {
    ensureSheetHeaders_(mealSheet, mealHeaders);
  }

  // 2. ValoracionesProveedor
  let provSheet = ss.getSheetByName('ValoracionesProveedor');
  const provHeaders = [
    'id', 'proveedor_periodo_id', 'proveedor_nombre', 'email_usuario', 'nombre_usuario', 'departamento',
    'puntuacion', 'comentario', 'version_voto', 'timestamp_creacion', 'timestamp_actualizacion'
  ];
  if (!provSheet) {
    provSheet = ss.insertSheet('ValoracionesProveedor');
    provSheet.getRange(1, 1, 1, provHeaders.length).setValues([provHeaders]);
    provSheet.getRange(1, 1, 1, provHeaders.length).setFontWeight('bold').setBackground('#f3f4f6');
    provSheet.setFrozenRows(1);
  } else {
    ensureSheetHeaders_(provSheet, provHeaders);
  }

  // 3. HistoricoValoracionesProveedor
  let histSheet = ss.getSheetByName('HistoricoValoracionesProveedor');
  const histHeaders = [
    'id', 'proveedor_periodo_id', 'proveedor_nombre', 'email_usuario', 'nombre_usuario', 'departamento',
    'puntuacion', 'comentario', 'timestamp'
  ];
  if (!histSheet) {
    histSheet = ss.insertSheet('HistoricoValoracionesProveedor');
    histSheet.getRange(1, 1, 1, histHeaders.length).setValues([histHeaders]);
    histSheet.getRange(1, 1, 1, histHeaders.length).setFontWeight('bold').setBackground('#f3f4f6');
    histSheet.setFrozenRows(1);
  } else {
    ensureSheetHeaders_(histSheet, histHeaders);
  }
}

function isMealRatingAllowed_(consumptionDateStr) {
  if (!consumptionDateStr) return false;
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  if (consumptionDateStr > todayStr) return false;
  if (consumptionDateStr === todayStr) {
    const hours = Number(Utilities.formatDate(now, tz, 'H'));
    if (hours < 12) return false;
  }
  return true;
}

function getUserMealRatingsMap_(email) {
  if (!email) return {};
  ensureRatingsSheets_();
  const sh = SpreadsheetApp.getActive().getSheetByName('ValoracionesComida');
  if (!sh) return {};
  const data = sh.getDataRange().getValues();
  const map = {};
  const targetEmail = String(email).toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][3] || '').toLowerCase();
    if (rowEmail === targetEmail) {
      const dateStr = formatDate_(new Date(data[i][2]));
      map[dateStr] = {
        id: String(data[i][0] || ''),
        pedidoId: String(data[i][1] || ''),
        date: dateStr,
        puntuacion: Number(data[i][6] || 0),
        comentario: String(data[i][7] || ''),
        platos: String(data[i][8] || ''),
        updatedAt: String(data[i][10] || '')
      };
    }
  }
  return map;
}

function getUserProviderRating_(email, periodId) {
  if (!email) return null;
  ensureRatingsSheets_();
  const sh = SpreadsheetApp.getActive().getSheetByName('ValoracionesProveedor');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const targetEmail = String(email).toLowerCase();
  const targetPeriodId = String(periodId || '').trim();

  for (let i = 1; i < data.length; i++) {
    const pId = String(data[i][1] || '').trim();
    const rowEmail = String(data[i][3] || '').toLowerCase();
    if (rowEmail === targetEmail && (!targetPeriodId || pId === targetPeriodId)) {
      return {
        id: String(data[i][0] || ''),
        periodId: pId,
        providerName: String(data[i][2] || ''),
        puntuacion: Number(data[i][6] || 0),
        comentario: String(data[i][7] || ''),
        version: Number(data[i][8] || 1),
        updatedAt: String(data[i][10] || '')
      };
    }
  }
  return null;
}

function apiSubmitMealRating(payload) {
  try {
    const activeUser = getUserInfo_();
    if (!activeUser || activeUser.estado !== 'ACTIVO') throw new Error("Usuario no autorizado.");

    const dateStr = String(payload.date || '').trim();
    if (!dateStr) throw new Error("Fecha de consumo requerida.");

    if (!isMealRatingAllowed_(dateStr)) {
      throw new Error("Solo puedes valorar la comida el mismo día después de las 12:00 PM o días posteriores.");
    }

    const rating = parseInt(payload.rating, 10);
    if (isNaN(rating) || rating < 1 || rating > 5) {
      throw new Error("La puntuación debe ser un número entre 1 y 5 estrellas.");
    }

    const comment = String(payload.comment || '').trim();
    if (comment.length > 500) {
      throw new Error("El comentario no puede exceder los 500 caracteres.");
    }

    ensureRatingsSheets_();
    const ss = SpreadsheetApp.getActive();

    // Verify user had an active order on this date
    const ordersData = readSheetValues_(ss.getSheetByName('Pedidos'), 9);
    let orderRow = null;
    for (let i = 1; i < ordersData.length; i++) {
      const rowDate = formatDate_(new Date(ordersData[i][2]));
      const rowEmail = String(ordersData[i][3] || '').toLowerCase();
      const rowStatus = String(ordersData[i][8] || '');
      if (rowEmail === activeUser.email.toLowerCase() && rowDate === dateStr && rowStatus !== 'CANCELADO') {
        orderRow = ordersData[i];
        break;
      }
    }

    if (!orderRow) {
      throw new Error("No se encontró un pedido activo para esta fecha.");
    }

    const orderId = String(orderRow[0] || '');
    const dishesSummary = String(orderRow[6] || '');
    const nowIso = new Date().toISOString();

    const sh = ss.getSheetByName('ValoracionesComida');
    const data = sh.getDataRange().getValues();
    let rowIndex = -1;
    let ratingId = '';

    for (let i = 1; i < data.length; i++) {
      const rEmail = String(data[i][3] || '').toLowerCase();
      const rDate = formatDate_(new Date(data[i][2]));
      if (rEmail === activeUser.email.toLowerCase() && rDate === dateStr) {
        rowIndex = i + 1;
        ratingId = String(data[i][0]);
        break;
      }
    }

    if (rowIndex > 0) {
      sh.getRange(rowIndex, 7).setValue(rating);
      sh.getRange(rowIndex, 8).setValue(comment);
      sh.getRange(rowIndex, 9).setValue(dishesSummary);
      sh.getRange(rowIndex, 11).setValue(nowIso);
    } else {
      ratingId = 'RAT_M_' + Utilities.getUuid();
      const newRow = [
        ratingId,
        orderId,
        dateStr,
        activeUser.email.toLowerCase(),
        activeUser.nombre,
        activeUser.departamento || '',
        rating,
        comment,
        dishesSummary,
        nowIso,
        nowIso
      ];
      sh.appendRow(newRow);
    }

    invalidateUserInitCache_();
    return {
      ok: true,
      rating: {
        id: ratingId,
        orderId: orderId,
        date: dateStr,
        puntuacion: rating,
        comentario: comment,
        platos: dishesSummary,
        updatedAt: nowIso
      },
      userMealRatings: getUserMealRatingsMap_(activeUser.email)
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSubmitProviderRating(payload) {
  try {
    const activeUser = getUserInfo_();
    if (!activeUser || activeUser.estado !== 'ACTIVO') throw new Error("Usuario no autorizado.");

    const rating = parseInt(payload.rating, 10);
    if (isNaN(rating) || rating < 1 || rating > 5) {
      throw new Error("La puntuación del proveedor debe ser un número entre 1 y 5 estrellas.");
    }

    const comment = String(payload.comment || '').trim();
    if (comment.length > 500) {
      throw new Error("El comentario no puede exceder los 500 caracteres.");
    }

    ensureRatingsSheets_();
    const provInfo = getProviderInfo_();
    const ss = SpreadsheetApp.getActive();
    const nowIso = new Date().toISOString();

    const shProv = ss.getSheetByName('ValoracionesProveedor');
    const data = shProv.getDataRange().getValues();
    let rowIndex = -1;
    let voteId = '';
    let currentVersion = 1;

    for (let i = 1; i < data.length; i++) {
      const pId = String(data[i][1] || '').trim();
      const uEmail = String(data[i][3] || '').toLowerCase();
      if (pId === provInfo.periodId && uEmail === activeUser.email.toLowerCase()) {
        rowIndex = i + 1;
        voteId = String(data[i][0]);
        currentVersion = Number(data[i][8] || 1) + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      shProv.getRange(rowIndex, 7).setValue(rating);
      shProv.getRange(rowIndex, 8).setValue(comment);
      shProv.getRange(rowIndex, 9).setValue(currentVersion);
      shProv.getRange(rowIndex, 11).setValue(nowIso);
    } else {
      voteId = 'RAT_P_' + Utilities.getUuid();
      const newRow = [
        voteId,
        provInfo.periodId,
        provInfo.name,
        activeUser.email.toLowerCase(),
        activeUser.nombre,
        activeUser.departamento || '',
        rating,
        comment,
        currentVersion,
        nowIso,
        nowIso
      ];
      shProv.appendRow(newRow);
    }

    // Append to HistoricoValoracionesProveedor for audit & trend analysis
    const shHist = ss.getSheetByName('HistoricoValoracionesProveedor');
    shHist.appendRow([
      voteId,
      provInfo.periodId,
      provInfo.name,
      activeUser.email.toLowerCase(),
      activeUser.nombre,
      activeUser.departamento || '',
      rating,
      comment,
      nowIso
    ]);

    invalidateUserInitCache_();
    return {
      ok: true,
      providerRating: {
        id: voteId,
        periodId: provInfo.periodId,
        providerName: provInfo.name,
        puntuacion: rating,
        comentario: comment,
        version: currentVersion,
        updatedAt: nowIso
      }
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiResetProviderPeriod(payload) {
  try {
    const admin = getUserInfo_();
    if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");

    const providerName = String(payload.providerName || '').trim();
    if (!providerName) throw new Error("El nombre del proveedor es requerido.");

    const todayStr = getTodayYmd_();
    const periodId = payload.periodId ? String(payload.periodId).trim() : ('PROV_' + todayStr.replace(/-/g, '_') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmm'));

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Config');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]);
      if (key === 'PROVIDER_NAME') sh.getRange(i + 1, 2).setValue(providerName);
      if (key === 'PROVIDER_PERIOD_ID') sh.getRange(i + 1, 2).setValue(periodId);
      if (key === 'PROVIDER_PERIOD_START') sh.getRange(i + 1, 2).setValue(todayStr);
    }

    _configCache = null;
    invalidateUserInitCache_();
    return {
      ok: true,
      providerInfo: {
        name: providerName,
        periodId: periodId,
        periodStart: todayStr
      }
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiGetRatingsSummary() {
  try {
    const user = getUserInfo_();
    if (!user || user.estado !== 'ACTIVO') {
      throw new Error("Debes iniciar sesión para consultar las valoraciones.");
    }

    ensureRatingsSheets_();
    const provInfo = getProviderInfo_();
    const ss = SpreadsheetApp.getActive();

    // 1. Provider ratings
    const provSh = ss.getSheetByName('ValoracionesProveedor');
    const provData = provSh.getDataRange().getValues();
    let provTotal = 0;
    let provSum = 0;
    const provStarsCount = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
    const recentProvFeedback = [];

    for (let i = 1; i < provData.length; i++) {
      const pId = String(provData[i][1] || '').trim();
      if (pId === provInfo.periodId) {
        const score = Number(provData[i][6]);
        if (score >= 1 && score <= 5) {
          provTotal++;
          provSum += score;
          provStarsCount[score] = (provStarsCount[score] || 0) + 1;
          recentProvFeedback.push({
            id: String(provData[i][0]),
            nombre: String(provData[i][4]),
            departamento: String(provData[i][5]),
            puntuacion: score,
            comentario: String(provData[i][7] || ''),
            version: Number(provData[i][8] || 1),
            fecha: String(provData[i][10] || '').substring(0, 10)
          });
        }
      }
    }

    const provAverage = provTotal > 0 ? (Math.round((provSum / provTotal) * 10) / 10) : 0;

    // 2. Meal ratings
    const mealSh = ss.getSheetByName('ValoracionesComida');
    const mealData = mealSh.getDataRange().getValues();
    let mealTotal = 0;
    let mealSum = 0;
    const mealStarsCount = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
    const recentMealFeedback = [];

    for (let i = 1; i < mealData.length; i++) {
      const score = Number(mealData[i][6]);
      if (score >= 1 && score <= 5) {
        mealTotal++;
        mealSum += score;
        mealStarsCount[score] = (mealStarsCount[score] || 0) + 1;
        recentMealFeedback.push({
          id: String(mealData[i][0]),
          fechaConsumo: formatDate_(new Date(mealData[i][2])),
          nombre: String(mealData[i][4]),
          departamento: String(mealData[i][5]),
          puntuacion: score,
          comentario: String(mealData[i][7] || ''),
          platos: String(mealData[i][8] || ''),
          fecha: String(mealData[i][10] || '').substring(0, 10)
        });
      }
    }

    const mealAverage = mealTotal > 0 ? (Math.round((mealSum / mealTotal) * 10) / 10) : 0;

    // Sort recent feedback by date desc
    recentProvFeedback.sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
    recentMealFeedback.sort((a, b) => (b.fechaConsumo || '').localeCompare(a.fechaConsumo || ''));

    return {
      ok: true,
      providerStats: {
        providerInfo: provInfo,
        total: provTotal,
        average: provAverage,
        starsCount: provStarsCount,
        recentFeedback: recentProvFeedback.slice(0, 50)
      },
      mealStats: {
        total: mealTotal,
        average: mealAverage,
        starsCount: mealStarsCount,
        recentFeedback: recentMealFeedback.slice(0, 50)
      }
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}
// === AUTOMATIZACIÓN (TRIGGERS) ===

function apiRecordChefGameEvent(payload) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(3000)) {
      throw new Error("El juego esta ocupado guardando otra jugada. Intenta de nuevo.");
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Usuarios');
    const colMap = ensureChefGameColumns_(sh);
    const usersData = sh.getDataRange().getValues();
    const deptMap = getDepartmentMap_();
    const activeUser = getUserInfo_(null, usersData, deptMap);
    if (!activeUser || activeUser.estado !== 'ACTIVO') throw new Error("Usuario no autorizado.");

    const targetUser = resolveChefGameTargetUser_(payload && payload.targetEmail, activeUser, usersData, deptMap);
    const context = getChefGameUserContext_(targetUser.email, sh, colMap, usersData);
    if (!context) throw new Error("Usuario no encontrado.");

    const eventType = normalizeChefGameEventType_(payload && payload.type);
    const now = new Date();
    const state = normalizeChefGameState_(context.row, colMap, now);
    const event = applyChefGameEvent_(state, eventType, payload, now);

    writeChefGameState_(sh, context.rowIndex, colMap, state, now);

    return {
      ok: true,
      state: state,
      event: event
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // Ignore lock release errors.
    }
  }
}

function apiGetChefGameRanking(targetEmail) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Usuarios');
    const colMap = ensureChefGameColumns_(sh);
    const data = sh.getDataRange().getValues();
    const deptMap = getDepartmentMap_();
    const activeUser = getUserInfo_(null, data, deptMap);
    if (!activeUser || activeUser.estado !== 'ACTIVO') throw new Error("Usuario no autorizado.");

    const targetUser = resolveChefGameTargetUser_(targetEmail, activeUser, data, deptMap);
    const now = new Date();
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4] || '').trim().toUpperCase() !== 'ACTIVO') continue;
      const game = normalizeChefGameState_(data[i], colMap, now);
      rows.push({
        _email: String(data[i][0] || '').toLowerCase(),
        nombre: data[i][1] || 'Usuario',
        departamento: deptMap[data[i][2]] || data[i][2] || '',
        score: game.score,
        hits: game.hits,
        misses: game.misses,
        bestStreak: game.bestStreak
      });
    }

    rows.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      if (b.hits !== a.hits) return b.hits - a.hits;
      if (a.misses !== b.misses) return a.misses - b.misses;
      return String(a.nombre).localeCompare(String(b.nombre));
    });

    rows.forEach((row, index) => {
      row.rank = index + 1;
      row.isCurrent = row._email === String(targetUser.email || '').toLowerCase();
      delete row._email;
    });

    const currentRow = rows.find(row => row.isCurrent) || null;
    return {
      ok: true,
      monthKey: getChefGameMonthKey_(now),
      monthLabel: getChefGameMonthLabel_(now),
      rows: rows.slice(0, 50),
      currentUser: currentRow
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSyncChefGameState(payload) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(3000)) {
      throw new Error("No se pudo guardar el juego ahora. Intenta de nuevo.");
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Usuarios');
    const colMap = ensureChefGameColumns_(sh);
    const usersData = sh.getDataRange().getValues();
    const deptMap = getDepartmentMap_();
    const activeUser = getUserInfo_(null, usersData, deptMap);
    if (!activeUser || activeUser.estado !== 'ACTIVO') throw new Error("Usuario no autorizado.");

    const targetUser = resolveChefGameTargetUser_(payload && payload.targetEmail, activeUser, usersData, deptMap);
    const context = getChefGameUserContext_(targetUser.email, sh, colMap, usersData);
    if (!context) throw new Error("Usuario no encontrado.");

    const now = new Date();
    const previousState = normalizeChefGameState_(context.row, colMap, now);
    const state = sanitizeChefGameSubmittedState_(payload && payload.state, now, previousState);
    writeChefGameState_(sh, context.rowIndex, colMap, state, now);
    return { ok: true, state: state };
  } catch (e) {
    return { ok: false, msg: e.message };
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // Ignore lock release errors.
    }
  }
}

function scheduledSendReminders() {
  // Only run on business days
  if (!isTodayBusinessDay_()) {
     console.log("Skipping scheduledSendReminders: Not a business day.");
     return;
  }

  const nextBusinessDay = getNextBusinessDay_(new Date());
  if (!nextBusinessDay) return;

  const dateStr = formatDate_(nextBusinessDay);

  // -- Check if Menu exists for target date --
  const mSh = SpreadsheetApp.getActive().getSheetByName('Menu');
  const mData = mSh.getDataRange().getValues();
  let hasMenu = false;
  for(let i=1; i<mData.length; i++) {
     if (formatDate_(new Date(mData[i][1])) === dateStr && String(mData[i][5]) === 'SI') {
        hasMenu = true;
        break;
     }
  }

  if (!hasMenu) {
     console.log(`Skipping reminders. No menu found for ${dateStr}.`);
     return;
  }

  const ss = SpreadsheetApp.getActive();
  const uSh = ss.getSheetByName('Usuarios');
  const uData = uSh.getDataRange().getValues();
  const pSh = ss.getSheetByName('Pedidos');
  const pData = pSh.getDataRange().getValues();

  const orderedEmails = new Set();
  for (let i = 1; i < pData.length; i++) {
    const rowDate = formatDate_(new Date(pData[i][2]));
    if (rowDate === dateStr && pData[i][8] !== 'CANCELADO') {
      orderedEmails.add(String(pData[i][3]).toLowerCase());
    }
  }

  uData.slice(1).forEach(row => {
    const email = String(row[0]).toLowerCase();
    const estado = row[4];
    const prefs = JSON.parse(row[5] || '{}');
    if (estado === 'ACTIVO' && !orderedEmails.has(email) && prefs.reminders !== false) {
      // Calculate closing time for display
      const envio = getConfigValue_('HORA_ENVIO') || '15:00';
      const mins = parseInt(getConfigValue_('MINUTOS_PREV_CIERRE') || '30', 10);

      // Calculate exact closing time for display
      let h = 15, m = 0;
      if (envio instanceof Date) { h = envio.getHours(); m = envio.getMinutes(); }
      else { const p = String(envio).split(':'); h = parseInt(p[0]||15); m = parseInt(p[1]||0); }

      const limitDate = new Date();
      limitDate.setHours(h, m, 0, 0);
      limitDate.setMinutes(limitDate.getMinutes() - mins);
      const limitStr = Utilities.formatDate(limitDate, Session.getScriptTimeZone(), 'hh:mm a');

      const appUrl = getAppUrl_();
      const userName = row[1] ? row[1].split(' ')[0] : 'Colaborador'; // First name

      const html = getEmailTemplate_({
         title: 'Recordatorio de Almuerzo',
         body: `
           <p>Hola <strong>${userName}</strong>,</p>
           <p>¿No pedirás nada? Hasta ahora no hemos recibido tu selección de almuerzo para el día de mañana (<b>${formatDisplayDate_(dateStr)}</b>).</p>
           <p>Si comerás aquí, por favor revisa la hoja de solicitudes.</p>
           <p style="background-color: #fff7ed; padding: 12px; border-left: 4px solid #f97316; margin: 16px 0; font-size: 14px; color: #9a3412;">
             ⚠️ Tienes hasta las <strong>${limitStr}</strong> de hoy para hacer tu pedido.
           </p>
           <p style="font-size: 12px; color: #6b7280; margin-top: 24px;">(Este es un mensaje automático, no es necesario responder).</p>
         `,
         cta: { text: 'Abrir App de Almuerzo', url: appUrl },
         footerNote: 'Si ya no deseas recibir estos recordatorios, puedes desactivarlos en la configuración de la App dando clic en el botón de notificaciones (🔔).'
      });

      sendEmail_(email, "Almuerzo Pre-empacado | Recordatorio de pedido", html);
    }
  });
}

function scheduledDailyClose() {
  const testRun = isTestEmailMode_();

  if (!testRun && !isTodayBusinessDay_()) {
     console.log("Skipping scheduledDailyClose: Not a business day.");
     return;
  }

  try {
    if (testRun) validateTestEmailConfig_();
    const dateStr = getDailyCloseTargetDate_();
    if (!dateStr) return;

    runDailyCloseReportEmails_(dateStr, {
      testRun: testRun,
      includeMaintenance: !testRun,
      source: 'scheduled'
    });
  } catch (e) {
    console.error("scheduledDailyClose error: " + e.message);
  }
}

function apiSendDailyCloseEmailsTest() {
  try {
    const admin = getUserInfo_();
    if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");

    validateTestEmailConfig_();
    const dateStr = getDailyCloseTargetDate_();
    if (!dateStr) throw new Error("No se encontro una fecha habil para probar.");

    return runDailyCloseReportEmails_(dateStr, {
      testRun: true,
      includeMaintenance: false,
      source: 'admin-ui'
    });
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function testSendDailyCloseEmails() {
  validateTestEmailConfig_();
  const dateStr = getDailyCloseTargetDate_();
  if (!dateStr) throw new Error("No se encontro una fecha habil para probar.");
  return runDailyCloseReportEmails_(dateStr, {
    testRun: true,
    includeMaintenance: false,
    source: 'manual'
  });
}

function getDailyCloseTargetDate_() {
  const nextBusinessDay = getNextBusinessDay_(new Date());
  return nextBusinessDay ? formatDate_(nextBusinessDay) : '';
}

function isTestEmailMode_() {
  return String(getConfigValue_('TEST_EMAIL_MODE') || '').trim().toUpperCase() === 'TRUE';
}

function validateTestEmailConfig_() {
  if (!isTestEmailMode_()) throw new Error("El modo prueba de correos no esta activo.");
  if (!String(getConfigValue_('TEST_EMAIL_DEST') || '').trim()) {
    throw new Error("Configura TEST_EMAIL_DEST antes de enviar correos de prueba.");
  }
}

function runDailyCloseReportEmails_(dateStr, options) {
  const opts = options || {};
  const testRun = opts.testRun === true;
  const normalizedDate = normalizeReportDate_(dateStr);
  const ss = SpreadsheetApp.getActive();
  const deptMap = getDepartmentMap_();
  const ordersData = readSheetValues_(ss.getSheetByName('Pedidos'), 11);
  const usersData = readSheetValues_(ss.getSheetByName('Usuarios'), 7);
  const codeMap = getUserCodeMap_(usersData);
  const orders = sortOrdersForGeneralReport_(getOrdersByDateDetailed_(normalizedDate, {
    ordersData: ordersData,
    deptMap: deptMap,
    codeMap: codeMap
  }));
  const byDept = groupOrdersByDepartment_(orders);
  const deptSummary = getDepartmentOrderSummary_(byDept, deptMap);
  const backupFolder = (!testRun && orders.length > 0) ? getDailyBackupFolder_(normalizedDate) : null;
  const deptAdminsMap = getDepartmentAdminsMap_(usersData);
  const formattedDate = Utilities.formatDate(new Date(normalizedDate + 'T12:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const fileDate = Utilities.formatDate(new Date(normalizedDate + 'T12:00:00'), Session.getScriptTimeZone(), 'dd-MM-yyyy');

  let departmentEmailsSent = 0;
  let departmentReportsSaved = 0;

  deptSummary.forEach(summary => {
    const deptId = summary.id;
    const deptName = summary.name;
    const deptOrders = byDept[deptId] || [];
    const recipients = getDepartmentReportRecipients_(deptId, deptAdminsMap);
    const fileName = `[${deptName} - ${fileDate}]`;
    let tempSS = null;

    try {
      tempSS = createReportFromTemplate_(deptName, normalizedDate, deptOrders);

      if (!testRun && backupFolder) {
        const pdfBlob = exportSheetToPdfBlob_(tempSS);
        pdfBlob.setName(`${fileName}.pdf`);
        backupFolder.createFile(pdfBlob);
        departmentReportsSaved++;
      }

      if (recipients.to || recipients.cc) {
        const excelBlob = exportSheetToExcelBlob_(tempSS);
        excelBlob.setName(`${fileName}.xlsx`);

        const toList = recipients.to || recipients.cc;
        const ccList = recipients.to ? recipients.cc : '';
        const html = getEmailTemplate_({
           title: `Reporte ${deptName}`,
           subtitle: `Pedidos para el ${formattedDate}`,
           body: `
             <p>Buenas tardes estimados,</p>
             <p>Hay <strong>${deptOrders.length}</strong> pedidos registrados del departamento de <strong>${escapeHtml_(deptName)}</strong> para el dia <strong>${formattedDate}</strong>.</p>
             <p>Favor revisar el archivo Excel adjunto para mas detalles sobre los platos solicitados.</p>
             <p>Cualquier duda, estamos a la orden.</p>
           `,
           footerNote: testRun ? 'Correo de prueba. No se genero respaldo ni cierre real.' : 'Este reporte se genera automaticamente al cierre de pedidos.'
        });

        sendEmail_(toList, `Almuerzo Pre-empacado | Reporte Almuerzo ${deptName} - ${normalizedDate}`, html, ccList, [excelBlob]);
        departmentEmailsSent++;
      } else {
        console.warn(`No recipients found for department ${deptName} (${deptId}).`);
      }
    } catch(e) {
      console.error(`Error processing report for ${deptName}: ${e.message}`);
    } finally {
      trashTempSpreadsheet_(tempSS);
    }
  });

  let generalExcelBlob = null;
  let generalReportSaved = false;
  if (orders.length > 0) {
    try {
      const generalArtifacts = createGeneralReportArtifactsWithRetry_(normalizedDate, fileDate, orders, byDept, deptSummary, backupFolder, testRun);
      generalExcelBlob = generalArtifacts.excelBlob;
      generalReportSaved = generalArtifacts.reportSaved;
    } catch (e) {
      console.error(`Error processing general report after retries: ${e.message}`);
    }
  }

  if (orders.length > 0 && String(getConfigValue_('ADMIN_EMAILS') || '').trim() && !generalExcelBlob) {
    throw new Error("No se pudo generar el Excel consolidado del resumen general. No se envio el resumen administrativo sin adjunto.");
  }

  const adminSummarySent = sendDailyAdminSummary_(normalizedDate, {
    orders: orders,
    deptSummary: deptSummary,
    attachments: generalExcelBlob ? [generalExcelBlob] : [],
    copyRecipients: getDailySummaryCopyRecipients_(),
    testRun: testRun
  });

  if (opts.includeMaintenance !== false && !testRun) {
    checkMenuIntegrity_();
  }

  return {
    ok: true,
    testRun: testRun,
    date: normalizedDate,
    orderCount: orders.length,
    departmentCount: deptSummary.length,
    departmentEmailsSent: departmentEmailsSent,
    departmentReportsSaved: departmentReportsSaved,
    generalReportSaved: generalReportSaved,
    adminSummarySent: adminSummarySent,
    msg: testRun
      ? `Prueba enviada para ${formattedDate}. No se guardaron respaldos ni se ejecuto mantenimiento.`
      : `Reportes enviados para ${formattedDate}.`
  };
}

// === ADMIN API ===

function apiGetAdminData() {
  try {
    const user = getUserInfo_();
    if (!user || (user.rol !== 'ADMIN_GEN' && user.rol !== 'ADMIN_DEP')) {
      return { ok: false, msg: "Acceso denegado." };
    }

    const cacheKey = getAdminCacheKey_(user);
    const cached = readJsonCache_(cacheKey);
    if (cached && cached.ok) return cached;

    const data = { ok: true, rol: user.rol, dept: user.departamentoId }; // Send ID
    const deptMap = getDepartmentMap_();

    // Users
    const uSh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
    data.users = uSh.getDataRange().getValues().slice(1).map(r => ({
      email: r[0], nombre: r[1],
      departamentoId: r[2], departamento: deptMap[r[2]] || r[2], // Resolve for display
      rol: r[3], estado: r[4], codigo: r[6] || ''
    })).filter(u => user.rol === 'ADMIN_GEN' || (u.departamentoId === user.departamentoId));

    // Orders
    const pSh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
    data.orders = pSh.getDataRange().getValues().slice(1)
      .filter(r => {
         if (!r[2]) return false;
         if (String(r[8]).toUpperCase() === 'CANCELADO') return false;
         try {
            const d = new Date(r[2]);
            if (isNaN(d.getTime())) return false;
            const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - 60);
            return d >= cutoff;
         } catch(e) { return false; }
      })
      .map(r => ({
         id: r[0], date: formatDate_(new Date(r[2])), email: r[3], nombre: r[4],
         deptId: r[5], dept: deptMap[r[5]] || r[5], // Resolve
         resumen: r[6], estado: r[8], creado_por: r[10] || ''
      }))
      .filter(o => user.rol === 'ADMIN_GEN' || (o.deptId === user.departamentoId));

    // Config & Holidays (Admin Gen only)
    if (user.rol === 'ADMIN_GEN') {
       ensureOperationalConfigKeys_();
       ensureConfigKey_('APP_URL', ScriptApp.getService().getUrl(), 'URL pública de la aplicación (Web App)');
       data.config = getConfigValue_('ALL');
       for (const k in data.config) {
          const val = data.config[k];
          if (val instanceof Date) {
             // Check if it's likely a time (Year 1899)
             if (val.getFullYear() === 1899) {
                data.config[k] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
             } else {
                data.config[k] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
             }
          } else {
             data.config[k] = String(val);
          }
       }
       data.configKeys = Object.keys(data.config);

       // Get descriptions
       const cSh = SpreadsheetApp.getActive().getSheetByName('Config');
       if (cSh) {
          data.configList = cSh.getDataRange().getValues().slice(1).map(r => ({
             key: String(r[0]), value: String(r[1]), desc: r[2]
          }));
       }

       data.holidays = getHolidaysList_();
       data.menuCategories = getMenuCategories_(true);
    }

    data.departments = getDepartmentsList_(); // Returns {id, nombre...}

    writeJsonCache_(cacheKey, data, 45);
    return data;
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSaveMenuCategory(categoryData) {
  try {
    const admin = getUserInfo_();
    if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error('Permiso denegado.');

    const name = normalizeMenuText_(categoryData && categoryData.nombre);
    const state = String(categoryData && categoryData.estado || 'ACTIVO').trim().toUpperCase();
    const order = Number(categoryData && categoryData.orden);
    if (!name) throw new Error('El nombre de la categoria es obligatorio.');
    if (name.length > 80) throw new Error('El nombre de la categoria no puede superar 80 caracteres.');
    if (!Number.isInteger(order) || order < 0 || order > 9999) throw new Error('El orden debe ser un numero entero entre 0 y 9999.');
    if (state !== 'ACTIVO' && state !== 'INACTIVO') throw new Error('El estado de la categoria no es valido.');

    const sheet = ensureMenuCategoriesSheet_();
    const data = readSheetValues_(sheet, 8);
    const id = String(categoryData && categoryData.id || '').trim();
    const nameKey = normalizeCategoryLookupKey_(name);
    const idKey = normalizeCategoryLookupKey_(id);
    const aliases = parseMenuCategoryAliases_(categoryData && (categoryData.aliasesText !== undefined ? categoryData.aliasesText : categoryData.aliases))
      .filter(alias => {
        const key = normalizeCategoryLookupKey_(alias);
        return key !== nameKey && (!idKey || key !== idKey);
      });
    if (aliases.length > 10 || aliases.some(alias => alias.length > 80)) {
      throw new Error('Puedes definir hasta 10 alias de 80 caracteres cada uno.');
    }

    const esCombinable = categoryData && (categoryData.es_combinable === false || String(categoryData.es_combinable).toUpperCase() === 'NO') ? 'NO' : 'SI';
    const combinableConList = parseCategoryCombinableWith_(categoryData && categoryData.combinable_con);
    const tipoSeleccion = String(categoryData && categoryData.tipo_seleccion || 'UNICA').trim().toUpperCase() === 'MULTIPLE' ? 'MULTIPLE' : 'UNICA';

    const candidateKeys = [name].concat(aliases).map(normalizeCategoryLookupKey_);
    let rowIndex = 0;

    for (let i = 1; i < data.length; i++) {
      const currentId = String(data[i][0] || '').trim();
      const currentName = normalizeMenuText_(data[i][1]);
      if (currentId === id) rowIndex = i + 1;
      if (currentId !== id) {
        const currentKeys = [currentId, currentName].concat(parseMenuCategoryAliases_(data[i][4])).map(normalizeCategoryLookupKey_);
        if (candidateKeys.some(key => currentKeys.indexOf(key) !== -1)) {
          throw new Error('El nombre o uno de los alias ya pertenece a otra categoria.');
        }
      }
    }

    const savedId = rowIndex ? id : 'CAT_' + Utilities.getUuid();
    const row = [savedId, name, order, state, aliases.join(', '), esCombinable, combinableConList.join(', '), tipoSeleccion];
    if (rowIndex) sheet.getRange(rowIndex, 1, 1, 8).setValues([row]);
    else sheet.appendRow(row);

    invalidateMenuCategoriesCache_();
    return {
      ok: true,
      category: {
        id: savedId,
        nombre: name,
        orden: order,
        estado: state,
        aliases: aliases,
        es_combinable: esCombinable === 'SI',
        combinable_con: combinableConList,
        tipo_seleccion: tipoSeleccion
      },
      categories: getMenuCategories_(true)
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSaveConfig(configData) {
   try {
     const admin = getUserInfo_();
     if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");

     ensureOperationalConfigKeys_();
     if (configData.RESPONSIBLES_EMAILS_JSON !== undefined) {
        const summaryCopies = normalizeDailySummaryCopyRecipients_(configData.RESPONSIBLES_EMAILS_JSON, { throwOnInvalid: true });
        configData.RESPONSIBLES_EMAILS_JSON = JSON.stringify(summaryCopies);
     }

     const ss = SpreadsheetApp.getActive();
     const sh = ss.getSheetByName('Config');
     const data = sh.getDataRange().getValues();

     let timeChanged = false;
     const currentMealPrice = getCurrentMealPrice_();
     const currentMealPriceHistory = getMealPriceHistory_();
     const hasIncomingMealPrice = configData.MEAL_PRICE_CURRENT !== undefined;
     const normalizedMealPrice = hasIncomingMealPrice ? normalizeMealPriceValue_(configData.MEAL_PRICE_CURRENT) : currentMealPrice;
     const mealPriceChanged = hasIncomingMealPrice && normalizedMealPrice !== currentMealPrice;
     let mealPriceHistoryRow = -1;

     for(let i=1; i<data.length; i++) {
        const key = String(data[i][0]);
        if (key === 'MEAL_PRICE_HISTORY_JSON') {
           mealPriceHistoryRow = i + 1;
           continue;
        }
        if (configData[key] !== undefined) {
           const val = key === 'MEAL_PRICE_CURRENT' ? normalizedMealPrice : configData[key];
           if ((key === 'HORA_RECORDATORIO' || key === 'HORA_ENVIO') && String(data[i][1]) !== String(val)) {
              timeChanged = true;
           }
           sh.getRange(i+1, 2).setValue(val);
        }
     }

     if (mealPriceChanged && mealPriceHistoryRow > 0) {
        const nextHistory = upsertMealPriceHistory_(currentMealPriceHistory, normalizedMealPrice, getTodayYmd_());
        sh.getRange(mealPriceHistoryRow, 2).setValue(JSON.stringify(nextHistory));
     }
     _configCache = null;

     if (timeChanged) {
        reinstallTimeTriggers_();
     }

     invalidateMenuDataCache_();
     invalidateUserInitCache_();
     return { ok: true };
   } catch (e) { return { ok: false, msg: e.message }; }
}

function apiSaveDepartment(dept) {
   try {
     const admin = getUserInfo_();
     if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");

     const ss = SpreadsheetApp.getActive();
     let sh = ss.getSheetByName('Departamentos');
     if (!sh) { sh = ss.insertSheet('Departamentos'); }

     const data = sh.getDataRange().getValues();

     // Check for duplicate name
     const normName = dept.nombre.trim().toLowerCase();
     for(let i=1; i<data.length; i++) {
        if (String(data[i][1]).trim().toLowerCase() === normName && String(data[i][0]) !== String(dept.id)) {
           throw new Error("Ya existe un departamento con ese nombre.");
        }
     }

     let rowIdx = -1;
     if (dept.id) {
        for(let i=1; i<data.length; i++) {
           if (String(data[i][0]) === String(dept.id)) { rowIdx = i+1; break; }
        }
     }

     const id = rowIdx > 0 ? dept.id : Utilities.getUuid();
     const rowContent = [id, dept.nombre, dept.admins, dept.estado || 'ACTIVO', '{}'];

     if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, rowContent.length).setValues([rowContent]);
     else sh.appendRow(rowContent);

     // Update User Roles
     if (dept.admins) {
        const uSh = ss.getSheetByName('Usuarios');
        const uData = uSh.getDataRange().getValues();
        const emails = dept.admins.split(',').map(e => e.trim().toLowerCase()).filter(e => e);

        // Map users by email
        const userMap = {};
        for(let i=1; i<uData.length; i++) userMap[String(uData[i][0]).toLowerCase()] = i + 1;

        emails.forEach(email => {
           const row = userMap[email];
           if (row) {
              const currentRol = uData[row-1][3];
              if (currentRol === 'ADMIN_GEN') {
                 // Skip
              } else {
                 // Update to ADMIN_DEP and set Dept ID
                 uSh.getRange(row, 3, 1, 2).setValues([[id, 'ADMIN_DEP']]);
              }
           }
        });

        // Remove these admins from any OTHER department to maintain consistency
        const dData = sh.getDataRange().getValues();
        for(let i=1; i<dData.length; i++) {
           if (String(dData[i][0]) === String(id)) continue; // Skip current

           let dAdmins = String(dData[i][2]).split(',').map(e => e.trim()).filter(e => e);
           const originalLen = dAdmins.length;

           dAdmins = dAdmins.filter(e => !emails.includes(e.toLowerCase()));

           if (dAdmins.length !== originalLen) {
              sh.getRange(i+1, 3).setValue(dAdmins.join(', '));
           }
        }
     }

     return { ok: true };
   } catch (e) { return { ok: false, msg: e.message }; }
}

function apiDeleteDepartment(deptId) {
   try {
     const admin = getUserInfo_();
     if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");
     const ss = SpreadsheetApp.getActive();
     const sh = ss.getSheetByName('Departamentos');
     const data = sh.getDataRange().getValues();
     for(let i=1; i<data.length; i++) {
        if (String(data[i][0]) === String(deptId)) {
           sh.deleteRow(i+1);
           invalidateUserInitCache_();
           return { ok: true };
        }
     }
     return { ok: false, msg: "No encontrado" };
   } catch (e) { return { ok: false, msg: e.message }; }
}

function apiAdminSaveUser(userData) {
  const admin = getUserInfo_();
  if (!admin || !['ADMIN_GEN', 'ADMIN_DEP'].includes(admin.rol)) throw new Error("Denegado");

  // If Admin Dep, force dept ID
  if (admin.rol === 'ADMIN_DEP') {
     userData.departamento = admin.departamentoId;
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Usuarios');
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  let prevStatus = null;

  for(let i=1; i<data.length; i++) {
     if (String(data[i][0]).toLowerCase() === String(userData.email).toLowerCase()) {
        rowIdx = i+1;
        prevStatus = data[i][4];
        break;
     }
  }

  // Validate Code
  if (!userData.codigo || !/^\d{4}$/.test(userData.codigo)) throw new Error("El código es obligatorio y debe tener 4 dígitos.");

  // Uniqueness Check
  for(let i=1; i<data.length; i++) {
     if (i+1 !== rowIdx) { // Skip self
        if (String(data[i][6]) === String(userData.codigo)) {
           throw new Error("El código " + userData.codigo + " ya pertenece a otro usuario.");
        }
     }
  }

  const rowContent = [
     userData.email.toLowerCase(),
     userData.nombre,
     userData.departamento, // This assumes ID is passed
     userData.rol || 'USER',
     userData.estado || 'ACTIVO',
     (rowIdx > 0 ? data[rowIdx-1][5] : '{}'),
     userData.codigo
  ];

  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, rowContent.length).setValues([rowContent]);
  else sh.appendRow(rowContent);

  // Send Notification if Activated
  if (userData.estado === 'ACTIVO' && prevStatus !== 'ACTIVO') {
     SpreadsheetApp.flush();
     const html = getEmailTemplate_({
        title: '¡Bienvenido!',
        subtitle: 'Acceso Aprobado',
        body: `
          <p>Hola <strong>${userData.nombre}</strong>,</p>
          <p>Tu cuenta ha sido activada exitosamente.</p>
          <p>Ya puedes ingresar al sistema para realizar tus pedidos de almuerzo.</p>
        `,
        cta: { text: 'Ingresar a la App', url: getAppUrl_() }
     });
     sendEmail_(userData.email, "Almuerzo Pre-empacado | Acceso Aprobado", html);
  }

  invalidateUserInitCache_();
  return { ok: true };
}

function apiAdminDeleteUser(email) {
   const admin = getUserInfo_();
   if (!admin || !['ADMIN_GEN', 'ADMIN_DEP'].includes(admin.rol)) throw new Error("Permiso denegado.");
   const ss = SpreadsheetApp.getActive();
   const sh = ss.getSheetByName('Usuarios');
   const data = sh.getDataRange().getValues();
   for(let i=1; i<data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(email).toLowerCase()) {
         if (admin.rol === 'ADMIN_DEP' && data[i][2] !== admin.departamentoId) throw new Error("Denegado");
         sh.getRange(i+1, 5).setValue('INACTIVO');
         invalidateUserInitCache_();
         return { ok: true };
      }
   }
   return { ok: false, msg: "No encontrado" };
}

function apiAdminCancelOrder(orderId) {
  const usersData = readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Usuarios'), 7);
  const admin = getUserAccessRecord_(null, usersData);
  if (!admin || !['ADMIN_GEN', 'ADMIN_DEP'].includes(admin.rol)) throw new Error("Permiso denegado.");

  const result = cancelOrderRecordById_(orderId, function(orderSnapshot) {
    return admin.rol === 'ADMIN_GEN' || orderSnapshot.departamentoId === admin.departamentoId;
  });

  if (!result.found) return { ok: false, msg: "Pedido no encontrado." };
  if (!result.allowed) throw new Error("Denegado: Pedido de otro departamento.");

  invalidateUserInitCache_();
  return { ok: true };
}

function isFutureDateString_(dateStr) {
  const d = new Date(String(dateStr || '') + 'T12:00:00');
  if (isNaN(d.getTime())) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(d);
  target.setHours(0, 0, 0, 0);
  return target > today;
}

function safeParseJsonObject_(raw) {
  if (!raw) return {};
  try {
    const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch(e) {
    return {};
  }
}

function getFirstName_(name) {
  const first = String(name || '').trim().split(/\s+/)[0];
  return first || 'Colaborador';
}

function getMenuSelectionKey_(cat, plato) {
  return String(cat || '').trim() + '\u0001' + normalizeMenuText_(plato);
}

function getMenuItemSnapshotFromRow_(row) {
  if (!row || !row[1]) return null;

  let dateStr = '';
  try {
    dateStr = formatMenuRowDate_(row[1]);
  } catch(e) {
    return null;
  }

  const cat = String(row[2] || '').trim();
  const plato = normalizeMenuText_(row[3]);
  if (!dateStr || !cat || !plato) return null;

  return {
    id: row[0],
    date: dateStr,
    cat: cat,
    plato: plato,
    desc: normalizeMenuText_(row[4]),
    enabled: String(row[5] || '').trim().toUpperCase() === 'SI',
    key: getMenuSelectionKey_(cat, plato)
  };
}

function createMenuDateSnapshot_() {
  return { items: [], keys: {} };
}

function addMenuItemToSnapshot_(snapshot, item) {
  if (!snapshot || !item || !item.enabled) return;
  snapshot.items.push(item);
  if (!snapshot.keys[item.key]) snapshot.keys[item.key] = [];
  snapshot.keys[item.key].push(item);
}

function buildActiveMenuSnapshotByDate_(menuData, targetDates) {
  const result = {};
  const datesSet = targetDates ? new Set(Array.from(targetDates)) : null;

  for (let i = 1; i < menuData.length; i++) {
    const item = getMenuItemSnapshotFromRow_(menuData[i]);
    if (!item || !item.enabled) continue;
    if (datesSet && !datesSet.has(item.date)) continue;
    if (!result[item.date]) result[item.date] = createMenuDateSnapshot_();
    addMenuItemToSnapshot_(result[item.date], item);
  }

  return result;
}

function createMenuItemFromPayload_(dateStr, itemData) {
  const cat = String(itemData && itemData.cat || '').trim();
  const plato = normalizeMenuText_(itemData && itemData.plato);
  if (!dateStr || !cat || !plato) return null;

  return {
    date: dateStr,
    cat: cat,
    plato: plato,
    desc: normalizeMenuText_(itemData && itemData.desc),
    enabled: true,
    key: getMenuSelectionKey_(cat, plato)
  };
}

function isMenuItemChanged_(oldItem, newItem) {
  if (!oldItem || !newItem) return false;
  return oldItem.date !== newItem.date ||
    oldItem.cat !== newItem.cat ||
    oldItem.plato !== newItem.plato ||
    oldItem.desc !== newItem.desc;
}

function addAffectedMenuKey_(affectedKeyMap, item) {
  if (!affectedKeyMap || !item || !item.key) return;
  affectedKeyMap[item.key] = {
    cat: item.cat,
    plato: item.plato
  };
}

function getAffectedMenuKeysByReplacement_(oldSnapshot, newSnapshot) {
  const affectedKeyMap = {};
  if (!oldSnapshot || !oldSnapshot.items || oldSnapshot.items.length === 0) return affectedKeyMap;

  const nextSnapshot = newSnapshot || createMenuDateSnapshot_();
  oldSnapshot.items.forEach(oldItem => {
    const candidates = nextSnapshot.keys[oldItem.key] || [];
    const unchanged = candidates.some(newItem => newItem.desc === oldItem.desc);
    if (!unchanged) addAffectedMenuKey_(affectedKeyMap, oldItem);
  });

  return affectedKeyMap;
}

function hasAffectedMenuKeys_(affectedKeyMap) {
  return affectedKeyMap && Object.keys(affectedKeyMap).length > 0;
}

function getOrderMenuKeyDetails_(detail) {
  const normalized = normalizeOrderDetail_(detail || {});
  const cats = Array.isArray(normalized.categorias) ? normalized.categorias : [];
  const items = Array.isArray(normalized.items) ? normalized.items : [];
  const result = [];
  const total = Math.max(cats.length, items.length);

  for (let i = 0; i < total; i++) {
    const cat = String(cats[i] || '').trim();
    const item = normalizeMenuText_(items[i]);
    if (!cat || !item) continue;
    result.push({
      key: getMenuSelectionKey_(cat, item),
      cat: cat,
      plato: item
    });
  }

  return result;
}

function uniqueStrings_(values) {
  const seen = {};
  const result = [];
  (values || []).forEach(value => {
    const normalized = String(value || '').trim();
    if (normalized && !seen[normalized]) {
      seen[normalized] = true;
      result.push(normalized);
    }
  });
  return result;
}

function getAffectedOrderItems_(detail, resumen, affectedKeyMap) {
  if (!hasAffectedMenuKeys_(affectedKeyMap)) return [];
  const matches = [];

  getOrderMenuKeyDetails_(detail).forEach(item => {
    if (affectedKeyMap[item.key]) matches.push(item.plato);
  });

  if (matches.length === 0 && resumen) {
    const normalizedSummary = normalizeMenuText_(resumen);
    Object.keys(affectedKeyMap).forEach(key => {
      const affectedItem = affectedKeyMap[key];
      if (affectedItem && affectedItem.plato && normalizedSummary.indexOf(affectedItem.plato) !== -1) {
        matches.push(affectedItem.plato);
      }
    });
  }

  return uniqueStrings_(matches);
}

function cancelActiveOrdersForPlans_(plans) {
  const validPlans = (plans || []).filter(plan => plan && plan.date && (plan.cancelAll || hasAffectedMenuKeys_(plan.affectedKeyMap)));
  if (validPlans.length === 0) return [];

  const planByDate = {};
  validPlans.forEach(plan => {
    planByDate[plan.date] = plan;
  });

  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = readSheetValues_(sh, 11);
  const affectedOrders = [];
  const statusRanges = [];
  const timestampRanges = [];
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    let rowDate = '';
    try {
      rowDate = formatDate_(new Date(data[i][2]));
    } catch(e) {
      continue;
    }

    const plan = planByDate[rowDate];
    if (!plan) continue;
    if (String(data[i][8] || '').trim().toUpperCase() === 'CANCELADO') continue;

    const detail = safeParseJsonObject_(data[i][7]);
    const affectedItems = plan.cancelAll
      ? []
      : getAffectedOrderItems_(detail, data[i][6], plan.affectedKeyMap);
    if (!plan.cancelAll && affectedItems.length === 0) continue;

    const rowIdx = i + 1;
    statusRanges.push(`I${rowIdx}`);
    timestampRanges.push(`J${rowIdx}`);
    affectedOrders.push({
      rowIdx: rowIdx,
      id: data[i][0],
      date: rowDate,
      email: String(data[i][3] || '').trim().toLowerCase(),
      nombre: data[i][4],
      departamentoId: data[i][5],
      resumen: normalizeMenuText_(data[i][6]),
      detail: normalizeOrderDetail_(detail),
      affectedItems: affectedItems
    });
  }

  if (affectedOrders.length > 0) {
    sh.getRangeList(statusRanges).setValue('CANCELADO');
    sh.getRangeList(timestampRanges).setValue(now);
  }

  return affectedOrders;
}

function groupAffectedOrdersByDate_(orders) {
  const grouped = {};
  (orders || []).forEach(order => {
    if (!grouped[order.date]) grouped[order.date] = [];
    grouped[order.date].push(order);
  });
  return grouped;
}

function buildMenuItemsHtml_(items) {
  const grouped = {};
  const categoryMap = getMenuCategoryMap_();

  (items || []).forEach(item => {
    if (!item || !item.cat || !item.plato) return;
    if (!grouped[item.cat]) grouped[item.cat] = [];
    grouped[item.cat].push(item);
  });

  const categories = Object.keys(grouped).sort((a, b) => {
    const aCategory = categoryMap[a];
    const bCategory = categoryMap[b];
    return (aCategory ? aCategory.orden : 999) - (bCategory ? bCategory.orden : 999) || a.localeCompare(b, 'es');
  });

  if (categories.length === 0) return '';

  const body = categories.map(cat => {
    const rows = grouped[cat].map(item => {
      const desc = item.desc ? `<span style="display:block;color:#6b7280;font-size:12px;margin-top:2px;">${escapeHtml_(item.desc)}</span>` : '';
      return `<li style="margin:6px 0;"><strong>${escapeHtml_(item.plato)}</strong>${desc}</li>`;
    }).join('');
    return `
      <div style="margin: 16px 0;">
        <p style="margin:0 0 6px;color:#1d4ed8;font-size:13px;font-weight:800;text-transform:uppercase;">${escapeHtml_((categoryMap[cat] && categoryMap[cat].nombre) || formatCatNameForEmail_(cat))}</p>
        <ul style="margin:0;padding-left:20px;">${rows}</ul>
      </div>
    `;
  }).join('');

  return `<div style="background:#f8fafc;border:1px solid #e5e7eb;border-radius:12px;padding:16px;margin:18px 0;">${body}</div>`;
}

function formatCatNameForEmail_(cat) {
  return String(cat || '').replace(/_/g, ' ');
}

function buildSimpleListHtml_(items) {
  const values = uniqueStrings_(items);
  if (values.length === 0) return '';
  return `<ul style="margin:8px 0 0;padding-left:20px;">${values.map(item => `<li>${escapeHtml_(item)}</li>`).join('')}</ul>`;
}

function getActiveNotificationUsers_(usersData) {
  const data = usersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Usuarios'), 7);
  const users = [];

  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][0] || '').trim().toLowerCase();
    const estado = String(data[i][4] || '').trim().toUpperCase();
    const prefs = safeParseJsonObject_(data[i][5]);
    if (estado === 'ACTIVO' && email && prefs.reminders !== false) {
      users.push({ email: email, nombre: data[i][1] });
    }
  }

  return users;
}

function notifyUsersMenuAvailable_(dateStr, menuItems, usersData) {
  if (!isFutureDateString_(dateStr)) return 0;

  const recipients = getActiveNotificationUsers_(usersData);
  if (recipients.length === 0) return 0;

  const dateLabel = formatDisplayDate_(dateStr);
  const menuHtml = buildMenuItemsHtml_(menuItems);
  const appUrl = getAppUrl_();

  recipients.forEach(user => {
    const html = getEmailTemplate_({
      title: 'Menu disponible',
      subtitle: `Almuerzo para ${dateLabel}`,
      body: `
        <p>Hola <strong>${escapeHtml_(getFirstName_(user.nombre))}</strong>,</p>
        <p>Ya esta disponible el menu de almuerzo preempacado para <strong>${escapeHtml_(dateLabel)}</strong>.</p>
        ${menuHtml}
        <p>Puedes entrar al sistema y realizar tu pedido mientras la fecha siga abierta.</p>
      `,
      cta: { text: 'Hacer pedido', url: appUrl },
      footerNote: 'Recibes este correo porque tienes las notificaciones activas en la app.'
    });

    sendEmail_(user.email, `Almuerzo Pre-empacado | Menu disponible ${dateStr}`, html);
  });

  return recipients.length;
}

function notifyUsersBulkMenuAvailable_(menuDays, usersData) {
  const days = (menuDays || [])
    .filter(day => day && day.date && isFutureDateString_(day.date) && day.items && day.items.length > 0)
    .slice()
    .sort((a, b) => String(a.date).localeCompare(String(b.date)));

  if (days.length === 0) return 0;

  const recipients = getActiveNotificationUsers_(usersData);
  if (recipients.length === 0) return 0;

  const firstDate = days[0].date;
  const lastDate = days[days.length - 1].date;
  const dateRangeLabel = days.length === 1
    ? formatDisplayDate_(firstDate)
    : `${formatDisplayDate_(firstDate)} - ${formatDisplayDate_(lastDate)}`;
  const appUrl = getAppUrl_();
  const daysHtml = days.map(day => `
    <div style="border:1px solid #dbeafe;border-radius:14px;overflow:hidden;margin:18px 0;background:#ffffff;">
      <div style="background:#eff6ff;padding:12px 16px;border-bottom:1px solid #dbeafe;">
        <p style="margin:0;color:#1e3a8a;font-size:15px;font-weight:800;">${escapeHtml_(formatDisplayDate_(day.date))}</p>
      </div>
      <div style="padding:0 14px 2px;">
        ${buildMenuItemsHtml_(day.items)}
      </div>
    </div>
  `).join('');

  recipients.forEach(user => {
    const html = getEmailTemplate_({
      title: 'Menu semanal disponible',
      subtitle: `Menus cargados: ${dateRangeLabel}`,
      body: `
        <p>Hola <strong>${escapeHtml_(getFirstName_(user.nombre))}</strong>,</p>
        <p>Ya estan disponibles los menus de almuerzo preempacado cargados para esta semana.</p>
        ${daysHtml}
        <p>Puedes entrar al sistema y realizar tus pedidos mientras las fechas sigan abiertas.</p>
      `,
      cta: { text: 'Planificar pedidos', url: appUrl },
      footerNote: 'Recibes este correo porque tienes las notificaciones activas en la app.'
    });

    sendEmail_(user.email, `Almuerzo Pre-empacado | Menu semanal disponible ${firstDate}`, html);
  });

  return recipients.length;
}

function notifyMenuChangedCancellations_(dateStr, orders) {
  const dateLabel = formatDisplayDate_(dateStr);
  const appUrl = getAppUrl_();
  let sent = 0;

  (orders || []).forEach(order => {
    if (!order.email) return;
    const affectedHtml = buildSimpleListHtml_(order.affectedItems);
    const html = getEmailTemplate_({
      title: 'Pedido cancelado',
      subtitle: `Cambio de menu para ${dateLabel}`,
      body: `
        <p>Hola <strong>${escapeHtml_(getFirstName_(order.nombre))}</strong>,</p>
        <p>Tu pedido para <strong>${escapeHtml_(dateLabel)}</strong> fue cancelado porque el menu de ese dia cambio.</p>
        <p>Debes volver a entrar al sistema y realizar tu pedido con las opciones actualizadas.</p>
        <div style="background:#fef2f2;border-left:4px solid #ef4444;padding:12px 14px;margin:16px 0;color:#991b1b;">
          <strong>Pedido cancelado:</strong><br>${escapeHtml_(order.resumen || 'Sin resumen')}
          ${affectedHtml ? `<div style="margin-top:10px;"><strong>Opciones afectadas:</strong>${affectedHtml}</div>` : ''}
        </div>
      `,
      cta: { text: 'Volver a pedir', url: appUrl },
      footerNote: 'Este aviso se envia aunque las notificaciones esten desactivadas porque afecta un pedido activo.'
    });

    sendEmail_(order.email, `Almuerzo Pre-empacado | Pedido cancelado por cambio de menu ${dateStr}`, html);
    sent++;
  });

  return sent;
}

function notifyLunchSuspendedCancellations_(dateStr, orders, reason) {
  const dateLabel = formatDisplayDate_(dateStr);
  let sent = 0;

  (orders || []).forEach(order => {
    if (!order.email) return;
    const reasonHtml = reason ? `<p><strong>Motivo registrado:</strong> ${escapeHtml_(reason)}</p>` : '';
    const html = getEmailTemplate_({
      title: 'Almuerzo suspendido',
      subtitle: `Suspension para ${dateLabel}`,
      body: `
        <p>Hola <strong>${escapeHtml_(getFirstName_(order.nombre))}</strong>,</p>
        <p>La administracion suspendio el almuerzo preempacado para <strong>${escapeHtml_(dateLabel)}</strong>.</p>
        ${reasonHtml}
        <div style="background:#fff7ed;border-left:4px solid #f97316;padding:12px 14px;margin:16px 0;color:#9a3412;">
          Tu pedido de ese dia fue cancelado automaticamente.
        </div>
        <p>No necesitas realizar ninguna accion adicional para esa fecha.</p>
      `,
      footerNote: 'Este aviso se envia aunque las notificaciones esten desactivadas porque afecta un pedido activo.'
    });

    sendEmail_(order.email, `Almuerzo Pre-empacado | Almuerzo suspendido ${dateStr}`, html);
    sent++;
  });

  return sent;
}

// === MENU MANAGEMENT API ===

function apiGetMenuDay(dateStr) {
   const admin = getUserInfo_();
   if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Denegado");

   const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
   const data = sh.getDataRange().getValues();
   const items = [];
   // Handle date string (YYYY-MM-DD) as local date to avoid timezone shift
   const fDate = formatDate_(new Date(dateStr + 'T12:00:00'));

   for(let i=1; i<data.length; i++) {
      let raw = data[i][1];
      // If raw is string YYYY-MM-DD, parse as local
      let dObj = (typeof raw === 'string' && raw.match(/^\d{4}-\d{2}-\d{2}$/)) ? new Date(raw + 'T12:00:00') : new Date(raw);
      const rowDate = formatDate_(dObj);
      if (rowDate === fDate) {
         items.push({
           id: data[i][0],
           cat: data[i][2],
           plato: normalizeMenuText_(data[i][3]),
           desc: normalizeMenuText_(data[i][4]),
           hab: data[i][5]
         });
      }
   }
   return { ok: true, items: items };
}

function apiSaveMenuItem(dateStr, cat, itemData) {
   const admin = getUserInfo_();
   if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Denegado");

   const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
   const data = sh.getDataRange().getValues();
   let rowIdx = -1;
   let oldItem = null;

   if (itemData.id) {
      for(let i=1; i<data.length; i++) {
         if (String(data[i][0]) === String(itemData.id)) {
            rowIdx = i+1;
            oldItem = getMenuItemSnapshotFromRow_(data[i]);
            break;
         }
      }
   }

   const category = getMenuCategoryById_(cat);
   if (!category) throw new Error('La categoria seleccionada no existe.');
   if (category.estado !== 'ACTIVO' && (!oldItem || oldItem.cat !== category.id)) {
      throw new Error('No puedes agregar platos a una categoria inactiva.');
   }

   const normalizedDate = formatDate_(new Date(dateStr + 'T12:00:00'));
   assertDateAllowedForMenuManagement_(normalizedDate, "editar");
   const menuSnapshot = buildActiveMenuSnapshotByDate_(data, new Set([normalizedDate]));
   const hadActiveMenuBefore = !!(menuSnapshot[normalizedDate] && menuSnapshot[normalizedDate].items.length > 0);
   const id = rowIdx > 0 ? itemData.id : Utilities.getUuid();
   // Save as Date object (local)
   const dateObj = new Date(normalizedDate + 'T12:00:00');
   const row = [id, dateObj, category.id, normalizeMenuText_(itemData.plato), normalizeMenuText_(itemData.desc), 'SI'];
   const newItem = createMenuItemFromPayload_(normalizedDate, { cat: category.id, plato: itemData.plato, desc: itemData.desc });

   if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
   else sh.appendRow(row);

   let cancellationCount = 0;
   let cancellationEmails = 0;
   if (oldItem && oldItem.enabled && isMenuItemChanged_(oldItem, newItem) && isFutureDateString_(oldItem.date)) {
      const affectedKeyMap = {};
      addAffectedMenuKey_(affectedKeyMap, oldItem);
      const affectedOrders = cancelActiveOrdersForPlans_([{ date: oldItem.date, affectedKeyMap: affectedKeyMap }]);
      cancellationCount = affectedOrders.length;
      cancellationEmails = notifyMenuChangedCancellations_(oldItem.date, affectedOrders);
   }

   let menuNotificationCount = 0;
   if (!oldItem && !hadActiveMenuBefore && newItem && isFutureDateString_(normalizedDate) && isDateOpenForOrdering_(normalizedDate)) {
      menuNotificationCount = notifyUsersMenuAvailable_(normalizedDate, [newItem]);
   }

   invalidateMenuDataCache_();
   invalidateUserInitCache_();
   return {
      ok: true,
      cancellations: cancellationCount,
      cancellationEmails: cancellationEmails,
      menuNotificationEmails: menuNotificationCount
   };
}

function apiDeleteMenuItem(id) {
   const admin = getUserInfo_();
   if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Denegado");
   const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
   const data = sh.getDataRange().getValues();
   for(let i=1; i<data.length; i++) {
      if (String(data[i][0]) === String(id)) {
         const oldItem = getMenuItemSnapshotFromRow_(data[i]);
         if (oldItem) assertDateAllowedForMenuManagement_(oldItem.date, "eliminar platos de");
         sh.deleteRow(i+1);
         let affectedOrders = [];
         let cancellationEmails = 0;
         if (oldItem && oldItem.enabled && isFutureDateString_(oldItem.date)) {
            const affectedKeyMap = {};
            addAffectedMenuKey_(affectedKeyMap, oldItem);
            affectedOrders = cancelActiveOrdersForPlans_([{ date: oldItem.date, affectedKeyMap: affectedKeyMap }]);
            cancellationEmails = notifyMenuChangedCancellations_(oldItem.date, affectedOrders);
         }
         invalidateMenuDataCache_();
         invalidateUserInitCache_();
         return {
            ok: true,
            cancellations: affectedOrders.length,
            cancellationEmails: cancellationEmails
         };
      }
   }
   return { ok: false, msg: "No encontrado" };
}

function apiSaveWeeklyMenu(menuData) {
   const admin = getUserInfo_();
   if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Denegado");

   const ss = SpreadsheetApp.getActive();
   const sh = ss.getSheetByName('Menu');
   const data = sh.getDataRange().getValues();
   const categoryMap = getMenuCategoryMap_();

   // Normalize keys to ensure matching
   const datesToUpdate = new Set();
   Object.keys(menuData).forEach(k => {
      const normalizedDate = formatDate_(new Date(k + 'T12:00:00'));
      assertDateAllowedForMenuManagement_(normalizedDate, "importar");
      const items = Array.isArray(menuData[k]) ? menuData[k] : [];
      items.forEach(item => {
         const category = categoryMap[String(item && item.cat || '').trim()];
         if (!category) throw new Error('La importacion incluye una categoria que no existe.');
         if (category.estado !== 'ACTIVO') throw new Error('La importacion incluye una categoria inactiva: ' + category.nombre + '.');
      });
      datesToUpdate.add(normalizedDate);
   });

   const oldMenuByDate = buildActiveMenuSnapshotByDate_(data, datesToUpdate);

   // 1. Identify rows to delete (indices, descending)
   const rowsToDelete = [];
   for (let i = data.length - 1; i >= 1; i--) {
      let raw = data[i][1];
      let dObj = (typeof raw === 'string' && raw.match(/^\d{4}-\d{2}-\d{2}$/)) ? new Date(raw + 'T12:00:00') : new Date(raw);
      const rowDate = formatDate_(dObj);
      if (datesToUpdate.has(rowDate)) {
         rowsToDelete.push(i + 1);
      }
   }

   // 2. Delete rows
   rowsToDelete.forEach(r => sh.deleteRow(r));

   // 3. Prepare new rows
   const allNewRows = [];
   const newMenuByDate = {};
   // Iterate over the keys provided by client to maintain association
   Object.keys(menuData).forEach(dateKey => {
      const normalizedDate = formatDate_(new Date(dateKey + 'T12:00:00'));
      // Only proceed if it was marked for update (double check)
      if (datesToUpdate.has(normalizedDate)) {
         const items = menuData[dateKey] || [];
         const dateObj = new Date(normalizedDate + 'T12:00:00');
         items.forEach(item => {
            const newItem = createMenuItemFromPayload_(normalizedDate, item);
            if (!newItem) return;
            if (!newMenuByDate[normalizedDate]) newMenuByDate[normalizedDate] = createMenuDateSnapshot_();
            addMenuItemToSnapshot_(newMenuByDate[normalizedDate], newItem);
            allNewRows.push([
              Utilities.getUuid(),
              dateObj,
              newItem.cat,
              newItem.plato,
              newItem.desc,
              'SI'
            ]);
         });
      }
   });

   // 4. Append
   if (allNewRows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, allNewRows.length, allNewRows[0].length).setValues(allNewRows);
   }

   const menuChangePlans = [];
   const newMenuDaysForBulkEmail = [];
   let newMenuNotificationEmails = 0;

   datesToUpdate.forEach(dateStr => {
      const oldSnapshot = oldMenuByDate[dateStr] || createMenuDateSnapshot_();
      const newSnapshot = newMenuByDate[dateStr] || createMenuDateSnapshot_();

      if (oldSnapshot.items.length === 0 && newSnapshot.items.length > 0 && isFutureDateString_(dateStr) && isDateOpenForOrdering_(dateStr)) {
         newMenuDaysForBulkEmail.push({ date: dateStr, items: newSnapshot.items });
         return;
      }

      const affectedKeyMap = getAffectedMenuKeysByReplacement_(oldSnapshot, newSnapshot);
      if (hasAffectedMenuKeys_(affectedKeyMap) && isFutureDateString_(dateStr)) {
         menuChangePlans.push({ date: dateStr, affectedKeyMap: affectedKeyMap });
      }
   });

   if (newMenuDaysForBulkEmail.length > 0) {
      const notificationUsersData = readSheetValues_(ss.getSheetByName('Usuarios'), 7);
      newMenuNotificationEmails = notifyUsersBulkMenuAvailable_(newMenuDaysForBulkEmail, notificationUsersData);
   }

   const affectedOrders = cancelActiveOrdersForPlans_(menuChangePlans);
   const affectedByDate = groupAffectedOrdersByDate_(affectedOrders);
   let cancellationEmails = 0;
   Object.keys(affectedByDate).forEach(dateStr => {
      cancellationEmails += notifyMenuChangedCancellations_(dateStr, affectedByDate[dateStr]);
   });

   invalidateMenuDataCache_();
   invalidateUserInitCache_();
   return {
      ok: true,
      cancellations: affectedOrders.length,
      cancellationEmails: cancellationEmails,
      menuNotificationEmails: newMenuNotificationEmails
   };
}

// === HOLIDAYS API ===

function apiGetHolidays() {
   return { ok: true, holidays: getHolidaysList_() };
}

function apiSaveHoliday(dateStr, desc) {
   const admin = getUserInfo_();
   if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Denegado");

   // Validate future date
   const d = new Date(dateStr + 'T12:00:00');
   const now = new Date();
   now.setHours(0,0,0,0);
   if (d < now) throw new Error("No puedes agregar días libres en el pasado.");

   const sh = SpreadsheetApp.getActive().getSheetByName('DiasLibres');
   const data = sh.getDataRange().getValues();
   let rowIdx = -1;

   // Check duplicate
   for(let i=1; i<data.length; i++) {
      if (formatDate_(new Date(data[i][0])) === dateStr) {
         rowIdx = i + 1;
         break;
      }
   }

   if (rowIdx > 0) sh.getRange(rowIdx, 2).setValue(desc);
   else sh.appendRow([dateStr, desc]);

   CacheService.getScriptCache().remove('HOLIDAYS_CACHE_V2');
   const suspendedOrders = cancelActiveOrdersForPlans_([{ date: dateStr, cancelAll: true }]);
   const suspensionEmails = notifyLunchSuspendedCancellations_(dateStr, suspendedOrders, desc);
   invalidateMenuDataCache_();
   invalidateUserInitCache_();
   return {
      ok: true,
      cancellations: suspendedOrders.length,
      cancellationEmails: suspensionEmails
   };
}

function apiDeleteHoliday(dateStr) {
   const admin = getUserInfo_();
   if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Denegado");
   const sh = SpreadsheetApp.getActive().getSheetByName('DiasLibres');
   const data = sh.getDataRange().getValues();
   for(let i=1; i<data.length; i++) {
      if (formatDate_(new Date(data[i][0])) === dateStr) {
         sh.deleteRow(i+1);
         CacheService.getScriptCache().remove('HOLIDAYS_CACHE_V2');
         invalidateMenuDataCache_();
         invalidateUserInitCache_();
         return { ok: true };
      }
   }
   return { ok: false };
}

// === UTILS ===

function apiHeartbeat() {
  const user = getUserInfo_();
  if (!user || user.estado !== 'ACTIVO') return { count: null };

  const lock = LockService.getScriptLock();
  // Background presence writes use a short lock to avoid cache write collisions.
  if (lock.tryLock(5000)) {
    try {
      const cache = CacheService.getScriptCache();
      const KEY = 'ACTIVE_SESSIONS_V2';
      const raw = cache.get(KEY);
      let sessions = raw ? JSON.parse(raw) : {};
      
      const now = Date.now();
      const TIME_WINDOW = 5 * 60 * 1000;

      sessions[String(user.email || '').toLowerCase()] = now;

      let count = 0;
      const cleanSessions = {};
      Object.keys(sessions).forEach(email => {
         const lastSeen = Number(sessions[email]) || 0;
         if (now - lastSeen < TIME_WINDOW) {
            cleanSessions[email] = lastSeen;
            count++;
         }
      });

      cache.put(KEY, JSON.stringify(cleanSessions), 21600);

      if (user.rol === 'ADMIN_GEN') {
         return { count: count };
      }
      
    } catch (e) {
      console.error('Error en heartbeat:', e);
    } finally {
      lock.releaseLock();
    }
  }
  
  return { count: null };
}

function getUserInfo_(targetEmail, usersData, deptMap) {
  const email = targetEmail ? targetEmail.toLowerCase() : Session.getActiveUser().getEmail().toLowerCase();
  const data = usersData || SpreadsheetApp.getActive().getSheetByName('Usuarios').getDataRange().getValues();
  const currentDeptMap = deptMap || getDepartmentMap_();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      const deptId = data[i][2];
      return {
        email: data[i][0],
        nombre: data[i][1],
        departamentoId: deptId,
        departamento: currentDeptMap[deptId] || deptId, // Resolve name
        rol: data[i][3],
        estado: data[i][4],
        codigo: data[i][6] || ''
      };
    }
  }
  return null;
}

function getDepartmentMap_() {
   const sh = SpreadsheetApp.getActive().getSheetByName('Departamentos');
  const map = {};
  if (sh) {
      const data = readSheetValues_(sh, 2);
      for(let i=1; i<data.length; i++) {
         map[data[i][0]] = data[i][1]; // ID -> Name
      }
   }
  return map;
}

function getDepartmentsList_() {
   const sh = SpreadsheetApp.getActive().getSheetByName('Departamentos');
   if (!sh) return [];
  const data = readSheetValues_(sh, 4);
   return data.slice(1).map(r => ({ id: r[0], nombre: r[1], admins: r[2], estado: r[3] }));
}

function getUsersByDept_(deptId, usersData) {
  const data = usersData || SpreadsheetApp.getActive().getSheetByName('Usuarios').getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === deptId && data[i][4] === 'ACTIVO') {
       users.push({ email: data[i][0], nombre: data[i][1] });
    }
  }
  return users;
}

function getChefGameHeaders_() {
  return [
    'juego_mes',
    'juego_puntos_mes',
    'juego_aciertos_mes',
    'juego_fallos_mes',
    'juego_tiempo_fecha',
    'juego_segundos_hoy',
    'juego_racha',
    'juego_racha_max',
    'juego_penalizacion_segundos',
    'juego_actualizado'
  ];
}

function ensureChefGameColumns_(sheet) {
  if (!sheet) throw new Error("Hoja Usuarios no encontrada.");

  const cache = CacheService.getScriptCache();
  const headers = getChefGameHeaders_();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  let current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const existing = {};
  current.forEach((header, index) => {
    if (header) existing[header] = index + 1;
  });

  const missing = headers.filter(header => !existing[header]);
  if (missing.length > 0) {
    sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, lastColumn + 1, 1, missing.length).setFontWeight('bold').setBackground('#f3f4f6');
    current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    cache.remove(CHEF_GAME_SCHEMA_CACHE_KEY);
  }

  const colMap = {};
  current.forEach((header, index) => {
    if (header) colMap[header] = index + 1;
  });
  cache.put(CHEF_GAME_SCHEMA_CACHE_KEY, '1', 3600);
  return colMap;
}

function getChefGameCell_(row, colMap, key) {
  const col = colMap[key];
  return col ? row[col - 1] : '';
}

function parseChefGameNumber_(value, fallback) {
  const parsed = Number(value);
  if (!isFinite(parsed)) return fallback || 0;
  return Math.max(0, Math.floor(parsed));
}

function getChefGameMonthKey_(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}

function normalizeChefGameMonthCell_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return getChefGameMonthKey_(value);
  }

  const raw = String(value || '').trim();
  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 7);

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) return getChefGameMonthKey_(parsed);
  return raw;
}

function normalizeChefGameDayCell_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const raw = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return raw;
}

function getChefGameMonthLabel_(date) {
  const d = date || new Date();
  const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  return months[d.getMonth()] + ' ' + d.getFullYear();
}

function createDefaultChefGameState_(now) {
  const current = now || new Date();
  return {
    monthKey: getChefGameMonthKey_(current),
    monthLabel: getChefGameMonthLabel_(current),
    dayKey: getTodayYmd_(),
    score: 0,
    hits: 0,
    misses: 0,
    usedSeconds: 0,
    remainingSeconds: CHEF_GAME_DAILY_LIMIT_SECONDS,
    dailyLimitSeconds: CHEF_GAME_DAILY_LIMIT_SECONDS,
    streak: 0,
    bestStreak: 0,
    penaltySeconds: 0
  };
}

function normalizeChefGameState_(row, colMap, now) {
  const current = now || new Date();
  const monthKey = getChefGameMonthKey_(current);
  const dayKey = getTodayYmd_();
  const rowMonth = normalizeChefGameMonthCell_(getChefGameCell_(row, colMap, 'juego_mes'));
  const rowDay = normalizeChefGameDayCell_(getChefGameCell_(row, colMap, 'juego_tiempo_fecha'));
  const sameMonth = rowMonth === monthKey;
  const sameDay = rowDay === dayKey;
  const usedSeconds = sameDay ? Math.min(CHEF_GAME_DAILY_LIMIT_SECONDS, parseChefGameNumber_(getChefGameCell_(row, colMap, 'juego_segundos_hoy'), 0)) : 0;
  const score = sameMonth ? parseChefGameNumber_(getChefGameCell_(row, colMap, 'juego_puntos_mes'), 0) : 0;

  return {
    monthKey: monthKey,
    monthLabel: getChefGameMonthLabel_(current),
    dayKey: dayKey,
    score: score,
    hits: sameMonth ? parseChefGameNumber_(getChefGameCell_(row, colMap, 'juego_aciertos_mes'), 0) : 0,
    misses: sameMonth ? parseChefGameNumber_(getChefGameCell_(row, colMap, 'juego_fallos_mes'), 0) : 0,
    usedSeconds: usedSeconds,
    remainingSeconds: Math.max(0, CHEF_GAME_DAILY_LIMIT_SECONDS - usedSeconds),
    dailyLimitSeconds: CHEF_GAME_DAILY_LIMIT_SECONDS,
    streak: sameDay ? parseChefGameNumber_(getChefGameCell_(row, colMap, 'juego_racha'), 0) : 0,
    bestStreak: sameMonth ? parseChefGameNumber_(getChefGameCell_(row, colMap, 'juego_racha_max'), 0) : 0,
    penaltySeconds: sameDay ? Math.min(30, parseChefGameNumber_(getChefGameCell_(row, colMap, 'juego_penalizacion_segundos'), 0)) : 0
  };
}

function sanitizeChefGameSubmittedState_(submittedState, now, previousState) {
  const current = now || new Date();
  const state = submittedState && typeof submittedState === 'object' ? submittedState : {};
  const previous = previousState && typeof previousState === 'object' ? previousState : createDefaultChefGameState_(current);
  const usedSeconds = Math.min(CHEF_GAME_DAILY_LIMIT_SECONDS, Math.max(parseChefGameNumber_(previous.usedSeconds, 0), parseChefGameNumber_(state.usedSeconds, 0)));
  const score = Math.min(999999, parseChefGameNumber_(state.score, parseChefGameNumber_(previous.score, 0)));
  const hits = Math.min(99999, parseChefGameNumber_(state.hits, 0));
  const misses = Math.min(99999, parseChefGameNumber_(state.misses, 0));
  const streak = Math.min(9999, parseChefGameNumber_(state.streak, 0));
  const bestStreak = Math.min(9999, Math.max(streak, parseChefGameNumber_(state.bestStreak, 0)));
  const penaltySeconds = Math.min(30, parseChefGameNumber_(state.penaltySeconds, 0));

  return {
    monthKey: getChefGameMonthKey_(current),
    monthLabel: getChefGameMonthLabel_(current),
    dayKey: getTodayYmd_(),
    score: score,
    hits: hits,
    misses: misses,
    usedSeconds: usedSeconds,
    remainingSeconds: Math.max(0, CHEF_GAME_DAILY_LIMIT_SECONDS - usedSeconds),
    dailyLimitSeconds: CHEF_GAME_DAILY_LIMIT_SECONDS,
    streak: streak,
    bestStreak: bestStreak,
    penaltySeconds: penaltySeconds
  };
}

function getChefGameUserContext_(email, sheet, colMap, usersData) {
  const targetEmail = String(email || '').toLowerCase();
  const data = usersData || sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').toLowerCase() === targetEmail) {
      return {
        rowIndex: i + 1,
        row: data[i],
        colMap: colMap
      };
    }
  }
  return null;
}

function getChefGameState_(email) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
    const colMap = ensureChefGameColumns_(sh);
    const context = getChefGameUserContext_(email, sh, colMap);
    if (!context) return createDefaultChefGameState_(new Date());
    return normalizeChefGameState_(context.row, colMap, new Date());
  } catch (e) {
    return createDefaultChefGameState_(new Date());
  }
}

function resolveChefGameTargetUser_(targetEmail, activeUser, usersData, deptMap) {
  const requestedEmail = String(targetEmail || '').trim().toLowerCase();
  if (!requestedEmail || requestedEmail === String(activeUser.email || '').toLowerCase()) return activeUser;

  if (!['ADMIN_GEN', 'ADMIN_DEP'].includes(activeUser.rol)) {
    throw new Error("No tienes permiso para registrar jugadas por otro usuario.");
  }

  const targetUser = getUserInfo_(requestedEmail, usersData, deptMap);
  if (!targetUser || targetUser.estado !== 'ACTIVO') throw new Error("Usuario objetivo no encontrado o inactivo.");
  if (activeUser.rol === 'ADMIN_DEP' && targetUser.departamentoId !== activeUser.departamentoId) {
    throw new Error("No puedes registrar jugadas para otro departamento.");
  }

  return targetUser;
}

function normalizeChefGameEventType_(value) {
  const type = String(value || '').trim().toUpperCase();
  if (type === 'CHEF' || type === 'HIT_CHEF') return 'CHEF';
  if (type === 'TRAP' || type === 'ONION' || type === 'PAN') return 'TRAP';
  if (type === 'MISS') return 'MISS';
  if (type === 'TICK') return 'TICK';
  throw new Error("Tipo de jugada invalido.");
}

function normalizeChefGameElapsed_(payload) {
  const raw = payload && payload.elapsedSeconds !== undefined ? Number(payload.elapsedSeconds) : 1;
  if (!isFinite(raw) || raw < 0) return 1;
  return Math.min(15, Math.ceil(raw));
}

function applyChefGameEvent_(state, eventType, payload, now) {
  const remainingBefore = state.remainingSeconds;
  const elapsedSeconds = Math.min(remainingBefore, normalizeChefGameElapsed_(payload));
  state.usedSeconds = Math.min(CHEF_GAME_DAILY_LIMIT_SECONDS, state.usedSeconds + elapsedSeconds);
  state.remainingSeconds = Math.max(0, CHEF_GAME_DAILY_LIMIT_SECONDS - state.usedSeconds);

  const event = {
    type: eventType,
    delta: 0,
    label: '',
    combo: '',
    elapsedSeconds: elapsedSeconds,
    cooldownSeconds: 0,
    capReached: state.remainingSeconds <= 0
  };

  if (remainingBefore <= 0) {
    event.label = 'Tiempo agotado';
    return event;
  }

  if (eventType === 'CHEF') {
    state.streak += 1;
    state.hits += 1;
    state.bestStreak = Math.max(state.bestStreak, state.streak);
    state.penaltySeconds = Math.max(0, state.penaltySeconds - 1);

    event.delta = CHEF_GAME_SCORE_CHEF;
    event.label = 'Chef atrapado';

    if (state.streak === 3) {
      event.delta += CHEF_GAME_SCORE_COMBO_3;
      event.combo = '3 chefs seguidos';
    } else if (state.streak === 5) {
      event.delta += CHEF_GAME_SCORE_COMBO_5;
      event.combo = '5 chefs seguidos';
    } else if (state.streak > 0 && state.streak % 10 === 0) {
      event.delta += CHEF_GAME_SCORE_PERFECT;
      event.combo = 'Racha perfecta';
    }
  } else if (eventType === 'TRAP') {
    const trapKind = String(payload && payload.kind || '').trim().toUpperCase();
    const trapPenalty = trapKind === 'PAN' ? 10 : 5;
    state.streak = 0;
    state.misses += 1;
    state.penaltySeconds = Math.min(30, state.penaltySeconds + trapPenalty);
    event.delta = trapKind === 'PAN' ? CHEF_GAME_SCORE_PAN : CHEF_GAME_SCORE_ONION;
    event.label = trapKind === 'PAN' ? 'Sarten quemado' : 'Cebolla podrida';
    event.cooldownSeconds = state.penaltySeconds;
  } else if (eventType === 'MISS') {
    state.streak = 0;
    state.misses += 1;
    state.penaltySeconds = Math.min(15, state.penaltySeconds + 1);
    event.delta = CHEF_GAME_SCORE_MISS;
    event.label = 'Chef escapado';
    event.cooldownSeconds = state.penaltySeconds;
  } else {
    event.label = 'Tiempo registrado';
  }

  if (event.delta !== 0) state.score = Math.max(0, state.score + event.delta);
  state.monthKey = getChefGameMonthKey_(now);
  state.monthLabel = getChefGameMonthLabel_(now);
  state.dayKey = getTodayYmd_();
  return event;
}

function writeChefGameState_(sheet, rowIndex, colMap, state, now) {
  const headers = getChefGameHeaders_();
  const cols = headers.map(header => colMap[header]).filter(Boolean);
  const minCol = Math.min.apply(null, cols);
  const maxCol = Math.max.apply(null, cols);
  const width = maxCol - minCol + 1;
  const values = sheet.getRange(rowIndex, minCol, 1, width).getValues()[0];
  const setValue = (key, value) => {
    const col = colMap[key];
    if (col) values[col - minCol] = value;
  };

  setValue('juego_mes', state.monthKey);
  setValue('juego_puntos_mes', state.score);
  setValue('juego_aciertos_mes', state.hits);
  setValue('juego_fallos_mes', state.misses);
  setValue('juego_tiempo_fecha', state.dayKey);
  setValue('juego_segundos_hoy', state.usedSeconds);
  setValue('juego_racha', state.streak);
  setValue('juego_racha_max', state.bestStreak);
  setValue('juego_penalizacion_segundos', state.penaltySeconds);
  setValue('juego_actualizado', now || new Date());

  if (colMap['juego_mes']) sheet.getRange(rowIndex, colMap['juego_mes']).setNumberFormat('@');
  if (colMap['juego_tiempo_fecha']) sheet.getRange(rowIndex, colMap['juego_tiempo_fecha']).setNumberFormat('@');
  sheet.getRange(rowIndex, minCol, 1, width).setValues([values]);
}

function isDateOpenForOrdering_(targetDateStr, holidaysSet) {
  if (!holidaysSet) holidaysSet = getHolidaysSet_();
  const now = new Date();
  const targetDate = new Date(targetDateStr + 'T12:00:00');

  // Past dates are closed
  const zeroNow = new Date(now); zeroNow.setHours(0,0,0,0);
  const zeroTarget = new Date(targetDate); zeroTarget.setHours(0,0,0,0);
  if (zeroTarget <= zeroNow) return false;

  const day = targetDate.getDay();
  if (day === 0 || day === 6) return false;
  if (holidaysSet.has(targetDateStr)) return false;

  const prevBizDay = getPreviousBusinessDay_(targetDate, holidaysSet);
  const prevBizDayStr = formatDate_(prevBizDay);
  const todayStr = formatDate_(now);

  // If today is the cutoff day
  if (todayStr === prevBizDayStr) {
    let envioTime = getConfigValue_('HORA_ENVIO') || '15:00';
    const minutesBefore = parseInt(getConfigValue_('MINUTOS_PREV_CIERRE') || '30', 10);

    let h, m;
    if (envioTime instanceof Date) {
       h = envioTime.getHours();
       m = envioTime.getMinutes();
    } else {
       const parts = String(envioTime).split(':');
       h = parseInt(parts[0], 10);
       m = parseInt(parts[1], 10);
    }

    // Fallback if config is invalid (e.g. "[]")
    if (isNaN(h) || isNaN(m)) { h = 15; m = 0; }
    if (isNaN(minutesBefore)) minutesBefore = 30;

    // Construct limit time using the SAME day as 'now'
    const limit = new Date(now);
    limit.setHours(h, m, 0, 0);
    limit.setMinutes(limit.getMinutes() - minutesBefore);

    if (now > limit) return false;
  }

  // If today is past the cutoff day
  const zeroPrev = new Date(prevBizDay); zeroPrev.setHours(0,0,0,0);
  if (zeroNow > zeroPrev) return false;

  return true;
}

function isDateAllowedForMenuManagement_(targetDateStr, holidaysSet) {
  if (!holidaysSet) holidaysSet = getHolidaysSet_();
  if (!targetDateStr) return false;

  const targetDate = new Date(targetDateStr + 'T12:00:00');
  if (isNaN(targetDate.getTime())) return false;

  const now = new Date();
  const zeroNow = new Date(now); zeroNow.setHours(0,0,0,0);
  const zeroTarget = new Date(targetDate); zeroTarget.setHours(0,0,0,0);
  if (zeroTarget <= zeroNow) return false;

  const day = targetDate.getDay();
  if (day === 0 || day === 6) return false;
  if (holidaysSet.has(targetDateStr)) return false;

  return true;
}

function assertDateAllowedForMenuManagement_(targetDateStr, actionLabel) {
  if (!isDateAllowedForMenuManagement_(targetDateStr)) {
    throw new Error("Solo puedes " + actionLabel + " menú para días laborables futuros.");
  }
}

function backupOrdersToDrive_(dateStr) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();

  const filtered = data.filter((row, i) => i === 0 || formatDate_(new Date(row[2])) === dateStr);
  if (filtered.length <= 1) return;

  try {
    const mFolder = getDailyBackupFolder_(dateStr);
    const tempSheet = SpreadsheetApp.create(`Pedidos_${dateStr}`);
    tempSheet.getSheets()[0].getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
    const tempFile = DriveApp.getFileById(tempSheet.getId());

    tempFile.moveTo(mFolder);
    const pdfBlob = tempFile.getAs('application/pdf');
    mFolder.createFile(pdfBlob).setName(`Pedidos_${dateStr}.pdf`);
  } catch (e) {
    console.error("Error backup: " + e.message);
  }
}

// Helpers reused...
let _configCache = null;
const OPERATIONAL_CONFIG_SCHEMA_CACHE_KEY = 'CONFIG_SCHEMA_READY_V3';

function readSheetValues_(sheet, columnCount) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  const totalColumns = columnCount || sheet.getLastColumn();
  if (totalColumns < 1) return [];
  return sheet.getRange(1, 1, lastRow, totalColumns).getValues();
}

function readJsonCache_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function writeJsonCache_(key, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), ttlSeconds);
  } catch (e) {
    // Ignore cache serialization/size failures and serve uncached data.
  }
}

function getRevisionValue_(key) {
  const props = PropertiesService.getScriptProperties();
  let revision = props.getProperty(key);
  if (!revision) {
    revision = '1';
    props.setProperty(key, revision);
  }
  return revision;
}

function bumpRevisionValue_(key) {
  const props = PropertiesService.getScriptProperties();
  const nextRevision = String(Number(props.getProperty(key) || '1') + 1);
  props.setProperty(key, nextRevision);
  return nextRevision;
}

function generateSecretToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function getInitCacheKey_(activeEmail, targetEmail, requestedDateStr) {
  return [
    'INIT_V2',
    getRevisionValue_('APP_INIT_REVISION'),
    String(activeEmail || '').toLowerCase(),
    String(targetEmail || '').toLowerCase(),
    requestedDateStr || 'AUTO'
  ].join(':');
}

function getAdminCacheKey_(user) {
  return [
    'ADMIN_V2',
    getRevisionValue_('APP_ADMIN_REVISION'),
    String(user.email || '').toLowerCase(),
    String(user.rol || ''),
    String(user.departamentoId || '')
  ].join(':');
}

function getDateViewCacheKey_(activeEmail, targetEmail, requestedDateStr) {
  return [
    'DATE_VIEW',
    getRevisionValue_('APP_INIT_REVISION'),
    String(activeEmail || '').toLowerCase(),
    String(targetEmail || '').toLowerCase(),
    requestedDateStr || ''
  ].join(':');
}

function getMenuBundleCacheKey_() {
  const bucket = Math.floor(Date.now() / (5 * 60 * 1000));
  return [
    'MENU_BUNDLE',
    getRevisionValue_('APP_MENU_REVISION'),
    bucket
  ].join(':');
}

function invalidateUserInitCache_() {
  bumpRevisionValue_('APP_INIT_REVISION');
  bumpRevisionValue_('APP_ADMIN_REVISION');
  _configCache = null;
}

function invalidateMenuDataCache_() {
  bumpRevisionValue_('APP_MENU_REVISION');
}

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
  if (key === 'ALL') return _configCache;
  return _configCache[key] !== undefined ? _configCache[key] : '';
}

function getOperationalConfigDefinitions_() {
  const defaultExpiry = formatDateWithOffset_(30);
  const todayStr = getTodayYmd_();
  const defaultAnnouncementPayload = JSON.stringify({
    slides: [
      {
        badge: "¡Novedad!",
        title: "Nuevo Sistema de Valoraciones",
        description: "Ahora puedes calificar tus comidas a partir de las 12:00 PM y valorar al proveedor de alimentos general.",
        icon: "fa-star",
        theme: "amber"
      },
      {
        badge: "Tu Opinión Cuenta",
        title: "Evalúa al Proveedor de Alimentos",
        description: "Califica y comenta el servicio del proveedor cuando quieras para ayudarnos a mantener y mejorar la calidad.",
        icon: "fa-award",
        theme: "indigo"
      }
    ]
  });

  return [
    { key: 'LOGO_ID', value: '', description: 'ID del archivo de imagen del Logo en Drive' },
    { key: 'APP_URL', value: ScriptApp.getService().getUrl(), description: 'URL publica de la aplicacion (Web App)' },
    { key: 'MEAL_PRICE_CURRENT', value: '57', description: 'Costo actual por almuerzo. Al cambiarlo se conserva historial automatico por fecha.' },
    { key: 'MEAL_PRICE_HISTORY_JSON', value: '[{"from":"1900-01-01","price":57}]', description: 'Historial auto-administrado del costo por almuerzo. No editar manualmente.' },
    { key: 'MENU_DAY_ENDPOINT_TOKEN', value: generateSecretToken_(), description: 'Token secreto para consumir el endpoint JSON de menu por fecha. Generar y compartir solo con TI.' },
    { key: 'RESPONSIBLES_EMAILS_JSON', value: '[]', description: 'JSON de correos externos en copia para el resumen diario general.' },
    { key: 'ANNOUNCEMENT_ENABLED', value: 'TRUE', description: 'Indica si el aviso general está activo para los usuarios (TRUE/FALSE)' },
    { key: 'ANNOUNCEMENT_ID', value: 'anuncio_v7_31_valoraciones', description: 'Identificador único del aviso activo. Al cambiarlo, todos los usuarios volverán a verlo.' },
    { key: 'ANNOUNCEMENT_EXPIRES_ON', value: defaultExpiry, description: 'Fecha límite para mostrar el aviso general (YYYY-MM-DD)' },
    { key: 'ANNOUNCEMENT_MAX_DISMISS', value: '3', description: 'Cantidad máxima de veces que el usuario puede cerrar el aviso antes de que no aparezca más.' },
    { key: 'ANNOUNCEMENT_PAYLOAD_JSON', value: defaultAnnouncementPayload, description: 'Contenido en formato JSON de los slides del aviso general.' },
    { key: 'PROVIDER_NAME', value: 'Proveedor de Alimentos', description: 'Nombre del proveedor de alimentos activo.' },
    { key: 'PROVIDER_PERIOD_ID', value: 'PROV_2026_01', description: 'Identificador del ciclo/período de evaluación del proveedor actual.' },
    { key: 'PROVIDER_PERIOD_START', value: todayStr, description: 'Fecha de inicio del ciclo del proveedor actual (YYYY-MM-DD).' }
  ];
}

function ensureConfigKeysBatch_(definitions) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Config');
    if (!sh || !definitions || definitions.length === 0) return;

    const data = readSheetValues_(sh, 3);
    const existing = {};
    for (let i = 1; i < data.length; i++) {
      existing[String(data[i][0])] = {
        row: i + 1,
        value: data[i][1]
      };
    }

    const missingRows = [];
    definitions.forEach(def => {
      if (!existing[def.key]) {
        missingRows.push([def.key, def.value, def.description]);
      } else if (def.key === 'MENU_DAY_ENDPOINT_TOKEN' && !String(existing[def.key].value || '').trim()) {
        sh.getRange(existing[def.key].row, 2).setValue(generateSecretToken_());
      }
    });

    if (missingRows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, missingRows.length, 3).setValues(missingRows);
      _configCache = null;
    }
  } catch (e) {
    console.error("Error ensuring config schema: " + e.message);
  }
}

function ensureOperationalConfigKeys_() {
  const cache = CacheService.getScriptCache();
  if (cache.get(OPERATIONAL_CONFIG_SCHEMA_CACHE_KEY)) return;
  ensureConfigKeysBatch_(getOperationalConfigDefinitions_());
  cache.put(OPERATIONAL_CONFIG_SCHEMA_CACHE_KEY, '1', 3600);
}

function ensureMealPriceConfig_() {
  ensureConfigKey_('MEAL_PRICE_CURRENT', '57', 'Costo actual por almuerzo. Al cambiarlo se conserva historial automatico por fecha.');
  ensureConfigKey_('MEAL_PRICE_HISTORY_JSON', '[{"from":"1900-01-01","price":57}]', 'Historial auto-administrado del costo por almuerzo. No editar manualmente.');
}

function getAnnouncementConfig_() {
  ensureOperationalConfigKeys_();
  const enabledRaw = String(getConfigValue_('ANNOUNCEMENT_ENABLED') || 'FALSE').trim().toUpperCase();
  const enabled = enabledRaw === 'TRUE' || enabledRaw === 'SI' || enabledRaw === '1';
  const id = String(getConfigValue_('ANNOUNCEMENT_ID') || 'anuncio_default').trim();
  const expiresOn = normalizeAnnouncementDate_(getConfigValue_('ANNOUNCEMENT_EXPIRES_ON'), formatDateWithOffset_(30));
  const maxDismiss = parsePositiveInt_(getConfigValue_('ANNOUNCEMENT_MAX_DISMISS'), 3);
  const rawPayload = getConfigValue_('ANNOUNCEMENT_PAYLOAD_JSON');

  let slides = [];
  try {
    const parsed = JSON.parse(rawPayload || '{}');
    if (Array.isArray(parsed.slides)) slides = parsed.slides;
    else if (Array.isArray(parsed)) slides = parsed;
  } catch (e) {
    slides = [];
  }

  if (slides.length === 0) {
    slides = [{
      badge: "Aviso",
      title: "Información importante",
      description: "Por favor revisa tus pedidos y mantén tus preferencias actualizadas.",
      icon: "fa-bullhorn",
      theme: "indigo"
    }];
  }

  return {
    enabled: enabled,
    id: id,
    expiresOn: expiresOn,
    maxDismiss: maxDismiss,
    slides: slides
  };
}

function normalizeAnnouncementDate_(value, fallback) {
  const normalized = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : fallback;
}

function getProviderInfo_() {
  ensureOperationalConfigKeys_();
  const name = String(getConfigValue_('PROVIDER_NAME') || 'Proveedor de Alimentos').trim();
  const periodId = String(getConfigValue_('PROVIDER_PERIOD_ID') || 'PROV_2026_01').trim();
  const periodStart = String(getConfigValue_('PROVIDER_PERIOD_START') || getTodayYmd_()).trim();
  return {
    name: name,
    periodId: periodId,
    periodStart: periodStart
  };
}

function parsePositiveInt_(value, fallback) {
  const parsed = parseInt(value, 10);
  return isNaN(parsed) || parsed <= 0 ? fallback : parsed;
}

function normalizeMealPriceValue_(value) {
  const rawValue = value === null || value === undefined ? '' : String(value);
  const parsed = Number(rawValue.replace(/[^0-9,.-]/g, '').replace(',', '.'));
  if (!isFinite(parsed) || parsed <= 0) {
    throw new Error("El costo por comida debe ser un numero mayor que cero.");
  }
  return Math.round(parsed * 100) / 100;
}

function getCurrentMealPrice_() {
  try {
    return normalizeMealPriceValue_(getConfigValue_('MEAL_PRICE_CURRENT') || '57');
  } catch (e) {
    return 57;
  }
}

function normalizeMealPriceHistory_(history) {
  if (!Array.isArray(history)) return [];

  const byDate = {};
  history.forEach(entry => {
    const from = entry && entry.from ? String(entry.from) : '';
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from)) return;

    try {
      byDate[from] = normalizeMealPriceValue_(entry.price);
    } catch (e) {
      // Ignore invalid history entries and keep the valid set.
    }
  });

  return Object.keys(byDate)
    .sort()
    .map(from => ({ from: from, price: byDate[from] }));
}

function getMealPriceHistory_() {
  const fallback = [{ from: '1900-01-01', price: getCurrentMealPrice_() }];
  const raw = getConfigValue_('MEAL_PRICE_HISTORY_JSON');
  if (!raw) return fallback;

  try {
    const normalized = normalizeMealPriceHistory_(JSON.parse(raw));
    return normalized.length > 0 ? normalized : fallback;
  } catch (e) {
    return fallback;
  }
}

function upsertMealPriceHistory_(history, price, effectiveDate) {
  const normalized = normalizeMealPriceHistory_(history);
  const next = {};
  normalized.forEach(entry => {
    next[entry.from] = entry.price;
  });
  next[effectiveDate] = normalizeMealPriceValue_(price);

  return Object.keys(next)
    .sort()
    .map(from => ({ from: from, price: next[from] }));
}

function getTodayYmd_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateWithOffset_(days) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeMenuText_(value) {
  if (value === null || value === undefined) return '';

  const minorWords = {
    de: true, del: true, la: true, las: true, el: true, los: true,
    y: true, e: true, o: true, u: true, con: true, al: true, en: true
  };

  const cleaned = String(value).trim().replace(/\s+/g, ' ');
  if (!cleaned) return '';

  return cleaned
    .toLowerCase()
    .split(' ')
    .map((word, index) => {
      return word
        .split(/([/-])/)
        .map(part => {
          if (!part || part === '/' || part === '-') return part;
          if (index > 0 && minorWords[part]) return part;
          return part.charAt(0).toUpperCase() + part.slice(1);
        })
        .join('');
    })
    .join(' ');
}

function normalizeOrderDetail_(detail) {
  const normalized = detail && typeof detail === 'object' ? Object.assign({}, detail) : {};
  if (Array.isArray(normalized.items)) {
    normalized.items = normalized.items.map(normalizeMenuText_);
  }
  return normalized;
}

function getOrderKitchenNote_(detail) {
  if (!detail || typeof detail !== 'object') return '';
  return String(detail.comentarios || '').trim();
}

function getHolidaysList_() {
   const sh = SpreadsheetApp.getActive().getSheetByName('DiasLibres');
   if(!sh) return [];
   return sh.getDataRange().getValues().slice(1)
      .filter(r => r[0])
      .map(r => {
         try { return { date: formatDate_(new Date(r[0])), desc: r[1] }; }
         catch(e) { return null; }
      })
      .filter(h => h && h.date >= formatDate_(new Date())); // Only future for list
}

function getHolidaysSet_() {
  const cache = CacheService.getScriptCache();
  const cachedHolidays = cache.get('HOLIDAYS_CACHE_V2');
  if (cachedHolidays) return new Set(JSON.parse(cachedHolidays));

  const set = new Set();
  const list = getHolidaysList_(); // uses sheet
  list.forEach(h => set.add(h.date));

  try {
    const calId = 'es.do#holiday@group.v.calendar.google.com';
    const now = new Date();
    const start = new Date(now.getTime() - 30 * 86400000);
    const end = new Date(now.getTime() + 365 * 86400000);
    CalendarApp.getCalendarById(calId).getEvents(start, end).forEach(e => set.add(formatDate_(e.getStartTime())));
  } catch (e) {}

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

function isTodayBusinessDay_() {
  const now = new Date();
  const day = now.getDay();
  // Weekend
  if (day === 0 || day === 6) return false;

  // Holidays
  const dateStr = formatDate_(now);
  const holidays = getHolidaysSet_();
  if (holidays.has(dateStr)) return false;

  return true;
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
  const fileId = getConfigValue_('FOOTER_SIGNATURE_ID');
  if (!fileId) return '';
  const cacheKey = `SIG_V4:${fileId}`;
  const cached = cache.get(cacheKey);
  if (cached) return cached;
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    const dataUrl = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
    cache.put(cacheKey, dataUrl, 21600);
    return dataUrl;
  } catch (blobError) {
    try {
      const file = Drive.Files.get(fileId, { fields: 'thumbnailLink' });
      if (!file || !file.thumbnailLink) return '';
      const imageUrl = file.thumbnailLink.replace(/=s\d+(-[a-z])?$/, '=s300');
      cache.put(cacheKey, imageUrl, 21600);
      return imageUrl;
    } catch (thumbnailError) {
      return '';
    }
  }
}

function getAvailableMenuDates_(fetchAll, menuData) {
  const data = menuData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Menu'), 6);
  // data[0] is header if raw fetch, but if logic assumes slicing elsewhere...
  // The original logic: data = menuSh.getRange(2, 1, ..., 2).getValues(); (No headers)
  // But generic 'getDataRange' includes headers.
  // We should loop from 1.

  const now = new Date();
  const todayStr = formatDate_(now);
  const datesSet = new Set();

  for(let i=1; i<data.length; i++) {
    const r = data[i];
    if(!r[1]) continue;
    const dStr = formatDate_(new Date(r[1]));
    if (dStr >= todayStr) datesSet.add(dStr);
  }

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

function getMenuBundle_() {
  const cacheKey = getMenuBundleCacheKey_();
  const cachedBundle = readJsonCache_(cacheKey);
  if (cachedBundle && cachedBundle.dates && cachedBundle.menusByDate) {
    return cachedBundle;
  }

  const menuData = readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Menu'), 6);
  const dates = getAvailableMenuDates_(true, menuData);
  const menusByDate = getAllMenus_(dates, menuData);
  const bundle = { dates: dates, menusByDate: menusByDate };
  writeJsonCache_(cacheKey, bundle, 300);
  return bundle;
}

function normalizeCategoryLookupKey_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}

function createSingleDateMap_(dateStr, menu) {
  if (!dateStr) return {};
  const map = {};
  map[dateStr] = menu || {};
  return map;
}

function getAllMenus_(availableDates, menuData) {
  const data = menuData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Menu'), 6);
  const menuMap = {};
  const validDates = new Set(availableDates.map(d => d.value));
  availableDates.forEach(d => { menuMap[d.value] = {}; });
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][1]));
    if (validDates.has(rowDate) && String(data[i][5]).toUpperCase() === 'SI') {
      const cat = data[i][2];
      const item = { id: data[i][0], plato: normalizeMenuText_(data[i][3]), desc: normalizeMenuText_(data[i][4]) };
      if (!menuMap[rowDate][cat]) menuMap[rowDate][cat] = [];
      menuMap[rowDate][cat].push(item);
    }
  }
  return menuMap;
}

function getAllUserOrders_(email, availableDates, ordersData) {
  const data = ordersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Pedidos'), 9);
  const ordersMap = {};
  const validDates = Array.isArray(availableDates) && availableDates.length > 0
    ? new Set(availableDates.map(d => d.value))
    : null;
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (validDates && !validDates.has(rowDate)) continue;
    if (String(data[i][3]).toLowerCase() === String(email).toLowerCase() && data[i][8] !== 'CANCELADO') {
      let detail = {};
      try {
        detail = JSON.parse(data[i][7] || '{}');
      } catch (e) {
        detail = {};
      }
      ordersMap[rowDate] = {
        id: data[i][0],
        resumen: normalizeMenuText_(data[i][6]),
        detalle: normalizeOrderDetail_(detail)
      };
    }
  }
  return ordersMap;
}

function getUserOrderByDate_(email, dateStr, ordersData) {
  const data = ordersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Pedidos'), 9);
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate !== dateStr) continue;
    if (String(data[i][3]).toLowerCase() !== String(email).toLowerCase()) continue;
    if (data[i][8] === 'CANCELADO') continue;

    let detail = {};
    try {
      detail = JSON.parse(data[i][7] || '{}');
    } catch (e) {
      detail = {};
    }

    return {
      id: data[i][0],
      resumen: normalizeMenuText_(data[i][6]),
      detalle: normalizeOrderDetail_(detail)
    };
  }
  return null;
}

function getUserAccessRecord_(targetEmail, usersData) {
  const email = targetEmail ? targetEmail.toLowerCase() : Session.getActiveUser().getEmail().toLowerCase();
  const data = usersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Usuarios'), 7);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      return {
        email: data[i][0],
        nombre: data[i][1],
        departamentoId: data[i][2],
        rol: data[i][3],
        estado: data[i][4],
        codigo: data[i][6] || ''
      };
    }
  }
  return null;
}

function getUserPrefs_(email, usersData) {
  const data = usersData || SpreadsheetApp.getActive().getSheetByName('Usuarios').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      return JSON.parse(data[i][5] || '{}');
    }
  }
  return {};
}

function getDepartmentStats_(dateStr, deptIdFilter, ordersData, deptMap) {
  const departmentStats = { total: 0, byUser: [] };
  const data = ordersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Pedidos'), 9);
  const currentDeptMap = deptMap || getDepartmentMap_();
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate !== dateStr || data[i][8] === 'CANCELADO') continue;
    if (deptIdFilter && data[i][5] !== deptIdFilter) continue;

    departmentStats.total++;
    departmentStats.byUser.push({
      nombre: data[i][4],
      pedido: normalizeMenuText_(data[i][6]),
      depto: currentDeptMap[data[i][5]] || data[i][5]
    });
  }
  return departmentStats;
}

function getOrdersByDate_(dateStr, ordersData, deptMap) {
  const data = ordersData || SpreadsheetApp.getActive().getSheetByName('Pedidos').getDataRange().getValues();
  const currentDeptMap = deptMap || getDepartmentMap_();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate === dateStr && data[i][8] !== 'CANCELADO') {
      list.push({
        nombre: data[i][4],
        departamentoId: data[i][5],
        departamento: currentDeptMap[data[i][5]] || data[i][5],
        resumen: normalizeMenuText_(data[i][6])
      });
    }
  }
  return list;
}

function validateOrderRules_(sel) {
  const cats = sel.categorias || [];
  const items = sel.items || [];
  if (!cats || cats.length === 0) return;

  const categoryMap = getMenuCategoryMap_();
  const uniqueCats = [...new Set(cats)];

  // 1. Validar reglas de combinabilidad dinámicas
  for (let i = 0; i < uniqueCats.length; i++) {
    const catId = uniqueCats[i];
    const cat = categoryMap[catId];
    if (!cat) continue;

    if (!cat.es_combinable && uniqueCats.length > 1) {
      const allowed = Array.isArray(cat.combinable_con) ? cat.combinable_con : [];
      const hasDisallowed = uniqueCats.some(otherId => otherId !== catId && !allowed.includes(otherId));
      if (hasDisallowed) {
        throw new Error("La categoría " + (cat.nombre || catId) + " no se puede combinar con el menú seleccionado.");
      }
    } else if (cat.es_combinable && Array.isArray(cat.combinable_con) && cat.combinable_con.length > 0) {
      const hasDisallowed = uniqueCats.some(otherId => otherId !== catId && !cat.combinable_con.includes(otherId));
      if (hasDisallowed) {
        throw new Error("La categoría " + (cat.nombre || catId) + " solo se puede combinar con las categorías permitidas.");
      }
    }
  }

  // 2. Reglas tradicionales de granos y arroz vs víveres
  if (cats.includes('Granos')) {
    const hasWhiteRice = items.some(i => String(i).toLowerCase().includes('arroz blanco'));
    if (!hasWhiteRice) throw new Error("Los granos requieren seleccionar Arroz Blanco.");
  }
  if (cats.includes('Arroces') && cats.includes('Viveres')) {
    throw new Error("No puedes combinar Arroz y Víveres.");
  }
}

function buildOrderRecordId_(email, dateStr) {
  return ['ORD', String(dateStr || ''), String(email || '').toLowerCase()].join('|');
}

function findOrderRowById_(sheet, orderId) {
  if (!sheet || !orderId) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const match = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(String(orderId))
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : 0;
}

function getOrderSnapshotByRow_(sheet, rowIdx) {
  if (!sheet || rowIdx < 2) return null;
  const row = sheet.getRange(rowIdx, 1, 1, 9).getValues()[0];
  if (!row || !row[0]) return null;
  return {
    rowIdx: rowIdx,
    id: row[0],
    date: formatDate_(new Date(row[2])),
    email: row[3],
    nombre: row[4],
    departamentoId: row[5],
    resumen: row[6],
    rawDetail: row[7],
    estado: row[8]
  };
}

function cancelOrderRecordById_(orderId, canCancelFn) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const rowIdx = findOrderRowById_(sh, orderId);
  if (!rowIdx) return { found: false };

  const snapshot = getOrderSnapshotByRow_(sh, rowIdx);
  if (!snapshot || snapshot.estado === 'CANCELADO') {
    return { found: false };
  }

  const allowed = typeof canCancelFn === 'function' ? !!canCancelFn(snapshot) : true;
  if (!allowed) {
    return { found: true, allowed: false, date: snapshot.date, snapshot: snapshot };
  }

  sh.getRange(rowIdx, 9, 1, 2).setValues([['CANCELADO', new Date()]]);
  return { found: true, allowed: true, date: snapshot.date, snapshot: snapshot };
}

function saveOrderToSheet_(user, dateStr, selection, creatorEmail) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const submittedOrderId = selection && selection.orderId ? String(selection.orderId) : '';
  const deterministicId = buildOrderRecordId_(user.email, dateStr);
  let rowIdx = findOrderRowById_(sh, submittedOrderId);
  if (!rowIdx && submittedOrderId !== deterministicId) {
    rowIdx = findOrderRowById_(sh, deterministicId);
  }
  const id = submittedOrderId || deterministicId;
  const now = new Date();
  const normalizedItems = (selection.items || []).map(normalizeMenuText_);
  const orderDetail = {
    categorias: Array.isArray(selection.categorias) ? selection.categorias.slice() : [],
    items: normalizedItems,
    comentarios: selection.comentarios || ''
  };

  // Save ID in col 6 (Index 5)
  const rowData = [
    id, now, dateStr, user.email, user.nombre, user.departamentoId,
    normalizedItems.join(', '), JSON.stringify(orderDetail), 'ACTIVO', now,
    creatorEmail || user.email
  ];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);

  return {
    id: id,
    resumen: normalizedItems.join(', '),
    detalle: normalizeOrderDetail_(orderDetail)
  };
}

function sendEmail_(to, subject, htmlBody, cc, attachments) {
  const testMode = isTestEmailMode_();
  const testDest = String(getConfigValue_('TEST_EMAIL_DEST') || '').trim();
  const senderName = getConfigValue_('MAIL_SENDER_NAME');

  const recipient = testMode ? testDest : to;
  if (!recipient) {
    if (testMode) console.warn("TEST_EMAIL_MODE activo sin TEST_EMAIL_DEST. Correo no enviado a destinatarios reales.");
    return;
  }

  const finalSubject = testMode ? `[TEST] ${subject}` : subject;

  // Signature is now handled by the template system (getEmailTemplate_)
  // We strictly send what we receive, assuming it's already formatted.

  const options = {
    to: recipient,
    subject: finalSubject,
    htmlBody: htmlBody,
    name: senderName
  };

  if (cc && !testMode) options.cc = cc;
  if (testMode) {
    const originalMeta = [
      to ? `<strong>Original TO:</strong> ${to}` : '',
      cc ? `<strong>Original CC:</strong> ${cc}` : ''
    ].filter(Boolean).join('<br>');
    if (originalMeta) {
      options.htmlBody = `<p style="background:#fef3c7;color:#92400e;padding:10px 12px;border-radius:8px;font-size:12px;">${originalMeta}</p>` + options.htmlBody;
    }
  }

  // Attachments handling (Array)
  if (attachments) options.attachments = attachments;

  // Inline Images (CID) Logic for Logo
  const logoId = getConfigValue_('LOGO_ID');
  if (logoId && htmlBody.includes('cid:appLogo')) {
     const logoBlob = getLogoBlob_(logoId);
     if (logoBlob) {
        if (!options.inlineImages) options.inlineImages = {};
        options.inlineImages['appLogo'] = logoBlob;
     }
  }

  try {
    MailApp.sendEmail(options);
  } catch(e) {
    console.error("Email error: " + e.message);
  }
}

// === EMAIL SYSTEM ===

function getEmailTemplate_(data) {
  // data: { title, subtitle, body, cta: {text, url}, footerNote }

  // Use CID for robust email support
  const appName = getConfigValue_('APP_TITLE') || 'Solicitud Almuerzo';

  // If LOGO_ID exists, we assume sendEmail_ will attach it as 'appLogo'
  const logoId = getConfigValue_('LOGO_ID');
  let logoHtml = '';
  if (logoId) {
     logoHtml = `<img src="cid:appLogo" alt="Logo" style="max-height: 80px; width: auto; margin-bottom: 20px; display: block; margin-left: auto; margin-right: auto;">`;
  }

  const primaryColor = '#2563eb'; // blue-600
  const grayBg = '#f9fafb';
  const white = '#ffffff';
  const textDark = '#111827';
  const textGray = '#4b5563';

  let ctaHtml = '';
  if (data.cta && data.cta.text && data.cta.url) {
     ctaHtml = `
       <div style="text-align: center; margin-top: 32px; margin-bottom: 32px;">
         <a href="${data.cta.url}" style="background-color: ${primaryColor}; color: ${white}; padding: 14px 28px; border-radius: 8px; text-decoration: none; font-weight: bold; font-family: sans-serif; font-size: 16px; display: inline-block; box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);">${data.cta.text}</a>
       </div>
     `;
  }

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body { margin: 0; padding: 0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background-color: ${grayBg}; }
        .container { width: 100%; background-color: ${grayBg}; padding: 40px 20px; box-sizing: border-box; }
        .card { background-color: ${white}; border-radius: 16px; max-width: 600px; margin: 0 auto; padding: 40px; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05); border: 1px solid #e5e7eb; }
        .header { text-align: center; margin-bottom: 30px; }
        .app-title { color: ${textDark}; font-size: 24px; font-weight: 800; margin: 0; letter-spacing: -0.5px; }
        .content { color: ${textGray}; font-size: 16px; line-height: 1.6; }
        .footer { text-align: center; margin-top: 40px; font-size: 12px; color: #9ca3af; }
        .footer a { color: #9ca3af; text-decoration: underline; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="card">
          <div class="header">
            ${logoHtml}
            <h1 class="app-title">${data.title || appName}</h1>
            ${data.subtitle ? `<p style="color: #6b7280; font-size: 14px; margin-top: 8px; font-weight: 500;">${data.subtitle}</p>` : ''}
          </div>

          <div class="content">
            ${data.body}
          </div>

          ${ctaHtml}

          <div class="footer">
             <p style="margin-bottom: 8px; font-weight: 600;">${appName}</p>
             <p>&copy; ${new Date().getFullYear()} Dirección de Innovación.</p>
             ${data.footerNote ? `<p style="margin-top: 16px; padding-top: 16px; border-top: 1px solid #f3f4f6;">${data.footerNote}</p>` : ''}
          </div>
        </div>
      </div>
    </body>
    </html>
  `;
}

function getLogoBlob_(fileId) {
  if (!fileId) return null;
  try {
    // We can fetch directly using DriveApp for internal scripts
    return DriveApp.getFileById(fileId).getBlob();
  } catch (e) {
     console.error("Error fetching logo blob: " + e.message);
     return null;
  }
}

function getLogoDataUrl_() {
   // Deprecated in favor of CID, but kept if needed for UI (not Email)
   // ...
   return null;
}

function ensureConfigKey_(key, defaultValue, description) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Config');
    if (!sh) return;

    // Check if exists using Cache or direct read (Direct is safer for Admin Panel load)
    const data = sh.getDataRange().getValues();
    let exists = false;
    for(let i=1; i<data.length; i++) {
       if(String(data[i][0]) === key) {
          exists = true;
          break;
       }
    }

    if (!exists) {
       sh.appendRow([key, defaultValue, description]);
       _configCache = null; // Invalidate cache
    }
  } catch(e) {
    console.error("Error ensuring config key: " + e.message);
  }
}

function checkMenuIntegrity_() {
  // Same logic as before
  const ss = SpreadsheetApp.getActive();
  const mSh = ss.getSheetByName('Menu');
  const data = mSh.getDataRange().getValues();
  const menuMap = {};
  for (let i = 1; i < data.length; i++) {
    const dStr = formatDate_(new Date(data[i][1]));
    const cat = data[i][2];
    const item = String(data[i][3]).toLowerCase();
    if (dStr > formatDate_(new Date())) {
      if (!menuMap[dStr]) menuMap[dStr] = { hasRice: false };
      if (cat === 'Arroces' && item.includes('arroz blanco')) menuMap[dStr].hasRice = true;
    }
  }
  const warnings = [];
  Object.keys(menuMap).forEach(d => { if (!menuMap[d].hasRice) warnings.push(d); });
  if (warnings.length > 0) {
    const admins = getConfigValue_('ADMIN_EMAILS');
    if (admins) {
       const html = getEmailTemplate_({
          title: '⚠️ Alerta de Menú',
          body: `
            <p>Se han detectado problemas de integridad en el menú cargado para las siguientes fechas:</p>
            <div style="background-color: #fef2f2; padding: 16px; border-left: 4px solid #ef4444; margin: 16px 0; color: #b91c1c;">
               <strong>Falta Arroz Blanco:</strong><br>
               ${warnings.join('<br>')}
            </div>
            <p>Por favor, revisa el menú y corrige estas fechas para evitar problemas con las validaciones de pedidos (Granos).</p>
          `,
          cta: { text: 'Revisar Menú', url: getAppUrl_() }
       });
       sendEmail_(admins, "Almuerzo Pre-empacado | Alerta: Integridad de Menú", html);
    }
  }
}

function sendDailyAdminSummary_(dateStr, context) {
  const admins = getConfigValue_('ADMIN_EMAILS');
  if (!admins) return false;
  const ctx = context || {};
  const orders = ctx.orders || getOrdersByDate_(dateStr);
  const count = orders.length;
  if (count > 0) {
     const deptSummary = ctx.deptSummary || getDepartmentOrderSummary_(groupOrdersByDepartment_(orders), getDepartmentMap_());
     const attachments = ctx.attachments || [];
     const copyRecipients = ctx.copyRecipients !== undefined
       ? joinEmailList_(ctx.copyRecipients)
       : getDailySummaryCopyRecipients_();
     const formattedDate = Utilities.formatDate(new Date(dateStr + 'T12:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy');
     const departmentTable = buildDepartmentSummaryTableHtml_(deptSummary);
     const html = getEmailTemplate_({
        title: 'Resumen Diario',
        subtitle: `Pedidos para el ${formattedDate}`,
        body: `
           <p>Resumen ejecutivo de la operación de almuerzo:</p>
           <div style="text-align: center; margin: 24px 0;">
              <span style="font-size: 48px; font-weight: 800; color: #111827;">${count}</span>
              <p style="color: #6b7280; margin-top: 8px;">Pedidos Totales</p>
           </div>
           ${departmentTable}
           ${ctx.testRun ? '<p>Esta ejecucion fue de prueba; no se guardaron respaldos en Google Drive.</p>' : '<p>Los respaldos detallados han sido generados y guardados en Google Drive.</p>'}
           ${attachments.length ? '<p>Se adjunta un Excel consolidado con el resumen general y una hoja por departamento.</p>' : ''}
        `,
        cta: { text: 'Ver Panel Administrativo', url: getAppUrl_() },
        footerNote: ctx.testRun ? 'Correo de prueba. No se generaron respaldos permanentes.' : ''
     });
    sendEmail_(admins, `Almuerzo Pre-empacado | Resumen Pedidos ${dateStr}`, html, copyRecipients, attachments);
    return true;
  }
  return false;
}

function getOrdersByDateDetailed_(dateStr, context) {
  const ctx = context || {};
  const data = ctx.ordersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Pedidos'), 11);
  const deptMap = ctx.deptMap || getDepartmentMap_();
  const codeMap = ctx.codeMap || getUserCodeMap_();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate === dateStr && data[i][8] !== 'CANCELADO') {
      let detail = {};
      try { detail = JSON.parse(data[i][7]); } catch(e){}
      const email = String(data[i][3]).toLowerCase();
      list.push({
        nombre: data[i][4],
        departamentoId: data[i][5],
        departamento: deptMap[data[i][5]] || data[i][5],
        resumen: data[i][6],
        detail: detail,
        notaCocina: getOrderKitchenNote_(detail),
        codigo: codeMap[email] || ''
      });
    }
  }
  return list;
}

function getUserCodeMap_(usersData) {
   const data = usersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Usuarios'), 7);
   const map = {};
   for(let i=1; i<data.length; i++) {
      const email = String(data[i][0]).toLowerCase();
      const code = data[i][6]; // Index 6 is Code
      if(code) map[email] = String(code);
   }
   return map;
}

function normalizeReportDate_(dateStr) {
  const normalized = String(dateStr || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    throw new Error("Fecha invalida para reportes.");
  }
  return normalized;
}

function sortOrdersForGeneralReport_(orders) {
  return (orders || []).slice().sort((a, b) => {
    const deptCompare = String(a.departamento || '').localeCompare(String(b.departamento || ''), 'es');
    if (deptCompare !== 0) return deptCompare;
    return String(a.nombre || '').localeCompare(String(b.nombre || ''), 'es');
  });
}

function groupOrdersByDepartment_(orders) {
  const grouped = {};
  (orders || []).forEach(order => {
    const deptId = order.departamentoId || 'Sin Depto';
    if (!grouped[deptId]) grouped[deptId] = [];
    grouped[deptId].push(order);
  });
  return grouped;
}

function getDepartmentOrderSummary_(byDept, deptMap) {
  const currentDeptMap = deptMap || getDepartmentMap_();
  return Object.keys(byDept || {}).map(deptId => {
    const orders = byDept[deptId] || [];
    return {
      id: deptId,
      name: currentDeptMap[deptId] || (orders[0] && orders[0].departamento) || deptId,
      count: orders.length
    };
  }).sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'es'));
}

function getReportRecipientsConfig_() {
  const rawConfig = getConfigValue_('RESPONSIBLES_EMAILS_JSON');
  try {
    return JSON.parse(rawConfig);
  } catch(e) {
    return rawConfig;
  }
}

function getDailySummaryCopyRecipients_() {
  const adminSet = {};
  uniqueEmailList_(getConfigValue_('ADMIN_EMAILS')).forEach(email => {
    adminSet[email] = true;
  });

  return joinEmailList_(normalizeDailySummaryCopyRecipients_(getReportRecipientsConfig_())
    .filter(email => !adminSet[email]));
}

function normalizeDailySummaryCopyRecipients_(value, options) {
  const opts = options || {};
  const candidates = collectReportRecipientEmailCandidates_(parseReportRecipientsValue_(value));
  const valid = [];
  const invalid = [];

  candidates.forEach(email => {
    if (isValidEmailAddress_(email)) {
      valid.push(email);
    } else {
      invalid.push(email);
    }
  });

  if (invalid.length > 0) {
    const msg = 'Corrige los correos de responsables del resumen diario: ' + uniqueEmailList_(invalid).join(', ');
    if (opts.throwOnInvalid) throw new Error(msg);
    console.warn(msg);
  }

  return uniqueEmailList_(valid);
}

function parseReportRecipientsValue_(value) {
  if (typeof value !== 'string') return value || [];
  const trimmed = value.trim();
  if (!trimmed) return [];

  try {
    return JSON.parse(trimmed);
  } catch(e) {
    return trimmed;
  }
}

function collectReportRecipientEmailCandidates_(value) {
  const result = [];

  const collect = item => {
    if (!item) return;

    if (typeof item === 'string') {
      String(item)
        .split(/[;,]/)
        .map(email => email.trim().toLowerCase())
        .filter(email => email)
        .forEach(email => result.push(email));
      return;
    }

    if (Array.isArray(item)) {
      item.forEach(collect);
      return;
    }

    if (typeof item === 'object') {
      const directEmail = getRecipientEmail_(item);
      if (directEmail) {
        collect(directEmail);
        return;
      }

      if (item.emails) {
        collect(item.emails);
        return;
      }

      Object.keys(item).forEach(key => {
        const nested = item[key];
        if (Array.isArray(nested) || (nested && typeof nested === 'object')) collect(nested);
      });
    }
  };

  collect(value);
  return result;
}

function isValidEmailAddress_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function getDepartmentReportRecipients_(deptId, deptAdminsMap) {
  return {
    to: joinEmailList_(deptAdminsMap[deptId] || []),
    cc: ''
  };
}

function getDepartmentAdminsMap_(usersData) {
  const map = {};
  const data = usersData || readSheetValues_(SpreadsheetApp.getActive().getSheetByName('Usuarios'), 7);

  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][0] || '').trim().toLowerCase();
    const deptId = data[i][2];
    const rol = String(data[i][3] || '').trim().toUpperCase();
    const estado = String(data[i][4] || '').trim().toUpperCase();

    if (email && deptId && rol === 'ADMIN_DEP' && estado === 'ACTIVO') {
      if (!map[deptId]) map[deptId] = [];
      map[deptId].push(email);
    }
  }

  Object.keys(map).forEach(deptId => {
    map[deptId] = uniqueEmailList_(map[deptId]);
  });
  return map;
}

function getRecipientEmail_(recipient) {
  if (!recipient) return '';
  if (typeof recipient === 'string') return recipient;
  return recipient.email || recipient.mail || '';
}

function splitEmailList_(value) {
  if (Array.isArray(value)) {
    return value.reduce((acc, item) => acc.concat(splitEmailList_(item)), []);
  }
  return String(value || '')
    .split(/[;,]/)
    .map(email => email.trim().toLowerCase())
    .filter(email => email);
}

function uniqueEmailList_(values) {
  const seen = {};
  const result = [];
  splitEmailList_(values).forEach(email => {
    if (!seen[email]) {
      seen[email] = true;
      result.push(email);
    }
  });
  return result;
}

function joinEmailList_(values) {
  return uniqueEmailList_(values).join(',');
}

function getUsedSheetNames_(ss) {
  const used = {};
  ss.getSheets().forEach(sheet => {
    used[sheet.getName()] = true;
  });
  return used;
}

function makeUniqueSheetName_(name, usedNames) {
  const used = usedNames || {};
  let base = String(name || 'Hoja')
    .replace(/[\[\]\*\/\\\?:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!base) base = 'Hoja';
  base = base.substring(0, 95);

  let candidate = base;
  let counter = 2;
  while (used[candidate]) {
    const suffix = ` ${counter}`;
    candidate = base.substring(0, 100 - suffix.length) + suffix;
    counter++;
  }
  return candidate;
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildDepartmentSummaryTableHtml_(deptSummary) {
  if (!deptSummary || deptSummary.length === 0) return '';

  const rows = deptSummary.map(item => `
    <tr>
      <td style="padding: 12px 14px; border-top: 1px solid #e5e7eb; color: #111827; font-weight: 700;">${escapeHtml_(item.name)}</td>
      <td style="padding: 12px 14px; border-top: 1px solid #e5e7eb; color: #2563eb; font-weight: 800; text-align: right;">${item.count}</td>
    </tr>
  `).join('');

  return `
    <div style="margin: 24px 0; border: 1px solid #dbeafe; border-radius: 14px; overflow: hidden; background: #ffffff;">
      <div style="background: #eff6ff; padding: 12px 14px; border-bottom: 1px solid #dbeafe;">
        <p style="margin: 0; color: #1e3a8a; font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.04em;">Pedidos por departamento</p>
      </div>
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse;">
        <thead>
          <tr>
            <th align="left" style="padding: 10px 14px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em;">Departamento</th>
            <th align="right" style="padding: 10px 14px; color: #6b7280; font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em;">Pedidos</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function trashTempSpreadsheet_(ss) {
  if (!ss) return;
  try {
    DriveApp.getFileById(ss.getId()).setTrashed(true);
  } catch (e) {
    console.error("Error cleaning temp report: " + e.message);
  }
}

function createGeneralReportArtifactsWithRetry_(dateStr, fileDate, orders, byDept, deptSummary, backupFolder, testRun) {
  return runSpreadsheetTransientRetry_('general daily report', function() {
    return createGeneralReportArtifacts_(dateStr, fileDate, orders, byDept, deptSummary, backupFolder, testRun);
  });
}

function createGeneralReportArtifacts_(dateStr, fileDate, orders, byDept, deptSummary, backupFolder, testRun) {
  let generalSS = null;
  try {
    generalSS = createAllDepartmentsReportFromTemplate_(dateStr, orders, byDept, deptSummary);
    const generalFileName = `[Resumen general - ${fileDate}]`;
    let excelBlob = null;
    let reportSaved = false;

    if (String(getConfigValue_('ADMIN_EMAILS') || '').trim()) {
      excelBlob = exportSheetToExcelBlob_(generalSS);
      excelBlob.setName(`${generalFileName}.xlsx`);
    }

    if (!testRun && backupFolder) {
      try {
        const pdfBlob = exportSheetToPdfBlob_(generalSS, generalSS.getSheets()[0]);
        pdfBlob.setName(`${generalFileName}.pdf`);
        backupFolder.createFile(pdfBlob);
        reportSaved = true;
      } catch (e) {
        console.error(`Error saving general report PDF backup: ${e.message}`);
      }
    }

    return {
      excelBlob: excelBlob,
      reportSaved: reportSaved
    };
  } finally {
    trashTempSpreadsheet_(generalSS);
  }
}

function runSpreadsheetTransientRetry_(label, operation, options) {
  const opts = options || {};
  const attempts = opts.attempts || SPREADSHEET_RETRY_ATTEMPTS;
  const delayMs = opts.delayMs || SPREADSHEET_RETRY_DELAY_MS;
  let lastError = null;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      return operation(attempt);
    } catch (e) {
      lastError = e;
      if (attempt >= attempts || !isRetryableSpreadsheetError_(e)) throw e;
      const errorMessage = e && e.message ? e.message : String(e);
      console.warn(`${label} attempt ${attempt} failed; retrying: ${errorMessage}`);
      Utilities.sleep(delayMs * attempt);
    }
  }

  throw lastError;
}

function isRetryableSpreadsheetError_(error) {
  const raw = String(error && error.message ? error.message : error || '').toLowerCase();
  const msg = raw.normalize ? raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : raw;
  return (
    msg.indexOf('no ha podido acceder') !== -1 ||
    msg.indexOf('no puede acceder') !== -1 ||
    msg.indexOf('could not access') !== -1 ||
    msg.indexOf('cannot access') !== -1 ||
    msg.indexOf('servicio hojas de calculo') !== -1 ||
    msg.indexOf('service spreadsheets') !== -1 ||
    msg.indexOf('returned code 5') !== -1 ||
    msg.indexOf('returned code 429') !== -1 ||
    msg.indexOf('timed out') !== -1 ||
    msg.indexOf('timeout') !== -1 ||
    msg.indexOf('backend error') !== -1 ||
    msg.indexOf('rate limit') !== -1
  );
}

function openSpreadsheetByIdWithRetry_(ssId, label) {
  return runSpreadsheetTransientRetry_(label || `open spreadsheet ${ssId}`, function() {
    return SpreadsheetApp.openById(ssId);
  }, {
    attempts: SPREADSHEET_RETRY_ATTEMPTS,
    delayMs: SPREADSHEET_RETRY_DELAY_MS
  });
}

function createReportSpreadsheetFromTemplate_(reportName) {
  const templateId = getConfigValue_('DAILY_REPORT_MODEL_ID');
  if (!templateId) throw new Error("Falta configurar DAILY_REPORT_MODEL_ID");

  const templateFile = DriveApp.getFileById(templateId);
  const newFile = templateFile.makeCopy(reportName);
  let ssId = newFile.getId();

  if (newFile.getMimeType() !== MimeType.GOOGLE_SHEETS) {
     const blob = newFile.getBlob();
     const config = {
        title: reportName,
        parents: [{id: 'root'}],
        mimeType: MimeType.GOOGLE_SHEETS
     };
     try {
       const resource = Drive.Files.create(config, blob, {convert: true});
       ssId = resource.id;
       newFile.setTrashed(true); // Delete the non-converted copy
     } catch(e) {
       newFile.setTrashed(true);
       throw new Error("Failed to convert Excel template: " + e.message);
     }
  }

  try {
    return openSpreadsheetByIdWithRetry_(ssId, `open report spreadsheet ${reportName}`);
  } catch (e) {
    try {
      DriveApp.getFileById(ssId).setTrashed(true);
    } catch (cleanupError) {
      console.error("Error cleaning inaccessible temp report: " + cleanupError.message);
    }
    throw e;
  }
}

function createReportFromTemplate_(deptName, dateStr, orders) {
  const ss = createReportSpreadsheetFromTemplate_(`Temp_Report_${deptName}_${dateStr}`);
  const sh = ss.getSheets()[0];
  fillReportSheet_(sh, deptName, dateStr, orders);
  SpreadsheetApp.flush();
  return ss;
}

function createAllDepartmentsReportFromTemplate_(dateStr, orders, byDept, deptSummary) {
  const ss = createReportSpreadsheetFromTemplate_(`Temp_Report_Resumen_General_${dateStr}`);
  const generalSheet = ss.getSheets()[0];
  const templateSheet = generalSheet.copyTo(ss);
  templateSheet.setName(makeUniqueSheetName_('__report_template__', getUsedSheetNames_(ss)));

  generalSheet.setName('Resumen general');
  fillReportSheet_(generalSheet, 'Resumen general', dateStr, orders, { preserveTitleCase: true });

  const usedNames = getUsedSheetNames_(ss);
  deptSummary.forEach(summary => {
    const sheet = templateSheet.copyTo(ss);
    const sheetName = makeUniqueSheetName_(summary.name, usedNames);
    sheet.setName(sheetName);
    usedNames[sheetName] = true;
    fillReportSheet_(sheet, summary.name, dateStr, byDept[summary.id] || []);
  });

  ss.deleteSheet(templateSheet);
  SpreadsheetApp.flush();
  return ss;
}

function fillReportSheet_(sh, deptName, dateStr, orders, options) {
  const opts = options || {};
  const title = opts.preserveTitleCase ? String(deptName) : String(deptName).toUpperCase();
  const reportCategories = getReportCategories_(orders);
  const totalColumns = 5 + reportCategories.length;
  const lastColumn = getSheetColumnLetter_(totalColumns);

  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  if (sh.getMaxColumns() < totalColumns) {
    sh.insertColumnsAfter(sh.getMaxColumns(), totalColumns - sh.getMaxColumns());
  }
  sh.getRange(3, 1, 2, sh.getMaxColumns()).breakApart();
  sh.getRange(5, 1, 1, sh.getMaxColumns()).breakApart();

  // 1. Set Dept Name
  sh.getRange('A3:' + lastColumn + '4').merge().setValue(title)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // 2. Set Date -> "PEDIDO ALMUERZO : 03/12/2025"
  const d = new Date(dateStr + 'T12:00:00');
  const fmtDate = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  sh.getRange('A5:' + lastColumn + '5').merge().setValue(`PEDIDO ALMUERZO : ${fmtDate}`)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold");

  // 3. Set Headers
  const headers = ['NO.', 'NOMBRE EMPLEADO', 'C\u00d3DIGO', 'DEPARTAMENTO']
    .concat(reportCategories.map(category => String(category.nombre || category.id).toUpperCase()))
    .concat(['NOTA PARA LA COCINA']);
  sh.getRange(7, 1, 1, totalColumns).setValues([headers])
    .setFontWeight("bold").setBorder(true, true, true, true, true, true);

  // 4. Populate Data
  const catMap = {};
  reportCategories.forEach((category, index) => { catMap[category.id] = index + 4; });

  const rows = [];
  orders.forEach((o, i) => {
     const row = new Array(totalColumns).fill('');
     row[0] = i + 1;
     row[1] = o.nombre;
     row[2] = o.codigo || '';
     row[3] = o.departamento;
     row[totalColumns - 1] = o.notaCocina || getOrderKitchenNote_(o.detail);

     const d = o.detail;
     if (d && d.categorias && d.items) {
        d.categorias.forEach((cat, idx) => {
           const colIdx = catMap[cat];
           if (colIdx !== undefined) {
              const item = d.items[idx];
              row[colIdx] = row[colIdx] ? row[colIdx] + ', ' + item : item;
           }
        });
     }
     rows.push(row);
  });

  if (rows.length > 0) {
     const range = sh.getRange(8, 1, rows.length, totalColumns); // Start A8
     range.setValues(rows);
     range.setBorder(true, true, true, true, true, true);
     range.setHorizontalAlignment("center");
     range.setVerticalAlignment("middle");
     range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
     sh.getRange(8, 2, rows.length, 1).setHorizontalAlignment("left");
  }

  applyReportSheetSizing_(sh, rows, reportCategories);
}

function getReportCategories_(orders) {
  const configured = getMenuCategories_(true);
  const byId = {};
  configured.forEach(category => { byId[category.id] = category; });
  const categoryIds = configured.map(category => category.id);

  (orders || []).forEach(order => {
    const detail = order && order.detail;
    (detail && Array.isArray(detail.categorias) ? detail.categorias : []).forEach(categoryId => {
      const id = String(categoryId || '').trim();
      if (!id || categoryIds.indexOf(id) !== -1) return;
      categoryIds.push(id);
      byId[id] = { id: id, nombre: formatCatNameForEmail_(id), orden: 999 };
    });
  });

  return categoryIds.map(id => byId[id]);
}

function getSheetColumnLetter_(columnNumber) {
  let result = '';
  let number = columnNumber;
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function applyReportSheetSizing_(sh, rows, reportCategories) {
  const widths = [42, 185, 78, 170]
    .concat((reportCategories || []).map(() => 145))
    .concat([220]);
  widths.forEach((width, idx) => sh.setColumnWidth(idx + 1, width));
  sh.setRowHeight(7, 26);

  if (!rows || rows.length === 0) return;

  SpreadsheetApp.flush();
  sh.autoResizeRows(8, rows.length);

  rows.forEach((row, idx) => {
    const sheetRow = idx + 8;
    const autoHeight = sh.getRowHeight(sheetRow);
    const targetHeight = Math.max(autoHeight, estimateReportRowHeight_(row, widths));
    sh.setRowHeight(sheetRow, targetHeight);
  });
}

function estimateReportRowHeight_(row, widths) {
  const lineHeight = 16;
  const verticalPadding = 10;
  let maxLines = 1;

  row.forEach((value, idx) => {
    const text = String(value || '').trim();
    if (!text) return;

    const width = widths[idx] || 120;
    const charsPerLine = Math.max(7, Math.floor(width / 7));
    const lines = text.split(/\r?\n/).reduce((total, line) => {
      return total + Math.max(1, Math.ceil(String(line).length / charsPerLine));
    }, 0);
    maxLines = Math.max(maxLines, lines);
  });

  return Math.min(110, Math.max(30, verticalPadding + (maxLines * lineHeight)));
}

function exportSheetToPdfBlob_(ss, sheet) {
  if (!sheet) {
    const file = DriveApp.getFileById(ss.getId());
    return file.getAs(MimeType.PDF);
  }

  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sheet.getSheetId()}&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  return response.getBlob();
}

function exportSheetToExcelBlob_(ss) {
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  return response.getBlob();
}

function getDailyBackupFolder_(dateStr) {
  let rootId = getConfigValue_('BACKUP_FOLDER_ID');
  if (!rootId) {
     const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
     const parents = ssFile.getParents();
     if (parents.hasNext()) {
        const parent = parents.next();
        const newFolder = parent.createFolder('Backups_Almuerzo');
        rootId = newFolder.getId();
        const cSh = SpreadsheetApp.getActive().getSheetByName('Config');
        const data = cSh.getDataRange().getValues();
        for(let i=1; i<data.length; i++) {
           if (data[i][0] === 'BACKUP_FOLDER_ID') {
              cSh.getRange(i+1, 2).setValue(rootId);
              break;
           }
        }
        _configCache = null;
     } else {
        throw new Error("No parent folder found");
     }
  }

  const rootFolder = DriveApp.getFolderById(rootId);
  const d = new Date(dateStr + 'T12:00:00');
  const year = String(d.getFullYear());
  const month = String(d.getMonth() + 1).padStart(2, '0');

  let yFolder = rootFolder.getFoldersByName(year).hasNext() ? rootFolder.getFoldersByName(year).next() : rootFolder.createFolder(year);
  let mFolder = yFolder.getFoldersByName(month).hasNext() ? yFolder.getFoldersByName(month).next() : yFolder.createFolder(month);

  return mFolder;
}

/**
 * Instala TODOS los triggers necesarios (Time-based y Edit-based).
 * Ejecutar esto manualmente una vez para inicializar.
 */
function installTriggers() {
  const ss = SpreadsheetApp.getActive();

  // 1. Manage Spreadsheet OnEdit Trigger (Persistent)
  const triggers = ScriptApp.getProjectTriggers();
  let onEditExists = false;

  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'onSpreadsheetEdit') {
      onEditExists = true;
    }
  });

  if (!onEditExists) {
    ScriptApp.newTrigger('onSpreadsheetEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    console.log("Trigger 'onSpreadsheetEdit' instalado.");
  }

  // 2. Install Time Triggers
  reinstallTimeTriggers_();
}

function reinstallTimeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  // Delete only time triggers (scheduledSendReminders, scheduledDailyClose)
  const targets = ['scheduledSendReminders', 'scheduledDailyClose'];

  triggers.forEach(t => {
    if (targets.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Helper to parse HH:mm
  const parseTime = (val, defH, defM) => {
     if (val instanceof Date) return { h: val.getHours(), m: val.getMinutes() };
     if (typeof val === 'string' && val.includes(':')) {
        const p = val.split(':');
        return { h: parseInt(p[0]||defH), m: parseInt(p[1]||defM) };
     }
     return { h: defH, m: defM };
  };
  
  // Recordatorios (HORA_RECORDATORIO)
  const recTime = parseTime(getConfigValue_('HORA_RECORDATORIO'), 13, 0); // Default 1:00 PM
  ScriptApp.newTrigger('scheduledSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(recTime.h)
    .nearMinute(recTime.m)
    .create();
    
  // Cierre y Reportes (HORA_ENVIO)
  const closeTime = parseTime(getConfigValue_('HORA_ENVIO'), 15, 0); // Default 3:00 PM
  ScriptApp.newTrigger('scheduledDailyClose')
    .timeBased()
    .everyDays(1)
    .atHour(closeTime.h)
    .nearMinute(closeTime.m)
    .create();
    
  console.log(`Triggers de tiempo reinstalados. Recordatorio: ${recTime.h}:${recTime.m}, Cierre: ${closeTime.h}:${closeTime.m}`);
}

function onSpreadsheetEdit(e) {
  // Check if edit is in Config sheet
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Config') return;

  // Check if edited column is Value (Col 2) or Key (Col 1)
  // We care if the Key (Col 1) corresponding to this row is HORA_RECORDATORIO or HORA_ENVIO
  const row = range.getRow();
  if (row <= 1) return; // Header

  const key = sheet.getRange(row, 1).getValue();
  if (key === 'HORA_RECORDATORIO' || key === 'HORA_ENVIO') {
    console.log(`Detectado cambio en ${key}. Reinstalando triggers...`);
    // Invalidate Cache
    _configCache = null;
    reinstallTimeTriggers_();
  }
}

function getAppUrl_() {
   const url = getConfigValue_('APP_URL');
   return url || ScriptApp.getService().getUrl();
}
