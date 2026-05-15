/**
 * Code.gs - Backend V5 (Refactor & New Features)
 */
const APP_VERSION = 'v7.9';

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
    const usersData = readSheetValues_(ss.getSheetByName('Usuarios'), 7);
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

    // Get Banner Text from Config
    const bannerText = getConfigValue_('PLAN_WEEK_TEXT') || 'Planifica tu semana';
    const bannerLimit = parseInt(getConfigValue_('PLAN_WEEK_LIMIT') || '5', 10);
    const mealPriceCurrent = getCurrentMealPrice_();
    const mealPriceHistory = getMealPriceHistory_();
    const hintConfig = getHintConfig_();
    const todayYmd = getTodayYmd_();

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
      bannerConfig: { text: bannerText, limit: bannerLimit },
      mealPricing: { current: mealPriceCurrent, history: mealPriceHistory },
      hintConfig: hintConfig,
      todayYmd: todayYmd,
      deptMap: deptMap
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

function apiDismissBanner() {
   const user = getUserInfo_();
   const prefs = getUserPrefs_(user.email);
   let count = prefs.banner_count || 0;
   count++;
   return apiSetUserPreference('banner_count', count);
}

function apiDismissSummaryCostHint() {
   const user = getUserInfo_();
   const prefs = getUserPrefs_(user.email);
   let count = prefs.summary_cost_hint_count || 0;
   count++;
   return apiSetUserPreference('summary_cost_hint_count', count);
}

function apiDismissCaldoMultiHint() {
   const user = getUserInfo_();
   const prefs = getUserPrefs_(user.email);
   let count = prefs.caldo_multi_hint_count || 0;
   count++;
   return apiSetUserPreference('caldo_multi_hint_count', count);
}
// === AUTOMATIZACIÓN (TRIGGERS) ===

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
  // Only run on business days
  if (!isTodayBusinessDay_()) {
     console.log("Skipping scheduledDailyClose: Not a business day.");
     return;
  }

  const now = new Date();
  const nextBusinessDay = getNextBusinessDay_(now);
  if (!nextBusinessDay) return;

  const dateStr = formatDate_(nextBusinessDay);

  // 1. Get Orders and Group by Dept
  const orders = getOrdersByDateDetailed_(dateStr);
  const byDept = {};
  orders.forEach(o => {
    const deptId = o.departamentoId || 'Sin Depto';
    if (!byDept[deptId]) byDept[deptId] = [];
    byDept[deptId].push(o);
  });

  // 2. Prepare Recipients Logic
  const rawConfig = getConfigValue_('RESPONSIBLES_EMAILS_JSON');
  let configRecipients = null;
  try { configRecipients = JSON.parse(rawConfig); } catch(e) {}

  const getRecipientsForDept = (deptId) => {
      let list = [];
      if (Array.isArray(configRecipients)) {
          list = configRecipients; // Global list
      } else if (configRecipients && typeof configRecipients === 'object') {
          // Map: try specific dept
          if (configRecipients[deptId] && Array.isArray(configRecipients[deptId])) {
             list = configRecipients[deptId];
          }
      }

      return {
          to: list.filter(r => r.type === 'TO').map(r => r.email).join(','),
          cc: list.filter(r => r.type === 'CC').map(r => r.email).join(',')
      };
  };

  const deptMap = getDepartmentMap_();
  const backupFolder = getDailyBackupFolder_(dateStr);

  // =================================================================
  // NUEVO: Cargar Admins de Departamento (Mapeo ID -> [Emails])
  // =================================================================
  const deptAdminsMap = {};
  const uSh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  const uData = uSh.getDataRange().getValues();
  
  // Empezamos en 1 para saltar encabezados
  for(let i=1; i<uData.length; i++) {
     const email = String(uData[i][0]).toLowerCase();
     const deptId = uData[i][2];
     const rol = uData[i][3];
     const estado = uData[i][4];
     
     // Filtramos solo ADMIN_DEP que estén ACTIVOS
     if (rol === 'ADMIN_DEP' && estado === 'ACTIVO') {
        if (!deptAdminsMap[deptId]) deptAdminsMap[deptId] = [];
        deptAdminsMap[deptId].push(email);
     }
  }
  // =================================================================

  // 3. Process each Department
  Object.keys(byDept).forEach(deptId => {
    const deptName = deptMap[deptId] || deptId;
    const deptOrders = byDept[deptId];

    // Resolve recipients for this department
    // CAMBIO: Usamos 'let' en lugar de 'const' para poder modificar ccList
    let { to: toList, cc: ccList } = getRecipientsForDept(deptId);

    // ===============================================================
    // NUEVO: Inyectar Admins del Depto actual en el CC
    // ===============================================================
    const admins = deptAdminsMap[deptId];
    if (admins && admins.length > 0) {
       // Convertimos el string actual de CC en array
       let ccArray = ccList ? ccList.split(',').map(e => e.trim()).filter(e => e) : [];
       
       // Agregamos los admins si no están ya en la lista
       admins.forEach(adminEmail => {
          if (!ccArray.includes(adminEmail)) {
             ccArray.push(adminEmail);
          }
       });
       
       // Volvemos a unir en string
       ccList = ccArray.join(',');
    }
    // ===============================================================

    // Format Filename Date: dd-MM-yyyy
    const d = new Date(dateStr + 'T12:00:00');
    const fDate = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    const fileName = `[${deptName} - ${fDate}]`;

    try {
      // Create Report SS from Template
      const tempSS = createReportFromTemplate_(deptName, dateStr, deptOrders);

      // Export PDF -> Backup
      const pdfBlob = exportSheetToPdfBlob_(tempSS);
      pdfBlob.setName(`${fileName}.pdf`);
      backupFolder.createFile(pdfBlob);

      // Export Excel -> Email
      if (toList || ccList) {
        const excelBlob = exportSheetToExcelBlob_(tempSS);
        excelBlob.setName(`${fileName}.xlsx`);

        const formattedDate = Utilities.formatDate(new Date(dateStr + 'T12:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy');

        const html = getEmailTemplate_({
           title: `Reporte ${deptName}`,
           subtitle: `Pedidos para el ${formattedDate}`,
           body: `
             <p>Buenas tardes estimados,</p>
             <p>Hay <strong>${deptOrders.length}</strong> pedidos registrados del departamento de <strong>${deptName}</strong> para el día <strong>${formattedDate}</strong>.</p>
             <p>Favor revisar el archivo Excel adjunto para más detalles sobre los platos solicitados.</p>
             <p>Cualquier duda, estamos a la orden.</p>
           `,
           footerNote: 'Este reporte se genera automáticamente al cierre de pedidos.'
        });

        sendEmail_(toList, `Almuerzo Pre-empacado | Reporte Almuerzo ${deptName} - ${dateStr}`, html, ccList, [excelBlob]);
      } else {
         console.warn(`No recipients found for department ${deptName} (${deptId}). Report saved to backup only.`);
      }

      // Cleanup
      DriveApp.getFileById(tempSS.getId()).setTrashed(true);

    } catch(e) {
      console.error(`Error processing report for ${deptName}: ${e.message}`);
    }
  });

  // 4. Maintenance & Admin Summary
  checkMenuIntegrity_();
  sendDailyAdminSummary_(dateStr);
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
    }

    data.departments = getDepartmentsList_(); // Returns {id, nombre...}

    writeJsonCache_(cacheKey, data, 45);
    return data;
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSaveConfig(configData) {
   try {
     const admin = getUserInfo_();
     if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");

     ensureOperationalConfigKeys_();
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

   if (!isDateOpenForOrdering_(dateStr)) {
      throw new Error("No puedes editar el menú de una fecha no hábil, pasada o cerrada.");
   }

   const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
   const data = sh.getDataRange().getValues();
   let rowIdx = -1;

   if (itemData.id) {
      for(let i=1; i<data.length; i++) {
         if (String(data[i][0]) === String(itemData.id)) { rowIdx = i+1; break; }
      }
   }

   const id = rowIdx > 0 ? itemData.id : Utilities.getUuid();
   // Save as Date object (local)
   const dateObj = new Date(dateStr + 'T12:00:00');
   const row = [id, dateObj, cat, normalizeMenuText_(itemData.plato), normalizeMenuText_(itemData.desc), 'SI'];

   if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
   else sh.appendRow(row);

   invalidateMenuDataCache_();
   invalidateUserInitCache_();
   return { ok: true };
}

function apiDeleteMenuItem(id) {
   const admin = getUserInfo_();
   if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Denegado");
   const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
   const data = sh.getDataRange().getValues();
   for(let i=1; i<data.length; i++) {
      if (String(data[i][0]) === String(id)) {
         sh.deleteRow(i+1);
         invalidateMenuDataCache_();
         invalidateUserInitCache_();
         return { ok: true };
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

   // Normalize keys to ensure matching
   const datesToUpdate = new Set();
   Object.keys(menuData).forEach(k => datesToUpdate.add(formatDate_(new Date(k + 'T12:00:00'))));

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
   // Iterate over the keys provided by client to maintain association
   Object.keys(menuData).forEach(dateKey => {
      const normalizedDate = formatDate_(new Date(dateKey + 'T12:00:00'));
      // Only proceed if it was marked for update (double check)
      if (datesToUpdate.has(normalizedDate)) {
         const items = menuData[dateKey] || [];
         const dateObj = new Date(normalizedDate + 'T12:00:00');
         items.forEach(item => {
            allNewRows.push([
              Utilities.getUuid(),
              dateObj,
              item.cat,
              normalizeMenuText_(item.plato),
              normalizeMenuText_(item.desc || ''),
              'SI'
            ]);
         });
      }
   });

   // 4. Append
   if (allNewRows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, allNewRows.length, allNewRows[0].length).setValues(allNewRows);
   }

   invalidateMenuDataCache_();
   invalidateUserInitCache_();
   return { ok: true };
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
   // Check duplicate
   for(let i=1; i<data.length; i++) {
      if (formatDate_(new Date(data[i][0])) === dateStr) {
         sh.getRange(i+1, 2).setValue(desc);
         return { ok: true };
      }
   }
   sh.appendRow([dateStr, desc]);
   CacheService.getScriptCache().remove('HOLIDAYS_CACHE_V2');
   invalidateMenuDataCache_();
   invalidateUserInitCache_();
   return { ok: true };
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
  // Si no hay usuario autenticado, salimos
  if (!user) return { count: null };

  const lock = LockService.getScriptLock();
  // Intentamos bloquear por 2s para evitar colisiones de escritura en la caché
  if (lock.tryLock(2000)) {
    try {
      const cache = CacheService.getScriptCache();
      const KEY = 'ACTIVE_SESSIONS_V1';
      const raw = cache.get(KEY);
      let sessions = raw ? JSON.parse(raw) : {};
      
      const now = Date.now();
      const TIME_WINDOW = 5 * 60 * 1000; // 5 minutos de inactividad para considerar "offline"

      // 1. Registrar/Actualizar al usuario actual
      sessions[user.email] = now;

      // 2. Limpiar usuarios antiguos y contar
      let count = 0;
      const cleanSessions = {};
      Object.keys(sessions).forEach(email => {
         if (now - sessions[email] < TIME_WINDOW) {
            cleanSessions[email] = sessions[email];
            count++;
         }
      });

      // 3. Guardar cambios (TTL de 6 horas para el contenedor)
      cache.put(KEY, JSON.stringify(cleanSessions), 21600);

      // 4. Retornar conteo SOLO si es Administrador General
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
    'INIT',
    getRevisionValue_('APP_INIT_REVISION'),
    String(activeEmail || '').toLowerCase(),
    String(targetEmail || '').toLowerCase(),
    requestedDateStr || 'AUTO'
  ].join(':');
}

function getAdminCacheKey_(user) {
  return [
    'ADMIN',
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
  return [
    { key: 'LOGO_ID', value: '', description: 'ID del archivo de imagen del Logo en Drive' },
    { key: 'APP_URL', value: ScriptApp.getService().getUrl(), description: 'URL publica de la aplicacion (Web App)' },
    { key: 'MEAL_PRICE_CURRENT', value: '57', description: 'Costo actual por almuerzo. Al cambiarlo se conserva historial automatico por fecha.' },
    { key: 'MEAL_PRICE_HISTORY_JSON', value: '[{"from":"1900-01-01","price":57}]', description: 'Historial auto-administrado del costo por almuerzo. No editar manualmente.' },
    { key: 'MENU_DAY_ENDPOINT_TOKEN', value: generateSecretToken_(), description: 'Token secreto para consumir el endpoint JSON de menu por fecha. Generar y compartir solo con TI.' },
    { key: 'SUMMARY_COST_HINT_LIMIT', value: '3', description: 'Cantidad maxima de cierres del hint del costo acumulado antes de ocultarlo.' },
    { key: 'SUMMARY_COST_HINT_EXPIRES_ON', value: defaultExpiry, description: 'Fecha limite para mostrar el hint del costo acumulado (YYYY-MM-DD).' },
    { key: 'CALDO_MULTI_HINT_LIMIT', value: '3', description: 'Cantidad maxima de cierres del hint de multiseleccion en Caldo.' },
    { key: 'CALDO_MULTI_HINT_EXPIRES_ON', value: defaultExpiry, description: 'Fecha limite para mostrar el hint de multiseleccion en Caldo (YYYY-MM-DD).' }
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

function ensureHintConfigKeys_() {
  const defaultExpiry = formatDateWithOffset_(30);
  ensureConfigKey_('SUMMARY_COST_HINT_LIMIT', '3', 'Cantidad maxima de cierres del hint del costo acumulado antes de ocultarlo.');
  ensureConfigKey_('SUMMARY_COST_HINT_EXPIRES_ON', defaultExpiry, 'Fecha limite para mostrar el hint del costo acumulado (YYYY-MM-DD).');
  ensureConfigKey_('CALDO_MULTI_HINT_LIMIT', '3', 'Cantidad maxima de cierres del hint de multiseleccion en Caldo.');
  ensureConfigKey_('CALDO_MULTI_HINT_EXPIRES_ON', defaultExpiry, 'Fecha limite para mostrar el hint de multiseleccion en Caldo (YYYY-MM-DD).');
}

function getHintConfig_() {
  return {
    summaryCost: {
      limit: parsePositiveInt_(getConfigValue_('SUMMARY_COST_HINT_LIMIT'), 3),
      expiresOn: normalizeHintDate_(getConfigValue_('SUMMARY_COST_HINT_EXPIRES_ON'), formatDateWithOffset_(30))
    },
    caldoMulti: {
      limit: parsePositiveInt_(getConfigValue_('CALDO_MULTI_HINT_LIMIT'), 3),
      expiresOn: normalizeHintDate_(getConfigValue_('CALDO_MULTI_HINT_EXPIRES_ON'), formatDateWithOffset_(30))
    }
  };
}

function normalizeHintDate_(value, fallback) {
  const normalized = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : fallback;
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
  const specialList = ['Vegetariana', 'Caldo', 'Opcion_Rapida'];
  const hasSpecial = cats.some(c => specialList.includes(c));
  if (hasSpecial && cats.length > 1) {
     const uniqueCats = [...new Set(cats)];
     if (uniqueCats.some(c => !specialList.includes(c))) throw new Error("Platos especiales no se pueden combinar con el menú regular.");
  }
  if (cats.includes('Granos')) {
    const hasWhiteRice = items.some(i => i.toLowerCase().includes('arroz blanco'));
    if (!hasWhiteRice) throw new Error("Los granos requieren seleccionar Arroz Blanco.");
  }
  if (cats.includes('Arroces') && cats.includes('Viveres')) throw new Error("No puedes combinar Arroz y Víveres.");
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
  const testMode = getConfigValue_('TEST_EMAIL_MODE') === 'TRUE';
  const testDest = getConfigValue_('TEST_EMAIL_DEST');
  const senderName = getConfigValue_('MAIL_SENDER_NAME');

  const recipient = testMode ? testDest : to;
  if (!recipient) return;

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
  if (testMode && cc) options.htmlBody = `<p><strong>[Original CC: ${cc}]</strong></p>` + options.htmlBody;

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

function sendDailyAdminSummary_(dateStr) {
  const admins = getConfigValue_('ADMIN_EMAILS');
  if (!admins) return;
  const orders = getOrdersByDate_(dateStr);
  const count = orders.length;
  if (count > 0) {
     const formattedDate = Utilities.formatDate(new Date(dateStr + 'T12:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy');
     const html = getEmailTemplate_({
        title: 'Resumen Diario',
        subtitle: `Pedidos para el ${formattedDate}`,
        body: `
           <p>Resumen ejecutivo de la operación de almuerzo:</p>
           <div style="text-align: center; margin: 24px 0;">
              <span style="font-size: 48px; font-weight: 800; color: #111827;">${count}</span>
              <p style="color: #6b7280; margin-top: 8px;">Pedidos Totales</p>
           </div>
           <p>Los respaldos detallados han sido generados y guardados en Google Drive.</p>
        `,
        cta: { text: 'Ver Panel Administrativo', url: getAppUrl_() }
     });
    sendEmail_(admins, `Almuerzo Pre-empacado | Resumen Pedidos ${dateStr}`, html);
  }
}

function getOrdersByDateDetailed_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const deptMap = getDepartmentMap_();
  const codeMap = getUserCodeMap_();
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
        codigo: codeMap[email] || ''
      });
    }
  }
  return list;
}

function getUserCodeMap_() {
   const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
   const data = sh.getDataRange().getValues();
   const map = {};
   for(let i=1; i<data.length; i++) {
      const email = String(data[i][0]).toLowerCase();
      const code = data[i][6]; // Index 6 is Code
      if(code) map[email] = String(code);
   }
   return map;
}

function createReportFromTemplate_(deptName, dateStr, orders) {
  const templateId = getConfigValue_('DAILY_REPORT_MODEL_ID');
  if (!templateId) throw new Error("Falta configurar DAILY_REPORT_MODEL_ID");

  // Copy Template (Handling .xlsx if needed by converting)
  const templateFile = DriveApp.getFileById(templateId);
  const newFile = templateFile.makeCopy(`Temp_Report_${deptName}_${dateStr}`);

  // If original is .xlsx, makeCopy keeps it as .xlsx (Blob) or Google Sheet depending on settings?
  // DriveApp.makeCopy of non-native file creates non-native copy.
  // We need to convert it.

  let ssId = newFile.getId();

  // Check if we need to convert (if mimeType is not Google Spreadsheet)
  if (newFile.getMimeType() !== MimeType.GOOGLE_SHEETS) {
     const blob = newFile.getBlob();
     const config = {
        title: `Temp_Report_${deptName}_${dateStr}`,
        parents: [{id: 'root'}], // Temporary location
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

  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheets()[0]; // Assume first sheet

  // 1. Set Dept Name (B3:M4)
  sh.getRange("B3:M4").merge().setValue(deptName.toUpperCase())
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // 2. Set Date (B5:M5) -> "PEDIDO ALMUERZO : 03/12/2025"
  const d = new Date(dateStr + 'T12:00:00');
  const fmtDate = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  sh.getRange("B5:M5").merge().setValue(`PEDIDO ALMUERZO : ${fmtDate}`)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold");

  // 3. Set Headers (B7:M7)
  const headers = ['NO.', 'NOMBRE EMPLEADO', 'CÓDIGO', 'DEPARTAMENTO', 'ARROCES', 'GRANOS', 'CARNES', 'VIVERES', 'ESPECIALIDADES', 'ENSALADAS', 'CALDO', 'OPCION RAPIDA'];
  sh.getRange("B7:M7").setValues([headers])
    .setFontWeight("bold").setBorder(true, true, true, true, true, true);

  // 4. Populate Data
  // Mapping categories to columns indices (relative to B, so 0-based index in values array)
  // Headers: No(0), Nombre(1), Cod(2), Dept(3), Arroz(4), Granos(5), Carnes(6), Viveres(7), Esp(8), Ens(9), Caldo(10), OpRap(11)

  const catMap = {
     'Arroces': 4,
     'Granos': 5,
     'Carnes': 6,
     'Viveres': 7,
     'Vegetariana': 8, // Especialidades
     'Ensaladas': 9,
     'Caldo': 10,
     'Opcion_Rapida': 11
  };

  const rows = [];
  orders.forEach((o, i) => {
     // Get user code if available (need to fetch from user object or pass it)
     // orders object from getOrdersByDateDetailed_ doesn't have code.
     // We might need to fetch it or ignore it.
     // Let's try to get it from cache or efficient lookup if possible.
     // For now, empty or fetch? fetching one by one is slow.
     // Optimization: getOrdersByDateDetailed_ could include code.

     // Let's assume we want to fix getOrdersByDateDetailed_ to include Code.
     // But for now, let's look at the current row structure.

     const row = new Array(12).fill('');
     row[0] = i + 1;
     row[1] = o.nombre;
     row[2] = o.codigo || ''; // Need to ensure 'codigo' is passed
     row[3] = o.departamento;

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
     const range = sh.getRange(8, 2, rows.length, 12); // Start B8
     range.setValues(rows);
     range.setBorder(true, true, true, true, true, true);
     range.setHorizontalAlignment("center");
     range.setVerticalAlignment("middle");
     range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }

  sh.setColumnWidth(2, 40); // Fix width for 'NO.' column
  sh.autoResizeColumns(3, 11); // Auto-resize rest (C-M)

  // Add padding to prevent tight text wrapping
  for(let i=3; i<=13; i++) {
     const w = sh.getColumnWidth(i);
     sh.setColumnWidth(i, w + 20);
  }

  SpreadsheetApp.flush(); // FORCE SAVE before export
  return ss;
}

function exportSheetToPdfBlob_(ss) {
  const file = DriveApp.getFileById(ss.getId());
  return file.getAs(MimeType.PDF);
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
