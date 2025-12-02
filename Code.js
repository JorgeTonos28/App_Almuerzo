/**
 * Code.gs - Backend V5 (Refactor & New Features)
 */
const APP_VERSION = 'v6.00';

// === RUTAS E INICIO ===

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('index');
  const user = getUserInfo_();

  // Inyectar firma
  t.signatureUrl = getSignatureDataUrl_();

  if (!user || user.estado !== 'ACTIVO') {
    const denied = HtmlService.createTemplateFromFile('Denied');
    denied.signatureUrl = t.signatureUrl;
    denied.email = Session.getActiveUser().getEmail().toLowerCase();
    denied.status = user ? user.estado : null;
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

function apiGetInitData(requestedDateStr, impersonateEmail) {
  try {
    const activeUser = getUserInfo_();
    if (!activeUser) throw new Error("Usuario no encontrado.");

    let targetUser = activeUser;
    let deptUsers = [];

    // Logic for Impersonation (ADMIN_DEP only)
    if (activeUser.rol === 'ADMIN_DEP') {
       // Filter out self
       deptUsers = getUsersByDept_(activeUser.departamentoId).filter(u => u.email.toLowerCase() !== activeUser.email.toLowerCase());

       if (impersonateEmail && impersonateEmail !== activeUser.email) {
          const checkUser = getUserInfo_(impersonateEmail);
          if (checkUser && checkUser.departamentoId === activeUser.departamentoId) {
             targetUser = checkUser;
          }
       }
    }

    const availableDates = getAvailableMenuDates_(true);
    if (availableDates.length === 0) {
      return { ok: true, empty: true, msg: "No hay menús disponibles." };
    }

    let targetDateStr = requestedDateStr;
    if (!targetDateStr || !availableDates.some(d => d.value === targetDateStr)) {
      targetDateStr = availableDates[0].value;
    }

    const allMenus = getAllMenus_(availableDates);
    const allOrders = getAllUserOrders_(targetUser.email, availableDates);

    const menu = allMenus[targetDateStr] || {};
    const existingOrder = allOrders[targetDateStr] || null;

    let adminSummary = null;
    if (activeUser.rol === 'ADMIN_GEN' || activeUser.rol === 'ADMIN_DEP') {
      adminSummary = getDepartmentStats_(targetDateStr, (activeUser.rol === 'ADMIN_GEN' ? null : activeUser.departamentoId));
    }

    const prefs = getUserPrefs_(targetUser.email);

    // Get Banner Text from Config
    const bannerText = getConfigValue_('PLAN_WEEK_TEXT') || 'Planifica tu semana';
    const bannerLimit = parseInt(getConfigValue_('PLAN_WEEK_LIMIT') || '5', 10);

    const nextBizDay = getNextBusinessDay_(new Date());

    return {
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
      deptMap: getDepartmentMap_()
    };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
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
        sendEmail_(admins, "Nueva Solicitud de Acceso",
           `El usuario <b>${data.name}</b> (${email}) ha solicitado acceso al sistema de almuerzo.<br>` +
           `Departamento: ${data.dept}<br><br>` +
           `Ingresa al Panel de Administración para aprobarlo.`
        );
     }

     return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function apiSubmitOrder(payload) {
  try {
    const activeUser = getUserInfo_();
    let targetUser = activeUser;

    if (payload.impersonateEmail && activeUser.rol === 'ADMIN_DEP') {
       const checkUser = getUserInfo_(payload.impersonateEmail);
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
    saveOrderToSheet_(targetUser, dateStr, payload, activeUser.email);

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
      // Rough calc for display
      sendEmail_(email, "Recordatorio de Almuerzo",
        `Hola ${row[1]},<br><br>Aún no has realizado tu pedido para el <b>${formatDisplayDate_(dateStr)}</b>.<br>` +
        `Recuerda pedir antes del cierre.<br><br>` +
        `<a href="${ScriptApp.getService().getUrl()}">Ir a la App</a>`
      );
    }
  });
}

function scheduledDailyClose() {
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

  // 3. Process each Department
  Object.keys(byDept).forEach(deptId => {
    const deptName = deptMap[deptId] || deptId;
    const deptOrders = byDept[deptId];

    // Resolve recipients for this department
    const { to: toList, cc: ccList } = getRecipientsForDept(deptId);

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

        sendEmail_(toList, `Reporte Almuerzo ${deptName} - ${dateStr}`,
          `<h3>Pedidos para ${formatDisplayDate_(dateStr)} - ${deptName}</h3>` +
          `<p>Total platos: ${deptOrders.length}</p>` +
          `<p>Se adjunta el reporte en Excel.</p>`,
          ccList,
          [excelBlob]
        );
      } else {
         console.warn(`No recipients found for department ${deptName} (${deptId}). Report saved to backup only.`);
      }

      // Cleanup
      DriveApp.getFileById(tempSS.getId()).setTrashed(true);

    } catch(e) {
      console.error(`Error processing report for ${deptName}: ${e.message}`);
    }
  });

  // 4. Maintenance
  checkMenuIntegrity_();
}

function scheduledDepartmentReports() {
  // Deprecated. Logic moved to scheduledDailyClose.
  console.log("Trigger scheduledDepartmentReports is deprecated.");
}

// === ADMIN API ===

function apiGetAdminData() {
  try {
    const user = getUserInfo_();
    if (!user || (user.rol !== 'ADMIN_GEN' && user.rol !== 'ADMIN_DEP')) {
      return { ok: false, msg: "Acceso denegado." };
    }

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
       data.config = getConfigValue_('ALL');
       for (const k in data.config) data.config[k] = String(data.config[k]);
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

    return data;
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

function apiSaveConfig(configData) {
   try {
     const admin = getUserInfo_();
     if (!admin || admin.rol !== 'ADMIN_GEN') throw new Error("Permiso denegado.");

     const ss = SpreadsheetApp.getActive();
     const sh = ss.getSheetByName('Config');
     const data = sh.getDataRange().getValues();

     for(let i=1; i<data.length; i++) {
        const key = String(data[i][0]);
        if (configData[key] !== undefined) {
           sh.getRange(i+1, 2).setValue(configData[key]);
        }
     }
     _configCache = null;
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
     sendEmail_(userData.email, "Acceso Aprobado - Almuerzo",
        `Hola ${userData.nombre},<br><br>` +
        `Tu cuenta ha sido activada exitosamente.<br>` +
        `Ya puedes ingresar al sistema para realizar tus pedidos.<br><br>` +
        `<a href="${ScriptApp.getService().getUrl()}">Ingresar a la App</a>`
     );
  }

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
         return { ok: true };
      }
   }
   return { ok: false, msg: "No encontrado" };
}

function apiAdminCancelOrder(orderId) {
  const admin = getUserInfo_();
  if (!admin || !['ADMIN_GEN', 'ADMIN_DEP'].includes(admin.rol)) throw new Error("Permiso denegado.");
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Pedidos');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(orderId)) {
       const orderDeptId = rows[i][5];
       if (admin.rol === 'ADMIN_DEP' && orderDeptId !== admin.departamentoId) {
          throw new Error("Denegado: Pedido de otro departamento.");
       }
       sh.deleteRow(i + 1);
       return { ok: true };
    }
  }
  return { ok: false, msg: "Pedido no encontrado." };
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
         items.push({ id: data[i][0], cat: data[i][2], plato: data[i][3], desc: data[i][4], hab: data[i][5] });
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
   const row = [id, dateObj, cat, itemData.plato, itemData.desc, 'SI'];

   if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
   else sh.appendRow(row);

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
            allNewRows.push([Utilities.getUuid(), dateObj, item.cat, item.plato, item.desc || '', 'SI']);
         });
      }
   });

   // 4. Append
   if (allNewRows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, allNewRows.length, allNewRows[0].length).setValues(allNewRows);
   }

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
         return { ok: true };
      }
   }
   return { ok: false };
}

// === UTILS ===

function getUserInfo_(targetEmail) {
  const email = targetEmail ? targetEmail.toLowerCase() : Session.getActiveUser().getEmail().toLowerCase();
  const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  const data = sh.getDataRange().getValues();
  const deptMap = getDepartmentMap_();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email) {
      const deptId = data[i][2];
      return {
        email: data[i][0],
        nombre: data[i][1],
        departamentoId: deptId,
        departamento: deptMap[deptId] || deptId, // Resolve name
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
      const data = sh.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
         map[data[i][0]] = data[i][1]; // ID -> Name
      }
   }
   return map;
}

function getDepartmentsList_() {
   const sh = SpreadsheetApp.getActive().getSheetByName('Departamentos');
   if (!sh) return [];
   const data = sh.getDataRange().getValues();
   return data.slice(1).map(r => ({ id: r[0], nombre: r[1], admins: r[2], estado: r[3] }));
}

function getUsersByDept_(deptId) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  const data = sh.getDataRange().getValues();
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

function getAllMenus_(availableDates) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Menu');
  const data = sh.getDataRange().getValues();
  const menuMap = {};
  const validDates = new Set(availableDates.map(d => d.value));
  availableDates.forEach(d => { menuMap[d.value] = {}; });
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

function getAllUserOrders_(email, availableDates) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const ordersMap = {};
  const validDates = new Set(availableDates.map(d => d.value));
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (validDates.has(rowDate) && String(data[i][3]).toLowerCase() === email && data[i][8] !== 'CANCELADO') {
      ordersMap[rowDate] = { id: data[i][0], resumen: data[i][6], detalle: JSON.parse(data[i][7] || '{}') };
    }
  }
  return ordersMap;
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

function getDepartmentStats_(dateStr, deptIdFilter) {
  const departmentStats = { total: 0, byUser: [] };
  const orders = getOrdersByDate_(dateStr);
  orders.forEach(o => {
    // Check ID
    if (!deptIdFilter || o.departamentoId === deptIdFilter) {
      departmentStats.total++;
      departmentStats.byUser.push({ nombre: o.nombre, pedido: o.resumen, depto: o.departamento });
    }
  });
  return departmentStats;
}

function getOrdersByDate_(dateStr) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  const data = sh.getDataRange().getValues();
  const deptMap = getDepartmentMap_();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatDate_(new Date(data[i][2]));
    if (rowDate === dateStr && data[i][8] !== 'CANCELADO') {
      list.push({
        nombre: data[i][4],
        departamentoId: data[i][5],
        departamento: deptMap[data[i][5]] || data[i][5],
        resumen: data[i][6]
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

function saveOrderToSheet_(user, dateStr, selection, creatorEmail) {
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

  // Save ID in col 6 (Index 5)
  const rowData = [
    id, now, dateStr, user.email, user.nombre, user.departamentoId,
    selection.items.join(', '), JSON.stringify(selection), 'ACTIVO', now,
    creatorEmail || user.email
  ];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  else sh.appendRow(rowData);
}

function sendEmail_(to, subject, htmlBody, cc, attachments) {
  const testMode = getConfigValue_('TEST_EMAIL_MODE') === 'TRUE';
  const testDest = getConfigValue_('TEST_EMAIL_DEST');
  const senderName = getConfigValue_('MAIL_SENDER_NAME');

  const recipient = testMode ? testDest : to;
  if (!recipient) return;

  const finalSubject = testMode ? `[TEST] ${subject}` : subject;

  const sigUrl = getSignatureDataUrl_();
  const signatureHtml = sigUrl ? `<br><br><img src="${sigUrl}" style="max-height:100px;">` : '';

  const options = {
    to: recipient,
    subject: finalSubject,
    htmlBody: htmlBody + signatureHtml,
    name: senderName
  };

  if (cc && !testMode) options.cc = cc;
  if (testMode && cc) options.htmlBody = `<p><strong>[Original CC: ${cc}]</strong></p>` + options.htmlBody;
  if (attachments) options.attachments = attachments;

  try {
    MailApp.sendEmail(options);
  } catch(e) {
    console.error("Email error: " + e.message);
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
    if (admins) sendEmail_(admins, "Alerta: Integridad de Menú", `Falta arroz blanco: ${warnings.join(', ')}`);
  }
}

function sendDailyAdminSummary_(dateStr) {
  const admins = getConfigValue_('ADMIN_EMAILS');
  if (!admins) return;
  const orders = getOrdersByDate_(dateStr);
  const count = orders.length;
  if (count > 0) {
    sendEmail_(admins, `Resumen Pedidos ${dateStr}`,
      `Se han registrado <b>${count}</b> pedidos para el día ${dateStr}.<br>Respaldo en Drive.`
    );
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
  sh.getRange("B3:M4").merge().setValue(deptName)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(14);

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
