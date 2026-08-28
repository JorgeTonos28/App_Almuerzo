const USER_GAME_HEADERS = [
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

/**
 * Setup.gs
 * Configura la base de datos (hojas y encabezados) y la configuración inicial.
 */
function setupSheetsAndConfig(){
  const must = [
    {
      name: 'Config', 
      headers: ['key', 'value', 'description']
    },
    {
      name: 'Usuarios', 
      headers: ['email', 'nombre', 'departamento', 'rol', 'estado', 'preferencias_json', 'codigo'].concat(USER_GAME_HEADERS)
      // Roles: USUARIO, ADMIN_DEP, ADMIN_GEN
    },
    {
      name: 'Departamentos',
      headers: ['id', 'nombre', 'admins', 'estado', 'preferencias_json']
      // admins: correos separados por coma (para notificaciones/reportes)
    },
    {
      name: 'Menu', 
      headers: ['id', 'fecha', 'categoria', 'plato', 'descripcion', 'habilitado']
    },
    {
      name: 'CategoriasMenu',
      headers: ['id', 'nombre', 'orden', 'estado', 'alias_importacion', 'es_combinable', 'combinable_con', 'tipo_seleccion']
    },
    {
      name: 'Pedidos', 
      headers: [
        'id', 'fecha_solicitud', 'fecha_consumo', 'email_usuario', 'nombre_usuario', 'departamento',
        'seleccion_resumen', // Texto legible ej: "Arroz B., Pollo, Ensalada"
        'json_detalle',      // Objeto JSON completo para re-edición
        'estado',            // ACTIVO, CANCELADO
        'timestamp_modificacion',
        'creado_por'         // Email de quien realizó la acción (trazabilidad proxy)
      ]
    },
    {
      name: 'ValoracionesComida',
      headers: [
        'id', 'pedido_id', 'fecha_consumo', 'email_usuario', 'nombre_usuario', 'departamento',
        'puntuacion', 'comentario', 'platos_resumen', 'timestamp_creacion', 'timestamp_actualizacion'
      ]
    },
    {
      name: 'ValoracionesProveedor',
      headers: [
        'id', 'proveedor_periodo_id', 'proveedor_nombre', 'email_usuario', 'nombre_usuario', 'departamento',
        'puntuacion', 'comentario', 'version_voto', 'timestamp_creacion', 'timestamp_actualizacion'
      ]
    },
    {
      name: 'HistoricoValoracionesProveedor',
      headers: [
        'id', 'proveedor_periodo_id', 'proveedor_nombre', 'email_usuario', 'nombre_usuario', 'departamento',
        'puntuacion', 'comentario', 'timestamp'
      ]
    },
    {
      name: 'DiasLibres',
      headers: ['fecha', 'motivo'] // Días libres institucionales (adicionales a feriados oficiales)
    },
  ];

  const ss = SpreadsheetApp.getActive();

  must.forEach(s => {
    let sh = ss.getSheetByName(s.name);
    if (!sh) {
      sh = ss.insertSheet(s.name);
      sh.getRange(1, 1, 1, s.headers.length).setValues([s.headers]);
      sh.getRange(1, 1, 1, s.headers.length).setFontWeight('bold').setBackground('#f3f4f6');
      sh.setFrozenRows(1);
    } else {
      // Validacion y migracion segura de encabezados existentes.
      const schema = ensureSheetHeaders_(sh, s.headers);
      if (!schema.prefixOk) {
        Logger.log('Aviso: Los encabezados de ' + s.name + ' pueden diferir.');
      }
    }
  });

  populateDefaultConfig_(ss.getSheetByName('Config'));
  ensureDefaultMenuCategories_(ss.getSheetByName('CategoriasMenu'));
  ensureMenuDayEndpointToken_(ss.getSheetByName('Config'));
  ensureBackupFolder_(ss.getSheetByName('Config'));
  populateSampleData_(ss); // Datos de prueba para que arranques rápido
  
  SpreadsheetApp.flush();
  Logger.log('Estructura de base de datos actualizada correctamente.');
  return 'OK';
}

function ensureSheetHeaders_(sheet, expectedHeaders) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const existing = {};
  current.forEach(header => {
    if (header) existing[header] = true;
  });

  const missing = expectedHeaders.filter(header => !existing[header]);
  if (missing.length > 0) {
    sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, lastColumn + 1, 1, missing.length).setFontWeight('bold').setBackground('#f3f4f6');
  }

  const compareLength = Math.min(current.length, expectedHeaders.length);
  const prefixOk = current.slice(0, compareLength).join() === expectedHeaders.slice(0, compareLength).join();
  return { prefixOk: prefixOk, missingCount: missing.length };
}

function ensureMenuDayEndpointToken_(configSheet) {
  if (!configSheet) return;

  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'MENU_DAY_ENDPOINT_TOKEN') {
      if (!String(data[i][1] || '').trim()) {
        configSheet.getRange(i + 1, 2).setValue(generateSecretToken_());
      }
      return;
    }
  }

  configSheet.appendRow([
    'MENU_DAY_ENDPOINT_TOKEN',
    generateSecretToken_(),
    'Token secreto para consumir el endpoint JSON de menu por fecha. Generar y compartir solo con TI.'
  ]);
}

function ensureBackupFolder_(configSheet) {
  if (!configSheet) return;
  const data = configSheet.getDataRange().getValues();
  let row = -1;
  let currentId = '';

  for(let i=1; i<data.length; i++) {
     if(data[i][0] === 'BACKUP_FOLDER_ID') {
        row = i+1;
        currentId = data[i][1];
        break;
     }
  }

  if (row > 0 && !currentId) {
     try {
        const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
        const parents = ssFile.getParents();
        if (parents.hasNext()) {
           const parent = parents.next();
           const folders = parent.getFoldersByName('Backups_Almuerzo');
           let folder;
           if (folders.hasNext()) folder = folders.next();
           else folder = parent.createFolder('Backups_Almuerzo');

           configSheet.getRange(row, 2).setValue(folder.getId());
           Logger.log('Carpeta backup creada/asignada: ' + folder.getId());
        }
     } catch(e) {
        Logger.log('Error creando carpeta backup: ' + e.message);
     }
  }
}

function populateDefaultConfig_(sheet){
  if (!sheet || sheet.getLastRow() > 1) return;
  const defaultExpiry = Utilities.formatDate(new Date(Date.now() + (30 * 24 * 60 * 60 * 1000)), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

  const defaults = [
    ['HORA_ENVIO', '15:00', 'Hora militar envío reportes a responsables (HH:MM)'],
    ['MINUTOS_PREV_CIERRE', '30', 'Minutos antes del envío para cerrar pedidos'],
    ['HORA_RECORDATORIO', '13:00', 'Hora envío correos recordatorios'],
    ['ADMIN_EMAILS', 'tu_correo@ejemplo.com', 'Correos admin general separados por ;'],
    ['MAIL_SENDER_NAME', 'Comedor Institucional', 'Nombre remitente correos'],
    ['APP_TITLE', 'Solicitud de Almuerzo', 'Título en la barra de navegación'],
    ['FOOTER_SIGNATURE_ID', '1SZlRhijFMv0V0jDlqtagChmGDEzGTv3R', 'ID de la imagen de firma en Drive'],
    ['BACKUP_FOLDER_ID', '', 'ID de carpeta Drive raíz para respaldos (Año/Mes/Semana)'],
    ['TEST_EMAIL_MODE', 'FALSE', 'Si es TRUE, todos los correos van a la dirección de prueba'],
    ['TEST_EMAIL_DEST', '', 'Correo de destino para modo de prueba'],
    ['RESPONSIBLES_EMAILS_JSON', '[]', 'JSON de correos externos en copia para el resumen diario general.'],
    ['ANNOUNCEMENT_ENABLED', 'TRUE', 'Indica si el aviso general está activo para los usuarios (TRUE/FALSE)'],
    ['ANNOUNCEMENT_ID', 'anuncio_v7_31_valoraciones', 'Identificador único del aviso activo. Al cambiarlo, todos los usuarios volverán a verlo.'],
    ['ANNOUNCEMENT_EXPIRES_ON', defaultExpiry, 'Fecha límite para mostrar el aviso general (YYYY-MM-DD)'],
    ['ANNOUNCEMENT_MAX_DISMISS', '3', 'Cantidad máxima de veces que el usuario puede cerrar el aviso antes de que no aparezca más.'],
    ['ANNOUNCEMENT_PAYLOAD_JSON', defaultAnnouncementPayload, 'Contenido en formato JSON de los slides del aviso general.'],
    ['PROVIDER_NAME', 'Proveedor de Alimentos', 'Nombre del proveedor de alimentos activo.'],
    ['PROVIDER_PERIOD_ID', 'PROV_2026_01', 'Identificador del ciclo/período de evaluación del proveedor actual.'],
    ['PROVIDER_PERIOD_START', todayStr, 'Fecha de inicio del ciclo del proveedor actual (YYYY-MM-DD).'],
    ['MEAL_PRICE_CURRENT', '57', 'Costo actual por almuerzo. Al cambiarlo se conserva historial automatico por fecha.'],
    ['MEAL_PRICE_HISTORY_JSON', '[{"from":"1900-01-01","price":57}]', 'Historial auto-administrado del costo por almuerzo. No editar manualmente.'],
    ['MENU_DAY_ENDPOINT_TOKEN', generateSecretToken_(), 'Token secreto para consumir el endpoint JSON de menu por fecha. Generar y compartir solo con TI.'],
    ['DAILY_REPORT_MODEL_ID', '', 'ID del archivo modelo Excel para reportes diarios'],
    ['LOGO_ID', '', 'ID del archivo de imagen del Logo en Drive'],
    ['APP_URL', '', 'URL pública de la aplicación (Web App) para enlaces en correos']
  ];
  sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
}

function populateSampleData_(ss){
  // Departamentos
  const dSh = ss.getSheetByName('Departamentos');
  // Usamos UUIDs fijos o generados para consistencia en la demo,
  // pero aquí generamos dinámicos para que sea un ejemplo válido.
  const deptTechId = Utilities.getUuid();
  const deptFinId = Utilities.getUuid();

  if (dSh.getLastRow() === 1) {
     dSh.appendRow([deptTechId, 'Tecnología', Session.getActiveUser().getEmail(), 'ACTIVO', '{}']);
     dSh.appendRow([deptFinId, 'Finanzas', 'jefe.demo@ejemplo.com', 'ACTIVO', '{}']);
  }

  // Usuarios
  const uSh = ss.getSheetByName('Usuarios');
  if (uSh.getLastRow() === 1) {
    uSh.appendRow([Session.getActiveUser().getEmail(), 'Admin Inicial', deptTechId, 'ADMIN_GEN', 'ACTIVO', '{}']);
    uSh.appendRow(['usuario.demo@ejemplo.com', 'Pepe Usuario', deptFinId, 'USUARIO', 'ACTIVO', '{}']);
    uSh.appendRow(['jefe.demo@ejemplo.com', 'Jefa Departamento', deptFinId, 'ADMIN_DEP', 'ACTIVO', '{}']);
  }

  // Menú de ejemplo (para mañana)
  const mSh = ss.getSheetByName('Menu');
  if (mSh.getLastRow() === 1) {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1); // Mañana
    // Ajuste simple si es fin de semana saltar al lunes (lógica básica setup)
    if (tomorrow.getDay() === 6) tomorrow.setDate(tomorrow.getDate() + 2);
    if (tomorrow.getDay() === 0) tomorrow.setDate(tomorrow.getDate() + 1);
    
    const ymd = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    const items = [
      ['M-001', ymd, 'Arroces', 'Arroz Blanco', '', 'SI'],
      ['M-002', ymd, 'Arroces', 'Moro de Guandules', '', 'SI'],
      ['M-003', ymd, 'Granos', 'Habichuelas Rojas', 'Guisadas', 'SI'],
      ['M-004', ymd, 'Carnes', 'Pollo al Horno', '', 'SI'],
      ['M-005', ymd, 'Carnes', 'Res Guisada', '', 'SI'],
      ['M-006', ymd, 'Ensaladas', 'Ensalada Verde', '', 'SI'],
      ['M-007', ymd, 'Ensaladas', 'Ensalada Rusa', '', 'SI'],
      ['M-008', ymd, 'Viveres', 'Yuca Encebollada', '', 'SI'],
      ['M-009', ymd, 'Vegetariana', 'Berenjenas a la Parmesana', 'Incluye guarnición', 'SI'],
      ['M-010', ymd, 'Opcion_Rapida', 'Sandwich de Jamón y Queso', '', 'SI']
    ];
    mSh.getRange(2, 1, items.length, 6).setValues(items);
  }
}

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
