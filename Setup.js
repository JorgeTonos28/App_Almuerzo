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
      headers: ['email', 'nombre', 'departamento', 'rol', 'estado', 'preferencias_json', 'codigo']
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
      // Categorías: Arroces, Granos, Carnes, Viveres, Ensaladas, Vegetariana, Caldo, Opcion_Rapida
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
      // Validación básica de encabezados existentes
      const current = sh.getRange(1, 1, 1, s.headers.length).getValues()[0];
      if (current.join() !== s.headers.join()) {
        Logger.log('Aviso: Los encabezados de ' + s.name + ' pueden diferir.');
      }
    }
  });

  populateDefaultConfig_(ss.getSheetByName('Config'));
  ensureBackupFolder_(ss.getSheetByName('Config'));
  populateSampleData_(ss); // Datos de prueba para que arranques rápido
  
  SpreadsheetApp.flush();
  Logger.log('Estructura de base de datos actualizada correctamente.');
  return 'OK';
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
    ['RESPONSIBLES_EMAILS_JSON', '{}', 'JSON mapeo DeptoID->Emails.'],
    ['PLAN_WEEK_TEXT', '¡Planifica tu semana! Ahora puedes adelantar tus pedidos para todos los días disponibles.', 'Texto del banner de planificación'],
    ['PLAN_WEEK_LIMIT', '5', 'Número de veces que se mostrará el banner al usuario']
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