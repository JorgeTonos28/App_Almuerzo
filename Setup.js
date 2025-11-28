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
      headers: ['email', 'nombre', 'departamento', 'rol', 'estado', 'preferencias_json'] 
      // Roles: USUARIO, ADMIN_DEP, ADMIN_GEN
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
        'timestamp_modificacion'
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
  populateSampleData_(ss); // Datos de prueba para que arranques rápido
  
  SpreadsheetApp.flush();
  Logger.log('Estructura de base de datos actualizada correctamente.');
  return 'OK';
}

function populateDefaultConfig_(sheet){
  if (!sheet || sheet.getLastRow() > 1) return;
  const defaults = [
    ['HORA_CIERRE', '14:00', 'Hora militar límite para pedidos del día siguiente'],
    ['HORA_RECORDATORIO', '13:00', 'Hora envío correos recordatorios'],
    ['ADMIN_EMAILS', 'tu_correo@ejemplo.com', 'Correos admin general separados por ;'],
    ['MAIL_SENDER_NAME', 'Comedor Institucional', 'Nombre remitente correos'],
    ['APP_TITLE', 'Solicitud de Almuerzo', 'Título en la barra de navegación'],
    ['FOOTER_SIGNATURE_ID', '1SZlRhijFMv0V0jDlqtagChmGDEzGTv3R', 'ID de la imagen de firma en Drive'],
    ['BACKUP_FOLDER_ID', '', 'ID de carpeta Drive raíz para respaldos (Año/Mes/Semana)'],
    ['TEST_EMAIL_MODE', 'FALSE', 'Si es TRUE, todos los correos van a la dirección de prueba'],
    ['TEST_EMAIL_DEST', '', 'Correo de destino para modo de prueba'],
    ['RESPONSIBLES_EMAILS_JSON', '{}', 'JSON mapeo Depto->Emails. Ej: {"Finanzas": "jefe@fin.com"}']
  ];
  sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
}

function populateSampleData_(ss){
  // Usuarios
  const uSh = ss.getSheetByName('Usuarios');
  if (uSh.getLastRow() === 1) {
    uSh.appendRow([Session.getActiveUser().getEmail(), 'Admin Inicial', 'Tecnología', 'ADMIN_GEN', 'ACTIVO', '{}']);
    uSh.appendRow(['usuario.demo@ejemplo.com', 'Pepe Usuario', 'Finanzas', 'USUARIO', 'ACTIVO', '{}']);
    uSh.appendRow(['jefe.demo@ejemplo.com', 'Jefa Departamento', 'Finanzas', 'ADMIN_DEP', 'ACTIVO', '{}']);
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