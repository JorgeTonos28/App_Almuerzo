# App Solicitud Almuerzo

Aplicacion web para gestionar pedidos de almuerzo institucional con Google Apps Script y Vue.js.

## Funcionalidades principales

- Pedidos de almuerzo por fecha con reglas de negocio dinámicas para combinaciones de menu y tipo de selección (única vs múltiple).
- Roles de `USUARIO`, `ADMIN_DEP` y `ADMIN_GEN`.
- Sistema de votaciones y valoraciones:
  - Valoración de comidas diarias (1 a 5 estrellas con comentario, disponible a partir de las 12:00 PM del mismo día o días posteriores).
  - Valoración del proveedor de alimentos activo (1 a 5 estrellas con comentarios, actualizable en el tiempo con historial de auditoría).
  - Panel administrativo de satisfacción con métricas KPI, distribución de estrellas y listado de opiniones.
- Mecanismo homologado de aviso general unificado:
  - Modal con portada/cover gradiente, badge, título, descripción e iconografía.
  - Soporte de múltiples slides/pantallas con navegación interactiva.
  - Control de frecuencia mediante fecha de expiración y límite máximo de cierres persistido en `Usuarios.preferencias_json`.
  - Configuración y publicación directa desde el panel de administración.
- Importación robusta de menú semanal desde Excel/TSV compatible con celdas multilínea entrecomilladas sin desalinear columnas ni generar opciones espurias.
- Gestion de usuarios, departamentos, menu, categorias de menu configurables, dias libres y configuracion del sistema.
- Recordatorios, cierre diario, respaldos en Drive y reportes por correo.
- Avisos por correo cuando se carga menu nuevo, cuando un cambio de menu cancela pedidos afectados y cuando Calendario suspende un almuerzo.
- Reporte diario consolidado con detalle por departamento y Excel general con hojas separadas.
- Resumen del usuario con costo acumulado por almuerzos segun el precio vigente en cada fecha.
- Mini juego "Mini chef" con puntaje mensual, ranking, cooldown diario y sincronizacion de progreso en la hoja `Usuarios`.

## Estructura base

La solucion usa estas hojas en la spreadsheet:

- `Config`
- `Usuarios`
- `Departamentos`
- `Menu`
- `CategoriasMenu`
- `Pedidos`
- `DiasLibres`
- `ValoracionesComida`
- `ValoracionesProveedor`
- `HistoricoValoracionesProveedor`

La hoja `CategoriasMenu` almacena las columnas: `id`, `nombre`, `orden`, `estado`, `alias_importacion`, `es_combinable`, `combinable_con`, `tipo_seleccion`.
La hoja `ValoracionesComida` almacena: `id`, `pedido_id`, `fecha_consumo`, `email_usuario`, `nombre_usuario`, `departamento`, `puntuacion`, `comentario`, `platos_resumen`, `timestamp_creacion`, `timestamp_actualizacion`.
La hoja `ValoracionesProveedor` almacena las calificaciones activas del período: `id`, `proveedor_periodo_id`, `proveedor_nombre`, `email_usuario`, `nombre_usuario`, `departamento`, `puntuacion`, `comentario`, `version_voto`, `timestamp_creacion`, `timestamp_actualizacion`.
La hoja `HistoricoValoracionesProveedor` registra el log inmutable de cada emisión o actualización de voto: `id`, `proveedor_periodo_id`, `proveedor_nombre`, `email_usuario`, `nombre_usuario`, `departamento`, `puntuacion`, `comentario`, `timestamp`.

## Setup inicial

1. Abre el editor de Apps Script.
2. Ejecuta `setupSheetsAndConfig()` desde `Setup.js`.
3. Verifica que se hayan creado y actualizado todas las hojas requeridas (`CategoriasMenu`, `ValoracionesComida`, `ValoracionesProveedor`, `HistoricoValoracionesProveedor`).
4. Revisa `Config` y completa los valores operativos necesarios.

Al ejecutar el setup en una instalacion existente se migran automáticamente los encabezados faltantes sin alterar datos existentes.

## Configuracion relevante

Claves importantes en `Config`:

- `HORA_ENVIO`
- `MINUTOS_PREV_CIERRE`
- `HORA_RECORDATORIO`
- `ADMIN_EMAILS`
- `RESPONSIBLES_EMAILS_JSON`
- `DAILY_REPORT_MODEL_ID`
- `BACKUP_FOLDER_ID`
- `TEST_EMAIL_MODE`
- `TEST_EMAIL_DEST`
- `MEAL_PRICE_CURRENT`
- `MEAL_PRICE_HISTORY_JSON`
- `MENU_DAY_ENDPOINT_TOKEN`
- `ANNOUNCEMENT_ENABLED`
- `ANNOUNCEMENT_ID`
- `ANNOUNCEMENT_EXPIRES_ON`
- `ANNOUNCEMENT_MAX_DISMISS`
- `ANNOUNCEMENT_PAYLOAD_JSON`
- `PROVIDER_NAME`
- `PROVIDER_PERIOD_ID`
- `PROVIDER_PERIOD_START`

Notas sobre valoraciones y proveedor:

- Las valoraciones de comida solo están habilitadas para días con pedidos activos. Para el día en curso, se habilitan a partir de las 12:00 PM (hora de almuerzo); para días pasados, están siempre disponibles.
- Las valoraciones de proveedor permiten a los usuarios calificar al proveedor contratado y actualizar su voto en cualquier momento.
- Cuando el Administrador General cambia de proveedor o inicia un nuevo período (`apiResetProviderPeriod`), se actualiza `PROVIDER_PERIOD_ID`, reiniciando los votos activos para el nuevo ciclo y conservando el historial previo.

Notas sobre aviso general:

- Reemplaza y unifica los antiguos banners y popovers dispersos (`PLAN_WEEK_TEXT`, `SUMMARY_COST_HINT`, etc.).
- `ANNOUNCEMENT_PAYLOAD_JSON` almacena la lista de slides (título, badge, descripción, icono, tema).
- El conteo de veces que un usuario ha descartado cada aviso se guarda en su `preferencias_json` bajo `announcements[announcement_id]`.

## Endpoint JSON de menu por fecha

La Web App expone un endpoint de solo lectura para integracion con TI:

`GET {APP_URL}?endpoint=menu-dia&fecha=YYYY-MM-DD&token={MENU_DAY_ENDPOINT_TOKEN}`

Detalles:

- `fecha` tambien puede enviarse como `date`.
- `token` debe coincidir con la clave `MENU_DAY_ENDPOINT_TOKEN` de la hoja `Config`.
- Si `MENU_DAY_ENDPOINT_TOKEN` esta vacio, el endpoint queda deshabilitado.
- La respuesta incluye solo platos habilitados (`habilitado = SI`) de la hoja `Menu`.
- El endpoint devuelve JSON con `ok`, `fecha`, `date`, `label`, `existeMenu`, `exists`, `menu`, `items`, `appVersion` y `generadoEn`.

## Reglas operativas de menu

- Las categorias del menu se administran desde `Administracion > Gestion de Menu > Configurar categorias`. Se pueden definir:
  - Nombre visible, orden y estado.
  - Tipo de selección: `UNICA` (un solo plato) o `MULTIPLE` (varios platos, como Caldos).
  - Combinabilidad: si es combinable con otras categorías y lista específica de categorías compatibles (`combinable_con`).
  - Alias de importación alternativos separados por coma.
- El parser de importación semanal procesa de forma segura texto tabulado de Excel (TSV) respetando saltos de línea internos en celdas con comillas y omitiendo comillas dobles residuales.
- Las validaciones de integridad se ejecutan tanto en frontend en tiempo real como en backend antes de persistir pedidos.

## Despliegue

1. Ejecuta `clasp push` para subir los archivos al proyecto de Apps Script.
2. Al modificar backend o archivos servidos por `doGet()` (`Code.js`, `index.html`, `js.html`, `css.html`), publica un nuevo deployment de la web app.

## Versionado

- Versión actual: `v7.34`.
- Todo cambio funcional debe incrementar `APP_VERSION` en `Code.js`.
- Sigue siempre las reglas operativas de `AGENTS.md`.
