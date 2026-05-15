# App Solicitud Almuerzo

Aplicacion web para gestionar pedidos de almuerzo institucional con Google Apps Script y Vue.js.

## Funcionalidades principales

- Pedidos de almuerzo por fecha con reglas de negocio para combinaciones de menu.
- Roles de `USUARIO`, `ADMIN_DEP` y `ADMIN_GEN`.
- Gestion de usuarios, departamentos, menu, dias libres y configuracion del sistema.
- Recordatorios, cierre diario, respaldos en Drive y reportes por correo.
- Reporte diario consolidado con detalle por departamento y Excel general con hojas separadas.
- Resumen del usuario con costo acumulado por almuerzos segun el precio vigente en cada fecha.

## Estructura base

La solucion usa estas hojas en la spreadsheet:

- `Config`
- `Usuarios`
- `Departamentos`
- `Menu`
- `Pedidos`
- `DiasLibres`

## Setup inicial

1. Abre el editor de Apps Script.
2. Ejecuta `setupSheetsAndConfig()` desde `Setup.js`.
3. Verifica que se hayan creado las hojas requeridas.
4. Revisa `Config` y completa los valores operativos necesarios.

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
- `PLAN_WEEK_TEXT`
- `PLAN_WEEK_LIMIT`
- `MEAL_PRICE_CURRENT`
- `MEAL_PRICE_HISTORY_JSON`
- `MENU_DAY_ENDPOINT_TOKEN`
- `SUMMARY_COST_HINT_LIMIT`
- `SUMMARY_COST_HINT_EXPIRES_ON`
- `CALDO_MULTI_HINT_LIMIT`
- `CALDO_MULTI_HINT_EXPIRES_ON`

Notas sobre costo por comida:

- `MEAL_PRICE_CURRENT` es el costo actual por almuerzo.
- `MEAL_PRICE_HISTORY_JSON` guarda el historial por fecha para calcular resumenes pasados correctamente.
- El historial se actualiza automaticamente al cambiar `MEAL_PRICE_CURRENT`. No debe editarse manualmente.
- El filtro del resumen usa el historial completo de pedidos del usuario, no solo las fechas abiertas para pedir.

Notas sobre hints:

- `SUMMARY_COST_HINT_LIMIT` y `CALDO_MULTI_HINT_LIMIT` controlan la cantidad maxima de cierres por usuario antes de ocultar cada hint.
- `SUMMARY_COST_HINT_EXPIRES_ON` y `CALDO_MULTI_HINT_EXPIRES_ON` limitan la vigencia del hint por fecha. El valor default se crea a 30 dias.
- Los contadores de dismiss viven en `Usuarios.preferencias_json`, no en la hoja `Config`.
- Los hints visibles se renderizan dentro de su propia seccion para que desaparezcan naturalmente al hacer scroll o cambiar de modulo.
- El hint de costo acumulado queda anclado al card de costo dentro del resumen semanal/diario.

## Endpoint JSON de menu por fecha

La Web App expone un endpoint de solo lectura para integracion con TI:

`GET {APP_URL}?endpoint=menu-dia&fecha=YYYY-MM-DD&token={MENU_DAY_ENDPOINT_TOKEN}`

Detalles:

- `fecha` tambien puede enviarse como `date`.
- `token` debe coincidir con la clave `MENU_DAY_ENDPOINT_TOKEN` de la hoja `Config`.
- Si `MENU_DAY_ENDPOINT_TOKEN` esta vacio, el endpoint queda deshabilitado.
- La respuesta incluye solo platos habilitados (`habilitado = SI`) de la hoja `Menu`.
- El endpoint devuelve JSON con `ok`, `fecha`, `date`, `label`, `existeMenu`, `exists`, `menu`, `items`, `appVersion` y `generadoEn`.
- Para consumo server-to-server desde ASP.NET, TI debe llamar este URL desde el backend y no desde el navegador, para no exponer el token.
- Si el deployment se mantiene con acceso `DOMAIN`, el consumidor debe poder autenticarse como usuario del dominio. Si ASP.NET no puede autenticarse contra Google, publica un deployment compatible con llamadas anonimas y protege el acceso con el token.

## Reglas operativas de menu

- `Caldo` permite seleccionar mas de una opcion dentro de la misma categoria.
- Los textos del menu se normalizan al guardar y al renderizar para evitar ALL CAPS.
- Las validaciones criticas siguen ejecutandose en backend antes de guardar pedidos.

## Reportes de cierre y modo prueba

- El cierre diario mantiene los correos por departamento con su Excel individual.
- Para esos correos, `RESPONSIBLES_EMAILS_JSON` aporta los destinatarios principales (`TO`).
- En los correos por departamento, la copia queda limitada a los administradores activos de ese departamento (`ADMIN_DEP`). Los administradores generales reciben el resumen consolidado.
- El resumen diario para `ADMIN_EMAILS` mantiene el total de pedidos y el CTA al panel administrativo, agrega una tabla de pedidos por departamento y adjunta un Excel consolidado.
- El Excel consolidado usa la plantilla de `DAILY_REPORT_MODEL_ID`: la primera hoja es `Resumen general` con todos los pedidos continuos, y las hojas siguientes separan los pedidos por departamento.
- En Drive se siguen guardando los PDF por departamento y ahora tambien se guarda un PDF del `Resumen general`.
- Si `TEST_EMAIL_MODE` esta en `TRUE`, los correos se redirigen a `TEST_EMAIL_DEST` y el flujo de prueba no guarda respaldos, no ejecuta mantenimiento y no deja cierre real. En el panel administrativo aparece un boton para enviar esos correos de prueba desde `CONFIG`.

## Arquitectura de rendimiento

- La app embebe el bootstrap inicial (`apiGetInitData`) directamente desde `doGet()`, evitando una llamada extra `google.script.run` al abrir la app por primera vez.
- Ese bootstrap precarga en una sola respuesta todos los menus abiertos del modulo principal, para que luego el cambio entre dias ocurra sin esperas ni recargas adicionales.
- Al confirmar o cancelar, la UI actualiza el estado local del pedido y solo refresca lo estrictamente necesario.
- La navegacion entre fechas reutiliza `allMenus` y `allOrders` ya cargados; el endpoint puntual por fecha queda como soporte y no como camino normal de navegacion.
- El bootstrap inicial y el panel administrativo usan cache corta en servidor con invalidacion por revision para reducir latencia repetida.
- El calculo de fechas abiertas y menus disponibles usa cache corta independiente para no reconstruir el bundle completo en cada request.
- La verificacion de claves operativas de `Config` se hace en lote y con cache corta, para no escanear la hoja varias veces por request.
- El panel administrativo se precarga en segundo plano para usuarios admin y asi la transicion a esa vista se siente mas rapida.
- El heartbeat de usuarios activos se reserva para `ADMIN_GEN`; no corre como roundtrip inicial para usuarios normales.
- El resumen semanal/diario entra en scroll vertical cuando supera 8 cards para evitar crecimiento excesivo de la pagina.
- El guardado de pedidos evita escanear toda la hoja `Pedidos` antes de escribir: usa un ID deterministico por usuario/fecha y lookup puntual sobre la columna de IDs.
- El detalle persistido de cada pedido guarda solo `categorias`, `items` y `comentarios`, reduciendo el peso del write y del parse posterior en bootstrap.
- La cancelacion y modificacion de pedidos reutilizan la misma fila del pedido cuando existe y marcan `CANCELADO` en sitio, evitando `deleteRow()` y reescrituras costosas de la hoja.
- La imagen decorativa del footer usa cache y prioriza un `data:` URL generado desde Drive, con fallback a `thumbnailLink` para no perder compatibilidad visual.
- El cierre diario reutiliza una sola lectura de pedidos, usuarios y departamentos para generar reportes por departamento, resumen consolidado y correo ejecutivo.

## Despliegue

1. Ejecuta `clasp push` para subir los archivos al proyecto de Apps Script.
2. Si cambias backend o archivos servidos por `doGet()`, publica un nuevo deployment de la web app.
3. Si cambias `appsscript.json`, vuelve a subirlo antes del deployment.

## Versionado

- Todo cambio funcional debe incrementar `APP_VERSION` en `Code.js`.
- Sigue tambien las reglas documentadas en `AGENTS.md`.
