# App Solicitud Almuerzo

Aplicacion web para gestionar pedidos de almuerzo institucional con Google Apps Script y Vue.js.

## Funcionalidades principales

- Pedidos de almuerzo por fecha con reglas de negocio para combinaciones de menu.
- Roles de `USUARIO`, `ADMIN_DEP` y `ADMIN_GEN`.
- Gestion de usuarios, departamentos, menu, dias libres y configuracion del sistema.
- Recordatorios, cierre diario, respaldos en Drive y reportes por correo.
- Avisos por correo cuando se carga menu nuevo, cuando un cambio de menu cancela pedidos afectados y cuando Calendario suspende un almuerzo.
- Reporte diario consolidado con detalle por departamento y Excel general con hojas separadas.
- Resumen del usuario con costo acumulado por almuerzos segun el precio vigente en cada fecha.
- Mini juego posterior a la confirmacion del pedido, con puntaje mensual, limite diario de juego y ranking.

## Estructura base

La solucion usa estas hojas en la spreadsheet:

- `Config`
- `Usuarios`
- `Departamentos`
- `Menu`
- `Pedidos`
- `DiasLibres`

La hoja `Usuarios` conserva sus primeras columnas historicas (`email`, `nombre`, `departamento`, `rol`, `estado`, `preferencias_json`, `codigo`) y agrega al final las columnas del mini juego: `juego_mes`, `juego_puntos_mes`, `juego_aciertos_mes`, `juego_fallos_mes`, `juego_tiempo_fecha`, `juego_segundos_hoy`, `juego_racha`, `juego_racha_max`, `juego_penalizacion_segundos` y `juego_actualizado`.

## Setup inicial

1. Abre el editor de Apps Script.
2. Ejecuta `setupSheetsAndConfig()` desde `Setup.js`.
3. Verifica que se hayan creado las hojas requeridas.
4. Revisa `Config` y completa los valores operativos necesarios.

Si ya existe la hoja `Usuarios`, `setupSheetsAndConfig()` agrega solo las columnas faltantes del mini juego al final de la hoja y conserva los datos existentes.

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

## Mini juego y ranking

- Al confirmar un pedido, el card de confirmacion entra con transicion y el usuario decide cuando iniciar o pausar el mini juego desde el control del recibo.
- Si no hay menu cargado, la app muestra un card de estado similar al recibo y tambien permite jugar Mini chef desde ahi.
- El usuario suma 1 punto al tocar al chef y pierde puntos si deja escapar al chef, toca una cebolla podrida o toca un sarten quemado.
- Si el chef se escapa, o si el usuario toca una trampa, se incrementa una penalizacion en segundos antes de la proxima aparicion. La UI muestra ese contador en una sola capsula.
- Mientras corre esa espera de penalizacion, no se puede pausar el juego, cambiar de fecha ni abrir el ranking para saltarse el contador.
- Si una cebolla podrida o un sarten quemado desaparece sin que el usuario lo toque, no hay penalizacion ni perdida de puntos.
- Las rachas otorgan combos pequenos para 3 chefs seguidos, 5 chefs seguidos y rachas perfectas cada 10 aciertos seguidos.
- Cada 8 chefs seguidos se prepara un chef bonus: el proximo chef aparece con brillo durante 3 segundos y cada toque suma 1 punto sin cortar el round.
- De forma muy poco frecuente puede salir un bizcocho disparado desde detras del card; si el usuario logra tocarlo antes de que se escape, suma 20 puntos.
- Cada usuario tiene 900 segundos (15 minutos) de juego por dia. El contador corre solo mientras el juego esta en play, se guarda en `Usuarios.juego_segundos_hoy` y se reinicia automaticamente al cambiar `juego_tiempo_fecha`.
- La economia de puntos esta calibrada para que los puntajes altos dependan de precision sostenida: los premios son pequenos, el bizcocho es raro y las penalizaciones pesan mas que un acierto normal.
- El puntaje mensual se guarda en `Usuarios.juego_puntos_mes` y se considera solo si `juego_mes` coincide con el mes actual (`YYYY-MM`). Al iniciar un mes nuevo, el ranking muestra puntaje 0 hasta que el usuario vuelva a jugar.
- Las jugadas se calculan en cliente durante la sesion para evitar latencia. El usuario puede guardar puntos manualmente y la app tambien sincroniza al agotarse el tiempo diario.
- El puntaje aparece en el header junto al nombre del usuario. Al hacer clic se abre una pantalla de ranking mensual con los usuarios activos.
- Cuando se agota el tiempo diario, la tarjeta del mini juego muestra un acceso directo al ranking mensual para que el usuario pueda revisar su posicion sin volver a buscar el boton del header.
- La pantalla de ranking mensual incluye un boton para desplegar las reglas del juego con un resumen visual de puntaje, trampas, limite diario y guardado.
- No requiere servicios avanzados ni scopes OAuth nuevos.

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
- Al cargar menu por primera vez para una fecha futura desde el modulo `Menu`, se envia un correo estilizado a los usuarios activos que mantienen las notificaciones activadas (`preferencias_json.reminders !== false`).
- Al importar menu semanal en bulk, los dias futuros que no tenian menu previo se agrupan en un solo correo por usuario con el detalle de todos los dias cargados.
- Si se edita, elimina o reemplaza una opcion/plato de una fecha futura y existen pedidos activos con esa opcion, solo esos pedidos se marcan como `CANCELADO` y sus usuarios reciben un correo personalizado para volver a pedir. Estos avisos se envian aunque el usuario tenga notificaciones desactivadas.
- Si la administracion suspende el almuerzo marcando la fecha como dia libre en `Calendario`, la fecha deja de aparecer como disponible, se cancelan los pedidos activos de ese dia y se notifica a todos los usuarios afectados aunque tengan notificaciones desactivadas.

## Reportes de cierre y modo prueba

- El cierre diario mantiene los correos por departamento con su Excel individual para los administradores activos de cada departamento (`ADMIN_DEP`).
- `ADMIN_EMAILS` define los destinatarios principales (`TO`) del resumen diario general y sigue controlando el acceso administrativo a la plataforma.
- `RESPONSIBLES_EMAILS_JSON` ahora guarda solo una lista JSON de correos externos que se agregan como copia (`CC`) al resumen diario general. Estos correos no otorgan acceso a la plataforma ni reciben solicitudes de acceso.
- Formato manual esperado para `RESPONSIBLES_EMAILS_JSON`: `["proveedor@ejemplo.com","cocina@ejemplo.com"]`.
- El resumen diario general mantiene el total de pedidos y el CTA al panel administrativo, agrega una tabla de pedidos por departamento y adjunta un Excel consolidado.
- El Excel consolidado usa la plantilla de `DAILY_REPORT_MODEL_ID`: la primera hoja es `Resumen general` con todos los pedidos continuos, y las hojas siguientes separan los pedidos por departamento.
- Los reportes diarios generados desde la plantilla incluyen una columna final `NOTA PARA LA COCINA`, tomada de la nota opcional del formulario de pedido.
- La generacion del Excel consolidado reintenta accesos transitorios a la hoja temporal antes de fallar. Si hay pedidos y no se logra generar el XLSX esperado, el resumen administrativo no se envia sin adjunto.
- Las hojas generadas escriben la tabla desde la columna `A`, no inmovilizan filas ni columnas, alinean `NOMBRE EMPLEADO` a la izquierda y calculan un alto minimo por fila para evitar truncar textos envueltos.
- En Drive se siguen guardando los PDF por departamento y ahora tambien se guarda un PDF del `Resumen general`.
- Si `TEST_EMAIL_MODE` esta en `TRUE` (sin importar mayusculas o espacios), todos los correos se redirigen solo a `TEST_EMAIL_DEST`; si falta ese destino, no se envian a destinatarios reales. El flujo de prueba no guarda respaldos, no ejecuta mantenimiento y no deja cierre real. En el panel administrativo aparece un boton para enviar esos correos de prueba desde `CONFIG`.

## Arquitectura de rendimiento

- La app embebe el bootstrap inicial (`apiGetInitData`) directamente desde `doGet()`, evitando una llamada extra `google.script.run` al abrir la app por primera vez.
- Ese bootstrap precarga en una sola respuesta todos los menus abiertos del modulo principal, para que luego el cambio entre dias ocurra sin esperas ni recargas adicionales.
- Al confirmar o cancelar, la UI actualiza el estado local del pedido y solo refresca lo estrictamente necesario.
- El mini juego acumula puntaje y tiempo en cliente durante la sesion y sincroniza en lote las columnas de juego del usuario, sin disparar un re-bootstrap global.
- La navegacion entre fechas reutiliza `allMenus` y `allOrders` ya cargados; el endpoint puntual por fecha queda como soporte y no como camino normal de navegacion.
- El bootstrap inicial y el panel administrativo usan cache corta en servidor con invalidacion por revision para reducir latencia repetida.
- El calculo de fechas abiertas y menus disponibles usa cache corta independiente para no reconstruir el bundle completo en cada request.
- La verificacion de claves operativas de `Config` se hace en lote y con cache corta, para no escanear la hoja varias veces por request.
- El panel administrativo se precarga en segundo plano para usuarios admin y asi la transicion a esa vista se siente mas rapida.
- El heartbeat registra en segundo plano a todos los usuarios activos para que el contador sea representativo; solo `ADMIN_GEN` recibe y ve el total.
- Las acciones administrativas que refrescan el panel despues de guardar mantienen el spinner hasta que termina tambien esa recarga.
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
