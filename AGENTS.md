# Reglas Operativas para Agentes

Estas reglas aplican a cualquier agente o desarrollador que modifique este proyecto.

## 1. Versionado primero

- Antes de cualquier cambio funcional o de logica, incrementa `APP_VERSION` al inicio de `Code.js`.
- El resumen final debe mencionar explicitamente la nueva version.
- Si el cambio impacta la web app, indica en el cierre que debe publicarse un nuevo deployment.
- Si se hace un commit que incluya cambios funcionales en `Code.js`, menciona la nueva version (`vX.Y`) en el mensaje.

## 2. README siempre al dia

- Si cambias comportamiento, setup, estructura de hojas, `Config`, manifest, scopes, templates, vistas o despliegue, actualiza `README.md` en la misma interaccion.
- No dejes deuda documental para despues.

## 3. Manifest y servicios

- Si agregas o quitas un servicio avanzado, actualiza `appsscript.json`.
- Si agregas o quitas scopes OAuth, actualiza `appsscript.json` y `README.md`.
- No uses servicios nuevos sin dejar el manifest listo para `clasp push`.
- Si el cambio depende de Drive, Sheets u otro servicio de Google, piensa tambien en permisos, errores de autorizacion y degradacion segura.

## 4. Config, hojas y setup

- Si agregas, renombras o eliminas claves de `Config`, actualiza `Setup.js` y deja una ruta segura para entornos existentes usando `ensureConfigKey_()` o un mecanismo equivalente.
- Si cambias headers o estructura de hojas, actualiza `setupSheetsAndConfig()` y preserva la compatibilidad con datos existentes.
- No rompas nombres de hojas, columnas o claves que ya use la app sin dejar migracion explicita.
- Si un cambio estructural requiere una accion manual posterior, indicalo claramente en el resumen final.

## 5. Responsive no negociable

- Toda modificacion de UI debe mantenerse usable en escritorio y movil.
- Breakpoints operativos:
- `<= 1024px`: sidebar como drawer, layouts apilados y tablas con overflow.
- `<= 720px`: compactacion adicional de acciones, formularios y tablas.
- En cambios de layout, valida scroll vertical real, `overflow` de tablas, espacios internos y que no queden botones o textos pegados a los bordes.

## 6. Spinner en todo proceso asincrono

- Toda llamada visible a `google.script.run` debe pasar por el mecanismo central de busy/loading correspondiente.
- Si un flujo encadena varias llamadas, el spinner debe cubrir todo el ciclo.
- No introduzcas procesos asincronos nuevos que no den feedback visual claro al usuario.

## 7. Hints, banners y estados descartables

- Todo banner, hint o resaltado descartable debe persistirse en `preferencias_json` del usuario.
- Si agregas un elemento descartable, define un limite de dismiss y evita que reaparezca indefinidamente.
- No sobrescribas otras preferencias existentes al guardar contadores o flags de UI.

## 8. Integridad de datos

- Manten validacion en cliente y servidor para todo dato critico.
- Si cambias headers o estructura de hojas, actualiza `setupSheetsAndConfig()` y preserva la migracion de datos existentes.
- Si agregas valores numericos configurables, valida rango, formato y fallback seguro.
- Si una regla de negocio existe en frontend, replica la validacion critica en backend.

## 9. Menus y texto visible

- Evita guardar nombres de platos o textos de menu en ALL CAPS.
- Si una fuente externa trae texto inconsistente, normalizalo al guardar y vuelve a normalizar al renderizar en vistas criticas.
- Al tocar reglas de seleccion de menu, revisa tambien resumenes, edicion, importacion y validaciones del backend.

## 10. Optimizacion y rendimiento

- Trata el rendimiento como un requisito de arquitectura, no como un ajuste posterior.
- Antes de implementar un cambio, evalua si introduce roundtrips innecesarios entre cliente y servidor.
- No uses un re-bootstrap global para reflejar cambios locales si el modulo puede refrescarse de forma incremental.
- Prioriza cache cliente y cache corta en servidor para vistas frecuentes como dashboard, detalle, edicion, impresion, catalogos y panel administrativo.
- Reutiliza datos ya cargados antes de volver a consultar Google Sheets o Drive.
- Evita lecturas y escrituras repetidas dentro de loops. Prefiere operaciones masivas en memoria.
- No agregues llamadas a `DriveApp`, `Drive API`, `SpreadsheetApp` o `getRange()` dentro de ciclos si pueden resolverse con una sola lectura o escritura.
- Toda validacion visual puede ocurrir en cliente, pero toda validacion de integridad debe mantenerse en servidor.
- El spinner global solo debe usarse en operaciones bloqueantes. Las tareas secundarias deben ir en segundo plano sin congelar la UI.
- Si una accion puede responder rapido y completar pasos secundarios despues, implementa esa estrategia.
- En flujos de impresion o vistas ya cargadas, prioriza cache local sobre una consulta nueva.
- Si un cambio afecta rendimiento, revisa tambien si debe actualizarse `README.md`.

## 11. Calidad minima antes de cerrar

- Revisa sintaxis de los archivos modificados si es posible.
- Piensa explicitamente en escenarios de error:
- usuario no autorizado,
- catalogos vacios,
- fecha invalida,
- configuracion faltante,
- permisos insuficientes,
- fallo en guardado, edicion o cancelacion,
- hoja inexistente o con estructura incompleta,
- archivo Drive invalido o inaccesible.
- Si no pudiste probar algo, dilo claramente en el resumen final.

## 12. Cierre y despliegue

- Si cambias la logica del backend o cualquier archivo servido por `doGet()` (`index.html`, `js.html`, `css.html`, `Denied.html`, `Code.js`), indica que hace falta un nuevo deployment de la web app.
- Si el cambio toca manifest, deja claro si basta con `clasp push` o si hace falta redeploy.
- Si el cambio toca setup, seeds, estructura de hojas o claves nuevas de `Config`, indica si debe ejecutarse `setupSheetsAndConfig()`.
