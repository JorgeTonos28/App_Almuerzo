# Instrucciones para Agentes de Desarrollo

Este archivo contiene reglas y directrices para mantener la consistencia y calidad del código en este proyecto. Cualquier desarrollador (humano o IA) debe seguir estas instrucciones.

## Control de Versiones

### Actualización de Versión
*   **Regla**: Siempre incrementar la constante `APP_VERSION` ubicada al inicio del archivo `Código.js` cada vez que se realice una modificación en el código que afecte la funcionalidad o lógica.
*   **Propósito**: Mantener un rastreo preciso de la versión desplegada y facilitar la depuración.

### Mensajes de Commit
*   **Regla**: Al realizar un commit que incluya cambios en `Código.js`, se debe mencionar la nueva versión (`vX.Y`) en el cuerpo o título del mensaje del commit.
*   **Ejemplo**: `fix: corregir validación de fecha (v3.2)`
