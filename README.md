# App Solicitud Almuerzo

Aplicación web para gestión de pedidos de almuerzo corporativo, construida con Google Apps Script y Vue.js.

## Características

*   **Pedidos:** Selección de menú diario con validaciones (reglas de negocio para granos, arroces, etc.).
*   **Roles:**
    *   **Usuario:** Realiza pedidos propios.
    *   **Admin Departamento:** Gestiona usuarios de su área y puede realizar pedidos por ellos (Proxy).
    *   **Admin General:** Gestión global de usuarios, departamentos, pedidos y configuración del sistema.
*   **Automatización:** Recordatorios automáticos, cierre diario, respaldos en Drive y reportes por correo.
*   **UI Moderna:** Diseño estilo "Uber Eats" con Tailwind CSS.

## Requisitos Previos

1.  Cuenta de Google Workspace.
2.  Hoja de Cálculo de Google (Base de Datos).
3.  Carpeta en Google Drive (para respaldos).

## Configuración Inicial (Setup)

El proyecto incluye un script de instalación automática:

1.  Abre el editor de Apps Script.
2.  Selecciona la función `setupSheetsAndConfig` en `Setup.js`.
3.  Ejecuta la función. Esto creará las hojas necesarias: `Config`, `Usuarios`, `Departamentos`, `Menu`, `Pedidos`, `DiasLibres`.
4.  **Importante:** Verifica la hoja `Config` y llena los valores clave (`HORA_CIERRE`, `ADMIN_EMAILS`, `BACKUP_FOLDER_ID`, etc.).

## Estructura de Datos

*   **Usuarios:** `email`, `nombre`, `departamento`, `rol` (USUARIO, ADMIN_DEP, ADMIN_GEN), `estado`.
*   **Departamentos:** `id`, `nombre`, `admins` (emails para reportes), `estado`.
*   **Pedidos:** Registro de solicitudes con trazabilidad (`creado_por`).

## Despliegue

1.  Clic en "Implementar" > "Nueva implementación".
2.  Tipo: "Aplicación web".
3.  Ejecutar como: "Yo" (propietario).
4.  Quién tiene acceso: "Cualquier usuario de..." (tu dominio).

## Control de Versiones

Este proyecto sigue las guías de `AGENTS.md`. Incrementar `APP_VERSION` en `Code.js` ante cambios lógicos.
