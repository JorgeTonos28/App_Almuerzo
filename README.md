# Solicitud Almuerzo

Aplicación web desarrollada en Google Apps Script para la gestión eficiente de pedidos de almuerzo en un entorno corporativo. Permite a los usuarios seleccionar su menú diario validando reglas de negocio, y ofrece a los administradores un resumen de los pedidos por departamento.

## Características Clave

*   **Autenticación Institucional**: Acceso restringido a usuarios del dominio con estado `ACTIVO` en la base de datos.
*   **Diseño Responsivo**: Interfaz adaptada para dispositivos móviles y escritorio.
*   **Menú Dinámico**: Carga de opciones de menú (platos, descripciones) directamente desde Google Sheets.
*   **Reglas de Negocio Inteligentes**: Validación automática de combinaciones de platos (ej. Granos requieren Arroz Blanco, exclusividad de platos especiales como "Opción Rápida").
*   **Cierre Automático**: Restricción de pedidos basada en hora límite (`HORA_CIERRE`) del día hábil anterior.
*   **Panel de Administración**: Visualización de estadísticas y resumen de pedidos por departamento para usuarios con rol `ADMIN_GEN` o `ADMIN_DEP`.
*   **Integración Multimedia**: Carga dinámica de firma institucional desde Google Drive.

## Estructura del Repositorio

| Archivo | Descripción |
| :--- | :--- |
| `Código.js` | Lógica principal del servidor (Backend), manejo de API, validaciones y rutas. |
| `Setup.js` | Scripts de instalación inicial para crear la estructura de la base de datos y datos de prueba. |
| `index.html` | Plantilla HTML principal de la aplicación. |
| `css.html` | Estilos CSS de la interfaz. |
| `js.html` | Lógica del lado del cliente (Frontend). |
| `appsscript.json` | Manifiesto del proyecto con configuración de zonas horarias y dependencias. |

## Requisitos Previos

1.  **Cuenta Google Workspace**: Permisos para crear scripts y hojas de cálculo.
2.  **Google Sheet**: Una hoja de cálculo que servirá como base de datos.
3.  **Servicios Avanzados**:
    *   **Google Drive API**: Debe estar habilitada en el proyecto de Apps Script para la carga de imágenes (firmas).

## Configuración Inicial (Setup)

Sigue estos pasos para desplegar la aplicación desde cero:

1.  **Crear Proyecto**: Crea una nueva Google Sheet y vincula un proyecto de Apps Script (Extensiones > Apps Script).
2.  **Copiar Código**: Copia los archivos de este repositorio (`Código.js`, `Setup.js`, `index.html`, etc.) a tu proyecto.
3.  **Instalar Base de Datos**:
    *   Abre el archivo `Setup.js`.
    *   Ejecuta la función `setupSheetsAndConfig()`. Esto creará automáticamente las pestañas necesarias (`Config`, `Usuarios`, `Menu`, `Pedidos`, `Feriados`) y cargará datos de ejemplo.
4.  **Habilitar API de Drive**:
    *   En el editor de Apps Script, ve a "Servicios" (+).
    *   Selecciona "Drive API" y añádela.
5.  **Ajustar Configuración**:
    *   Ve a la pestaña `Config` en tu Google Sheet.
    *   Actualiza `ADMIN_EMAILS`, `HORA_CIERRE` y otros parámetros según tu necesidad.
    *   (Opcional) Si deseas usar una firma en el pie de página, sube la imagen a Drive, obtén su ID y colócalo en `FOOTER_SIGNATURE_ID`.
6.  **Gestionar Usuarios**:
    *   En la pestaña `Usuarios`, asegúrate de que tu correo esté registrado con rol `ADMIN_GEN` y estado `ACTIVO`.

## Uso de la Aplicación

### Usuarios
1.  Ingresa a la URL de la aplicación web.
2.  Selecciona la fecha para la cual deseas ordenar (se muestran solo fechas habilitadas).
3.  Elige los componentes de tu almuerzo (Arroz, Carnes, Ensaladas, etc.). La app validará tu selección.
4.  Haz clic en "Enviar Pedido".

### Administradores
1.  Si tienes rol de administrador, verás un panel de estadísticas en la parte superior o inferior (según implementación).
2.  Puedes consultar los reportes agregados por departamento en la hoja de cálculo.

## Guía de Despliegue

Para publicar cambios o una nueva versión:

1.  En el editor de Apps Script, haz clic en **Implementar** > **Nueva implementación**.
2.  Selecciona el tipo **Aplicación web**.
3.  Configura:
    *   **Ejecutar como**: *Yo* (propietario del script).
    *   **Quién tiene acceso**: *Cualquier usuario de [Tu Organización]* (o según política interna).
4.  Haz clic en **Implementar**.
5.  Copia la URL proporcionada y compártela con los usuarios.

> **Nota**: Recuerda actualizar `APP_VERSION` en `Código.js` antes de desplegar para mantener el control de versiones.

## Contacto

Equipo de Desarrollo Interno.
Para soporte técnico, contactar a: [Correo de Soporte]
