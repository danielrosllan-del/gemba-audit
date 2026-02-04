# Gestión de Reuniones - Seguimiento de Acciones

Sistema web desarrollado en Google Apps Script para gestionar y dar seguimiento a acciones derivadas de reuniones, con integración a **Microsoft Outlook** para notificaciones por correo electrónico.

![Dashboard Preview](docs/dashboard-preview.png)

## Características

### 📊 Dashboard
- KPIs de cumplimiento en tiempo real
- Gráficos de cumplimiento por indicador (P, Q, C, D, S, M, E)
- Tabla de responsables con métricas
- Gráficos de cumplimiento por área y sector
- Filtros dinámicos por planta, gerencia, área y tipo de reunión

### 📝 Registro de Acciones
- Formulario completo para registro individual
- Carga masiva mediante archivo Excel
- Descarga de plantilla maestra
- Selección múltiple de indicadores
- Envío automático de notificaciones

### 📋 Seguimiento
- Tabla con todas las acciones
- Filtros avanzados (planta, área, sector, estado, etc.)
- Estados visuales (Concluido, En Proceso, Retrasado)
- Vista detallada de cada acción
- Marcar acciones como concluidas

### 👔 Gerencia
- Acceso restringido por credenciales
- KPIs gerenciales
- Lista de responsables con bajo cumplimiento
- Acciones próximas a vencer
- Envío de recordatorios masivos

### 📈 Análisis
- Matriz de herramientas (ACR, AIQ, AFP, CAPDO)
- Matriz de planes 5W2H
- Gráfico de tendencia de cumplimiento
- Distribución por herramienta

### 📧 Integración con Outlook
- Notificaciones automáticas al asignar acciones
- Recordatorios de acciones próximas a vencer
- Plantillas de correo HTML profesionales

---

## Requisitos

- Cuenta de Google (Gmail o Google Workspace)
- Cuenta de Microsoft 365 con acceso a Azure AD (para notificaciones Outlook)
- Google Sheets para almacenamiento de datos

---

## Instalación

### 1. Crear proyecto en Google Apps Script

1. Ir a [Google Apps Script](https://script.google.com)
2. Crear un nuevo proyecto
3. Nombrar el proyecto: "Gestión de Reuniones"

### 2. Estructura de archivos

Crear los siguientes archivos en el proyecto:

```
📁 Proyecto
├── 📄 Code.gs              (Copiar de src/server/Code.gs)
├── 📄 OutlookConfig.gs     (Copiar de src/config/OutlookConfig.gs)
├── 📄 Index.html           (Copiar de src/client/html/Index.html)
├── 📄 Styles.html          (Copiar de src/client/css/Styles.html)
├── 📄 Dashboard.html       (Copiar de src/client/html/Dashboard.html)
├── 📄 Registro.html        (Copiar de src/client/html/Registro.html)
├── 📄 Seguimiento.html     (Copiar de src/client/html/Seguimiento.html)
├── 📄 Gerencia.html        (Copiar de src/client/html/Gerencia.html)
├── 📄 Analisis.html        (Copiar de src/client/html/Analisis.html)
└── 📄 Scripts.html         (Copiar de src/client/js/Scripts.html)
```

### 3. Configurar Google Sheets

1. Ejecutar la función `setupSpreadsheet()` para crear la hoja de cálculo
2. Copiar el ID del Spreadsheet generado
3. Pegar el ID en `CONFIG.SPREADSHEET_ID` en `Code.gs`

### 4. Configurar Azure AD (para Outlook)

#### 4.1 Registrar aplicación en Azure

1. Ir a [Azure Portal](https://portal.azure.com)
2. Navegar a **Azure Active Directory** > **App registrations**
3. Click en **New registration**
4. Configurar:
   - **Name**: Gestión de Reuniones - Apps Script
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Web - `https://script.google.com`
5. Click en **Register**

#### 4.2 Configurar permisos de API

1. En la aplicación creada, ir a **API permissions**
2. Click en **Add a permission**
3. Seleccionar **Microsoft Graph**
4. Seleccionar **Application permissions**
5. Agregar los siguientes permisos:
   - `Mail.Send`
   - `Mail.ReadWrite`
   - `User.Read.All`
6. Click en **Grant admin consent**

#### 4.3 Crear secreto de cliente

1. Ir a **Certificates & secrets**
2. Click en **New client secret**
3. Agregar descripción y seleccionar expiración
4. **IMPORTANTE**: Copiar el valor del secreto (solo se muestra una vez)

#### 4.4 Configurar credenciales en el proyecto

En el archivo `OutlookConfig.gs`, actualizar:

```javascript
const OUTLOOK_CONFIG = {
  CLIENT_ID: 'tu-client-id-aqui',
  CLIENT_SECRET: 'tu-client-secret-aqui',
  TENANT_ID: 'tu-tenant-id-aqui',
  SENDER_EMAIL: 'notificaciones@tuempresa.com',
  // ...
};
```

### 5. Desplegar como Web App

1. En Google Apps Script, ir a **Deploy** > **New deployment**
2. Seleccionar tipo: **Web app**
3. Configurar:
   - **Description**: Gestión de Reuniones v1.0
   - **Execute as**: Me
   - **Who has access**: Anyone / Anyone within organization
4. Click en **Deploy**
5. Autorizar los permisos solicitados
6. Copiar la URL de la web app

---

## Configuración de Catálogos

### Hoja "Catalogos" en Google Sheets

Crear columnas con los valores de cada catálogo:

| plantas | gerencias | areas | sectores | tiposReunion | herramientas | pilaresTPM |
|---------|-----------|-------|----------|--------------|--------------|------------|
| Planta Lima | Gerencia Ops | Producción | Sector A | Diaria | ACR | MA |
| Planta Arequipa | Gerencia Cal | Calidad | Sector B | Semanal | AIQ | MP |
| ... | ... | ... | ... | ... | ... | ... |

### Hoja "Usuarios" (para acceso gerencial)

| DNI | Password | Rol | Nombre | Email |
|-----|----------|-----|--------|-------|
| 12345678 | pass123 | gerente | Juan Pérez | juan@empresa.com |

---

## Uso

### Dashboard
1. Acceder a la URL de la web app
2. La vista Dashboard carga automáticamente
3. Usar filtros para segmentar datos
4. Click en "Limpiar" para resetear filtros

### Registro de Acciones
1. Click en pestaña **REGISTRO**
2. Completar todos los campos requeridos (*)
3. Seleccionar indicadores (P, Q, C, D, S, M, E)
4. Click en **Guardar**
5. Se enviará notificación automática al responsable

### Carga Masiva
1. Click en **Plantilla Maestra** para descargar el formato
2. Completar el archivo Excel
3. Click en **Subir Excel**
4. Revisar resultados de la carga

### Seguimiento
1. Click en pestaña **SEGUIMIENTO**
2. Usar filtros para encontrar acciones
3. Click en **Ver** para ver detalle
4. Click en ✓ para marcar como concluido

### Acceso Gerencial
1. Click en pestaña **GERENCIA**
2. Click en **Iniciar Sesión**
3. Ingresar DNI y contraseña
4. Acceder a funciones avanzadas

---

## Triggers Automáticos

Ejecutar `configurarTriggers()` para configurar:

- **Recordatorios diarios**: 8:00 AM - Envía recordatorios de acciones próximas a vencer
- **Actualización de estados**: Cada hora - Actualiza estados de acciones retrasadas

---

## Estructura del Proyecto

```
gemba-audit/
├── src/
│   ├── server/
│   │   └── Code.gs          # Lógica del servidor
│   ├── client/
│   │   ├── html/
│   │   │   ├── Index.html   # Página principal
│   │   │   ├── Dashboard.html
│   │   │   ├── Registro.html
│   │   │   ├── Seguimiento.html
│   │   │   ├── Gerencia.html
│   │   │   └── Analisis.html
│   │   ├── css/
│   │   │   └── Styles.html  # Estilos CSS
│   │   └── js/
│   │       └── Scripts.html # JavaScript principal
│   └── config/
│       └── OutlookConfig.gs # Configuración Outlook
└── README.md
```

---

## Tecnologías Utilizadas

- **Google Apps Script** - Backend y hosting
- **Google Sheets** - Base de datos
- **HTML5/CSS3** - Interfaz de usuario
- **JavaScript** - Lógica del cliente
- **Chart.js** - Gráficos
- **Microsoft Graph API** - Integración con Outlook

---

## Indicadores (PQCDSME)

| Código | Significado |
|--------|-------------|
| P | Productividad |
| Q | Quality (Calidad) |
| C | Cost (Costo) |
| D | Delivery (Entrega) |
| S | Safety (Seguridad) |
| M | Morale (Moral) |
| E | Environment (Ambiente) |

---

## Herramientas de Mejora Continua

| Herramienta | Descripción |
|-------------|-------------|
| ACR | Análisis de Causa Raíz |
| AIQ | Análisis de Incidentes de Calidad |
| AFP | Análisis de Fallas de Proceso |
| CAPDO | Check-Act-Plan-Do |
| 5W2H | What, Who, Where, When, Why, How, How Much |

---

## Solución de Problemas

### Error de autenticación con Outlook
1. Verificar que las credenciales de Azure AD sean correctas
2. Confirmar que los permisos de API estén otorgados
3. Ejecutar `testOutlookConnection()` para diagnosticar

### No se cargan los datos
1. Verificar el ID del Spreadsheet
2. Confirmar permisos de acceso a la hoja
3. Revisar la consola de Apps Script para errores

### Error al desplegar
1. Verificar que todos los archivos estén creados
2. Confirmar que no haya errores de sintaxis
3. Revisar los logs de ejecución

---

## Contribuir

1. Fork del repositorio
2. Crear rama de feature (`git checkout -b feature/NuevaCaracteristica`)
3. Commit de cambios (`git commit -m 'Agregar nueva característica'`)
4. Push a la rama (`git push origin feature/NuevaCaracteristica`)
5. Crear Pull Request

---

## Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para detalles.

---

## Contacto

Para soporte o consultas, contactar al equipo de desarrollo.

---

*Desarrollado con ❤️ para la gestión eficiente de acciones*
