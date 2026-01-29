/**
 * CONFIGURACIÓN DE INTEGRACIÓN CON MICROSOFT OUTLOOK
 *
 * Este archivo contiene las funciones y configuraciones necesarias
 * para integrar la aplicación con Microsoft Outlook a través de
 * Microsoft Graph API.
 */

// ============================================
// CONFIGURACIÓN DE AZURE AD
// ============================================

/**
 * Configuración de la aplicación registrada en Azure AD
 *
 * PASOS PARA CONFIGURAR:
 *
 * 1. Ir a Azure Portal (https://portal.azure.com)
 * 2. Navegar a "Azure Active Directory" > "App registrations"
 * 3. Crear nueva aplicación:
 *    - Nombre: "Gestión de Reuniones - Apps Script"
 *    - Tipo de cuenta: "Cuentas en este directorio organizacional"
 *    - URI de redirección: https://script.google.com/macros/d/{SCRIPT_ID}/usercallback
 *
 * 4. Configurar permisos de API:
 *    - Microsoft Graph > Application permissions:
 *      - Mail.Send
 *      - Mail.ReadWrite
 *      - User.Read.All
 *    - Otorgar consentimiento de administrador
 *
 * 5. Crear secreto de cliente:
 *    - Certificates & secrets > New client secret
 *    - Guardar el valor del secreto (solo se muestra una vez)
 *
 * 6. Copiar los valores en las constantes de abajo
 */

const OUTLOOK_CONFIG = {
  // ID de la aplicación (Client ID)
  CLIENT_ID: 'TU_CLIENT_ID_AQUI',

  // Secreto de cliente
  CLIENT_SECRET: 'TU_CLIENT_SECRET_AQUI',

  // ID del tenant de Azure AD
  TENANT_ID: 'TU_TENANT_ID_AQUI',

  // Email del remitente (debe tener licencia de Outlook/Exchange)
  SENDER_EMAIL: 'notificaciones@tuempresa.com',

  // Scopes requeridos
  SCOPES: [
    'https://graph.microsoft.com/.default'
  ],

  // URLs de Microsoft Graph
  TOKEN_URL: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token',
  GRAPH_URL: 'https://graph.microsoft.com/v1.0'
};

// ============================================
// FUNCIONES DE AUTENTICACIÓN
// ============================================

/**
 * Obtener token de acceso usando credenciales de cliente
 * (Client Credentials Flow - para aplicaciones daemon/server)
 */
function getOutlookAccessToken() {
  const tokenUrl = OUTLOOK_CONFIG.TOKEN_URL.replace('{tenant}', OUTLOOK_CONFIG.TENANT_ID);

  const payload = {
    'client_id': OUTLOOK_CONFIG.CLIENT_ID,
    'client_secret': OUTLOOK_CONFIG.CLIENT_SECRET,
    'scope': OUTLOOK_CONFIG.SCOPES.join(' '),
    'grant_type': 'client_credentials'
  };

  // Convertir payload a formato URL encoded
  const payloadString = Object.keys(payload)
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(payload[key]))
    .join('&');

  const options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payloadString,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      const json = JSON.parse(responseText);
      return {
        success: true,
        token: json.access_token,
        expiresIn: json.expires_in
      };
    } else {
      console.error('Error obteniendo token:', responseText);
      return {
        success: false,
        error: responseText
      };
    }
  } catch (error) {
    console.error('Error de conexión:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================
// FUNCIONES DE ENVÍO DE CORREO
// ============================================

/**
 * Enviar correo a través de Microsoft Graph API
 *
 * @param {string} destinatario - Email del destinatario
 * @param {string} asunto - Asunto del correo
 * @param {string} contenidoHTML - Cuerpo del correo en HTML
 * @param {Array} adjuntos - Array de objetos {nombre, contenidoBase64, tipoMime}
 */
function enviarCorreoOutlook(destinatario, asunto, contenidoHTML, adjuntos = []) {
  // Obtener token de acceso
  const tokenResult = getOutlookAccessToken();

  if (!tokenResult.success) {
    console.error('No se pudo obtener token de acceso');
    return { success: false, error: 'Error de autenticación' };
  }

  // Construir el mensaje
  const mensaje = {
    message: {
      subject: asunto,
      body: {
        contentType: 'HTML',
        content: contenidoHTML
      },
      toRecipients: [
        {
          emailAddress: {
            address: destinatario
          }
        }
      ]
    },
    saveToSentItems: true
  };

  // Agregar adjuntos si existen
  if (adjuntos.length > 0) {
    mensaje.message.attachments = adjuntos.map(adj => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: adj.nombre,
      contentType: adj.tipoMime,
      contentBytes: adj.contenidoBase64
    }));
  }

  // Endpoint de envío
  const sendMailUrl = `${OUTLOOK_CONFIG.GRAPH_URL}/users/${OUTLOOK_CONFIG.SENDER_EMAIL}/sendMail`;

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': `Bearer ${tokenResult.token}`
    },
    'payload': JSON.stringify(mensaje),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(sendMailUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 202) {
      console.log('Correo enviado exitosamente a:', destinatario);
      return { success: true };
    } else {
      const errorText = response.getContentText();
      console.error('Error enviando correo:', errorText);
      return { success: false, error: errorText };
    }
  } catch (error) {
    console.error('Error de conexión:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Enviar correo con copia y copia oculta
 */
function enviarCorreoOutlookConCopias(destinatario, asunto, contenidoHTML, cc = [], bcc = []) {
  const tokenResult = getOutlookAccessToken();

  if (!tokenResult.success) {
    return { success: false, error: 'Error de autenticación' };
  }

  const mensaje = {
    message: {
      subject: asunto,
      body: {
        contentType: 'HTML',
        content: contenidoHTML
      },
      toRecipients: [
        { emailAddress: { address: destinatario } }
      ],
      ccRecipients: cc.map(email => ({ emailAddress: { address: email } })),
      bccRecipients: bcc.map(email => ({ emailAddress: { address: email } }))
    },
    saveToSentItems: true
  };

  const sendMailUrl = `${OUTLOOK_CONFIG.GRAPH_URL}/users/${OUTLOOK_CONFIG.SENDER_EMAIL}/sendMail`;

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': `Bearer ${tokenResult.token}`
    },
    'payload': JSON.stringify(mensaje),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(sendMailUrl, options);
    return { success: response.getResponseCode() === 202 };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================
// PLANTILLAS DE CORREO
// ============================================

/**
 * Generar plantilla de notificación de nueva acción
 */
function generarPlantillaAccionNueva(accion) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      line-height: 1.6;
      color: #333;
      margin: 0;
      padding: 0;
    }
    .container {
      max-width: 600px;
      margin: 0 auto;
      background: #ffffff;
    }
    .header {
      background: linear-gradient(135deg, #1a365d 0%, #2c5282 100%);
      color: white;
      padding: 30px;
      text-align: center;
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .header p {
      margin: 10px 0 0;
      opacity: 0.9;
    }
    .content {
      padding: 30px;
    }
    .info-box {
      background: #f7fafc;
      border-left: 4px solid #3182ce;
      padding: 20px;
      margin: 20px 0;
      border-radius: 0 8px 8px 0;
    }
    .field {
      margin-bottom: 15px;
    }
    .field-label {
      font-size: 12px;
      font-weight: 600;
      color: #718096;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .field-value {
      font-size: 15px;
      color: #2d3748;
      margin-top: 4px;
    }
    .highlight {
      background: #ebf8ff;
      padding: 15px;
      border-radius: 8px;
      margin: 20px 0;
    }
    .highlight strong {
      color: #2c5282;
    }
    .btn {
      display: inline-block;
      padding: 12px 30px;
      background: #3182ce;
      color: white;
      text-decoration: none;
      border-radius: 6px;
      font-weight: 500;
      margin-top: 20px;
    }
    .footer {
      background: #f7fafc;
      padding: 20px;
      text-align: center;
      font-size: 12px;
      color: #718096;
    }
    .badge {
      display: inline-block;
      padding: 4px 12px;
      border-radius: 4px;
      font-size: 12px;
      font-weight: 600;
    }
    .badge-warning {
      background: #fefcbf;
      color: #d69e2e;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Nueva Acción Asignada</h1>
      <p>Gestión de Reuniones - Seguimiento de Acciones</p>
    </div>
    <div class="content">
      <p>Estimado/a <strong>${accion.responsables}</strong>,</p>
      <p>Se le ha asignado una nueva acción que requiere su atención:</p>

      <div class="info-box">
        <div class="field">
          <div class="field-label">Motivo</div>
          <div class="field-value"><strong>${accion.motivo || 'No especificado'}</strong></div>
        </div>
        <div class="field">
          <div class="field-label">Plan de Acción</div>
          <div class="field-value">${accion.plan || 'No especificado'}</div>
        </div>
      </div>

      <div class="highlight">
        <div class="field">
          <div class="field-label">Fecha Compromiso</div>
          <div class="field-value"><strong>${accion.fechaCompromiso}</strong> <span class="badge badge-warning">Importante</span></div>
        </div>
      </div>

      <table style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="padding: 8px 0;">
            <span class="field-label">Herramienta:</span>
            <span class="field-value">${accion.herramienta}</span>
          </td>
          <td style="padding: 8px 0;">
            <span class="field-label">Indicadores:</span>
            <span class="field-value">${accion.indicadores.join(', ')}</span>
          </td>
        </tr>
        <tr>
          <td style="padding: 8px 0;">
            <span class="field-label">Reportado por:</span>
            <span class="field-value">${accion.reportadoPor}</span>
          </td>
          <td style="padding: 8px 0;">
            <span class="field-label">Planta:</span>
            <span class="field-value">${accion.planta}</span>
          </td>
        </tr>
      </table>

      <p style="margin-top: 30px;">Por favor, revise la acción y actualice su estado cuando corresponda.</p>

      <center>
        <a href="#" class="btn">Ver en el Sistema</a>
      </center>
    </div>
    <div class="footer">
      <p>Este es un mensaje automático del sistema de Gestión de Reuniones.</p>
      <p>Por favor, no responda directamente a este correo.</p>
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * Generar plantilla de recordatorio
 */
function generarPlantillaRecordatorio(accion, diasRestantes) {
  const urgencia = diasRestantes <= 1 ? 'URGENTE' : diasRestantes <= 3 ? 'Próximo a vencer' : 'Recordatorio';
  const colorUrgencia = diasRestantes <= 1 ? '#e53e3e' : diasRestantes <= 3 ? '#d69e2e' : '#3182ce';

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Segoe UI', sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; }
    .header { background: ${colorUrgencia}; color: white; padding: 20px; text-align: center; }
    .content { padding: 30px; background: #fff; }
    .countdown { font-size: 48px; font-weight: bold; color: ${colorUrgencia}; text-align: center; }
    .footer { background: #f7fafc; padding: 15px; text-align: center; font-size: 12px; color: #718096; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>⚠️ ${urgencia}: Acción Pendiente</h1>
    </div>
    <div class="content">
      <div class="countdown">${diasRestantes} día${diasRestantes !== 1 ? 's' : ''}</div>
      <p style="text-align: center; color: #718096;">restantes para la fecha compromiso</p>

      <h3 style="color: #2d3748; margin-top: 30px;">Detalles de la acción:</h3>
      <p><strong>Motivo:</strong> ${accion.motivo}</p>
      <p><strong>Plan:</strong> ${accion.plan}</p>
      <p><strong>Fecha Compromiso:</strong> ${accion.fechaCompromiso}</p>

      <p style="margin-top: 30px;">Por favor, actualice el estado de esta acción lo antes posible.</p>
    </div>
    <div class="footer">
      <p>Sistema de Gestión de Reuniones - Seguimiento de Acciones</p>
    </div>
  </div>
</body>
</html>
  `;
}

// ============================================
// FUNCIONES DE PRUEBA
// ============================================

/**
 * Probar conexión con Microsoft Graph
 */
function testOutlookConnection() {
  const result = getOutlookAccessToken();

  if (result.success) {
    console.log('✅ Conexión exitosa con Microsoft Graph');
    console.log('Token expira en:', result.expiresIn, 'segundos');
    return true;
  } else {
    console.error('❌ Error de conexión:', result.error);
    return false;
  }
}

/**
 * Enviar correo de prueba
 */
function testSendEmail() {
  const resultado = enviarCorreoOutlook(
    'tu-email@empresa.com',
    'Prueba - Gestión de Reuniones',
    '<h1>Correo de Prueba</h1><p>Este es un correo de prueba del sistema.</p>'
  );

  console.log('Resultado:', resultado);
}
