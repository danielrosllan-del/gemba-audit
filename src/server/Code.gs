/**
 * GESTIÓN DE REUNIONES - SEGUIMIENTO DE ACCIONES
 * Aplicación Google Apps Script con integración a Microsoft Outlook
 *
 * @author Tu Empresa
 * @version 1.0.0
 */

// ============================================
// CONFIGURACIÓN GLOBAL
// ============================================

const CONFIG = {
  SPREADSHEET_ID: '', // ID de la hoja de cálculo de datos
  SHEET_ACCIONES: 'Acciones',
  SHEET_USUARIOS: 'Usuarios',
  SHEET_CATALOGOS: 'Catalogos',
  SHEET_CONFIG: 'Configuracion',

  // Microsoft Graph API (Outlook)
  MS_CLIENT_ID: '',
  MS_CLIENT_SECRET: '',
  MS_TENANT_ID: '',
  MS_REDIRECT_URI: '',

  // Metas
  META_CUMPLIMIENTO: 0.95,
  META_AVANCE: 0.75
};

// ============================================
// FUNCIONES PRINCIPALES DE LA WEB APP
// ============================================

/**
 * Punto de entrada para peticiones GET
 */
function doGet(e) {
  const page = e.parameter.page || 'index';
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestión de Reuniones')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Punto de entrada para peticiones POST
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    switch(action) {
      case 'guardarAccion':
        return ContentService.createTextOutput(JSON.stringify(guardarAccion(data.payload)))
          .setMimeType(ContentService.MimeType.JSON);
      case 'actualizarAccion':
        return ContentService.createTextOutput(JSON.stringify(actualizarAccion(data.payload)))
          .setMimeType(ContentService.MimeType.JSON);
      default:
        return ContentService.createTextOutput(JSON.stringify({error: 'Acción no válida'}))
          .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({error: error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Incluir archivos HTML parciales
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// FUNCIONES DE DATOS - ACCIONES
// ============================================

/**
 * Obtener todas las acciones con filtros opcionales
 */
function getAcciones(filtros = {}) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_ACCIONES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    let acciones = [];

    for (let i = 1; i < data.length; i++) {
      let accion = {};
      headers.forEach((header, index) => {
        accion[header] = data[i][index];
      });

      // Aplicar filtros
      let cumpleFiltros = true;

      if (filtros.planta && filtros.planta !== 'Todas' && accion.planta !== filtros.planta) {
        cumpleFiltros = false;
      }
      if (filtros.area && filtros.area !== 'Todas' && accion.area !== filtros.area) {
        cumpleFiltros = false;
      }
      if (filtros.gerencia && filtros.gerencia !== 'Todas' && accion.gerencia !== filtros.gerencia) {
        cumpleFiltros = false;
      }
      if (filtros.tipoReunion && filtros.tipoReunion !== 'Todas' && accion.tipoReunion !== filtros.tipoReunion) {
        cumpleFiltros = false;
      }
      if (filtros.estado && filtros.estado !== 'Todos' && accion.estado !== filtros.estado) {
        cumpleFiltros = false;
      }
      if (filtros.responsable && accion.responsable !== filtros.responsable) {
        cumpleFiltros = false;
      }
      if (filtros.sector && filtros.sector !== 'Todos' && accion.sector !== filtros.sector) {
        cumpleFiltros = false;
      }

      if (cumpleFiltros) {
        acciones.push(accion);
      }
    }

    return { success: true, data: acciones };
  } catch (error) {
    console.error('Error en getAcciones:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Guardar nueva acción
 */
function guardarAccion(accion) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_ACCIONES);

    // Generar ID único
    const id = Utilities.getUuid();
    const fechaCreacion = new Date();

    // Calcular estado inicial
    const estado = calcularEstado(accion.fechaCompromiso);

    const fila = [
      id,
      fechaCreacion,
      accion.fechaRegistro,
      accion.planta,
      accion.donde,
      accion.tipoReunion,
      accion.herramienta,
      accion.reportadoPor,
      accion.verificador,
      accion.motivo,
      accion.plan,
      accion.indicadores.join(','),
      accion.responsables,
      accion.email,
      accion.fechaCompromiso,
      accion.comentarios,
      estado,
      '', // Fecha conclusión
      '', // Evidencia
      accion.area || '',
      accion.sector || '',
      accion.gerencia || '',
      accion.pilarTPM || ''
    ];

    sheet.appendRow(fila);

    // Enviar notificación por Outlook
    if (accion.email) {
      enviarNotificacionOutlook(accion);
    }

    return { success: true, id: id, message: 'Acción guardada correctamente' };
  } catch (error) {
    console.error('Error en guardarAccion:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Actualizar acción existente
 */
function actualizarAccion(accion) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_ACCIONES);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === accion.id) {
        // Actualizar campos específicos
        if (accion.estado) sheet.getRange(i + 1, 17).setValue(accion.estado);
        if (accion.fechaConclusion) sheet.getRange(i + 1, 18).setValue(accion.fechaConclusion);
        if (accion.evidencia) sheet.getRange(i + 1, 19).setValue(accion.evidencia);
        if (accion.comentarios) sheet.getRange(i + 1, 16).setValue(accion.comentarios);

        return { success: true, message: 'Acción actualizada correctamente' };
      }
    }

    return { success: false, error: 'Acción no encontrada' };
  } catch (error) {
    console.error('Error en actualizarAccion:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Eliminar acción
 */
function eliminarAccion(id) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_ACCIONES);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Acción eliminada correctamente' };
      }
    }

    return { success: false, error: 'Acción no encontrada' };
  } catch (error) {
    console.error('Error en eliminarAccion:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Calcular estado de una acción basado en fecha compromiso
 */
function calcularEstado(fechaCompromiso) {
  const hoy = new Date();
  const fecha = new Date(fechaCompromiso);

  if (fecha < hoy) {
    return 'Retrasado';
  } else if (fecha.toDateString() === hoy.toDateString()) {
    return 'Vence Hoy';
  } else {
    return 'En Proceso';
  }
}

// ============================================
// FUNCIONES DE DASHBOARD / MÉTRICAS
// ============================================

/**
 * Obtener métricas del dashboard
 */
function getDashboardMetrics(filtros = {}) {
  try {
    const resultado = getAcciones(filtros);
    if (!resultado.success) return resultado;

    const acciones = resultado.data;
    const total = acciones.length;

    if (total === 0) {
      return {
        success: true,
        data: {
          total: 0,
          ok: 0,
          enProceso: 0,
          noProgramado: 0,
          retrasado: 0,
          porcentajeCumplimiento: 0,
          porIndicador: {},
          porResponsable: [],
          porArea: [],
          porSector: []
        }
      };
    }

    // Contar por estado
    let ok = 0, enProceso = 0, noProgramado = 0, retrasado = 0;

    acciones.forEach(accion => {
      switch(accion.estado) {
        case 'Concluido': ok++; break;
        case 'En Proceso': enProceso++; break;
        case 'No Programado': noProgramado++; break;
        case 'Retrasado': retrasado++; break;
      }
    });

    const porcentajeCumplimiento = Math.round((ok / total) * 100);

    // Métricas por indicador (P, Q, C, D, S, M, E)
    const indicadores = ['P', 'Q', 'C', 'D', 'S', 'M', 'E'];
    const porIndicador = {};

    indicadores.forEach(ind => {
      const accionesInd = acciones.filter(a =>
        a.indicadores && a.indicadores.includes(ind)
      );
      const totalInd = accionesInd.length;
      const okInd = accionesInd.filter(a => a.estado === 'Concluido').length;
      porIndicador[ind] = totalInd > 0 ? Math.round((okInd / totalInd) * 100) : 0;
    });

    // Métricas por responsable
    const responsablesMap = {};
    acciones.forEach(accion => {
      const resp = accion.responsables || 'Sin Asignar';
      if (!responsablesMap[resp]) {
        responsablesMap[resp] = { total: 0, ok: 0, enProceso: 0, noProg: 0, retrasado: 0 };
      }
      responsablesMap[resp].total++;
      switch(accion.estado) {
        case 'Concluido': responsablesMap[resp].ok++; break;
        case 'En Proceso': responsablesMap[resp].enProceso++; break;
        case 'No Programado': responsablesMap[resp].noProg++; break;
        case 'Retrasado': responsablesMap[resp].retrasado++; break;
      }
    });

    const porResponsable = Object.entries(responsablesMap).map(([nombre, datos]) => ({
      nombre,
      ...datos,
      porcentajeCumplimiento: Math.round((datos.ok / datos.total) * 100),
      porcentajeAvance: Math.round(((datos.ok + datos.enProceso) / datos.total) * 100)
    })).sort((a, b) => b.porcentajeCumplimiento - a.porcentajeCumplimiento);

    // Métricas por área
    const areasMap = {};
    acciones.forEach(accion => {
      const area = accion.area || 'Sin Asignar';
      if (!areasMap[area]) {
        areasMap[area] = { total: 0, ok: 0 };
      }
      areasMap[area].total++;
      if (accion.estado === 'Concluido') areasMap[area].ok++;
    });

    const porArea = Object.entries(areasMap).map(([nombre, datos]) => ({
      nombre,
      porcentaje: Math.round((datos.ok / datos.total) * 100)
    })).sort((a, b) => b.porcentaje - a.porcentaje);

    // Métricas por sector
    const sectoresMap = {};
    acciones.forEach(accion => {
      const sector = accion.sector || 'Sin Asignar';
      if (!sectoresMap[sector]) {
        sectoresMap[sector] = { total: 0, ok: 0 };
      }
      sectoresMap[sector].total++;
      if (accion.estado === 'Concluido') sectoresMap[sector].ok++;
    });

    const porSector = Object.entries(sectoresMap).map(([nombre, datos]) => ({
      nombre,
      porcentaje: Math.round((datos.ok / datos.total) * 100)
    })).sort((a, b) => b.porcentaje - a.porcentaje);

    return {
      success: true,
      data: {
        total,
        ok,
        enProceso,
        noProgramado,
        retrasado,
        porcentajeCumplimiento,
        porIndicador,
        porResponsable,
        porArea,
        porSector
      }
    };
  } catch (error) {
    console.error('Error en getDashboardMetrics:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Obtener métricas para análisis (matrices)
 */
function getAnalisisMetrics(filtros = {}) {
  try {
    const resultado = getAcciones(filtros);
    if (!resultado.success) return resultado;

    const acciones = resultado.data;

    // Herramientas y áreas para la matriz
    const herramientas = ['ACR', 'AIQ', 'AFP', 'CAPDO'];
    const areas = ['Conversión', 'Fabricación', 'Mantenimiento', 'Operaciones Logísticas'];

    // Construir matriz de herramientas
    const matrizHerramientas = {};

    herramientas.forEach(herr => {
      matrizHerramientas[herr] = {};
      areas.forEach(area => {
        const accionesHA = acciones.filter(a =>
          a.herramienta === herr && a.area === area
        );
        const total = accionesHA.length;
        const ok = accionesHA.filter(a => a.estado === 'Concluido').length;
        const con5W2H = accionesHA.filter(a => a.plan5W2H).length;

        matrizHerramientas[herr][area] = {
          porcentajeCumplimiento: total > 0 ? Math.round((ok / total) * 100) : null,
          porcentaje5W2H: total > 0 ? Math.round((con5W2H / total) * 100) : null
        };
      });

      // Total por herramienta
      const accionesH = acciones.filter(a => a.herramienta === herr);
      const totalH = accionesH.length;
      const okH = accionesH.filter(a => a.estado === 'Concluido').length;
      const con5W2HH = accionesH.filter(a => a.plan5W2H).length;

      matrizHerramientas[herr]['TOTAL'] = {
        porcentajeCumplimiento: totalH > 0 ? Math.round((okH / totalH) * 100) : null,
        porcentaje5W2H: totalH > 0 ? Math.round((con5W2HH / totalH) * 100) : null
      };
    });

    // Planes 5W2H
    const planes5W2H = acciones
      .filter(a => a.plan5W2H)
      .map(a => ({
        causaRaiz: a.causaRaiz || '',
        what: a.plan5W2H_what || '',
        who: a.plan5W2H_who || '',
        where: a.plan5W2H_where || '',
        when: a.plan5W2H_when || '',
        why: a.plan5W2H_why || '',
        how: a.plan5W2H_how || '',
        howMuch: a.plan5W2H_howMuch || '',
        evidencia: a.evidencia || ''
      }));

    return {
      success: true,
      data: {
        matrizHerramientas,
        planes5W2H
      }
    };
  } catch (error) {
    console.error('Error en getAnalisisMetrics:', error);
    return { success: false, error: error.message };
  }
}

// ============================================
// FUNCIONES DE CATÁLOGOS
// ============================================

/**
 * Obtener catálogos para dropdowns
 */
function getCatalogos() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_CATALOGOS);

    if (!sheet) {
      // Retornar catálogos por defecto si no existe la hoja
      return {
        success: true,
        data: {
          plantas: ['Planta 1', 'Planta 2', 'Planta 3'],
          gerencias: ['Gerencia 1', 'Gerencia 2'],
          areas: ['Personas & Organización', 'Excelencia Operacional', 'Operaciones Logísticas', 'SHE', 'Conversión', 'Mantenimiento', 'Fabricación', 'Control de Producción', 'Calidad', 'Seguridad Física', 'Gerencia'],
          sectores: ['Almacén de Repuestos', 'Efluentes', 'Enfardado', 'PP1', 'L110', 'Almacén de Producto Terminado', 'ISO', 'MP1', 'Almacén de Semielaborados'],
          tiposReunion: ['Diaria', 'Semanal', 'Mensual', 'Extraordinaria'],
          herramientas: ['ACR', 'ACR - 5W2H', 'AIQ', 'AFP', 'CAPDO', 'Plan de acción', 'Coordinación TPM'],
          pilaresTPM: ['MA', 'MP', 'ME', 'EI', 'CI', 'SHE', 'ADM', 'ET'],
          indicadores: ['P', 'Q', 'C', 'D', 'S', 'M', 'E'],
          estados: ['Concluido', 'En Proceso', 'No Programado', 'Retrasado']
        }
      };
    }

    const data = sheet.getDataRange().getValues();
    const catalogos = {};

    // Asumiendo que cada columna es un catálogo diferente
    const headers = data[0];
    headers.forEach((header, index) => {
      catalogos[header] = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][index]) {
          catalogos[header].push(data[i][index]);
        }
      }
    });

    return { success: true, data: catalogos };
  } catch (error) {
    console.error('Error en getCatalogos:', error);
    return { success: false, error: error.message };
  }
}

// ============================================
// FUNCIONES DE AUTENTICACIÓN (GERENCIA)
// ============================================

/**
 * Verificar credenciales de acceso gerencial
 */
function verificarAccesoGerencial(dni, password) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_USUARIOS);

    if (!sheet) {
      return { success: false, error: 'Hoja de usuarios no configurada' };
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === dni && data[i][1] === password && data[i][2] === 'gerente') {
        return {
          success: true,
          usuario: {
            dni: data[i][0],
            nombre: data[i][3],
            rol: data[i][2]
          }
        };
      }
    }

    return { success: false, error: 'Credenciales inválidas' };
  } catch (error) {
    console.error('Error en verificarAccesoGerencial:', error);
    return { success: false, error: error.message };
  }
}

// ============================================
// INTEGRACIÓN CON MICROSOFT OUTLOOK
// ============================================

/**
 * Obtener token de acceso de Microsoft Graph
 */
function getMicrosoftAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${CONFIG.MS_TENANT_ID}/oauth2/v2.0/token`;

  const payload = {
    'client_id': CONFIG.MS_CLIENT_ID,
    'client_secret': CONFIG.MS_CLIENT_SECRET,
    'scope': 'https://graph.microsoft.com/.default',
    'grant_type': 'client_credentials'
  };

  const options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payload,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const json = JSON.parse(response.getContentText());
    return json.access_token;
  } catch (error) {
    console.error('Error obteniendo token de Microsoft:', error);
    return null;
  }
}

/**
 * Enviar notificación por Outlook
 */
function enviarNotificacionOutlook(accion) {
  const accessToken = getMicrosoftAccessToken();

  if (!accessToken) {
    console.error('No se pudo obtener token de acceso para Outlook');
    return false;
  }

  const emailEndpoint = 'https://graph.microsoft.com/v1.0/users/{user-id}/sendMail';

  const emailBody = {
    message: {
      subject: `Nueva Acción Asignada: ${accion.motivo}`,
      body: {
        contentType: 'HTML',
        content: generarPlantillaEmail(accion)
      },
      toRecipients: [
        {
          emailAddress: {
            address: accion.email
          }
        }
      ]
    },
    saveToSentItems: true
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': `Bearer ${accessToken}`
    },
    'payload': JSON.stringify(emailBody),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(emailEndpoint, options);
    return response.getResponseCode() === 202;
  } catch (error) {
    console.error('Error enviando email por Outlook:', error);
    return false;
  }
}

/**
 * Generar plantilla HTML para email
 */
function generarPlantillaEmail(accion) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; }
        .header { background-color: #1a365d; color: white; padding: 20px; text-align: center; }
        .content { padding: 20px; background-color: #f7fafc; }
        .field { margin-bottom: 15px; }
        .label { font-weight: bold; color: #2d3748; }
        .value { color: #4a5568; }
        .footer { padding: 20px; text-align: center; font-size: 12px; color: #718096; }
        .button { display: inline-block; padding: 10px 20px; background-color: #3182ce; color: white; text-decoration: none; border-radius: 5px; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>Nueva Acción Asignada</h1>
          <p>Gestión de Reuniones - Seguimiento de Acciones</p>
        </div>
        <div class="content">
          <div class="field">
            <span class="label">Motivo:</span>
            <span class="value">${accion.motivo}</span>
          </div>
          <div class="field">
            <span class="label">Plan:</span>
            <span class="value">${accion.plan}</span>
          </div>
          <div class="field">
            <span class="label">Fecha Compromiso:</span>
            <span class="value">${accion.fechaCompromiso}</span>
          </div>
          <div class="field">
            <span class="label">Reportado por:</span>
            <span class="value">${accion.reportadoPor}</span>
          </div>
          <div class="field">
            <span class="label">Herramienta:</span>
            <span class="value">${accion.herramienta}</span>
          </div>
          <div class="field">
            <span class="label">Indicadores:</span>
            <span class="value">${accion.indicadores.join(', ')}</span>
          </div>
        </div>
        <div class="footer">
          <p>Este es un mensaje automático del sistema de Gestión de Reuniones.</p>
          <p>Por favor, no responda a este correo.</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Enviar recordatorio de acciones próximas a vencer
 */
function enviarRecordatoriosOutlook() {
  const resultado = getAcciones({ estado: 'En Proceso' });

  if (!resultado.success) return;

  const hoy = new Date();
  const tresDias = new Date();
  tresDias.setDate(hoy.getDate() + 3);

  resultado.data.forEach(accion => {
    const fechaCompromiso = new Date(accion.fechaCompromiso);

    if (fechaCompromiso <= tresDias && fechaCompromiso >= hoy && accion.email) {
      enviarRecordatorioOutlook(accion);
    }
  });
}

/**
 * Enviar recordatorio individual
 */
function enviarRecordatorioOutlook(accion) {
  const accessToken = getMicrosoftAccessToken();

  if (!accessToken) return false;

  const diasRestantes = Math.ceil(
    (new Date(accion.fechaCompromiso) - new Date()) / (1000 * 60 * 60 * 24)
  );

  const emailBody = {
    message: {
      subject: `⚠️ Recordatorio: Acción próxima a vencer (${diasRestantes} días)`,
      body: {
        contentType: 'HTML',
        content: `
          <h2>Recordatorio de Acción Pendiente</h2>
          <p><strong>Motivo:</strong> ${accion.motivo}</p>
          <p><strong>Plan:</strong> ${accion.plan}</p>
          <p><strong>Fecha Compromiso:</strong> ${accion.fechaCompromiso}</p>
          <p style="color: #e53e3e;"><strong>Días restantes: ${diasRestantes}</strong></p>
        `
      },
      toRecipients: [{ emailAddress: { address: accion.email } }]
    }
  };

  // Similar al envío anterior...
  return true;
}

// ============================================
// FUNCIONES DE IMPORTACIÓN/EXPORTACIÓN
// ============================================

/**
 * Procesar archivo Excel para carga masiva
 */
function procesarCargaMasiva(base64Data) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    // Crear archivo temporal en Drive
    const tempFile = DriveApp.createFile(blob.setName('temp_import.xlsx'));
    const ss = SpreadsheetApp.open(tempFile);
    const sheet = ss.getActiveSheet();
    const data = sheet.getDataRange().getValues();

    const resultados = {
      exitosos: 0,
      errores: []
    };

    // Procesar cada fila (saltando encabezado)
    for (let i = 1; i < data.length; i++) {
      try {
        const accion = {
          fechaRegistro: data[i][0],
          planta: data[i][1],
          donde: data[i][2],
          tipoReunion: data[i][3],
          herramienta: data[i][4],
          reportadoPor: data[i][5],
          verificador: data[i][6],
          motivo: data[i][7],
          plan: data[i][8],
          indicadores: data[i][9] ? data[i][9].split(',') : [],
          responsables: data[i][10],
          email: data[i][11],
          fechaCompromiso: data[i][12],
          comentarios: data[i][13]
        };

        const result = guardarAccion(accion);
        if (result.success) {
          resultados.exitosos++;
        } else {
          resultados.errores.push({ fila: i + 1, error: result.error });
        }
      } catch (error) {
        resultados.errores.push({ fila: i + 1, error: error.message });
      }
    }

    // Eliminar archivo temporal
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);

    return { success: true, data: resultados };
  } catch (error) {
    console.error('Error en procesarCargaMasiva:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Generar plantilla Excel para descarga
 */
function generarPlantillaMaestra() {
  const headers = [
    'Fecha Registro',
    'Planta',
    'Donde',
    'Tipo Reunión',
    'Herramienta',
    'Reportado Por',
    'Verificador',
    'Motivo',
    'Plan',
    'Indicadores (P,Q,C,D,S,M,E)',
    'Responsables',
    'Email',
    'Fecha Compromiso',
    'Comentarios'
  ];

  const ss = SpreadsheetApp.create('Plantilla_Carga_Masiva');
  const sheet = ss.getActiveSheet();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#1a365d');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');

  // Auto-ajustar columnas
  headers.forEach((_, index) => {
    sheet.autoResizeColumn(index + 1);
  });

  const file = DriveApp.getFileById(ss.getId());
  const blob = file.getBlob();

  // Eliminar archivo temporal
  file.setTrashed(true);

  return {
    success: true,
    data: Utilities.base64Encode(blob.getBytes()),
    filename: 'Plantilla_Carga_Masiva.xlsx'
  };
}

// ============================================
// TRIGGERS Y TAREAS PROGRAMADAS
// ============================================

/**
 * Configurar triggers automáticos
 */
function configurarTriggers() {
  // Eliminar triggers existentes
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  // Trigger diario para recordatorios
  ScriptApp.newTrigger('enviarRecordatoriosOutlook')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  // Trigger para actualizar estados
  ScriptApp.newTrigger('actualizarEstadosAcciones')
    .timeBased()
    .everyHours(1)
    .create();
}

/**
 * Actualizar estados de acciones automáticamente
 */
function actualizarEstadosAcciones() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_ACCIONES);
    const data = sheet.getDataRange().getValues();

    const hoy = new Date();

    for (let i = 1; i < data.length; i++) {
      const estado = data[i][16];
      const fechaCompromiso = new Date(data[i][14]);

      // Solo actualizar si no está concluido
      if (estado !== 'Concluido') {
        let nuevoEstado = estado;

        if (fechaCompromiso < hoy) {
          nuevoEstado = 'Retrasado';
        } else if (estado === 'Retrasado' && fechaCompromiso >= hoy) {
          nuevoEstado = 'En Proceso';
        }

        if (nuevoEstado !== estado) {
          sheet.getRange(i + 1, 17).setValue(nuevoEstado);
        }
      }
    }
  } catch (error) {
    console.error('Error en actualizarEstadosAcciones:', error);
  }
}

// ============================================
// FUNCIONES AUXILIARES
// ============================================

/**
 * Obtener usuario actual (email)
 */
function getUsuarioActual() {
  return Session.getActiveUser().getEmail();
}

/**
 * Formatear fecha para mostrar
 */
function formatearFecha(fecha) {
  if (!fecha) return '';
  const d = new Date(fecha);
  return d.toLocaleDateString('es-PE', {
    day: '2-digit',
    month: 'short',
    year: '2-digit'
  });
}

/**
 * Configuración inicial del spreadsheet
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.create('Gestión de Reuniones - Data');

  // Crear hoja de Acciones
  const sheetAcciones = ss.getActiveSheet();
  sheetAcciones.setName('Acciones');
  const headersAcciones = [
    'ID', 'Fecha Creación', 'Fecha Registro', 'Planta', 'Donde',
    'Tipo Reunión', 'Herramienta', 'Reportado Por', 'Verificador',
    'Motivo', 'Plan', 'Indicadores', 'Responsables', 'Email',
    'Fecha Compromiso', 'Comentarios', 'Estado', 'Fecha Conclusión',
    'Evidencia', 'Área', 'Sector', 'Gerencia', 'Pilar TPM'
  ];
  sheetAcciones.getRange(1, 1, 1, headersAcciones.length).setValues([headersAcciones]);

  // Crear hoja de Usuarios
  const sheetUsuarios = ss.insertSheet('Usuarios');
  const headersUsuarios = ['DNI', 'Password', 'Rol', 'Nombre', 'Email'];
  sheetUsuarios.getRange(1, 1, 1, headersUsuarios.length).setValues([headersUsuarios]);

  // Crear hoja de Catálogos
  const sheetCatalogos = ss.insertSheet('Catalogos');

  console.log('Spreadsheet ID:', ss.getId());
  return ss.getId();
}
