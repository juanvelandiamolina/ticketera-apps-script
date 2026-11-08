/**
 * Ticketera del equipo de Data · Google Apps Script
 *
 * Aplicación web para registrar y seguir requerimientos con tres roles:
 *   - Solicitante: crea tickets y consulta su estado.
 *   - Aprobador:   aprueba o rechaza las solicitudes de prioridad alta.
 *   - Responsable: atiende los tickets, avanza su estado y deja observaciones.
 *
 * Los datos viven en una hoja de cálculo (pestaña "Tickets"), los adjuntos en una
 * carpeta de Drive y los avisos salen por correo con MailApp. El acceso es con
 * correo y contraseña propios (PBKDF2 + reto HMAC + sesión firmada), así que la
 * web puede publicarse sin exigir cuenta de Google.
 */

const CONFIG = {
  SHEET_NAME: "Tickets",
  // Quién atiende los tickets (rol Responsable) y quién aprueba las prioridades altas.
  RESPONSIBLE_EMAIL: "responsable@example.com",
  RESPONSIBLE_NAME: "Persona Responsable",
  APPROVER_EMAILS: ["aprobador@example.com", "responsable@example.com"],
  APPROVER_NAME: "Persona Aprobadora",
  SESSION_SECONDS: 43200,
  SESSION_VERSION: 2,
  PASSWORD_HASH_ITERATIONS: 60000,
  LOGIN_FAILURE_LIMIT: 5,
  LOGIN_LOCK_SECONDS: 900,
  // Se usan cuando el correo no existe, para que la respuesta no delate qué cuentas hay.
  // Genera los tuyos con herramientas/generar_cuenta.html y una contraseña cualquiera.
  DUMMY_PASSWORD_SALT: "0d61aed73d999cb28331c0c4b09c3a52",
  DUMMY_PASSWORD_HASH: "09eced15f2e310b96d98b8c8fc33900da6980864d3fd107c602740da71c90731",
  // Cuentas autorizadas. salt y hash se generan con herramientas/generar_cuenta.html:
  // la contraseña no se guarda en ninguna parte ni viaja al servidor.
  //   canSeeInternalViews  ve las vistas Responsable y Prioridad alta
  //   canResponsible       cambia estados y agrega observaciones
  //   canApprove           aprueba o rechaza prioridades altas
  ACCOUNTS: [
    { email: "solicitante@example.com", salt: "REEMPLAZAR", hash: "REEMPLAZAR", canSeeInternalViews: false, canResponsible: false, canApprove: false },
    { email: "aprobador@example.com", salt: "REEMPLAZAR", hash: "REEMPLAZAR", canSeeInternalViews: true, canResponsible: false, canApprove: true },
    { email: "responsable@example.com", salt: "REEMPLAZAR", hash: "REEMPLAZAR", canSeeInternalViews: true, canResponsible: true, canApprove: true }
  ],
  HEADERS: [
    "ID_Ticket",
    "Fecha_Solicitud",
    "Cliente",
    "País",
    "Tipo de Solicitud",
    "Detalles",
    "Solicitante",
    "Equipo/Área",
    "Prioridad",
    "Estado",
    "Responsable",
    "Fecha Inicio Atención",
    "Fecha Fin Atención",
    "Tiempo Atención",
    "Prioridad Alta Aceptada por",
    "Comentarios",
    "Archivos adjuntos",
    "Token aprobación",
    "Correo solicitante"
  ],
  // Quién puede figurar como solicitante y a qué equipo pertenece.
  APPLICANTS: [
    { name: "Ana Pérez", team: "Equipo Comercial" },
    { name: "Luis Gómez", team: "Equipo Comercial" },
    { name: "María Rojas", team: "Cuentas" },
    { name: "Carlos Díaz", team: "Cuentas" },
    { name: "Laura Torres", team: "Dirección" }
  ],
  REQUEST_TYPES: [
    "Otro",
    "Cambio de gráfico",
    "Cambio de texto/insight",
    "Cambio de logo",
    "Cambio de Fecha",
    "Nuevo reporte",
    "Crear PDF"
  ],
  COUNTRIES: ["COL", "PER"],
  PRIORITIES: ["Baja", "Media", "Alta"],
  STATES: [
    "Esperando Aprobación",
    "Pendiente",
    "En Proceso",
    "Esperando Validación",
    "Completado",
    "Rechazado"
  ],
  ATTACHMENT_COUNT_LIMIT: 3,
  ATTACHMENT_SIZE_LIMIT: 5 * 1024 * 1024,
  ATTACHMENT_TOTAL_LIMIT: 12 * 1024 * 1024
};

const COL = {
  ID: 0,
  REQUESTED_AT: 1,
  CLIENT: 2,
  COUNTRY: 3,
  REQUEST_TYPE: 4,
  DETAILS: 5,
  APPLICANT: 6,
  TEAM: 7,
  PRIORITY: 8,
  STATUS: 9,
  RESPONSIBLE: 10,
  STARTED_AT: 11,
  FINISHED_AT: 12,
  DURATION_MINUTES: 13,
  HIGH_PRIORITY_APPROVER: 14,
  COMMENTS: 15,
  ATTACHMENTS: 16,
  APPROVAL_TOKEN: 17,
  REQUESTER_EMAIL: 18
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Ticketera | Equipo de Data")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Ticketera")
    .addItem("Configurar ticketera", "setupTicketera_")
    .addToUi();
}

/**
 * Ejecutar una vez desde el proyecto vinculado a la hoja que conservará el enlace.
 * Guarda el ID del spreadsheet en propiedades y prepara la pestaña Tickets.
 */
function setupTicketera_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Abre el proyecto desde la hoja de cálculo que usará la ticketera y vuelve a ejecutar Configurar ticketera.");
  }
  MailApp.getRemainingDailyQuota();

  const props = PropertiesService.getScriptProperties();
  props.setProperty("SPREADSHEET_ID", ss.getId());
  if (!props.getProperty("AUTH_SIGNING_SECRET")) {
    props.setProperty("AUTH_SIGNING_SECRET", Utilities.getUuid() + Utilities.getUuid());
  }
  ss.setSpreadsheetTimeZone("America/Bogota");

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).getDisplayValues()[0];
    if (JSON.stringify(existingHeaders) !== JSON.stringify(CONFIG.HEADERS)) {
      throw new Error(
        "La pestaña Tickets ya contiene encabezados distintos. No se modificó su información. Revisa la pestaña antes de ejecutar la configuración."
      );
    }
  }

  sheet.setFrozenRows(1);
  sheet.setTabColor("#26724D");
  sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
    .setBackground("#143F32")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setWrap(true);
  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 145);
  sheet.setColumnWidth(2, 155);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 205);
  sheet.setColumnWidth(6, 360);
  sheet.setColumnWidth(7, 175);
  sheet.setColumnWidth(8, 165);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 175);
  sheet.setColumnWidth(11, 155);
  sheet.setColumnWidth(12, 165);
  sheet.setColumnWidth(13, 165);
  sheet.setColumnWidth(14, 145);
  sheet.setColumnWidth(15, 230);
  sheet.setColumnWidth(16, 350);
  sheet.setColumnWidth(17, 320);
  sheet.setColumnWidth(18, 250);
  sheet.setColumnWidth(19, 250);
  sheet.getRange("B:B").setNumberFormat("dd/MM/yyyy HH:mm");
  sheet.getRange("L:M").setNumberFormat("dd/MM/yyyy HH:mm");
  sheet.getRange("N:N").setNumberFormat("0");
  const editableRows = Math.max(sheet.getMaxRows() - 1, 1);
  setListValidation_(sheet, 4, CONFIG.COUNTRIES, editableRows);
  setListValidation_(sheet, 5, CONFIG.REQUEST_TYPES, editableRows);
  setListValidation_(sheet, 7, applicantNames_(), editableRows);
  setListValidation_(sheet, 9, CONFIG.PRIORITIES, editableRows);
  setListValidation_(sheet, 10, CONFIG.STATES, editableRows);
  setListValidation_(sheet, 11, [CONFIG.RESPONSIBLE_NAME], editableRows);

  if (lastRow > 1 && !sheet.getFilter()) {
    sheet.getRange(1, 1, lastRow, CONFIG.HEADERS.length).createFilter();
  }

  let folderId = props.getProperty("UPLOAD_FOLDER_ID");
  let folder;
  if (folderId) {
    try {
      folder = DriveApp.getFolderById(folderId);
      folder.getName();
    } catch (error) {
      folder = null;
    }
  }
  if (!folder) {
    folder = DriveApp.createFolder("Ticketera - Adjuntos de tickets");
    folderId = folder.getId();
    props.setProperty("UPLOAD_FOLDER_ID", folderId);
  }

  const summary = {
    spreadsheetUrl: ss.getUrl(),
    spreadsheetId: ss.getId(),
    sheetName: CONFIG.SHEET_NAME,
    uploadFolderUrl: folder.getUrl(),
    uploadFolderId: folderId
  };
  console.log(JSON.stringify(summary));
  return summary;
}

function getAppData(authToken) {
  const email = requireAuthenticated_(authToken);
  const account = getAccountByEmail_(email);
  const records = readTicketRecords_().reverse();

  const result = {
    userEmail: email,
    canSeeInternalViews: account.canSeeInternalViews,
    canResponsible: account.canResponsible,
    canApprove: account.canApprove,
    applicants: applicantNames_(),
    countries: CONFIG.COUNTRIES,
    requestTypes: CONFIG.REQUEST_TYPES,
    priorities: CONFIG.PRIORITIES,
    publicTickets: records.map(toPublicRecord_)
  };

  if (account.canSeeInternalViews) {
    result.managerTickets = records.map(toManagerRecord_);
    result.pendingApprovals = records
      .filter(function(record) {
        return record.priority === "Alta" && record.status === "Esperando Aprobación";
      })
      .map(function(record) {
        const item = toApprovalRecord_(record);
        item.token = record.token;
        return item;
      });
  }

  return result;
}

function createTicket(form, authToken) {
  const email = requireAuthenticated_(authToken);
  const input = validateTicketInput_(form || {});
  const preparedFiles = validateAttachments_(form && form.attachments);
  const createdFiles = [];
  let ticketId = "";
  let token = "";

  try {
    const folder = getUploadFolder_();
    preparedFiles.forEach(function(file) {
      const bytes = Utilities.base64Decode(file.data);
      const blob = Utilities.newBlob(bytes, file.mimeType, file.name);
      const driveFile = folder.createFile(blob);
      createdFiles.push({
        name: file.name,
        id: driveFile.getId()
      });
    });

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sheet = getTicketSheet_();
      ticketId = nextTicketId_(sheet);
      const timestamp = new Date();
      const team = getTeamForApplicant_(input.applicant);
      const highPriority = input.priority === "Alta";
      token = highPriority ? Utilities.getUuid() : "";
      const fileCell = createdFiles.map(function(file) {
        return file.name + " — " + file.id;
      }).join("\n");

      const row = [
        ticketId,
        timestamp,
        safeCellText_(input.client),
        input.country,
        input.requestType,
        safeCellText_(input.details),
        input.applicant,
        team,
        input.priority,
        highPriority ? "Esperando Aprobación" : "Pendiente",
        CONFIG.RESPONSIBLE_NAME,
        "",
        "",
        "",
        "",
        "",
        safeCellText_(fileCell),
        token,
        email
      ];

      sheet.getRange(sheet.getLastRow() + 1, 1, 1, CONFIG.HEADERS.length).setValues([row]);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    createdFiles.forEach(function(file) {
      try {
        DriveApp.getFileById(file.id).setTrashed(true);
      } catch (cleanupError) {
        console.warn("No se pudo eliminar un adjunto incompleto: " + cleanupError.message);
      }
    });
    throw error;
  }

  let notification = "";
  try {
    sendTicketCreatedNotification_(ticketId, input, email);
  } catch (error) {
    notification = " El ticket quedó guardado, pero no se pudo enviar el correo a la persona responsable.";
    console.error(error);
  }
  if (input.priority === "Alta") {
    try {
      sendHighPriorityRequest_(ticketId, input, email, token);
    } catch (error) {
      notification += " No se pudo enviar el correo de aprobación; la solicitud aparece en la vista de aprobación.";
      console.error(error);
    }
  }

  return {
    ok: true,
    ticketId: ticketId,
    status: input.priority === "Alta" ? "Esperando Aprobación" : "Pendiente",
    message: "Ticket creado." + notification
  };
}

function downloadAttachment(fileIdInput, authToken) {
  requireAuthenticated_(authToken);
  const fileId = String(fileIdInput || "").trim();
  if (!/^[A-Za-z0-9_-]{10,}$/.test(fileId)) {
    throw new Error("El archivo solicitado no es válido.");
  }

  const isAttached = readTicketRecords_().some(function(record) {
    return record.attachments.some(function(file) { return file.fileId === fileId; });
  });
  if (!isAttached) {
    throw new Error("No encontramos ese archivo en la ticketera.");
  }

  const file = DriveApp.getFileById(fileId);
  if (file.isTrashed()) {
    throw new Error("El archivo ya no está disponible.");
  }
  const blob = file.getBlob();
  return {
    name: file.getName(),
    mimeType: blob.getContentType() || "application/octet-stream",
    data: Utilities.base64Encode(blob.getBytes())
  };
}

function updateTicketStatus(ticketId, nextStatus, authToken) {
  const email = requireResponsible_(authToken);
  if (CONFIG.STATES.indexOf(nextStatus) < 0 || nextStatus === "Esperando Aprobación") {
    throw new Error("Ese estado no se puede asignar desde la vista Responsable.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getTicketSheet_();
    const rowNumber = findTicketRowById_(sheet, ticketId);
    const row = sheet.getRange(rowNumber, 1, 1, CONFIG.HEADERS.length).getValues()[0];
    const currentStatus = String(row[COL.STATUS] || "");
    const allowedNext = {
      "Pendiente": ["En Proceso", "Rechazado"],
      "En Proceso": ["Esperando Validación", "Rechazado"],
      "Esperando Validación": ["Completado", "Rechazado"]
    };

    if (!allowedNext[currentStatus] || allowedNext[currentStatus].indexOf(nextStatus) < 0) {
      throw new Error("El estado debe seguir la secuencia de atención; Rechazado solo se permite antes de completar el ticket.");
    }

    const now = new Date();
    if (nextStatus === "En Proceso" && !row[COL.STARTED_AT]) {
      sheet.getRange(rowNumber, COL.STARTED_AT + 1).setValue(now);
    }

    if (nextStatus === "Esperando Validación") {
      let startedAt = row[COL.STARTED_AT];
      if (!startedAt) {
        startedAt = now;
        sheet.getRange(rowNumber, COL.STARTED_AT + 1).setValue(startedAt);
      }
      sheet.getRange(rowNumber, COL.FINISHED_AT + 1).setValue(now);
      const minutes = Math.max(0, Math.floor((now.getTime() - new Date(startedAt).getTime()) / 60000));
      sheet.getRange(rowNumber, COL.DURATION_MINUTES + 1).setValue(minutes);
    }

    if (nextStatus === "Rechazado") {
      if (row[COL.STARTED_AT] && !row[COL.FINISHED_AT]) {
        sheet.getRange(rowNumber, COL.FINISHED_AT + 1).setValue(now);
        const minutes = Math.max(0, Math.floor((now.getTime() - new Date(row[COL.STARTED_AT]).getTime()) / 60000));
        sheet.getRange(rowNumber, COL.DURATION_MINUTES + 1).setValue(minutes);
      }
      appendCommentToRow_(sheet, rowNumber, "Ticket marcado como Rechazado por " + email + " · " + formatTimestamp_(now));
    }

    sheet.getRange(rowNumber, COL.STATUS + 1).setValue(nextStatus);
    SpreadsheetApp.flush();
    return { ok: true, ticketId: ticketId, status: nextStatus, changedBy: email };
  } finally {
    lock.releaseLock();
  }
}

function addResponsibleComment(ticketId, comment, authToken) {
  const email = requireResponsible_(authToken);
  const message = cleanText_(comment, 2000, "Comentario");
  if (!message) {
    throw new Error("Escribe una observación antes de guardarla.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getTicketSheet_();
    const rowNumber = findTicketRowById_(sheet, ticketId);
    appendCommentToRow_(sheet, rowNumber, email + ": " + message);
    SpreadsheetApp.flush();
    return { ok: true, ticketId: ticketId };
  } finally {
    lock.releaseLock();
  }
}

function decideHighPriority(token, decision, authToken) {
  const email = requireApprover_(authToken);
  if (decision !== "approve" && decision !== "reject") {
    throw new Error("La decisión no es válida.");
  }
  const approvalToken = String(token || "").trim();
  if (!approvalToken) {
    throw new Error("Falta el identificador de aprobación.");
  }

  let requesterEmail = "";
  let ticketId = "";
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getTicketSheet_();
    const rowNumber = findTicketRowByToken_(sheet, approvalToken);
    const row = sheet.getRange(rowNumber, 1, 1, CONFIG.HEADERS.length).getValues()[0];

    if (String(row[COL.STATUS] || "") !== "Esperando Aprobación" ||
        String(row[COL.PRIORITY] || "") !== "Alta") {
      throw new Error("Esta solicitud ya fue resuelta o ya no espera aprobación.");
    }

    ticketId = String(row[COL.ID] || "");
    requesterEmail = String(row[COL.REQUESTER_EMAIL] || "");
    const now = new Date();
    sheet.getRange(rowNumber, COL.STATUS + 1).setValue("Pendiente");

    if (decision === "approve") {
      sheet.getRange(rowNumber, COL.HIGH_PRIORITY_APPROVER + 1).setValue(email);
      appendCommentToRow_(sheet, rowNumber, "Prioridad alta aprobada por " + email + " · " + formatTimestamp_(now));
    } else {
      sheet.getRange(rowNumber, COL.PRIORITY + 1).setValue("Media");
      appendCommentToRow_(sheet, rowNumber, "Prioridad alta no aprobada por " + email + "; prioridad ajustada a Media · " + formatTimestamp_(now));
    }

    sheet.getRange(rowNumber, COL.APPROVAL_TOKEN + 1).clearContent();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  let notification = "";
  if (requesterEmail) {
    try {
      const approverLabel = email === CONFIG.RESPONSIBLE_EMAIL.toLowerCase() ? CONFIG.RESPONSIBLE_NAME : CONFIG.APPROVER_NAME;
      const subject = decision === "approve"
        ? "Prioridad alta aprobada: " + ticketId
        : "Prioridad revisada: " + ticketId;
      const body = decision === "approve"
        ? approverLabel + " aprobó la prioridad alta de tu ticket " + ticketId + ". La solicitud pasó a Pendiente."
        : approverLabel + " no aprobó la prioridad alta de tu ticket " + ticketId + ". La prioridad se ajustó a Media y la solicitud pasó a Pendiente.";
      const outcome = decision === "approve" ? "Aprobada" : "Ajustada a prioridad media";
      const htmlBody = emailFrame_(
        "Actualización de prioridad", "Ticket " + ticketId,
        decision === "approve" ? "La solicitud de prioridad alta fue aprobada." : "La prioridad alta no fue aprobada; el ticket continuará con prioridad media.",
        [emailField_("Resultado", outcome), emailField_("Revisó", approverLabel), emailField_("Estado", "Pendiente")],
        ""
      );
      MailApp.sendEmail(requesterEmail, subject, body, { htmlBody: htmlBody, name: "Ticketera · Data" });
    } catch (error) {
      notification = " La decisión quedó guardada, pero no se pudo notificar al solicitante.";
      console.error(error);
    }
  }

  return {
    ok: true,
    ticketId: ticketId,
    decision: decision,
    message: (decision === "approve" ? "Prioridad alta aprobada." : "Prioridad alta no aprobada; el ticket quedó en prioridad Media.") + notification
  };
}

function sendHighPriorityRequest_(ticketId, input, requesterEmail, token) {
  const appUrl = ScriptApp.getService().getUrl();
  const approveUrl = appUrl ? appUrl + "?token=" + encodeURIComponent(token) + "&decision=approve" : "";
  const rejectUrl = appUrl ? appUrl + "?token=" + encodeURIComponent(token) + "&decision=reject" : "";
  const body = [
    "Se requiere revisar una solicitud de prioridad alta.", "",
    "Ticket: " + ticketId, "Cliente: " + input.client, "País: " + input.country,
    "Solicitud: " + input.requestType, "Detalles: " + input.details,
    "Solicitante: " + input.applicant, "Correo solicitante: " + requesterEmail,
    "", "Aprobar: " + (approveUrl || "Despliega la aplicación web para habilitar este enlace."),
    "No aprobar: " + (rejectUrl || "Despliega la aplicación web para habilitar este enlace.")
  ].join("\n");
  const htmlBody = emailFrame_(
    "Revisión requerida", "Prioridad alta",
    "Hay una solicitud que necesita aprobación antes de iniciar.",
    [emailField_("Ticket", ticketId), emailField_("Cliente", input.client), emailField_("País", input.country),
      emailField_("Solicitud", input.requestType), emailField_("Solicitante", input.applicant),
      emailField_("Correo", requesterEmail), emailField_("Detalles", input.details)],
    '<a href="' + escapeHtml_(approveUrl || "#") + '" style="display:inline-block;padding:13px 18px;background:#143f32;color:#ffffff;text-decoration:none;font-weight:700;border-radius:5px;margin-right:8px">Aprobar prioridad alta</a>' +
    '<a href="' + escapeHtml_(rejectUrl || "#") + '" style="display:inline-block;padding:13px 18px;background:#ffffff;color:#8b342e;text-decoration:none;font-weight:700;border:1px solid #d9aaa6;border-radius:5px">No aprobar</a>'
  );
  MailApp.sendEmail(CONFIG.APPROVER_EMAILS.join(","), "Aprobación de prioridad alta · " + ticketId, body, { htmlBody: htmlBody, name: "Ticketera · Data" });
}

function sendTicketCreatedNotification_(ticketId, input, requesterEmail) {
  const appUrl = ScriptApp.getService().getUrl() || "";
  const body = [
    "Se registró un nuevo ticket.", "", "Ticket: " + ticketId,
    "Cliente: " + input.client, "País: " + input.country, "Solicitud: " + input.requestType,
    "Prioridad: " + input.priority, "Solicitante: " + input.applicant,
    "Correo solicitante: " + requesterEmail, "", "Detalles: " + input.details,
    appUrl ? "\nAbrir ticketera: " + appUrl : ""
  ].join("\n");
  const htmlBody = emailFrame_(
    "Nuevo ticket", "Solicitud " + ticketId,
    "Se registró una nueva solicitud para el equipo.",
    [emailField_("Cliente", input.client), emailField_("País", input.country),
      emailField_("Solicitud", input.requestType), emailField_("Prioridad", input.priority),
      emailField_("Solicitante", input.applicant), emailField_("Correo", requesterEmail),
      emailField_("Detalles", input.details)],
    appUrl ? '<a href="' + escapeHtml_(appUrl) + '" style="display:inline-block;padding:13px 20px;background:#143f32;color:#ffffff;text-decoration:none;font-weight:700;border-radius:5px">Abrir ticketera</a>' : ""
  );
  MailApp.sendEmail(CONFIG.RESPONSIBLE_EMAIL, "Nuevo ticket · " + ticketId, body, { htmlBody: htmlBody, name: "Ticketera · Data" });
}

function emailField_(label, value) {
  return '<tr><td style="padding:8px 0;color:#66776e;font-size:12px;vertical-align:top;width:120px">' + escapeHtml_(label) +
    '</td><td style="padding:8px 0;color:#182a22;font-size:13px;line-height:1.5;white-space:pre-wrap">' + escapeHtml_(value || "—") + '</td></tr>';
}

function emailFrame_(eyebrow, title, intro, rows, action) {
  return '<div style="margin:0;padding:28px 12px;background:#f3f7f4;font-family:Arial,sans-serif;color:#182a22">' +
    '<div style="max-width:600px;margin:0 auto;background:#ffffff;border:1px solid #dce6df;border-radius:8px;overflow:hidden">' +
    '<div style="padding:22px 26px;background:#143f32;color:#ffffff"><div style="font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#d8ebad">TICKETERA · DATA</div>' +
    '<div style="margin-top:10px;font-size:12px;color:#d8e6dc">' + escapeHtml_(eyebrow) + '</div><h1 style="margin:5px 0 0;font-size:22px;line-height:1.25">' + escapeHtml_(title) + '</h1></div>' +
    '<div style="padding:24px 26px"><p style="margin:0 0 18px;color:#52645a;font-size:14px;line-height:1.5">' + escapeHtml_(intro) + '</p>' +
    '<table role="presentation" style="width:100%;border-collapse:collapse">' + rows.join("") + '</table>' +
    (action ? '<div style="padding-top:22px">' + action + '</div>' : "") +
    '<p style="margin:24px 0 0;padding-top:14px;border-top:1px solid #e8eee9;color:#66776e;font-size:11px;line-height:1.5">Notificación automática de la ticketera del equipo de Data.</p></div></div></div>';
}

function escapeHtml_(value) {
  return String(value == null ? "" : value).replace(/[&<>"']/g, function(character) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character];
  });
}

function getLoginChallenge(emailInput) {
  const email = String(emailInput || "").trim().toLowerCase();
  const genericError = "Correo o contraseña incorrectos.";
  if (!email || email.length > 254) throw new Error(genericError);

  const cache = CacheService.getScriptCache();
  const failureKey = loginFailureKey_(email);
  const failures = Number(cache.get(failureKey) || 0);
  if (failures >= CONFIG.LOGIN_FAILURE_LIMIT) {
    throw new Error("Se alcanzó el límite de intentos. Vuelve a intentarlo en 15 minutos.");
  }

  const account = getAccountByEmail_(email);
  const challengeId = Utilities.getUuid() + Utilities.getUuid();
  const challenge = {
    email: email,
    verifier: account ? account.hash : CONFIG.DUMMY_PASSWORD_HASH
  };
  cache.put("tk_login_challenge_" + challengeId, JSON.stringify(challenge), 180);
  return {
    challengeId: challengeId,
    salt: account ? account.salt : CONFIG.DUMMY_PASSWORD_SALT,
    iterations: CONFIG.PASSWORD_HASH_ITERATIONS
  };
}

function loginWithChallenge(emailInput, challengeIdInput, proofInput) {
  const email = String(emailInput || "").trim().toLowerCase();
  const challengeId = String(challengeIdInput || "");
  const proof = String(proofInput || "").toLowerCase();
  const genericError = "Correo o contraseña incorrectos.";
  const cache = CacheService.getScriptCache();
  const failureKey = loginFailureKey_(email);
  const failures = Number(cache.get(failureKey) || 0);
  if (failures >= CONFIG.LOGIN_FAILURE_LIMIT) {
    throw new Error("Se alcanzó el límite de intentos. Vuelve a intentarlo en 15 minutos.");
  }

  const challengeKey = "tk_login_challenge_" + challengeId;
  const challengeText = challengeId.length <= 100 ? cache.get(challengeKey) : "";
  if (challengeText) cache.remove(challengeKey);
  let challenge = null;
  try { challenge = challengeText ? JSON.parse(challengeText) : null; } catch (error) {}

  const account = getAccountByEmail_(email);
  if (!challenge || challenge.email !== email || !/^[0-9a-f]{64}$/.test(proof)) {
    recordLoginFailure_(cache, failureKey);
    throw new Error(genericError);
  }
  const expectedProof = bytesToHex_(Utilities.computeHmacSha256Signature(
    Utilities.newBlob(challengeId).getBytes(),
    hexToBytes_(challenge.verifier)
  ));
  if (!account || !constantTimeEqual_(expectedProof, proof)) {
    recordLoginFailure_(cache, failureKey);
    throw new Error(genericError);
  }

  cache.remove(failureKey);
  return { authToken: createAppSession_(email), userEmail: email };
}

function recordLoginFailure_(cache, failureKey) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const currentFailures = Number(cache.get(failureKey) || 0);
    const safeFailures = Number.isFinite(currentFailures) && currentFailures > 0 ? currentFailures : 0;
    cache.put(failureKey, String(Math.min(safeFailures + 1, CONFIG.LOGIN_FAILURE_LIMIT)), CONFIG.LOGIN_LOCK_SECONDS);
  } finally {
    lock.releaseLock();
  }
}

function hexToBytes_(hex) {
  const value = String(hex || "");
  if (!/^(?:[0-9a-f]{2})+$/i.test(value)) throw new Error("La configuración de acceso no es válida.");
  const bytes = [];
  for (let index = 0; index < value.length; index += 2) {
    bytes.push(parseInt(value.substring(index, index + 2), 16));
  }
  return bytes;
}

function constantTimeEqual_(left, right) {
  const a = String(left || "");
  const b = String(right || "");
  let difference = a.length ^ b.length;
  const length = Math.max(a.length, b.length);
  for (let index = 0; index < length; index += 1) {
    difference |= (a.charCodeAt(index) || 0) ^ (b.charCodeAt(index) || 0);
  }
  return difference === 0;
}

function loginFailureKey_(email) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email, Utilities.Charset.UTF_8);
  return "tk_login_fail_" + bytesToHex_(digest);
}

function getAccountByEmail_(email) {
  const normalized = String(email || "").trim().toLowerCase();
  return CONFIG.ACCOUNTS.filter(function(account) { return account.email === normalized; })[0] || null;
}
function createAppSession_(email) {
  const secret = PropertiesService.getScriptProperties().getProperty("AUTH_SIGNING_SECRET");
  if (!secret) throw new Error("Ejecuta Configurar ticketera para inicializar la sesión segura.");
  const normalizedEmail = String(email).toLowerCase();
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    v: CONFIG.SESSION_VERSION,
    email: normalizedEmail,
    sub: normalizedEmail,
    iat: now,
    exp: now + CONFIG.SESSION_SECONDS,
    nonce: Utilities.getUuid()
  };
  const payloadPart = base64UrlEncodeBytes_(Utilities.newBlob(JSON.stringify(payload)).getBytes());
  const signature = Utilities.computeHmacSha256Signature(payloadPart, secret, Utilities.Charset.UTF_8);
  return payloadPart + "." + base64UrlEncodeBytes_(signature);
}

function requireAuthenticated_(authToken) {
  const token = String(authToken || "");
  if (token.length > 5000) throw new Error("Tu sesión venció. Vuelve a iniciar sesión.");
  const parts = token.split(".");
  const secret = PropertiesService.getScriptProperties().getProperty("AUTH_SIGNING_SECRET");
  if (parts.length !== 2 || !secret) {
    throw new Error("Tu sesión venció. Vuelve a iniciar sesión.");
  }
  const expected = base64UrlEncodeBytes_(Utilities.computeHmacSha256Signature(parts[0], secret, Utilities.Charset.UTF_8));
  if (!constantTimeEqual_(expected, parts[1])) throw new Error("Tu sesión venció. Vuelve a iniciar sesión.");

  let session;
  try {
    session = JSON.parse(base64UrlDecodeText_(parts[0]));
  } catch (error) {
    throw new Error("Tu sesión venció. Vuelve a iniciar sesión.");
  }
  const now = Math.floor(Date.now() / 1000);
  const expiresAt = Number(session.exp);
  const issuedAt = Number(session.iat);
  const email = String(session.email || "").toLowerCase();
  if (session.v !== CONFIG.SESSION_VERSION || !email || session.sub !== email || !getAccountByEmail_(email) ||
      !Number.isFinite(expiresAt) || !Number.isFinite(issuedAt) || expiresAt <= now || issuedAt > now + 60) {
    throw new Error("Tu sesión venció. Vuelve a iniciar sesión.");
  }
  return email;
}
function base64UrlEncodeBytes_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, "");
}

function base64UrlDecodeBytes_(value) {
  let normalized = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  while (normalized.length % 4) normalized += "=";
  return Utilities.base64Decode(normalized);
}

function base64UrlDecodeText_(value) {
  return Utilities.newBlob(base64UrlDecodeBytes_(value)).getDataAsString("UTF-8");
}

function bytesToHex_(bytes) {
  return bytes.map(function(byte) {
    return (Number(byte) & 255).toString(16).padStart(2, "0");
  }).join("");
}

function requireResponsible_(authToken) {
  const email = requireAuthenticated_(authToken);
  const account = getAccountByEmail_(email);
  if (!account.canResponsible || email !== CONFIG.RESPONSIBLE_EMAIL.toLowerCase()) {
    throw new Error("Esta acción está disponible solo para la persona responsable.");
  }
  return email;
}

function requireApprover_(authToken) {
  const email = requireAuthenticated_(authToken);
  const account = getAccountByEmail_(email);
  if (!account.canApprove || CONFIG.APPROVER_EMAILS.indexOf(email) < 0) {
    throw new Error("Esta acción está disponible solo para quien aprueba prioridades altas.");
  }
  return email;
}
function getSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    throw new Error("Primero ejecuta Configurar ticketera desde la hoja vinculada.");
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function getTicketSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error("No se encontró la pestaña " + CONFIG.SHEET_NAME + ". Ejecuta Configurar ticketera desde la hoja vinculada.");
  }
  const actualHeaders = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).getDisplayValues()[0];
  if (JSON.stringify(actualHeaders) !== JSON.stringify(CONFIG.HEADERS)) {
    throw new Error("Los encabezados de Tickets no coinciden con el esquema esperado. No se escribieron cambios.");
  }
  return sheet;
}

function getUploadFolder_() {
  const folderId = PropertiesService.getScriptProperties().getProperty("UPLOAD_FOLDER_ID");
  if (!folderId) {
    throw new Error("No se configuró la carpeta de adjuntos. Ejecuta Configurar ticketera desde la hoja vinculada.");
  }
  return DriveApp.getFolderById(folderId);
}

function readTicketRecords_() {
  const sheet = getTicketSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const count = lastRow - 1;
  const values = sheet.getRange(2, 1, count, CONFIG.HEADERS.length).getValues();
  const display = sheet.getRange(2, 1, count, CONFIG.HEADERS.length).getDisplayValues();
  const records = [];

  for (let index = 0; index < values.length; index += 1) {
    const raw = values[index];
    const shown = display[index];
    if (!shown[COL.ID]) {
      continue;
    }
    records.push({
      id: shown[COL.ID],
      requestedAt: shown[COL.REQUESTED_AT],
      client: shown[COL.CLIENT],
      country: shown[COL.COUNTRY],
      requestType: shown[COL.REQUEST_TYPE],
      details: shown[COL.DETAILS],
      applicant: shown[COL.APPLICANT],
      team: shown[COL.TEAM],
      priority: shown[COL.PRIORITY],
      status: shown[COL.STATUS],
      responsible: shown[COL.RESPONSIBLE],
      startedAt: shown[COL.STARTED_AT],
      finishedAt: shown[COL.FINISHED_AT],
      duration: shown[COL.DURATION_MINUTES],
      highPriorityApprover: shown[COL.HIGH_PRIORITY_APPROVER],
      comments: shown[COL.COMMENTS],
      attachments: parseAttachments_(shown[COL.ATTACHMENTS]),
      token: String(raw[COL.APPROVAL_TOKEN] || ""),
      requesterEmail: shown[COL.REQUESTER_EMAIL]
    });
  }
  return records;
}

function toPublicRecord_(record) {
  return {
    id: record.id,
    requestedAt: record.requestedAt,
    client: record.client,
    country: record.country,
    requestType: record.requestType,
    details: record.details,
    applicant: record.applicant,
    team: record.team,
    priority: record.priority,
    status: record.status,
    finishedAt: record.finishedAt,
    comments: record.comments,
    attachments: record.attachments
  };
}

function toManagerRecord_(record) {
  return {
    id: record.id,
    requestedAt: record.requestedAt,
    client: record.client,
    country: record.country,
    requestType: record.requestType,
    details: record.details,
    applicant: record.applicant,
    team: record.team,
    priority: record.priority,
    status: record.status,
    responsible: record.responsible,
    startedAt: record.startedAt,
    finishedAt: record.finishedAt,
    duration: record.duration,
    highPriorityApprover: record.highPriorityApprover,
    comments: record.comments,
    attachments: record.attachments,
    token: record.token,
    requesterEmail: record.requesterEmail
  };
}

function toApprovalRecord_(record) {
  return {
    id: record.id,
    requestedAt: record.requestedAt,
    client: record.client,
    country: record.country,
    requestType: record.requestType,
    details: record.details,
    applicant: record.applicant,
    requesterEmail: record.requesterEmail,
    comments: record.comments,
    attachments: record.attachments
  };
}

function parseAttachments_(cell) {
  return String(cell || "").split("\n").filter(function(line) {
    return line.trim() !== "";
  }).map(function(line) {
    const separator = line.lastIndexOf(" — ");
    if (separator < 0) {
      return { name: line, fileId: "" };
    }
    const reference = line.substring(separator + 3);
    const queryMatch = reference.match(/[?&]id=([^&]+)/);
    const pathMatch = reference.match(/\/d\/([A-Za-z0-9_-]+)/);
    let fileId = queryMatch ? queryMatch[1] : pathMatch ? pathMatch[1] : reference;
    try { fileId = decodeURIComponent(fileId); } catch (error) {}
    return {
      name: line.substring(0, separator),
      fileId: /^[A-Za-z0-9_-]{10,}$/.test(fileId) ? fileId : ""
    };
  });
}

function validateTicketInput_(form) {
  const client = cleanText_(form.client, 120, "Cliente");
  const details = cleanText_(form.details, 5000, "Detalles");
  const country = String(form.country || "");
  const requestType = String(form.requestType || "");
  const applicant = String(form.applicant || "");
  const priority = String(form.priority || "");

  if (!client) throw new Error("Escribe el nombre del cliente.");
  if (!details) throw new Error("Describe la solicitud.");
  if (CONFIG.COUNTRIES.indexOf(country) < 0) throw new Error("Selecciona un país válido.");
  if (CONFIG.REQUEST_TYPES.indexOf(requestType) < 0) throw new Error("Selecciona un tipo de solicitud válido.");
  if (applicantNames_().indexOf(applicant) < 0) throw new Error("Selecciona un solicitante válido.");
  if (CONFIG.PRIORITIES.indexOf(priority) < 0) throw new Error("Selecciona una prioridad válida.");

  return {
    client: client,
    details: details,
    country: country,
    requestType: requestType,
    applicant: applicant,
    priority: priority
  };
}

function validateAttachments_(attachments) {
  const files = Array.isArray(attachments) ? attachments : [];
  if (files.length > CONFIG.ATTACHMENT_COUNT_LIMIT) {
    throw new Error("Puedes adjuntar hasta " + CONFIG.ATTACHMENT_COUNT_LIMIT + " archivos.");
  }

  let totalSize = 0;
  const prepared = [];
  files.forEach(function(file) {
    const name = safeFileName_(file && file.name);
    const mimeType = String((file && file.mimeType) || "application/octet-stream").slice(0, 120);
    const data = String((file && file.data) || "");
    if (!name || !data || data.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(data)) {
      throw new Error("Uno de los archivos adjuntos no se pudo leer.");
    }
    const padding = (data.match(/=*$/) || [""])[0].length;
    const size = Math.floor(data.length * 3 / 4) - padding;
    if (size > CONFIG.ATTACHMENT_SIZE_LIMIT) {
      throw new Error(name + " supera el límite de 5 MB.");
    }
    totalSize += size;
    if (totalSize > CONFIG.ATTACHMENT_TOTAL_LIMIT) {
      throw new Error("El tamaño combinado de los adjuntos supera 12 MB.");
    }
    prepared.push({ name: name, mimeType: mimeType, data: data, size: size });
  });
  return prepared;
}

function safeFileName_(name) {
  return String(name || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/[\r\n\t]/g, " ")
    .trim()
    .slice(0, 180);
}

function cleanText_(value, maxLength, label) {
  const text = String(value == null ? "" : value).replace(/\u0000/g, "").trim();
  if (text.length > maxLength) {
    throw new Error(label + " supera el máximo de " + maxLength + " caracteres.");
  }
  return text;
}

function safeCellText_(value) {
  const text = String(value == null ? "" : value);
  return /^[\s]*[=+\-@]/.test(text) ? "'" + text : text;
}

function applicantNames_() {
  return CONFIG.APPLICANTS.map(function(applicant) { return applicant.name; });
}

function getTeamForApplicant_(applicant) {
  const match = CONFIG.APPLICANTS.filter(function(item) { return item.name === applicant; })[0];
  return match ? match.team : "Sin equipo";
}

function nextTicketId_(sheet) {
  const lastRow = sheet.getLastRow();
  let maxNumber = 0;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, COL.ID + 1, lastRow - 1, 1).getDisplayValues();
    ids.forEach(function(row) {
      const match = String(row[0] || "").match(/^TKT-(\d+)$/);
      if (match) {
        maxNumber = Math.max(maxNumber, Number(match[1]));
      }
    });
  }
  return "TKT-" + String(maxNumber + 1).padStart(6, "0");
}

function findTicketRowById_(sheet, ticketId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No se encontraron tickets.");
  const ids = sheet.getRange(2, COL.ID + 1, lastRow - 1, 1).getDisplayValues();
  for (let index = 0; index < ids.length; index += 1) {
    if (String(ids[index][0]) === String(ticketId)) {
      return index + 2;
    }
  }
  throw new Error("No se encontró el ticket " + ticketId + ".");
}

function findTicketRowByToken_(sheet, token) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No hay solicitudes pendientes de aprobación.");
  const tokens = sheet.getRange(2, COL.APPROVAL_TOKEN + 1, lastRow - 1, 1).getValues();
  for (let index = 0; index < tokens.length; index += 1) {
    if (String(tokens[index][0] || "") === token) {
      return index + 2;
    }
  }
  throw new Error("El enlace de aprobación ya no es válido.");
}

function appendCommentToRow_(sheet, rowNumber, message) {
  const oldComment = String(sheet.getRange(rowNumber, COL.COMMENTS + 1).getDisplayValue() || "");
  const entry = formatTimestamp_(new Date()) + " · " + message;
  const updated = oldComment ? oldComment + "\n" + entry : entry;
  sheet.getRange(rowNumber, COL.COMMENTS + 1).setValue(safeCellText_(updated));
}

function formatTimestamp_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone() || "America/Bogota", "dd/MM/yyyy HH:mm");
}

function setListValidation_(sheet, column, values, rowCount) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, column, rowCount, 1).setDataValidation(rule);
}
