/*********************************
 * NEXUS | Code.gs
 *********************************/
const CFG = {
  SHEET_NAME: "nexus",
  UPDATES_SHEET_NAME: "nexus_updates",
  AUDIT_SHEET_NAME: "nexus_audit",
  AUDIT_STRICT: false,
  FOLDER_ID: "YOUR_ATTACHMENT_FOLDER_ID",

  ID_PREFIX: "NXS",
  ID_PAD: 7,
  MAX_FILES: 5,

  MAX_FILE_BYTES: 10 * 1024 * 1024,

  DT_FMT: "yyyy-MM-dd HH:mm:ss",
  DURATION_FMT: "HH:mm:ss",
  EMAIL_DT_FMT: "MMMM d, yyyy h:mm a",

  EMAIL_FROM_ALIAS: "nexus@internal.example.com",
  EMAIL_FROM_NAME: "NEXUS",
  DEPARTMENT_DLS: {
    "ACCOUNTING": "accounting@internal.example.com",
    "ADMINISTRATIVE": "administrative@internal.example.com",
    "BUSINESS SYSTEMS": "systems@internal.example.com",
    "OPERATIONS": "support@operations.example.com"
  }
};


const ALLOWED_DOMAINS = ["internal.example.com", "operations.example.com"];


// assignees by department
const TEAM_ASSIGNEES = {
  "ACCOUNTING": [
    "alex.morgan@internal.example.com",
    "jamie.lee@internal.example.com",
    "morgan.reed@internal.example.com",
    "natalie.brooks@internal.example.com",
    "ryan.cooper@internal.example.com"
  ],


  "BUSINESS SYSTEMS": [
    "sam.rivera@internal.example.com",
    "taylor.chen@internal.example.com"
  ],


  "ADMINISTRATIVE": [
    "casey.brooks@internal.example.com",
    "drew.miller@internal.example.com",
    "jordan.hayes@internal.example.com",
    "riley.evans@internal.example.com",
    "avery.turner@internal.example.com",
    "cameron.gray@internal.example.com",
    "harper.king@internal.example.com",
    "quinn.ward@internal.example.com"
  ],


  "OPERATIONS": [
    "bailey.parker@operations.example.com",
    "devon.blake@operations.example.com",
    "emerson.carter@operations.example.com",
    "finley.moore@operations.example.com",
    "hayden.scott@operations.example.com",
    "kendall.young@operations.example.com",
    "reese.hall@operations.example.com"
  ]
};


// super admins
const SUPER_ADMINS = [
  "admin@internal.example.com",
  "sam.rivera@internal.example.com",
  "taylor.chen@internal.example.com"
].map(e => String(e).toLowerCase().trim());


// all assignees for dropdowns
function getAllAssignees_() {
  const set = new Set();
  Object.keys(TEAM_ASSIGNEES).forEach(team => {
    TEAM_ASSIGNEES[team].forEach(e => set.add(String(e).toLowerCase().trim()));
  });
  return Array.from(set).sort();
}

/*********************************
 * EMAIL NOTIFICATIONS
 * - Uses nexus + nexus_updates only
 * - No audit dependency
 *********************************/
function getDepartmentDl_(department) {
  return CFG.DEPARTMENT_DLS[String(department || "").trim()] || "";
}

function getTicketRecordById_(ticketIdRaw) {
  const ticketId = String(ticketIdRaw || "").trim();
  if (!ticketId) throw new Error("Missing ticketId.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sheet) throw new Error("Missing sheet tab: " + CFG.SHEET_NAME);

  const map = getHeaderMap_(sheet);
  const row = findRowByTicketId_(sheet, ticketId);
  if (row === -1) throw new Error("Ticket not found: " + ticketId);

  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  return {
    sheet,
    map,
    row,
    values,
    ticket: rowToTicketObject_(values, map),
    raw: function (name) {
      const col = map[name];
      if (!col) return "";
      return values[col - 1];
    }
  };
}

function getUpdateRecordByRow_(updateRowIndexRaw) {
  const updateRowIndex = parseInt(updateRowIndexRaw, 10);
  if (!updateRowIndex || updateRowIndex < 2) throw new Error("Invalid update row.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
  if (!sh) throw new Error("Missing sheet tab: " + CFG.UPDATES_SHEET_NAME);

  const lastRow = sh.getLastRow();
  if (updateRowIndex > lastRow) throw new Error("Update row not found.");

  const row = sh.getRange(updateRowIndex, 1, 1, 9).getValues()[0];

  return {
    rowIndex: updateRowIndex,
    update_timestamp_raw: row[0],
    update_ticket_id: String(row[1] || "").trim(),
    update_author_email: String(row[2] || "").toLowerCase().trim(),
    update_message: String(row[3] || ""),
    update_att_1: String(row[4] || ""),
    update_att_2: String(row[5] || ""),
    update_att_3: String(row[6] || ""),
    update_att_4: String(row[7] || ""),
    update_att_5: String(row[8] || "")
  };
}

function emailFormatDateTime_(val) {
  if (!val || val === "") return "—";

  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), CFG.EMAIL_DT_FMT);
  }

  const d = new Date(val);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), CFG.EMAIL_DT_FMT);
  }

  return String(val);
}

function emailEscape_(s) {
  return String(s || "").replace(/[&<>"']/g, function (c) {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "\"": "&quot;",
      "'": "&#039;"
    }[c];
  });
}

function emailNl2br_(s) {
  return emailEscape_(s).replace(/\n/g, "<br>");
}

function emailValue_(v, emptyLabel) {
  const out = String(v || "").trim();
  return out || (emptyLabel || "—");
}

function collectTicketAttachmentLinks_(ticket) {
  const out = [];
  for (let i = 1; i <= CFG.MAX_FILES; i++) {
    const url = String(ticket["attachment_" + i] || "").trim();
    if (url) out.push({ label: "Attachment " + i, url: url });
  }
  return out;
}

function collectUpdateAttachmentLinks_(updateObj) {
  const out = [];
  for (let i = 1; i <= CFG.MAX_FILES; i++) {
    const url = String(updateObj["update_att_" + i] || "").trim();
    if (url) out.push({ label: "Attachment " + i, url: url });
  }
  return out;
}

function renderLinkListHtml_(links) {
  if (!links || !links.length) {
    return `<div style="margin:6px 0 0 16px; color:#6b7280;">None</div>`;
  }

  return links.map(function (x) {
    return `<div style="margin:6px 0 0 16px;">• <a href="${emailEscape_(x.url)}" target="_blank">${emailEscape_(x.label)}</a></div>`;
  }).join("");
}

function renderLinkListText_(links) {
  if (!links || !links.length) return "None";
  return links.map(function (x) {
    return "  - " + x.label + ": " + x.url;
  }).join("\n");
}

function buildEmailShell_(title, innerHtml) {
  return `
    <!doctype html>
    <html>
      <body style="margin:0; padding:24px; background:#f3f4f6; font-family:Arial,Helvetica,sans-serif; color:#111827;">
        <div style="max-width:720px; margin:0 auto; background:#ffffff; border:1px solid #e5e7eb; border-radius:16px; overflow:hidden;">
          <div style="padding:20px 24px; background:#312e81; color:#ffffff;">
            <div style="font-size:28px; font-weight:800; letter-spacing:.3px;">NEXUS</div>
            <div style="margin-top:6px; font-size:14px; opacity:.92;">${emailEscape_(title)}</div>
          </div>

          <div style="padding:24px; font-size:14px; line-height:1.65;">
            ${innerHtml}
          </div>

          <div style="padding:14px 24px; background:#f9fafb; border-top:1px solid #e5e7eb; font-size:12px; color:#6b7280; text-align:center;">
            NEXUS | Internal Support Platform
          </div>
        </div>
      </body>
    </html>
  `;
}

function getTicketRecipients_(ticket) {
  const to = String(ticket.email_address || "").trim();
  const cc = getDepartmentDl_(ticket.department);

  if (!to) throw new Error("Ticket requester email is blank.");
  if (!cc) throw new Error("No department DL configured for: " + ticket.department);

  return { to, cc };
}

function sendNexusEmail_(to, cc, subject, htmlBody, plainBody) {
  const wanted = String(CFG.EMAIL_FROM_ALIAS || "").toLowerCase().trim();
  const aliases = GmailApp.getAliases().map(function (a) {
    return String(a || "").toLowerCase().trim();
  });

  if (!aliases.includes(wanted)) {
    throw new Error("Configured alias is not available in Gmail: " + CFG.EMAIL_FROM_ALIAS);
  }

  GmailApp.sendEmail(to, subject, plainBody, {
    htmlBody: htmlBody,
    name: CFG.EMAIL_FROM_NAME,
    from: CFG.EMAIL_FROM_ALIAS,
    cc: cc || ""
  });
}

function sendOnceWithProperty_(propertyKey, senderFn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(propertyKey) === "1") {
      return { sent: false, skipped: true };
    }

    senderFn();
    props.setProperty(propertyKey, "1");
    return { sent: true, skipped: false };

  } finally {
    lock.releaseLock();
  }
}

function buildSubmissionEmail_(record) {
  const ticket = record.ticket;
  const recipients = getTicketRecipients_(ticket);
  const links = collectTicketAttachmentLinks_(ticket);

  const subject = `[NEXUS] Ticket Received — ${ticket.ticket_id}`;

  const html = buildEmailShell_("Ticket Received", `
    <p style="margin:0 0 16px;">Hi,</p>

    <p style="margin:0 0 16px;">
      Your request has been successfully submitted to <strong>NEXUS</strong>.
    </p>

    <div style="margin:0 0 16px;">
      <div><strong>Ticket ID:</strong> ${emailEscape_(ticket.ticket_id)}</div>
      <div><strong>Date Submitted:</strong> ${emailEscape_(emailFormatDateTime_(record.raw("timestamp")))}</div>
      <div><strong>Department:</strong> ${emailEscape_(ticket.department)}</div>
      <div><strong>Category:</strong> ${emailEscape_(ticket.category)}</div>
      <div><strong>Priority:</strong> ${emailEscape_(ticket.priority)}</div>
      <div><strong>Status:</strong> ${emailEscape_(ticket.status)}</div>
    </div>

    <div style="margin:0 0 8px;"><strong>Summary:</strong></div>
    <div style="margin:0 0 16px;">${emailEscape_(ticket.summary)}</div>

    <div style="margin:0 0 8px;"><strong>Description:</strong></div>
    <div style="margin:0 0 16px; padding:12px; background:#f9fafb; border:1px solid #e5e7eb; border-radius:10px;">
      ${emailNl2br_(ticket.description)}
    </div>

    <div style="margin:0 0 8px;"><strong>Attachments:</strong></div>
    ${renderLinkListHtml_(links)}

    <p style="margin:16px 0 0;">
      You may monitor updates and add comments through <strong>NEXUS</strong>.
    </p>

    <p style="margin:20px 0 0;">
      Regards,<br>
      <strong>NEXUS</strong><br>
      nexus@internal.example.com
    </p>
  `);

  const text = [
    "Hi,",
    "",
    "Your request has been successfully submitted to NEXUS.",
    "",
    "Ticket ID: " + ticket.ticket_id,
    "Date Submitted: " + emailFormatDateTime_(record.raw("timestamp")),
    "Department: " + ticket.department,
    "Category: " + ticket.category,
    "Priority: " + ticket.priority,
    "Status: " + ticket.status,
    "",
    "Summary:",
    ticket.summary || "",
    "",
    "Description:",
    ticket.description || "",
    "",
    "Attachments:",
    renderLinkListText_(links),
    "",
    "You may monitor updates and add comments through NEXUS.",
    "",
    "Regards,",
    "NEXUS",
    "nexus@internal.example.com"
  ].join("\n");

  return {
    to: recipients.to,
    cc: recipients.cc,
    subject: subject,
    html: html,
    text: text
  };
}

function buildFieldChangeEmail_(record, changes) {
  const ticket = record.ticket;
  const recipients = getTicketRecipients_(ticket);
  const updatedOn = emailFormatDateTime_(record.raw("last_updated") || record.raw("timestamp"));

  const changesHtml = (changes || []).map(function (c) {
    return `<div style="margin:6px 0 0 16px;">• <strong>${emailEscape_(c.label)}:</strong> ${emailEscape_(c.oldValue)} → ${emailEscape_(c.newValue)}</div>`;
  }).join("");

  const changesText = (changes || []).map(function (c) {
    return "  - " + c.label + ": " + c.oldValue + " -> " + c.newValue;
  }).join("\n");

  const subject = `[NEXUS] Ticket Updated — ${ticket.ticket_id}`;

  const html = buildEmailShell_("Ticket Updated", `
    <p style="margin:0 0 16px;">Hi,</p>

    <p style="margin:0 0 16px;">
      There is a new update on your <strong>NEXUS</strong> ticket.
    </p>

    <div style="margin:0 0 16px;">
      <div><strong>Ticket ID:</strong> ${emailEscape_(ticket.ticket_id)}</div>
      <div><strong>Summary:</strong> ${emailEscape_(ticket.summary)}</div>
      <div><strong>Latest Status:</strong> ${emailEscape_(ticket.status)}</div>
      <div><strong>Assigned To:</strong> ${emailEscape_(emailValue_(ticket.assigned, "Unassigned"))}</div>
      <div><strong>Updated On:</strong> ${emailEscape_(updatedOn)}</div>
    </div>

    <div style="margin:0 0 8px;"><strong>What Changed:</strong></div>
    ${changesHtml}

    <p style="margin:16px 0 0;">
      You may reply through <strong>NEXUS</strong> if you need to add more details.
    </p>

    <p style="margin:20px 0 0;">
      Regards,<br>
      <strong>NEXUS</strong><br>
      nexus@internal.example.com
    </p>
  `);

  const text = [
    "Hi,",
    "",
    "There is a new update on your NEXUS ticket.",
    "",
    "Ticket ID: " + ticket.ticket_id,
    "Summary: " + ticket.summary,
    "Latest Status: " + ticket.status,
    "Assigned To: " + emailValue_(ticket.assigned, "Unassigned"),
    "Updated On: " + updatedOn,
    "",
    "What Changed:",
    changesText || "None",
    "",
    "You may reply through NEXUS if you need to add more details.",
    "",
    "Regards,",
    "NEXUS",
    "nexus@internal.example.com"
  ].join("\n");

  return {
    to: recipients.to,
    cc: recipients.cc,
    subject: subject,
    html: html,
    text: text
  };
}

function buildThreadUpdateEmail_(record, updateObj) {
  const ticket = record.ticket;
  const recipients = getTicketRecipients_(ticket);
  const links = collectUpdateAttachmentLinks_(updateObj);
  const updatedOn = emailFormatDateTime_(updateObj.update_timestamp_raw);

  const subject = `[NEXUS] Ticket Updated — ${ticket.ticket_id}`;

  const html = buildEmailShell_("Ticket Updated", `
    <p style="margin:0 0 16px;">Hi,</p>

    <p style="margin:0 0 16px;">
      There is a new update on your <strong>NEXUS</strong> ticket.
    </p>

    <div style="margin:0 0 16px;">
      <div><strong>Ticket ID:</strong> ${emailEscape_(ticket.ticket_id)}</div>
      <div><strong>Summary:</strong> ${emailEscape_(ticket.summary)}</div>
      <div><strong>Latest Status:</strong> ${emailEscape_(ticket.status)}</div>
      <div><strong>Assigned To:</strong> ${emailEscape_(emailValue_(ticket.assigned, "Unassigned"))}</div>
      <div><strong>Updated On:</strong> ${emailEscape_(updatedOn)}</div>
      <div><strong>Updated By:</strong> ${emailEscape_(updateObj.update_author_email || "—")}</div>
    </div>

    <div style="margin:0 0 8px;"><strong>New Update:</strong></div>
    <div style="margin:0 0 16px; padding:12px; background:#f9fafb; border:1px solid #e5e7eb; border-radius:10px;">
      ${emailNl2br_(updateObj.update_message)}
    </div>

    <div style="margin:0 0 8px;"><strong>Update Attachments:</strong></div>
    ${renderLinkListHtml_(links)}

    <p style="margin:16px 0 0;">
      You may reply through <strong>NEXUS</strong> if you need to add more details.
    </p>

    <p style="margin:20px 0 0;">
      Regards,<br>
      <strong>NEXUS</strong><br>
      nexus@internal.example.com
    </p>
  `);

  const text = [
    "Hi,",
    "",
    "There is a new update on your NEXUS ticket.",
    "",
    "Ticket ID: " + ticket.ticket_id,
    "Summary: " + ticket.summary,
    "Latest Status: " + ticket.status,
    "Assigned To: " + emailValue_(ticket.assigned, "Unassigned"),
    "Updated On: " + updatedOn,
    "Updated By: " + (updateObj.update_author_email || "—"),
    "",
    "New Update:",
    updateObj.update_message || "",
    "",
    "Update Attachments:",
    renderLinkListText_(links),
    "",
    "You may reply through NEXUS if you need to add more details.",
    "",
    "Regards,",
    "NEXUS",
    "nexus@internal.example.com"
  ].join("\n");

  return {
    to: recipients.to,
    cc: recipients.cc,
    subject: subject,
    html: html,
    text: text
  };
}

function buildResolvedEmail_(record) {
  const ticket = record.ticket;
  const recipients = getTicketRecipients_(ticket);

  const resolvedOn = emailFormatDateTime_(record.raw("time_resolved") || record.raw("last_updated"));
  const handledBy = emailValue_(ticket.updated_by, "—");
  const resolutionText = emailValue_(ticket.update_latest, "No additional resolution note was provided.");

  const subject = `[NEXUS] Ticket Resolved — ${ticket.ticket_id}`;

  const html = buildEmailShell_("Ticket Resolved", `
    <p style="margin:0 0 16px;">Hi,</p>

    <p style="margin:0 0 16px;">
      Your <strong>NEXUS</strong> ticket has been marked as <strong>Resolved / Closed</strong>.
    </p>

    <div style="margin:0 0 16px;">
      <div><strong>Ticket ID:</strong> ${emailEscape_(ticket.ticket_id)}</div>
      <div><strong>Summary:</strong> ${emailEscape_(ticket.summary)}</div>
      <div><strong>Final Status:</strong> ${emailEscape_(ticket.status)}</div>
      <div><strong>Resolved On:</strong> ${emailEscape_(resolvedOn)}</div>
      <div><strong>Handled By:</strong> ${emailEscape_(handledBy)}</div>
    </div>

    <div style="margin:0 0 8px;"><strong>Resolution Details:</strong></div>
    <div style="margin:0 0 16px; padding:12px; background:#f9fafb; border:1px solid #e5e7eb; border-radius:10px;">
      ${emailNl2br_(resolutionText)}
    </div>

    <p style="margin:16px 0 0;">
      If the issue persists or recurs, please create a new ticket or post a follow-up through <strong>NEXUS</strong>.
    </p>

    <p style="margin:20px 0 0;">
      Regards,<br>
      <strong>NEXUS</strong><br>
      nexus@internal.example.com
    </p>
  `);

  const text = [
    "Hi,",
    "",
    "Your NEXUS ticket has been marked as Resolved / Closed.",
    "",
    "Ticket ID: " + ticket.ticket_id,
    "Summary: " + ticket.summary,
    "Final Status: " + ticket.status,
    "Resolved On: " + resolvedOn,
    "Handled By: " + handledBy,
    "",
    "Resolution Details:",
    resolutionText,
    "",
    "If the issue persists or recurs, please create a new ticket or post a follow-up through NEXUS.",
    "",
    "Regards,",
    "NEXUS",
    "nexus@internal.example.com"
  ].join("\n");

  return {
    to: recipients.to,
    cc: recipients.cc,
    subject: subject,
    html: html,
    text: text
  };
}

function finalizeSubmissionNotification(ticketIdRaw) {
  const ticketId = String(ticketIdRaw || "").trim();
  if (!ticketId) return { success: false, error: "Missing ticketId." };

  try {
    const result = sendOnceWithProperty_("NXS_SUBMIT_MAIL_SENT_" + ticketId, function () {
      const record = getTicketRecordById_(ticketId);
      const mail = buildSubmissionEmail_(record);
      sendNexusEmail_(mail.to, mail.cc, mail.subject, mail.html, mail.text);
    });

    return {
      success: true,
      sent: result.sent,
      skipped: result.skipped
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function finalizeThreadUpdateNotification(updateRowIndexRaw) {
  const updateRowIndex = parseInt(updateRowIndexRaw, 10);
  if (!updateRowIndex || updateRowIndex < 2) {
    return { success: false, error: "Invalid update row." };
  }

  try {
    const updateObj = getUpdateRecordByRow_(updateRowIndex);
    const propertyKey = buildThreadUpdateEmailPropertyKey_(updateObj);

    const result = sendOnceWithProperty_(propertyKey, function () {
      const record = getTicketRecordById_(updateObj.update_ticket_id);
      const mail = buildThreadUpdateEmail_(record, updateObj);
      sendNexusEmail_(mail.to, mail.cc, mail.subject, mail.html, mail.text);
    });

    return {
      success: true,
      sent: result.sent,
      skipped: result.skipped,
      propertyKey: propertyKey
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function buildThreadUpdateEmailPropertyKey_(updateObj) {
  const rawKey = [
    normalizeThreadUpdateEmailKeyPart_(updateObj.update_timestamp_raw),
    normalizeThreadUpdateEmailKeyPart_(updateObj.update_ticket_id),
    normalizeThreadUpdateEmailKeyPart_(updateObj.update_author_email),
    normalizeThreadUpdateEmailKeyPart_(updateObj.update_message)
  ].join("♦");

  return "NXS_THREAD_MAIL_SENT_V2_" + sha256Hex_(rawKey);
}

function normalizeThreadUpdateEmailKeyPart_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return String(value.getTime());
  }

  return String(value ?? "").trim();
}

function sha256Hex_(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || ""),
    Utilities.Charset.UTF_8
  );

  return bytes
    .map(function (b) {
      const v = (b < 0) ? b + 256 : b;
      return v.toString(16).padStart(2, "0");
    })
    .join("");
}

function sendFieldChangeNotification_(ticketIdRaw, changes) {
  try {
    const record = getTicketRecordById_(ticketIdRaw);
    const mail = buildFieldChangeEmail_(record, changes || []);
    sendNexusEmail_(mail.to, mail.cc, mail.subject, mail.html, mail.text);
    return { success: true };
  } catch (e) {
    Logger.log("sendFieldChangeNotification_ error: " + e);
    return { success: false, error: e.toString() };
  }
}

function sendResolvedNotification_(ticketIdRaw) {
  try {
    const record = getTicketRecordById_(ticketIdRaw);
    const mail = buildResolvedEmail_(record);
    sendNexusEmail_(mail.to, mail.cc, mail.subject, mail.html, mail.text);
    return { success: true };
  } catch (e) {
    Logger.log("sendResolvedNotification_ error: " + e);
    return { success: false, error: e.toString() };
  }
}


/*********************************
 * AUDIT TRAIL (QMS)
 * - Append-only audit events
 *********************************/
function ensureAuditSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.AUDIT_SHEET_NAME);


  if (!sh) {
    sh = ss.insertSheet(CFG.AUDIT_SHEET_NAME);
    sh.appendRow([
      "audit_ts",
      "ticket_id",
      "actor_email",
      "actor_type",
      "source_view",
      "event_type",
      "field_name",
      "old_value",
      "new_value",
      "update_row_index",
      "attachment_slot",
      "file_url",
      "meta_json"
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}


function safeStr_(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date && !isNaN(v)) return formatDateTime_(v);
  if (typeof v === "object") {
    try { return JSON.stringify(v); } catch (e) { return String(v); }
  }
  return String(v);
}


function actorType_() {
  const emailLower = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
  if (isSuperAdmin_(emailLower)) return "SUPER_ADMIN";
  if (isAdmin_()) return "ADMIN";
  return "USER";
}


function logAudit_(ticketId, sourceView, eventType, fieldName, oldValue, newValue, metaObj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);


  try {
    const sh = ensureAuditSheet_();


    const actor = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
    const actorType = actorType_();


    let metaEnriched = {};
    try {
      metaEnriched = metaObj ? JSON.parse(safeStr_(metaObj) || "{}") : {};
    } catch (e) {
      metaEnriched = {};
    }




    try {
      if (!metaEnriched.department && ticketId) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const main = ss.getSheetByName(CFG.SHEET_NAME);
        if (main) {
          const map = getHeaderMap_(main);
          const row = findRowByTicketId_(main, String(ticketId).trim());
          if (row !== -1 && map["department"]) {
            metaEnriched.department = String(main.getRange(row, map["department"]).getValue() || "").trim();
          }
        }
      }
    } catch (e) {
      // audit is best-effort
    }


    const metaJson = safeStr_(metaEnriched);


    const metaCapped = metaJson.length > 45000 ? metaJson.slice(0, 45000) : metaJson;


    const row = [
      new Date(),
      safeStr_(ticketId),
      actor,
      actorType,
      safeStr_(sourceView),
      safeStr_(eventType),
      safeStr_(fieldName),
      safeStr_(oldValue),
      safeStr_(newValue),
      safeStr_(metaObj?.updateRowIndex || ""),
      safeStr_(metaObj?.slot || ""),
      safeStr_(metaObj?.fileUrl || ""),
      metaCapped
    ];


    sh.appendRow(row);
    return true;


  } catch (e) {
    if (CFG.AUDIT_STRICT) throw e;
    return false;
  } finally {
    lock.releaseLock();
  }
}




/*********************************
 * Routing
 *********************************/
function doGet(e) {
  var view = e?.parameter?.view;


  if (view === "admin") {
    // Admin page: only team members + super admins
    if (!isAdmin_()) {
      return HtmlService.createHtmlOutput(
        `<div style="font-family:Arial;padding:24px;">
          <h2>Access denied</h2>
          <p>This page is restricted.</p>
          <p><b>Signed in as:</b> ${sanitize_(Session.getActiveUser().getEmail())}</p>
        </div>`
      ).setTitle("NEXUS | Admin");
    }


    return HtmlService.createTemplateFromFile("admin")
      .evaluate()
      .setTitle("NEXUS | Admin")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }


  if (view === "my") {
    // My tickets page: any allowed domain user can open, but data is still enforced server-side
    const email = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
    const domain = email.split("@").pop();


    if (!email || !ALLOWED_DOMAINS.includes(domain)) {
      return HtmlService.createHtmlOutput(
        `<div style="font-family:Arial;padding:24px;">
          <h2>Access denied</h2>
          <p>This page is restricted.</p>
          <p><b>Signed in as:</b> ${sanitize_(Session.getActiveUser().getEmail())}</p>
        </div>`
      ).setTitle("NEXUS | My Tickets");
    }


    return HtmlService.createTemplateFromFile("myticketviewer")
      .evaluate()
      .setTitle("NEXUS | My Tickets")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }


  if (view === "audit") {
    enforceSuperAdmin_();


    return HtmlService.createTemplateFromFile("audit")
      .evaluate()
      .setTitle("NEXUS | Audit")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (view === "metrics") {
    // Metrics page: same access as Admin
    enforceAdmin_();

    return HtmlService.createTemplateFromFile("metrics")
      .evaluate()
      .setTitle("NEXUS | Metrics")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }


  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("NEXUS | Ticket Portal")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}




/*********************************
 * USER PORTAL
 *********************************/
function createTicket(formData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);


  try {
    var sheet = getNexusSheet_();


    if (!formData || !formData.summary || !formData.description || !formData.priority || !formData.category) {
      return { success: false, error: "Missing required fields." };
    }


    var nextId = generateNextTicketId_(sheet);
    var department = parseDepartment_(formData.category);


    var timestamp = new Date();


    var email = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
    var domain = email.split("@").pop();
    if (!ALLOWED_DOMAINS.includes(domain)) {
      return { success: false, error: "Access denied." };
    }




    var rowData = [
      nextId,
      timestamp,
      email,
      department,
      formData.category || "",
      formData.priority || "",
      "Open",
      formData.summary || "",
      formData.description || "",
      "", "", "", "", ""
    ];


    sheet.appendRow(rowData);


    logAudit_(nextId, "portal", "CREATE_TICKET", "", "", "", {
      department: department,
      category: formData.category || "",
      priority: formData.priority || "",
      summary: (formData.summary || "").slice(0, 200)
    });


    return { success: true, ticketId: nextId };


  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}


function uploadSingleFile(fileObj, ticketId, slotIndex) {
  try {
    if (!ticketId) return { success: false, error: "Missing ticketId." };
    if (slotIndex === null || slotIndex === undefined) return { success: false, error: "Missing slotIndex." };
    if (slotIndex < 0 || slotIndex >= CFG.MAX_FILES) return { success: false, error: "Invalid slotIndex." };


    if (!fileObj) return { success: true, url: "", fileName: "" };


    var approxBytes = Math.floor((fileObj.content?.length || 0) * 0.75);
    if (approxBytes > CFG.MAX_FILE_BYTES) {
      return { success: false, error: "File exceeds 10MB limit (backend safeguard)." };
    }


    var sheet = getNexusSheet_();
    var rowIndex = findRowByTicketId_(sheet, ticketId);
    if (rowIndex === -1) return { success: false, error: "Ticket not found." };


    var folder = DriveApp.getFolderById(CFG.FOLDER_ID);


    var blob = Utilities.newBlob(
      Utilities.base64Decode(fileObj.content),
      fileObj.mimeType || "application/octet-stream",
      fileObj.name || "attachment"
    );


    var saved = folder.createFile(blob);
    saved.setName(ticketId + "_" + saved.getName());


    var url = saved.getUrl();


    var col = 10 + slotIndex;
    var cell = sheet.getRange(rowIndex, col);
    var oldUrl = cell.getValue();
    cell.setValue(url);


    logAudit_(ticketId, "portal", "UPLOAD_TICKET_ATTACHMENT", "attachment_" + (slotIndex + 1), oldUrl, url, {
      slot: slotIndex + 1,
      fileUrl: url,
      fileName: saved.getName(),
      fileId: saved.getId()
    });


    return { success: true, url: url, fileName: saved.getName() };




  } catch (e) {
    return { success: false, error: e.toString() };
  }
}


function getAssigneesByDept_() {
  const out = {};
  Object.keys(TEAM_ASSIGNEES).forEach(dept => {
    out[dept] = (TEAM_ASSIGNEES[dept] || [])
      .map(e => String(e).toLowerCase().trim())
      .sort();
  });
  return out;
}


/*********************************
 * ADMIN API
 *********************************/
function adminInit() {
  enforceAdmin_();


  const assignees = getAllAssignees_();
  const assigneesByDept = getAssigneesByDept_();
  const tickets = adminGetTickets_();


  return {
    success: true,
    assignees,
    assigneesByDept,
    priorities: ["Low", "Medium", "High"],
    statuses: ["Open", "In Progress", "Waiting", "Closed"],
    tickets
  };
}


function auditInit() {
  enforceSuperAdmin_();


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.AUDIT_SHEET_NAME);
  if (!sh) return { success: true, eventTypes: [], fieldNames: [] };


  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: true, eventTypes: [], fieldNames: [] };


  const take = Math.min(2000, lastRow - 1);
  const data = sh.getRange(lastRow - take + 1, 1, take, 13).getValues();


  const ev = new Set();
  const fn = new Set();


  data.forEach(r => {
    const eventType = String(r[5] || "").trim();
    const fieldName = String(r[6] || "").trim();
    if (eventType) ev.add(eventType);
    if (fieldName) fn.add(fieldName);
  });


  return {
    success: true,
    eventTypes: Array.from(ev).sort(),
    fieldNames: Array.from(fn).sort(),
    sourceViews: ["portal","admin","my"],
    actorTypes: ["USER","ADMIN","SUPER_ADMIN"]
  };
}


function auditSearch(filters, pageRaw, pageSizeRaw) {
  enforceSuperAdmin_();


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.AUDIT_SHEET_NAME);
  if (!sh) return { success: true, rows: [], total: 0 };


  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: true, rows: [], total: 0 };


  const pageSize = Math.max(1, Math.min(200, parseInt(pageSizeRaw, 10) || 50));
  const page = Math.max(1, parseInt(pageRaw, 10) || 1);
  const startIndex = (page - 1) * pageSize;


  const f = filters || {};


  const fromMs = f.fromMs ? Number(f.fromMs) : null;
  const toMs   = f.toMs ? Number(f.toMs) : null;


  const ticketQ = String(f.ticketQ || "").toLowerCase().trim();
  const actorQ  = String(f.actorQ || "").toLowerCase().trim();
  const actorType = String(f.actorType || "").trim();
  const sourceView = String(f.sourceView || "").trim();
  const eventType  = String(f.eventType || "").trim();
  const fieldName  = String(f.fieldName || "").trim();
  const dept       = String(f.department || "").trim();
  const textQ      = String(f.textQ || "").toLowerCase().trim();


  const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();


  const matches = [];


  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];


    const ts = r[0];
    const ticketId = String(r[1] || "");
    const actorEmail = String(r[2] || "").toLowerCase();
    const aType = String(r[3] || "");
    const src = String(r[4] || "");
    const ev = String(r[5] || "");
    const fn = String(r[6] || "");
    const oldV = String(r[7] || "");
    const newV = String(r[8] || "");
    const updIdx = String(r[9] || "");
    const slot = String(r[10] || "");
    const fileUrl = String(r[11] || "");
    const metaJson = String(r[12] || "");


    const tsMs = toMillis_(ts);


    if (fromMs !== null && tsMs && tsMs < fromMs) continue;
    if (toMs !== null && tsMs && tsMs > toMs) continue;


    if (ticketQ && !ticketId.toLowerCase().includes(ticketQ)) continue;
    if (actorQ && !actorEmail.includes(actorQ)) continue;


    if (actorType && aType !== actorType) continue;
    if (sourceView && src !== sourceView) continue;
    if (eventType && ev !== eventType) continue;
    if (fieldName && fn !== fieldName) continue;


    if (dept) {
      let deptVal = "";
      try {
        const obj = JSON.parse(metaJson || "{}");
        deptVal = String(obj.department || "").trim();
      } catch (e) {}
      if (deptVal !== dept) continue;
    }


    if (textQ) {
      const hay = (oldV + " " + newV + " " + fileUrl + " " + metaJson).toLowerCase();
      if (!hay.includes(textQ)) continue;
    }


    matches.push({
      audit_ts: formatDateTime_(ts),
      audit_ts_ms: tsMs,
      ticket_id: ticketId,
      actor_email: actorEmail,
      actor_type: aType,
      source_view: src,
      event_type: ev,
      field_name: fn,
      old_value: oldV,
      new_value: newV,
      update_row_index: updIdx,
      attachment_slot: slot,
      file_url: fileUrl,
      meta_json: metaJson
    });
  }


  const total = matches.length;
  const rows = matches.slice(startIndex, startIndex + pageSize);


  return { success: true, rows, total };
}






function adminGetUpdates(ticketIdRaw, limitRaw, beforeMsRaw) {
  enforceAdmin_();


  const ticketId = String(ticketIdRaw || "").trim();
  if (!ticketId) return { success: false, error: "Missing ticketId." };


  let limit = parseInt(limitRaw, 10);
  if (!limit || limit < 1) limit = 10;
  if (limit > 50) limit = 50;


  const beforeMs = (beforeMsRaw === null || beforeMsRaw === undefined || beforeMsRaw === "")
    ? null
    : Number(beforeMsRaw);


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(CFG.SHEET_NAME);
  const mainMap = getHeaderMap_(main);
  const mainRow = findRowByTicketId_(main, ticketId);
  if (mainRow === -1) return { success: false, error: "Ticket not found." };
  enforceTicketScope_(main, mainMap, mainRow);


  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
  if (!updates) return { success: true, updates: [], nextCursor: null };


  const lastRow = updates.getLastRow();
  if (lastRow < 2) return { success: true, updates: [], nextCursor: null };


  const data = updates.getRange(2, 1, lastRow - 1, 9).getValues();


  const out = [];
  let oldestReturnedMs = null;
  let hasMore = false;


  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const ts = r[0];
    const tId = String(r[1] || "").trim();
    if (tId !== ticketId) continue;


    const tsMs = toMillis_(ts);
    if (!tsMs) continue;


    if (beforeMs !== null && !isNaN(beforeMs) && tsMs >= beforeMs) continue;


    out.push({
      update_timestamp: formatDateTime_(ts),
      update_timestamp_ms: tsMs,
      update_ticket_id: tId,
      update_author_email: String(r[2] || "").toLowerCase().trim(),
      update_message: String(r[3] || ""),
      update_att_1: String(r[4] || ""),
      update_att_2: String(r[5] || ""),
      update_att_3: String(r[6] || ""),
      update_att_4: String(r[7] || ""),
      update_att_5: String(r[8] || "")
    });


    oldestReturnedMs = tsMs;


    if (out.length >= limit) {
      for (let j = i - 1; j >= 0; j--) {
        const rr = data[j];
        const tid2 = String(rr[1] || "").trim();
        if (tid2 !== ticketId) continue;
        const ms2 = toMillis_(rr[0]);
        if (ms2 && ms2 < oldestReturnedMs) {
          hasMore = true;
          break;
        }
      }
      break;
    }
  }


  const nextCursor = (out.length > 0 && hasMore) ? oldestReturnedMs : null;


  return { success: true, updates: out, nextCursor: nextCursor };
}


function adminUpdateAssigned(ticketId, newAssignedRaw) {
  enforceAdmin_();

  const newAssigned = String(newAssignedRaw || "").toLowerCase().trim();
  const all = getAllAssignees_();
  if (newAssigned && !all.includes(newAssigned)) {
    return { success: false, error: "Invalid assignee." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  const map = getHeaderMap_(sheet);

  const row = findRowByTicketId_(sheet, ticketId);
  if (row === -1) return { success: false, error: "Ticket not found." };

  enforceTicketScope_(sheet, map, row);

  const now = new Date();
  const actor = Session.getActiveUser().getEmail();

  const assignedCol = map["assigned"];
  const statusCol = map["status"];
  const timeAssignedCol = map["time_assigned"];
  const initRespCol = map["initial_response_duration"];
  const tsCol = map["timestamp"];
  const updatedByCol = map["updated_by"];
  const lastUpdatedCol = map["last_updated"];

  const oldAssigned = String(sheet.getRange(row, assignedCol).getValue() || "").toLowerCase().trim();
  const oldStatus = String(sheet.getRange(row, statusCol).getValue() || "").trim();
  const created = sheet.getRange(row, tsCol).getValue();

  sheet.getRange(row, assignedCol).setValue(newAssigned);

  const existingTimeAssigned = sheet.getRange(row, timeAssignedCol).getValue();
  if (!existingTimeAssigned && newAssigned) {
    sheet.getRange(row, timeAssignedCol).setValue(now);

    if (created && created instanceof Date) {
      const diffMs = now.getTime() - created.getTime();
      const diffFraction = diffMs / 86400000;
      sheet.getRange(row, initRespCol).setValue(diffFraction);
    }
  }

  if (!oldAssigned && newAssigned && oldStatus === "Open") {
    sheet.getRange(row, statusCol).setValue("In Progress");
  }

  sheet.getRange(row, updatedByCol).setValue(actor);
  sheet.getRange(row, lastUpdatedCol).setValue(now);

  const finalStatus = String(sheet.getRange(row, statusCol).getValue() || "").trim();

  if (oldAssigned !== newAssigned) {
    logAudit_(ticketId, "admin", "CHANGE_ASSIGNED", "assigned", oldAssigned, newAssigned, {});
  }

  if (oldStatus !== finalStatus) {
    logAudit_(ticketId, "admin", "CHANGE_STATUS", "status", oldStatus, finalStatus, {
      reason: (!oldAssigned && newAssigned && oldStatus === "Open") ? "auto_in_progress_on_first_assign" : "side_effect"
    });
  }

  const ticket = adminGetTicketById_(ticketId);

  const changes = [];
  if (oldAssigned !== newAssigned) {
    changes.push({
      label: "Assigned To",
      oldValue: oldAssigned || "Unassigned",
      newValue: newAssigned || "Unassigned"
    });
  }
  if (oldStatus !== finalStatus) {
    changes.push({
      label: "Status",
      oldValue: oldStatus || "—",
      newValue: finalStatus || "—"
    });
  }

  if (changes.length) {
    sendFieldChangeNotification_(ticketId, changes);
  }

  return { success: true, ticket };
}


function adminUpdateStatus(ticketId, newStatusRaw) {
  enforceAdmin_();

  const newStatus = String(newStatusRaw || "").trim();
  const allowed = ["Open", "In Progress", "Waiting", "Closed"];
  if (!allowed.includes(newStatus)) return { success: false, error: "Invalid status." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  const map = getHeaderMap_(sheet);

  const row = findRowByTicketId_(sheet, ticketId);
  if (row === -1) return { success: false, error: "Ticket not found." };

  enforceTicketScope_(sheet, map, row);

  const now = new Date();
  const actor = Session.getActiveUser().getEmail();

  const statusCol = map["status"];
  const tsCol = map["timestamp"];
  const waitStartCol = map["wait_start"];
  const waitEndCol = map["wait_end"];
  const totalWaitMsCol = map["total_wait_ms"];
  const timeResolvedCol = map["time_resolved"];
  const timeToResolveCol = map["time_to_resolve"];
  const updatedByCol = map["updated_by"];
  const lastUpdatedCol = map["last_updated"];

  if (!totalWaitMsCol) {
    return { success: false, error: "Missing required column: total_wait_ms" };
  }

  const oldStatus = String(sheet.getRange(row, statusCol).getValue() || "").trim();
  const created = sheet.getRange(row, tsCol).getValue();

  let totalWaitMs = Number(sheet.getRange(row, totalWaitMsCol).getValue() || 0);
  if (!isFinite(totalWaitMs) || totalWaitMs < 0) totalWaitMs = 0;

  const waitStartVal = sheet.getRange(row, waitStartCol).getValue();

  // Entering Waiting: start a fresh current/latest wait cycle
  if (oldStatus !== "Waiting" && newStatus === "Waiting") {
    sheet.getRange(row, waitStartCol).setValue(now);
    sheet.getRange(row, waitEndCol).setValue("");
  }

  // Leaving Waiting: accumulate this finished wait cycle into total_wait_ms
  if (oldStatus === "Waiting" && newStatus !== "Waiting") {
    if (waitStartVal && waitStartVal instanceof Date) {
      totalWaitMs += Math.max(0, now.getTime() - waitStartVal.getTime());
      sheet.getRange(row, totalWaitMsCol).setValue(totalWaitMs);
    }
    sheet.getRange(row, waitEndCol).setValue(now);
  }

  sheet.getRange(row, statusCol).setValue(newStatus);

  if (newStatus === "Closed") {
    sheet.getRange(row, timeResolvedCol).setValue(now);

    if (created && created instanceof Date) {
      const ms = computeResolveMsExcludingWait_(sheet, row, map, created, now);
      sheet.getRange(row, timeToResolveCol).setValue(ms / 86400000);
    }
  } else {
    const tr = sheet.getRange(row, timeResolvedCol).getValue();
    if (tr && tr instanceof Date && created && created instanceof Date) {
      const ms = computeResolveMsExcludingWait_(sheet, row, map, created, tr);
      sheet.getRange(row, timeToResolveCol).setValue(ms / 86400000);
    }
  }

  sheet.getRange(row, updatedByCol).setValue(actor);
  sheet.getRange(row, lastUpdatedCol).setValue(now);

  if (oldStatus !== newStatus) {
    logAudit_(ticketId, "admin", "CHANGE_STATUS", "status", oldStatus, newStatus, {});
  }

  const ticket = adminGetTicketById_(ticketId);

  if (oldStatus !== newStatus) {
    if (newStatus === "Closed") {
      sendResolvedNotification_(ticketId);
    } else {
      sendFieldChangeNotification_(ticketId, [{
        label: "Status",
        oldValue: oldStatus || "—",
        newValue: newStatus || "—"
      }]);
    }
  }

  return { success: true, ticket };
}


function adminCreateUpdate(ticketId, messageRaw) {
  enforceAdmin_();


  const message = String(messageRaw || "").trim();
  if (!message) return { success: false, error: "Update message is empty." };


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nexus = ss.getSheetByName(CFG.SHEET_NAME);
  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);


  if (!updates) return { success: false, error: "Missing sheet tab: nexus_updates" };


  const map = getHeaderMap_(nexus);
  const row = findRowByTicketId_(nexus, ticketId);
  if (row === -1) return { success: false, error: "Ticket not found." };


  enforceTicketScope_(nexus, map, row);


  const now = new Date();
  const actor = Session.getActiveUser().getEmail();


  const updateRowValues = [
    now,
    ticketId,
    actor,
    message,
    "", "", "", "", ""
  ];
  updates.appendRow(updateRowValues);


  nexus.getRange(row, map["update_latest"]).setValue(message);
  nexus.getRange(row, map["updated_by"]).setValue(actor);
  nexus.getRange(row, map["last_updated"]).setValue(now);


  const latest = adminGetLatestUpdate_(ticketId);
  const ticket = adminGetTicketById_(ticketId);
  const updateRowIndex = updates.getLastRow();


  logAudit_(ticketId, "admin", "ADD_UPDATE", "update_message", "", message, {
    updateRowIndex: updateRowIndex,
    messagePreview: message.slice(0, 200)
  });


  return { success: true, ticket, latestUpdate: latest, updateRowIndex };
}


function adminUploadUpdateAttachment(updateRowIndex, ticketId, slotIndex, fileObj) {
  enforceAdmin_();


  if (!updateRowIndex || updateRowIndex < 2) return { success: false, error: "Invalid update row." };
  if (slotIndex < 0 || slotIndex >= CFG.MAX_FILES) return { success: false, error: "Invalid slotIndex." };
  if (!fileObj) return { success: true, url: "" };


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
  const nexus = ss.getSheetByName(CFG.SHEET_NAME);


  if (!updates) return { success: false, error: "Missing sheet tab: nexus_updates" };


  const rowTicket = String(updates.getRange(updateRowIndex, 2).getValue() || "").trim();
  if (rowTicket !== ticketId) return { success: false, error: "Update row does not match ticket." };


  const mainMap = getHeaderMap_(nexus);
  const mainRow = findRowByTicketId_(nexus, ticketId);
  if (mainRow === -1) return { success: false, error: "Ticket not found." };
  enforceTicketScope_(nexus, mainMap, mainRow);


  var approxBytes = Math.floor((fileObj.content?.length || 0) * 0.75);
  if (approxBytes > CFG.MAX_FILE_BYTES) {
    return { success: false, error: "File exceeds 10MB limit (backend safeguard)." };
  }


  const folder = DriveApp.getFolderById(CFG.FOLDER_ID);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(fileObj.content),
    fileObj.mimeType || "application/octet-stream",
    fileObj.name || "update_attachment"
  );


  const saved = folder.createFile(blob);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  saved.setName(ticketId + "_UPD_" + stamp + "_" + saved.getName());


  const url = saved.getUrl();


  const col = 5 + slotIndex;
  updates.getRange(updateRowIndex, col).setValue(url);


  logAudit_(ticketId, "admin", "UPLOAD_UPDATE_ATTACHMENT", "update_att_" + (slotIndex + 1), "", url, {
    updateRowIndex: updateRowIndex,
    slot: slotIndex + 1,
    fileUrl: url,
    fileName: saved.getName(),
    fileId: saved.getId()
  });




  const now = new Date();
  const actor = Session.getActiveUser().getEmail();
  nexus.getRange(mainRow, mainMap["updated_by"]).setValue(actor);
  nexus.getRange(mainRow, mainMap["last_updated"]).setValue(now);


  return { success: true, url };
}


/*********************************
 * Admin data getters
 *********************************/
function adminGetTicketById_(ticketId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  const map = getHeaderMap_(sheet);


  const row = findRowByTicketId_(sheet, ticketId);
  if (row === -1) return null;


  enforceTicketScope_(sheet, map, row);


  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const ticket = rowToTicketObject_(values, map);


  const latestUpdateMap = buildLatestUpdateMap_();
  ticket.latest_update = latestUpdateMap[ticket.ticket_id] || null;


  return ticket;
}


function adminGetTickets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  const map = getHeaderMap_(sheet);


  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];


  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();


  const latestUpdateMap = buildLatestUpdateMap_();


  const tickets = data.map(row => {
    const t = rowToTicketObject_(row, map);
    t.latest_update = latestUpdateMap[t.ticket_id] || null;
    return t;
  });


  const scope = getAdminScope_();
  if (scope.isSuperAdmin) return tickets;


  return tickets.filter(t => scope.allowedDepartments.includes(String(t.department || "").trim()));
}


function adminGetLatestUpdate_(ticketId) {
  const map = buildLatestUpdateMap_();
  return map[ticketId] || null;
}


function buildLatestUpdateMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
  if (!updates) return {};


  const lastRow = updates.getLastRow();
  if (lastRow < 2) return {};


  const data = updates.getRange(2, 1, lastRow - 1, 9).getValues();


  const best = {};
  data.forEach(r => {
    const ts = r[0];
    const ticketId = String(r[1] || "").trim();
    if (!ticketId) return;


    const tsMs = toMillis_(ts);
    const obj = {
      update_timestamp: formatDateTime_(ts),
      update_ticket_id: ticketId,
      update_author_email: String(r[2] || "").toLowerCase().trim(),
      update_message: String(r[3] || ""),
      update_att_1: String(r[4] || ""),
      update_att_2: String(r[5] || ""),
      update_att_3: String(r[6] || ""),
      update_att_4: String(r[7] || ""),
      update_att_5: String(r[8] || "")
    };


    if (!best[ticketId] || tsMs > best[ticketId].tsMs) {
      best[ticketId] = { tsMs, obj };
    }
  });


  const out = {};
  Object.keys(best).forEach(k => out[k] = best[k].obj);
  return out;
}


/*********************************
 * Scope + Access control helpers
 *********************************/
function getTeamsForEmail_(emailLower) {
  const teams = [];
  Object.keys(TEAM_ASSIGNEES).forEach(team => {
    if (TEAM_ASSIGNEES[team].includes(emailLower)) teams.push(team);
  });
  return teams;
}


function isSuperAdmin_(emailLower) {
  return SUPER_ADMINS.includes(emailLower);
}




function isAdmin_() {
  const email = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
  if (!email) return false;


  const domain = email.split("@").pop();
  if (!ALLOWED_DOMAINS.includes(domain)) return false;


  if (isSuperAdmin_(email)) return true;


  const teams = getTeamsForEmail_(email);
  return teams.length > 0;
}




function enforceAdmin_() {
  if (!isAdmin_()) throw new Error("Access denied.");
}


function enforceSuperAdmin_() {
  const email = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
  if (!email) throw new Error("Access denied.");
  if (!isSuperAdmin_(email)) throw new Error("Access denied.");
}




// Scope:
// - super admins: all tickets
// - team members: only their department(s)
function getAdminScope_() {
  const email = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
  const teams = getTeamsForEmail_(email);
  const superAdmin = isSuperAdmin_(email);


  return {
    email,
    isSuperAdmin: superAdmin,
    allowedDepartments: superAdmin ? [] : teams
  };
}


function enforceTicketScope_(sheet, headerMap, rowIndex) {
  const scope = getAdminScope_();
  if (scope.isSuperAdmin) return;


  const deptCol = headerMap["department"];
  const dept = String(sheet.getRange(rowIndex, deptCol).getValue() || "").trim();


  if (!scope.allowedDepartments.includes(dept)) {
    throw new Error("Access denied for this ticket's department.");
  }
}


/*********************************
 * Helpers (private)
 *********************************/
function getNexusSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found: " + CFG.SHEET_NAME);
  return sheet;
}


function generateNextTicketId_(sheet) {
  var lastRow = sheet.getLastRow();
  var nextId = CFG.ID_PREFIX + ("0000000" + 1).slice(-CFG.ID_PAD);


  if (lastRow > 1) {
    var lastId = sheet.getRange(lastRow, 1).getValue();
    var n = parseInt(String(lastId).replace(CFG.ID_PREFIX, ""), 10);
    if (!isNaN(n)) {
      n++;
      nextId = CFG.ID_PREFIX + ("0000000" + n).slice(-CFG.ID_PAD);
    }
  }
  return nextId;
}


function parseDepartment_(category) {
  var department = "General";
  var str = String(category || "");
  if (str.indexOf(":") > -1) {
    var prefix = str.split(":")[0].trim();
    department = (prefix === "ADMIN") ? "ADMINISTRATIVE" : prefix;
  }
  return department;
}


function findRowByTicketId_(sheet, ticketId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;


  var range = sheet.getRange(2, 1, lastRow - 1, 1);
  var finder = range.createTextFinder(ticketId).matchEntireCell(true);
  var found = finder.findNext();
  if (!found) return -1;


  return found.getRow();
}


function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || "").trim();
    if (key) map[key] = i + 1;
  });
  return map;
}


function rowToTicketObject_(row, map) {
  const get = (name) => row[(map[name] || 1) - 1];


  return {
    ticket_id: String(get("ticket_id") || ""),
    timestamp: formatDateTime_(get("timestamp")),
    email_address: String(get("email_address") || "").toLowerCase().trim(),
    department: String(get("department") || ""),
    category: String(get("category") || ""),
    priority: String(get("priority") || ""),
    status: String(get("status") || ""),
    summary: String(get("summary") || ""),
    description: String(get("description") || ""),


    attachment_1: String(get("attachment_1") || ""),
    attachment_2: String(get("attachment_2") || ""),
    attachment_3: String(get("attachment_3") || ""),
    attachment_4: String(get("attachment_4") || ""),
    attachment_5: String(get("attachment_5") || ""),


    assigned: String(get("assigned") || "").toLowerCase().trim(),
    update_latest: String(get("update_latest") || ""),


    time_assigned: formatDateTime_(get("time_assigned")),
    initial_response_duration: formatDuration_(get("initial_response_duration")),


    wait_start: formatDateTime_(get("wait_start")),
    wait_end: formatDateTime_(get("wait_end")),


    time_resolved: formatDateTime_(get("time_resolved")),
    time_to_resolve: formatDuration_(get("time_to_resolve")),


    updated_by: String(get("updated_by") || "").toLowerCase().trim(),
    last_updated: formatDateTime_(get("last_updated"))
  };
}


function formatDateTime_(val) {
  if (!val || val === "") return "";
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), CFG.DT_FMT);
  }
  return String(val);
}


function formatDuration_(val) {
  if (!val || val === "") return "";
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
    const diffFraction = (val.getHours() / 24) + (val.getMinutes() / 1440) + (val.getSeconds() / 86400);
    return secondsToHms_(Math.floor(diffFraction * 86400));
  }
  if (typeof val === "number") {
    return secondsToHms_(Math.floor(val * 86400));
  }
  return String(val);
}


function secondsToHms_(totalSeconds) {
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  return [h.toString().padStart(2, "0"), m.toString().padStart(2, "0"), s.toString().padStart(2, "0")].join(":");
}


function computeResolveMsExcludingWait_(sheet, row, map, created, resolved) {
  let ms = resolved.getTime() - created.getTime();

  const totalWaitMsCol = map["total_wait_ms"];
  let totalWaitMs = 0;

  if (totalWaitMsCol) {
    totalWaitMs = Number(sheet.getRange(row, totalWaitMsCol).getValue() || 0);
    if (!isFinite(totalWaitMs) || totalWaitMs < 0) totalWaitMs = 0;
  }

  ms -= totalWaitMs;

  return Math.max(0, ms);
}


function toMillis_(val) {
  if (!val) return 0;
  if (val instanceof Date) return val.getTime();
  const d = new Date(val);
  if (!isNaN(d.getTime())) return d.getTime();
  return 0;
}


function sanitize_(s) {
  return String(s || "").replace(/[<>&'\"]/g, "");
}


/*********************************
 * MY TICKETS VIEWER (MyTicketViewer)
 * - Shows tickets where email_address == active user
 * - Users can view updates + post updates + upload update attachments
 *********************************/


function getActiveEmailLower_() {
  return String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
}


function enforceAllowedDomainOrThrow_() {
  const email = getActiveEmailLower_();
  const domain = email.split("@").pop();
  if (!email || !ALLOWED_DOMAINS.includes(domain)) throw new Error("Access denied.");
  return email;
}


function enforceTicketOwner_(sheet, headerMap, rowIndex, emailLower) {
  const col = headerMap["email_address"];
  if (!col) throw new Error("Missing column: email_address");
  const owner = String(sheet.getRange(rowIndex, col).getValue() || "").toLowerCase().trim();
  if (owner !== emailLower) throw new Error("Access denied for this ticket.");
}


function myInit() {
  const emailLower = enforceAllowedDomainOrThrow_();


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sheet) return { success: false, error: "Sheet not found: " + CFG.SHEET_NAME };


  const map = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, tickets: [], defaultPageSize: 10 };
  }


  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();


  const tickets = data
    .map(r => rowToTicketObject_(r, map))
    .filter(t => String(t.email_address || "").toLowerCase().trim() === emailLower);


  return { success: true, tickets, defaultPageSize: 10 };
}


function myGetUpdates(ticketIdRaw, limitRaw, beforeMsRaw) {
  const emailLower = enforceAllowedDomainOrThrow_();


  const ticketId = String(ticketIdRaw || "").trim();
  if (!ticketId) return { success: false, error: "Missing ticketId." };


  let limit = parseInt(limitRaw, 10);
  if (!limit || limit < 1) limit = 10;
  if (limit > 50) limit = 50;


  const beforeMs = (beforeMsRaw === null || beforeMsRaw === undefined || beforeMsRaw === "")
    ? null
    : Number(beforeMsRaw);


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(CFG.SHEET_NAME);
  if (!main) return { success: false, error: "Sheet not found: " + CFG.SHEET_NAME };


  const mainMap = getHeaderMap_(main);
  const mainRow = findRowByTicketId_(main, ticketId);
  if (mainRow === -1) return { success: false, error: "Ticket not found." };


  enforceTicketOwner_(main, mainMap, mainRow, emailLower);


  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
  if (!updates) return { success: true, updates: [], nextCursor: null };


  const lastRow = updates.getLastRow();
  if (lastRow < 2) return { success: true, updates: [], nextCursor: null };


  const data = updates.getRange(2, 1, lastRow - 1, 9).getValues();


  const out = [];
  let oldestReturnedMs = null;
  let hasMore = false;


  for (let i = data.length - 1; i >= 0; i--) {
    const r = data[i];
    const ts = r[0];
    const tId = String(r[1] || "").trim();
    if (tId !== ticketId) continue;


    const tsMs = toMillis_(ts);
    if (!tsMs) continue;


    if (beforeMs !== null && !isNaN(beforeMs) && tsMs >= beforeMs) continue;


    out.push({
      update_timestamp: formatDateTime_(ts),
      update_timestamp_ms: tsMs,
      update_ticket_id: tId,
      update_author_email: String(r[2] || "").toLowerCase().trim(),
      update_message: String(r[3] || ""),
      update_att_1: String(r[4] || ""),
      update_att_2: String(r[5] || ""),
      update_att_3: String(r[6] || ""),
      update_att_4: String(r[7] || ""),
      update_att_5: String(r[8] || "")
    });


    oldestReturnedMs = tsMs;


    if (out.length >= limit) {
      for (let j = i - 1; j >= 0; j--) {
        const rr = data[j];
        const tid2 = String(rr[1] || "").trim();
        if (tid2 !== ticketId) continue;
        const ms2 = toMillis_(rr[0]);
        if (ms2 && ms2 < oldestReturnedMs) {
          hasMore = true;
          break;
        }
      }
      break;
    }
  }


  const nextCursor = (out.length > 0 && hasMore) ? oldestReturnedMs : null;
  return { success: true, updates: out, nextCursor };
}


function myCreateUpdate(ticketIdRaw, messageRaw) {
  const emailLower = enforceAllowedDomainOrThrow_();


  const ticketId = String(ticketIdRaw || "").trim();
  const message = String(messageRaw || "").trim();


  if (!ticketId) return { success: false, error: "Missing ticketId." };
  if (!message) return { success: false, error: "Update message is empty." };


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nexus = ss.getSheetByName(CFG.SHEET_NAME);
  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);


  if (!nexus) return { success: false, error: "Sheet not found: " + CFG.SHEET_NAME };
  if (!updates) return { success: false, error: "Missing sheet tab: " + CFG.UPDATES_SHEET_NAME };


  const map = getHeaderMap_(nexus);
  const row = findRowByTicketId_(nexus, ticketId);
  if (row === -1) return { success: false, error: "Ticket not found." };


  enforceTicketOwner_(nexus, map, row, emailLower);


  const now = new Date();
  const actor = Session.getActiveUser().getEmail();


  const updateRowValues = [ now, ticketId, actor, message, "", "", "", "", "" ];
  updates.appendRow(updateRowValues);


  if (map["update_latest"]) nexus.getRange(row, map["update_latest"]).setValue(message);
  if (map["updated_by"]) nexus.getRange(row, map["updated_by"]).setValue(actor);
  if (map["last_updated"]) nexus.getRange(row, map["last_updated"]).setValue(now);


  const updateRowIndex = updates.getLastRow();


  logAudit_(ticketId, "my", "ADD_UPDATE", "update_message", "", message, {
    updateRowIndex: updateRowIndex,
    messagePreview: message.slice(0, 200)
  });




  const latestUpdate = {
    update_timestamp: formatDateTime_(now),
    update_ticket_id: ticketId,
    update_author_email: String(actor || "").toLowerCase().trim(),
    update_message: message,
    update_att_1: "",
    update_att_2: "",
    update_att_3: "",
    update_att_4: "",
    update_att_5: ""
  };


  return { success: true, latestUpdate, updateRowIndex };
}


function myUploadUpdateAttachment(updateRowIndex, ticketIdRaw, slotIndex, fileObj) {
  const emailLower = enforceAllowedDomainOrThrow_();


  const ticketId = String(ticketIdRaw || "").trim();
  if (!ticketId) return { success: false, error: "Missing ticketId." };
  if (!updateRowIndex || updateRowIndex < 2) return { success: false, error: "Invalid update row." };
  if (slotIndex < 0 || slotIndex >= CFG.MAX_FILES) return { success: false, error: "Invalid slotIndex." };
  if (!fileObj) return { success: true, url: "" };


  const approxBytes = Math.floor((fileObj.content?.length || 0) * 0.75);
  if (approxBytes > CFG.MAX_FILE_BYTES) {
    return { success: false, error: "File exceeds 10MB limit (backend safeguard)." };
  }


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
  const nexus = ss.getSheetByName(CFG.SHEET_NAME);


  if (!updates) return { success: false, error: "Missing sheet tab: " + CFG.UPDATES_SHEET_NAME };
  if (!nexus) return { success: false, error: "Sheet not found: " + CFG.SHEET_NAME };


  const rowTicket = String(updates.getRange(updateRowIndex, 2).getValue() || "").trim();
  if (rowTicket !== ticketId) return { success: false, error: "Update row does not match ticket." };


  const mainMap = getHeaderMap_(nexus);
  const mainRow = findRowByTicketId_(nexus, ticketId);
  if (mainRow === -1) return { success: false, error: "Ticket not found." };
  enforceTicketOwner_(nexus, mainMap, mainRow, emailLower);


  const folder = DriveApp.getFolderById(CFG.FOLDER_ID);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(fileObj.content),
    fileObj.mimeType || "application/octet-stream",
    fileObj.name || "update_attachment"
  );


  const saved = folder.createFile(blob);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  saved.setName(ticketId + "_MYUPD_" + stamp + "_" + saved.getName());


  const url = saved.getUrl();


  const col = 5 + slotIndex;
  updates.getRange(updateRowIndex, col).setValue(url);


  logAudit_(ticketId, "my", "UPLOAD_UPDATE_ATTACHMENT", "update_att_" + (slotIndex + 1), "", url, {
    updateRowIndex: updateRowIndex,
    slot: slotIndex + 1,
    fileUrl: url,
    fileName: saved.getName(),
    fileId: saved.getId()
  });




  const now = new Date();
  const actor = Session.getActiveUser().getEmail();
  if (mainMap["updated_by"]) nexus.getRange(mainRow, mainMap["updated_by"]).setValue(actor);
  if (mainMap["last_updated"]) nexus.getRange(mainRow, mainMap["last_updated"]).setValue(now);


  return { success: true, url };
}



/*********************************
 * METRICS (Attention Queue)
 * - Uses ticket + update data only
 *********************************/

function metricsInit() {
  enforceAdmin_();

  const emails = getAllAssignees_();
  const map = metricsBuildLocalDomainMap_(emails.concat(SUPER_ADMINS));

  const assignees = emails.map(e => ({
    email: String(e).toLowerCase().trim(),
    label: metricsEmailLabel_(e, map)
  }));

  return {
    success: true,
    departments: Object.keys(TEAM_ASSIGNEES).sort(),
    priorities: ["Low", "Medium", "High"],
    statuses: ["Open", "In Progress", "Waiting", "Closed"],
    assignees: assignees
  };
}

function metricsGet(filters) {
  enforceAdmin_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tSh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!tSh) return { success: false, error: "Missing sheet tab: " + CFG.SHEET_NAME };

  const uSh = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);

  const now = new Date();
  const nowMs = now.getTime();

  // Metrics is intentionally global for all admins who can access metrics.html
  const tMap = getHeaderMap_(tSh);
  const lastRow = tSh.getLastRow();
  const lastCol = tSh.getLastColumn();
  const data = lastRow >= 2 ? tSh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const tickets = [];
  const assignedSeen = new Set();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];

    const ticketId = String(r[(tMap["ticket_id"] || 1) - 1] || "").trim();
    if (!ticketId) continue;

    const dept = String(r[(tMap["department"] || 1) - 1] || "").trim();
    const priority = String(r[(tMap["priority"] || 1) - 1] || "").trim();
    const status = String(r[(tMap["status"] || 1) - 1] || "").trim();
    const assigned = String(r[(tMap["assigned"] || 1) - 1] || "").toLowerCase().trim();
    const summary = String(r[(tMap["summary"] || 1) - 1] || "").trim();

    const createdMs = toMillis_(r[(tMap["timestamp"] || 1) - 1]);
    const lastMsRaw = toMillis_(r[(tMap["last_updated"] || 1) - 1]);
    const lastMs = lastMsRaw || createdMs;

    const initRespSec = metricsDurationSeconds_(r[(tMap["initial_response_duration"] || 1) - 1]);
    const resolveSec = metricsDurationSeconds_(r[(tMap["time_to_resolve"] || 1) - 1]);

    if (assigned) assignedSeen.add(assigned);

    const ageMs = createdMs ? (nowMs - createdMs) : 0;
    const sinceMs = lastMs ? (nowMs - lastMs) : 0;

    tickets.push({
      ticket_id: ticketId,
      department: dept,
      priority: priority,
      status: status,
      assigned: assigned,
      summary: summary,
      age_ms: ageMs,
      since_ms: sinceMs,
      initial_response_sec: initRespSec,
      time_to_resolve_sec: resolveSec
    });
  }

  let scoped = tickets;

  const f = filters || {};
  const fDept = String(f.department || "").trim();
  const fStatus = String(f.status || "").trim();
  const fPriority = String(f.priority || "").trim();
  const fAssigned = String(f.assigned || "").toLowerCase().trim();
  const includeWaiting = String(f.includeWaiting).toLowerCase() !== "false";

  let filtered = scoped.filter(t => {
    if (fDept && t.department !== fDept) return false;
    if (fStatus && t.status !== fStatus) return false;
    if (fPriority && t.priority !== fPriority) return false;
    if (fAssigned && t.assigned !== fAssigned) return false;
    if (!includeWaiting && t.status === "Waiting") return false;
    return true;
  });

  const baseEmails = getAllAssignees_().concat(SUPER_ADMINS).map(e => String(e).toLowerCase().trim());
  assignedSeen.forEach(e => baseEmails.push(String(e).toLowerCase().trim()));
  const labelMap = metricsBuildLocalDomainMap_(baseEmails);

  const labelEmail = (email) => metricsEmailLabel_(email, labelMap);

  const totalTickets = filtered.length;
  const openBacklog = filtered.filter(t => t.status !== "Closed").length;
  const unassignedOpen = filtered.filter(t => t.status !== "Closed" && !t.assigned).length;
  const waitingOpen = filtered.filter(t => t.status === "Waiting").length;

  const medianFirstAssignSec = metricsMedian_(filtered.map(t => t.initial_response_sec));
  const medianResolveSec = metricsMedian_(filtered.filter(t => t.status === "Closed").map(t => t.time_to_resolve_sec));

  const openOnly = filtered.filter(t => t.status !== "Closed");

  const oldestOutstanding = openOnly
    .slice()
    .sort((a,b) => (b.age_ms || 0) - (a.age_ms || 0))
    .slice(0, 10)
    .map(t => ({
      ticket_id: t.ticket_id,
      summary: t.summary,
      assigned_display: t.assigned ? labelEmail(t.assigned) : "—",
      age_ms: t.age_ms,
      since_ms: t.since_ms
    }));

  const stalest = openOnly
    .slice()
    .sort((a,b) => (b.since_ms || 0) - (a.since_ms || 0))
    .slice(0, 10)
    .map(t => ({
      ticket_id: t.ticket_id,
      summary: t.summary,
      assigned_display: t.assigned ? labelEmail(t.assigned) : "—",
      age_ms: t.age_ms,
      since_ms: t.since_ms
    }));

  const oldestUnassigned = openOnly
    .filter(t => !t.assigned)
    .slice()
    .sort((a,b) => (b.age_ms || 0) - (a.age_ms || 0))
    .slice(0, 10)
    .map(t => ({
      ticket_id: t.ticket_id,
      summary: t.summary,
      status: t.status,
      age_ms: t.age_ms
    }));

  const peopleEmails = getAllAssignees_().map(e => String(e).toLowerCase().trim());

  const ticketIdSet = new Set(filtered.map(t => t.ticket_id));
  const updStats = metricsBuildUpdateStats_(uSh, ticketIdSet);

  const people = peopleEmails.map(email => {
    const openMine = openOnly.filter(t => t.assigned === email);

    const openCount = openMine.length;
    const oldestAgeMs = metricsMax_(openMine.map(x => x.age_ms));
    const longestSinceMs = metricsMax_(openMine.map(x => x.since_ms));
    const avgSinceMs = metricsAvg_(openMine.map(x => x.since_ms));

    const closedMine = filtered.filter(t => t.assigned === email && t.status === "Closed");
    const closedCount = closedMine.length;

    const avgResolveSec = metricsAvg_(closedMine.map(x => x.time_to_resolve_sec));

    const stats = updStats[email] || { updates: 0, atts: 0 };

    return {
      email: email,
      display: labelEmail(email),

      open_count: openCount,
      closed_count: closedCount,

      avg_resolve_sec: avgResolveSec,

      oldest_age_ms: oldestAgeMs,
      longest_since_ms: longestSinceMs,
      avg_since_ms: avgSinceMs,

      updates_count: stats.updates,
      update_att_count: stats.atts
    };
  });

  return {
    success: true,
    updatedAt: formatDateTime_(now),
    kpis: {
      totalTickets,
      openBacklog,
      unassignedOpen,
      waitingOpen,
      medianFirstAssignSec,
      medianResolveSec
    },
    attention: {
      oldestOutstanding,
      stalest,
      oldestUnassigned
    },
    people
  };
}

/** Build map {localPart: {domain:true}} for collision-safe labels */
function metricsBuildLocalDomainMap_(emails) {
  const map = {};
  (emails || []).forEach(e => {
    const email = String(e || "").toLowerCase().trim();
    if (!email.includes("@")) return;
    const [local, domain] = email.split("@");
    if (!local || !domain) return;
    if (!map[local]) map[local] = {};
    map[local][domain] = true;
  });
  return map;
}

/** Display local-part only. Add domain label if needed. */
function metricsEmailLabel_(email, localDomainMap) {
  const em = String(email || "").toLowerCase().trim();
  if (!em) return "";
  if (!em.includes("@")) return em;

  const [local, domain] = em.split("@");
  if (!local) return em;

  const doms = Object.keys((localDomainMap && localDomainMap[local]) ? localDomainMap[local] : {});
  if (doms.length > 1) {
    const short = domain.startsWith("operations.") ? "operations" : "internal";
    return `${local} (${short})`;
  }
  return local;
}

/** Convert stored duration to seconds */
function metricsDurationSeconds_(val) {
  if (val === null || val === undefined || val === "") return 0;

  if (val instanceof Date && !isNaN(val)) {
    return (val.getHours() * 3600) + (val.getMinutes() * 60) + (val.getSeconds());
  }

  if (typeof val === "number" && isFinite(val)) {
    return Math.round(val * 86400);
  }

  const s = String(val).trim();
  const m = s.match(/^(\d{1,3}):(\d{2}):(\d{2})$/);
  if (m) return (Number(m[1]) * 3600) + (Number(m[2]) * 60) + Number(m[3]);

  return 0;
}

function metricsMedian_(arr) {
  const a = (arr || [])
    .map(n => Number(n))
    .filter(n => isFinite(n) && n > 0)
    .sort((x,y) => x - y);
  if (!a.length) return 0;
  const mid = Math.floor(a.length / 2);
  return (a.length % 2) ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function metricsMax_(arr) {
  const a = (arr || []).map(n => Number(n)).filter(n => isFinite(n) && n > 0);
  if (!a.length) return 0;
  return Math.max.apply(null, a);
}

function metricsAvg_(arr) {
  const a = (arr || []).map(n => Number(n)).filter(n => isFinite(n) && n > 0);
  if (!a.length) return 0;
  const sum = a.reduce((acc, n) => acc + n, 0);
  return sum / a.length;
}

/** Count updates + update-attachments per author */
function metricsBuildUpdateStats_(updatesSheet, ticketIdSet) {
  const out = {};

  if (!updatesSheet) return out;

  const lastRow = updatesSheet.getLastRow();
  if (lastRow < 2) return out;

  const data = updatesSheet.getRange(2, 1, lastRow - 1, 9).getValues();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const ticketId = String(r[1] || "").trim();
    if (!ticketId) continue;
    if (ticketIdSet && !ticketIdSet.has(ticketId)) continue;

    const author = String(r[2] || "").toLowerCase().trim();
    if (!author) continue;

    if (!out[author]) out[author] = { updates: 0, atts: 0 };
    out[author].updates++;

    for (let c = 4; c <= 8; c++) {
      if (String(r[c] || "").trim()) out[author].atts++;
    }
  }

  return out;
}


/*********************************
 * NEXUS LIVE CLEANUP
 * Removes Closed tickets from LIVE NEXUS
 * plus related nexus_updates and nexus_audit rows.
 *********************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("NEXUS Live Cleanup")
    .addItem("1) Preview Closed Ticket Cleanup", "nexusPreviewClosedTicketCleanup")
    .addItem("2) Delete Closed Tickets + Related Records", "nexusDeleteClosedTicketsAndRelatedRecords")
    .addToUi();
}

function nexusPreviewClosedTicketCleanup() {
  enforceSuperAdmin_();

  const plan = nexusBuildClosedTicketCleanupPlan_();

  const sampleIds = plan.closedTicketIds.slice(0, 20).join(", ");
  const moreText = plan.closedTicketIds.length > 20
    ? `\n\nFirst 20 ticket IDs:\n${sampleIds}\n\n...and ${plan.closedTicketIds.length - 20} more.`
    : plan.closedTicketIds.length
      ? `\n\nTicket IDs:\n${sampleIds}`
      : "";

  SpreadsheetApp.getUi().alert(
    [
      "NEXUS Live Cleanup Preview",
      "",
      `Closed tickets to delete from ${CFG.SHEET_NAME}: ${plan.ticketRows.length}`,
      `Related update rows to delete from ${CFG.UPDATES_SHEET_NAME}: ${plan.updateRows.length}`,
      `Related audit rows to delete from ${CFG.AUDIT_SHEET_NAME}: ${plan.auditRows.length}`,
      moreText,
      "",
      "Nothing has been deleted yet."
    ].join("\n")
  );
}

function nexusDeleteClosedTicketsAndRelatedRecords() {
  enforceSuperAdmin_();

  const ui = SpreadsheetApp.getUi();

  let plan = nexusBuildClosedTicketCleanupPlan_();

  if (!plan.ticketRows.length) {
    ui.alert("No Closed tickets found. Nothing to delete.");
    return;
  }

  const response = ui.prompt(
    "Confirm NEXUS Live Cleanup",
    [
      "This will permanently delete from LIVE NEXUS:",
      "",
      `Closed tickets: ${plan.ticketRows.length}`,
      `Related update rows: ${plan.updateRows.length}`,
      `Related audit rows: ${plan.auditRows.length}`,
      "",
      "This will NOT delete Drive attachment files.",
      "",
      "Type DELETE CLOSED to continue."
    ].join("\n"),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Cleanup cancelled.");
    return;
  }

  if (String(response.getResponseText() || "").trim() !== "DELETE CLOSED") {
    ui.alert("Cleanup cancelled. Confirmation text did not match.");
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    plan = nexusBuildClosedTicketCleanupPlan_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nexus = ss.getSheetByName(CFG.SHEET_NAME);
    const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
    const audit = ss.getSheetByName(CFG.AUDIT_SHEET_NAME);

    nexusDeleteRowsByNumbers_(updates, plan.updateRows);
    nexusDeleteRowsByNumbers_(audit, plan.auditRows);

    nexusDeleteRowsByNumbers_(nexus, plan.ticketRows);

    ui.alert(
      [
        "NEXUS Live Cleanup completed.",
        "",
        `Deleted Closed tickets: ${plan.ticketRows.length}`,
        `Deleted related update rows: ${plan.updateRows.length}`,
        `Deleted related audit rows: ${plan.auditRows.length}`
      ].join("\n")
    );

  } finally {
    lock.releaseLock();
  }
}

function nexusBuildClosedTicketCleanupPlan_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const nexus = ss.getSheetByName(CFG.SHEET_NAME);
  const updates = ss.getSheetByName(CFG.UPDATES_SHEET_NAME);
  const audit = ss.getSheetByName(CFG.AUDIT_SHEET_NAME);

  if (!nexus) throw new Error("Missing sheet tab: " + CFG.SHEET_NAME);
  if (!updates) throw new Error("Missing sheet tab: " + CFG.UPDATES_SHEET_NAME);
  if (!audit) throw new Error("Missing sheet tab: " + CFG.AUDIT_SHEET_NAME);

  const nexusMap = getHeaderMap_(nexus);
  const ticketIdCol = nexusRequireCol_(nexusMap, "ticket_id", CFG.SHEET_NAME);
  const statusCol = nexusRequireCol_(nexusMap, "status", CFG.SHEET_NAME);

  const closedTicketIds = [];
  const closedSet = new Set();
  const ticketRows = [];

  const nexusLastRow = nexus.getLastRow();
  const nexusLastCol = nexus.getLastColumn();

  if (nexusLastRow >= 2) {
    const data = nexus.getRange(2, 1, nexusLastRow - 1, nexusLastCol).getValues();

    data.forEach((row, i) => {
      const ticketId = String(row[ticketIdCol - 1] || "").trim();
      const status = String(row[statusCol - 1] || "").trim().toLowerCase();

      if (ticketId && status === "closed") {
        closedTicketIds.push(ticketId);
        closedSet.add(ticketId);
        ticketRows.push(i + 2);
      }
    });
  }

  const updateRows = nexusFindRowsByTicketSet_(
    updates,
    "update_ticket_id",
    closedSet,
    CFG.UPDATES_SHEET_NAME
  );

  const auditRows = nexusFindRowsByTicketSet_(
    audit,
    "ticket_id",
    closedSet,
    CFG.AUDIT_SHEET_NAME
  );

  return {
    closedTicketIds,
    ticketRows,
    updateRows,
    auditRows
  };
}

function nexusFindRowsByTicketSet_(sheet, ticketHeaderName, ticketSet, sheetName) {
  if (!ticketSet || ticketSet.size === 0) return [];

  const map = getHeaderMap_(sheet);
  const ticketCol = nexusRequireCol_(map, ticketHeaderName, sheetName);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];

  data.forEach((row, i) => {
    const ticketId = String(row[ticketCol - 1] || "").trim();

    if (ticketSet.has(ticketId)) {
      rows.push(i + 2);
    }
  });

  return rows;
}

function nexusRequireCol_(map, headerName, sheetName) {
  const col = map[headerName];

  if (!col) {
    throw new Error(`Missing required header "${headerName}" in sheet "${sheetName}".`);
  }

  return col;
}

function nexusDeleteRowsByNumbers_(sheet, rowNumbers) {
  if (!sheet || !rowNumbers || !rowNumbers.length) return;

  const rows = Array.from(new Set(rowNumbers))
    .filter(r => Number(r) >= 2)
    .sort((a, b) => b - a);

  if (!rows.length) return;

  let blockStart = null;
  let blockCount = 0;

  rows.forEach(row => {
    if (blockStart === null) {
      blockStart = row;
      blockCount = 1;
      return;
    }

    if (row === blockStart - blockCount) {
      blockCount++;
      return;
    }

    sheet.deleteRows(blockStart - blockCount + 1, blockCount);

    blockStart = row;
    blockCount = 1;
  });

  if (blockStart !== null && blockCount > 0) {
    sheet.deleteRows(blockStart - blockCount + 1, blockCount);
  }
}
