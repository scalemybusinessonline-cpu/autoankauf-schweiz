/**
 * Referenz-Kopie des Google Apps Script Web-App-Backends für das Offerte-Formular
 * (script.google.com/macros/s/AKfycbxnOeWjoHunSxaz13rdtQfp-GmGphvuf5TfSYy5aDntM3MTzCorzxBfpMso0mocpCOV/exec).
 *
 * Diese Datei liegt hier nur als Backup/Versionierung — deployed wird sie manuell
 * über den Apps-Script-Editor (script.google.com), nicht automatisch aus diesem Repo.
 *
 * Änderungen in dieser Version:
 * - Feld-Mapping an das Formular angepasst: p.name (statt p.vorname/p.nachname),
 *   p.plz_ort (statt p.plz/p.ort) — das Formular sendet diese Felder seit dem
 *   Merge von Vorname+Nachname bzw. PLZ+Ort nur noch kombiniert.
 * - Header-Check robuster: vergleicht die komplette Kopfzeile statt nur Zelle A1,
 *   damit Spaltenänderungen (wie diese) den Header zuverlässig aktualisieren.
 * - Google-Drive-Backup: pro Anfrage wird ein Unterordner mit den Formulardaten
 *   (als Textdatei) und allen hochgeladenen Bildern angelegt — unabhängig von
 *   Gmail/Sheet, als zusätzliche Sicherung.
 * - E-Mail-Anhang-Limit: ab MAX_EMAIL_ATTACHMENT_BYTES (18 MB) werden Bilder
 *   nicht mehr an die Benachrichtigungs-Mail angehängt (Gmail-Limit), sondern
 *   nur noch der Drive-Link verschickt. Drive-Backup bekommt immer alle Bilder,
 *   unabhängig von der Grösse.
 * - Direct-to-Drive-Upload für grosse Anhänge (>30 MB gesamt): das Formular
 *   fragt vorher per doGet (action=createUploadSession) eine Drive-Resumable-
 *   Upload-Session an und lädt die Datei direkt vom Browser zu Google hoch —
 *   das umgeht das ~50-MB-POST-Body-Limit von Apps-Script-Web-Apps. Nur die
 *   resultierenden Drive-Datei-IDs kommen dann im normalen Formular-POST als
 *   bilder_ids an (parallel zum alten bilder_data-Feld für kleine Uploads).
 */

// Fester Ziel-Ordner: "Anfragen autoankauf-schweiz.ch" in Google Drive
var DRIVE_ROOT_FOLDER_ID = '1B9eXRoC2HwQQ8ZiBWMuPKdCy7XdCj7AN';

// Make.com Custom-Webhook-URL ("AutoAnkauf Schweiz – Neue Anfrage" Szenario)
// Leer lassen, um den Webhook-Versand zu deaktivieren.
var AUTOMATION_WEBHOOK_URL = 'https://hook.eu1.make.com/drcu2cx3a4wf65j9q8lioxwkygysfbkh';

var HEADERS = ['Lfd. Nr.', 'Zeitstempel', 'Name / Vorname', 'Firma', 'Straße & Nr.', 'PLZ / Ort', 'E-Mail', 'Telefon',
               'Automarke', 'Typ', 'Typenscheinnummer', 'Erste Inverkehrssetzung', 'Kilometerstand', 'Farbe',
               'Preisvorstellung', 'Fahrzeugart', 'Bemerkungen', 'Kanton-Seite',
               'utm_source', 'utm_medium', 'utm_campaign', 'utm_term', 'gclid', 'fbid', 'Drive-Ordner'];

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'createUploadSession') {
    return createUploadSession(e.parameter.filename, e.parameter.mimeType);
  }
  return ContentService.createTextOutput(JSON.stringify({ error: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Erstellt eine Google-Drive-Resumable-Upload-Session im Staging-Ordner und
 * gibt die Session-URL als JSON zurück. Der Browser lädt die Bilddatei danach
 * per PUT direkt an diese URL hoch (an googleapis.com, nicht an dieses
 * Apps-Script) — das umgeht das ~50-MB-Limit für normale doPost-Requests bei
 * grossen Uploads. Wird per einfachem GET aufgerufen (kein Preflight nötig).
 */
function createUploadSession(filename, mimeType) {
  var staging = getOrCreateStagingFolder();
  var metadata = {
    name: filename || 'upload',
    parents: [staging.getId()]
  };

  var response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id', {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'X-Upload-Content-Type': mimeType || 'application/octet-stream'
    },
    payload: JSON.stringify(metadata),
    muteHttpExceptions: true
  });

  var headers = response.getHeaders();
  var uploadUrl = headers['Location'] || headers['location'] || '';

  return ContentService.createTextOutput(JSON.stringify({ uploadUrl: uploadUrl }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureHeaders(sheet);

  // Header steht jetzt sicher in Zeile 1 -> aktuelle getLastRow() ist die laufende Nummer der neuen Zeile
  var nextNumber = sheet.getLastRow();

  var p = e.parameter;
  var driveResult = saveToDrive(p, nextNumber);
  var attachments = driveResult.attachments;
  var driveFolderUrl = driveResult.folderUrl;

  var row = [
    nextNumber,
    new Date(),
    p.name || '',
    p.firma || '',
    p.strasse || '',
    p.plz_ort || '',
    p.email || '',
    p.telefon || '',
    p.marke || '',
    p.modell || '',
    p.typenscheinnummer || '',
    p.jahrgang || '',
    p.km || '',
    p.farbe || '',
    p.preisvorstellung || '',
    p.typ || '',
    p.bemerkungen || '',
    p.kanton || '',
    p.utm_source || '',
    p.utm_medium || '',
    p.utm_campaign || '',
    p.utm_term || '',
    p.gclid || '',
    p.fbid || '',
    driveFolderUrl || ''
  ];

  // Neue Zeile direkt unter der Kopfzeile einfügen, statt ans Ende anzuhängen
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, row.length).setValues([row]);

  // Benachrichtigung per E-Mail (inkl. Bilder als Anhang, falls vorhanden)
  sendNotificationEmail(p, nextNumber, attachments, driveFolderUrl);

  // Webhook für Automation (Make.com/Zapier) — läuft parallel zu Sheet/E-Mail/Drive
  sendToAutomationWebhook(p, nextNumber, driveFolderUrl);

  return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Schreibt die Kopfzeile neu, falls sie fehlt oder nicht mehr zur aktuellen
 * Spaltenstruktur passt (nicht nur Zelle A1 prüfen — sonst bleiben Zeilen bei
 * einer Spaltenanzahl-Änderung wie dieser stumm verschoben).
 */
function ensureHeaders(sheet) {
  var currentHeaderRange = sheet.getRange(1, 1, 1, HEADERS.length);
  var currentHeaders = currentHeaderRange.getValues()[0];
  var matches = HEADERS.every(function (h, i) { return currentHeaders[i] === h; });
  if (!matches) {
    currentHeaderRange.setValues([HEADERS]);
  }
}

function parseImageAttachments(bilderDataJson) {
  if (!bilderDataJson) return [];
  var items;
  try {
    items = JSON.parse(bilderDataJson);
  } catch (e) {
    return [];
  }
  if (!Array.isArray(items)) return [];

  return items.map(function (item) {
    var bytes = Utilities.base64Decode(item.data);
    return Utilities.newBlob(bytes, item.type || 'application/octet-stream', item.name || 'bild.jpg');
  });
}

/**
 * Legt pro Anfrage einen Unterordner in Google Drive an mit den Formulardaten
 * (als Textdatei) und allen hochgeladenen Bildern — als Sicherung unabhängig
 * von Gmail (Anhang-Limit) und dem Sheet (Zell-Zeichenlimit für Bilder).
 * Schlägt das Drive-Backup fehl, wird das geloggt, aber Sheet-Eintrag und
 * E-Mail-Versand laufen trotzdem weiter (kein harter Abbruch) — die inline per
 * Base64 mitgeschickten Bilder werden dann trotzdem noch an die E-Mail gehängt.
 *
 * Führt zwei Anhang-Quellen zusammen: klassische Base64-Bilder (bilder_data,
 * für kleine Uploads) und bereits per Resumable-Session zu Drive hochgeladene
 * Dateien (bilder_ids, für grosse Uploads > 30 MB) — letztere liegen im
 * Staging-Ordner und werden hier in den finalen Anfrage-Ordner verschoben.
 */
function saveToDrive(p, nextNumber) {
  var inlineBlobs = parseImageAttachments(p.bilder_data);

  try {
    var root = getOrCreateRootFolder();
    var datum = Utilities.formatDate(new Date(), 'Europe/Zurich', 'yyyy-MM-dd_HH-mm');
    var nameFuerOrdner = String(p.name || 'Unbekannt').replace(/[\\\/:*?"<>|]/g, '').trim() || 'Unbekannt';
    var folder = root.createFolder('#' + nextNumber + ' – ' + datum + ' – ' + nameFuerOrdner);

    folder.createFile('anfrage.txt', buildDetailsText(p, nextNumber), MimeType.PLAIN_TEXT);

    inlineBlobs.forEach(function (blob) {
      folder.createFile(blob);
    });

    var uploadedBlobs = resolveDriveUploadedAttachments(p.bilder_ids, folder);

    return { folderUrl: folder.getUrl(), attachments: inlineBlobs.concat(uploadedBlobs) };
  } catch (err) {
    Logger.log('Drive-Backup fehlgeschlagen: ' + err);
    return { folderUrl: '', attachments: inlineBlobs };
  }
}

/**
 * Verschiebt Dateien, die der Browser per Resumable-Upload-Session bereits
 * direkt zu Drive hochgeladen hat (siehe createUploadSession), aus dem
 * Staging-Ordner in den finalen Anfrage-Ordner. Gibt sie als Blobs zurück,
 * damit sendNotificationEmail sie wie die Base64-Bilder behandeln kann
 * (gleiche 18-MB-Schwelle für Mail-Anhang vs. nur Drive-Link).
 */
function resolveDriveUploadedAttachments(bilderIdsJson, zielOrdner) {
  if (!bilderIdsJson) return [];
  var items;
  try {
    items = JSON.parse(bilderIdsJson);
  } catch (e) {
    return [];
  }
  if (!Array.isArray(items)) return [];

  var staging = getOrCreateStagingFolder();
  var blobs = [];
  items.forEach(function (item) {
    if (!item || !item.id) return;
    try {
      var file = DriveApp.getFileById(item.id);
      zielOrdner.addFile(file);
      staging.removeFile(file);
      blobs.push(file.getBlob());
    } catch (err) {
      Logger.log('Hochgeladene Datei konnte nicht verschoben werden (' + item.id + '): ' + err);
    }
  });
  return blobs;
}

function getOrCreateRootFolder() {
  return DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
}

/**
 * Zwischenablage für Dateien, die der Browser per Resumable-Upload-Session
 * direkt zu Drive hochgeladen hat, bevor das eigentliche Formular abgeschickt
 * wurde (die Ziel-Ordner-Nummer steht erst in doPost fest). Bricht ein Nutzer
 * nach dem Hochladen ab, bleiben Dateien hier liegen — Aufräumen (z. B. per
 * zeitgesteuertem Trigger) ist bewusst nicht Teil dieser Version.
 */
function getOrCreateStagingFolder() {
  var root = getOrCreateRootFolder();
  var name = '_Eingehende-Uploads';
  var existing = root.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return root.createFolder(name);
}

/**
 * Einmalig manuell im Editor ausführen (Funktion im Dropdown auswählen -> ▶ Ausführen),
 * um alle benötigten Berechtigungen zu erteilen (Drive lesen+schreiben, externe
 * Webhook-Requests). Das doPost-Formular selbst kann den Autorisierungs-Dialog
 * nicht auslösen (läuft als Web App, nicht interaktiv) — ohne diesen einmaligen
 * manuellen Lauf schlagen DriveApp- bzw. UrlFetchApp-Aufrufe mit einer
 * "You do not have permission to call ..."-Meldung fehl.
 */
function authorizeDriveAccess() {
  var folder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  var testFolder = folder.createFolder('__autorisierungstest__');
  testFolder.setTrashed(true); // sofort wieder löschen, diente nur dem Schreibzugriffs-Check
  Logger.log('Drive-Zugriff (lesen + schreiben) OK: ' + folder.getName());

  if (AUTOMATION_WEBHOOK_URL) {
    UrlFetchApp.fetch(AUTOMATION_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ test: true, hinweis: 'Autorisierungs-Test aus authorizeDriveAccess()' }),
      muteHttpExceptions: true
    });
    Logger.log('Webhook-Zugriff OK.');
  }
}

function buildDetailsText(p, nextNumber) {
  return 'Anfrage #' + nextNumber + ' — ' + (p.kanton || '-') + '\n\n' +
    '--- Kontakt ---\n' +
    'Name / Vorname: ' + (p.name || '-') + '\n' +
    'Firma: ' + (p.firma || '-') + '\n' +
    'Straße & Nr.: ' + (p.strasse || '-') + '\n' +
    'PLZ / Ort: ' + (p.plz_ort || '-') + '\n' +
    'E-Mail: ' + (p.email || '-') + '\n' +
    'Telefon: ' + (p.telefon || '-') + '\n\n' +
    '--- Fahrzeug ---\n' +
    'Automarke: ' + (p.marke || '-') + '\n' +
    'Typ: ' + (p.modell || '-') + '\n' +
    'Typenscheinnummer: ' + (p.typenscheinnummer || '-') + '\n' +
    'Erste Inverkehrssetzung: ' + (p.jahrgang || '-') + '\n' +
    'Kilometerstand: ' + (p.km || '-') + '\n' +
    'Farbe: ' + (p.farbe || '-') + '\n' +
    'Preisvorstellung: ' + (p.preisvorstellung || '-') + '\n' +
    'Fahrzeugart: ' + (p.typ || '-') + '\n' +
    'Bemerkungen: ' + (p.bemerkungen || '-') + '\n\n' +
    '--- Herkunft ---\n' +
    'Kanton-Seite: ' + (p.kanton || '-') + '\n' +
    'utm_source: ' + (p.utm_source || '-') + '\n' +
    'utm_medium: ' + (p.utm_medium || '-') + '\n' +
    'utm_campaign: ' + (p.utm_campaign || '-') + '\n' +
    'utm_term: ' + (p.utm_term || '-') + '\n' +
    'gclid: ' + (p.gclid || '-') + '\n' +
    'fbid: ' + (p.fbid || '-');
}

// Ab dieser Gesamtgrösse (Rohbytes) werden Bilder nicht mehr an die Mail
// angehängt, sondern nur noch der Drive-Link verschickt — Gmail-Anhanglimit
// liegt bei 25 MB pro Mail, hier bewusst Sicherheitsmarge eingeplant.
var MAX_EMAIL_ATTACHMENT_BYTES = 18 * 1024 * 1024;

function sendNotificationEmail(p, nextNumber, attachments, driveFolderUrl) {
  var to = 'scale.my.business.online@gmail.com';
  var subject = 'AAS Neue Anfrage #' + nextNumber + ': ' + (p.name || '') + ' (' + (p.kanton || '') + ')';

  var attachmentBytes = (attachments || []).reduce(function (sum, blob) {
    return sum + blob.getBytes().length;
  }, 0);
  var tooBigForEmail = attachmentBytes > MAX_EMAIL_ATTACHMENT_BYTES;

  var bilderZeile;
  if (!attachments || attachments.length === 0) {
    bilderZeile = 'Keine Bilder hochgeladen.';
  } else if (tooBigForEmail) {
    bilderZeile = attachments.length + ' Bild(er) — zu gross für Mail-Anhang, siehe Drive-Link unten.';
  } else {
    bilderZeile = attachments.length + ' Bild(er) angehängt.';
  }

  var body = buildDetailsText(p, nextNumber) + '\n\n' +
    '--- Bilder ---\n' +
    bilderZeile +
    (driveFolderUrl ? '\n\n--- Drive-Sicherung ---\n' + driveFolderUrl : '');

  MailApp.sendEmail(to, subject, body, {
    attachments: tooBigForEmail ? [] : (attachments || [])
  });
}

/**
 * Sendet die Anfrage als JSON an einen externen Automation-Webhook (Make.com/Zapier/n8n …).
 * Läuft parallel zu Sheet/E-Mail/Drive und blockiert diese bei einem Fehler nicht (try/catch).
 * Bilder werden NICHT mitgeschickt (zu groß für Webhook-Payloads) — bei Bedarf kann
 * driveFolderUrl im Ziel-Tool nachgeladen werden, um an die Originalbilder zu kommen.
 */
function sendToAutomationWebhook(p, nextNumber, driveFolderUrl) {
  if (!AUTOMATION_WEBHOOK_URL) return;

  var payload = {
    lfdNr: nextNumber,
    zeitstempel: new Date().toISOString(),
    name: p.name || '',
    firma: p.firma || '',
    strasse: p.strasse || '',
    plzOrt: p.plz_ort || '',
    email: p.email || '',
    telefon: p.telefon || '',
    automarke: p.marke || '',
    typ: p.modell || '',
    typenscheinnummer: p.typenscheinnummer || '',
    ersteInverkehrssetzung: p.jahrgang || '',
    kilometerstand: p.km || '',
    farbe: p.farbe || '',
    preisvorstellung: p.preisvorstellung || '',
    fahrzeugart: p.typ || '',
    bemerkungen: p.bemerkungen || '',
    kantonSeite: p.kanton || '',
    utmSource: p.utm_source || '',
    utmMedium: p.utm_medium || '',
    utmCampaign: p.utm_campaign || '',
    utmTerm: p.utm_term || '',
    gclid: p.gclid || '',
    fbid: p.fbid || '',
    driveOrdner: driveFolderUrl || ''
  };

  try {
    UrlFetchApp.fetch(AUTOMATION_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('Automation-Webhook fehlgeschlagen: ' + err);
  }
}
