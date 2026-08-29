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
 * - Grosse Anhänge (>30 MB gesamt) laufen nicht mehr über dieses Apps-Script
 *   (POST-Body-Limit von ~50 MB), sondern werden vom Browser direkt zu
 *   Cloudflare R2 hochgeladen (presigned URLs, ausgestellt vom Cloudflare
 *   Worker der Website — siehe worker/index.js im Repo). Dieses Backend
 *   bekommt dafür im Formular-POST nur noch das Feld bilder_r2 mit
 *   {name, type, downloadUrl}-Einträgen und lädt die Bytes per UrlFetchApp
 *   von der jeweiligen (befristet gültigen) R2-Download-URL nach.
 *   (Ein früherer Versuch, Bilder direkt zu Google Drive hochzuladen, scheiterte
 *   an fehlender CORS-Unterstützung von Googles Resumable-Upload-Endpoint.)
 * - Benachrichtigungs-Mail nutzt jetzt immer eine HTML-Ansicht (buildDetailsHtml)
 *   mit fett dargestellten Feld-Bezeichnern, damit Schlüssel und Wert sich
 *   schneller unterscheiden lassen. Reihenfolge und Werte sind unverändert
 *   identisch zur bisherigen Klartext-Mail (buildDetailsText bleibt als
 *   Plain-Text-Fallback für Mail-Clients ohne HTML-Darstellung erhalten).
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
 * für kleine Uploads) und Bilder, die der Browser für grosse Uploads (>30 MB)
 * bereits direkt zu Cloudflare R2 hochgeladen hat (bilder_r2) — letztere werden
 * hier per befristet gültiger Download-URL abgeholt.
 */
function saveToDrive(p, nextNumber) {
  var inlineBlobs = parseImageAttachments(p.bilder_data);
  var r2Blobs = resolveR2Attachments(p.bilder_r2);
  var allBlobs = inlineBlobs.concat(r2Blobs);

  try {
    var root = getOrCreateRootFolder();
    var datum = Utilities.formatDate(new Date(), 'Europe/Zurich', 'yyyy-MM-dd_HH-mm');
    var nameFuerOrdner = String(p.name || 'Unbekannt').replace(/[\\\/:*?"<>|]/g, '').trim() || 'Unbekannt';
    var folder = root.createFolder('#' + nextNumber + ' – ' + datum + ' – ' + nameFuerOrdner);

    folder.createFile('anfrage.txt', buildDetailsText(p, nextNumber), MimeType.PLAIN_TEXT);

    allBlobs.forEach(function (blob) {
      folder.createFile(blob);
    });

    return { folderUrl: folder.getUrl(), attachments: allBlobs };
  } catch (err) {
    Logger.log('Drive-Backup fehlgeschlagen: ' + err);
    return { folderUrl: '', attachments: allBlobs };
  }
}

/**
 * Holt Bilder, die der Browser für grosse Uploads (>30 MB) bereits direkt zu
 * Cloudflare R2 hochgeladen hat (siehe worker/index.js), per befristet
 * gültiger presigned Download-URL. Schlägt eine einzelne Datei fehl (z. B.
 * abgelaufene URL), wird das geloggt und nur diese Datei übersprungen —
 * kein harter Abbruch der ganzen Anfrage.
 */
function resolveR2Attachments(bilderR2Json) {
  if (!bilderR2Json) return [];
  var items;
  try {
    items = JSON.parse(bilderR2Json);
  } catch (e) {
    return [];
  }
  if (!Array.isArray(items)) return [];

  var blobs = [];
  items.forEach(function (item) {
    if (!item || !item.downloadUrl) return;
    try {
      var response = UrlFetchApp.fetch(item.downloadUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        Logger.log('R2-Download fehlgeschlagen (' + item.name + '): HTTP ' + response.getResponseCode());
        return;
      }
      var blob = response.getBlob();
      blob.setName(item.name || 'bild.jpg');
      if (item.type) blob.setContentType(item.type);
      blobs.push(blob);
    } catch (err) {
      Logger.log('R2-Download fehlgeschlagen (' + item.name + '): ' + err);
    }
  });
  return blobs;
}

function getOrCreateRootFolder() {
  return DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
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

/**
 * HTML-Variante von buildDetailsText: identische Reihenfolge und Werte,
 * aber die Feld-Bezeichner (Schlüssel) fett dargestellt, damit sie sich in
 * der Mail schneller vom Wert unterscheiden lassen. Werte werden nur für
 * HTML escaped, inhaltlich nicht verändert.
 */
function buildDetailsHtml(p, nextNumber) {
  function line(label, value) {
    return '<b>' + escapeHtml(label) + ':</b> ' + escapeHtml(value || '-') + '<br>';
  }
  function section(title) {
    return '<br><b>--- ' + escapeHtml(title) + ' ---</b><br>';
  }

  return '<b>Anfrage #' + nextNumber + ' — ' + escapeHtml(p.kanton || '-') + '</b><br>' +
    section('Kontakt') +
    line('Name / Vorname', p.name) +
    line('Firma', p.firma) +
    line('Straße & Nr.', p.strasse) +
    line('PLZ / Ort', p.plz_ort) +
    line('E-Mail', p.email) +
    line('Telefon', p.telefon) +
    section('Fahrzeug') +
    line('Automarke', p.marke) +
    line('Typ', p.modell) +
    line('Typenscheinnummer', p.typenscheinnummer) +
    line('Erste Inverkehrssetzung', p.jahrgang) +
    line('Kilometerstand', p.km) +
    line('Farbe', p.farbe) +
    line('Preisvorstellung', p.preisvorstellung) +
    line('Fahrzeugart', p.typ) +
    line('Bemerkungen', p.bemerkungen) +
    section('Herkunft') +
    line('Kanton-Seite', p.kanton) +
    line('utm_source', p.utm_source) +
    line('utm_medium', p.utm_medium) +
    line('utm_campaign', p.utm_campaign) +
    line('utm_term', p.utm_term) +
    line('gclid', p.gclid) +
    line('fbid', p.fbid);
}

// Ab dieser Gesamtgrösse (Rohbytes) werden Bilder nicht mehr an die Mail
// angehängt, sondern nur noch der Drive-Link verschickt — Gmail-Anhanglimit
// liegt bei 25 MB pro Mail, hier bewusst Sicherheitsmarge eingeplant.
var MAX_EMAIL_ATTACHMENT_BYTES = 18 * 1024 * 1024;

// Nur genutzt, wenn die vollen Bilder wegen MAX_EMAIL_ATTACHMENT_BYTES nicht
// angehängt werden — gibt trotzdem einen kleinen visuellen Eindruck direkt in
// der Mail, ohne die Gmail-25-MB-Grenze zu gefährden. Sind die Bilder ohnehin
// als normaler Anhang dabei, zeigt Gmail dafür schon von sich aus Thumbnails an,
// eine zusätzliche Inline-Vorschau würde die Bytes nur doppelt zählen.
var MAX_INLINE_PREVIEW_BYTES = 6 * 1024 * 1024;

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Bettet die ersten Bilder (bis zu einem Byte-Budget) direkt per CID ins
 * HTML der Mail ein, klein dargestellt. Bleibt das Budget bei 0 (siehe
 * Aufrufer), wird gar nichts eingebettet.
 */
function buildInlinePreview(attachments, budgetBytes) {
  var inlineImages = {};
  var htmlParts = [];
  var usedBytes = 0;

  (attachments || []).forEach(function (blob, i) {
    if (usedBytes >= budgetBytes) return;
    var type = (blob.getContentType() || '').toLowerCase();
    if (type.indexOf('image/') !== 0) return;
    var bytes = blob.getBytes().length;
    if (usedBytes + bytes > budgetBytes) return;

    usedBytes += bytes;
    var cid = 'previewImg' + i;
    inlineImages[cid] = blob;
    htmlParts.push('<img src="cid:' + cid + '" width="150" style="margin:4px; border-radius:4px; border:1px solid #ddd;">');
  });

  return { html: htmlParts.join(''), inlineImages: inlineImages };
}

function sendNotificationEmail(p, nextNumber, attachments, driveFolderUrl) {
  // Mehrere Empfänger: MailApp.sendEmail akzeptiert eine kommagetrennte Liste
  var to = 'scale.my.business.online@gmail.com,info@autoankauf-schweiz.ch';
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

  var preview = buildInlinePreview(attachments, tooBigForEmail ? MAX_INLINE_PREVIEW_BYTES : 0);

  var htmlBody = '<div style="font-family:Arial,sans-serif;">'
    + buildDetailsHtml(p, nextNumber)
    + '<br><b>--- Bilder ---</b><br>' + escapeHtml(bilderZeile) + '<br>'
    + (driveFolderUrl ? '<br><b>--- Drive-Sicherung ---</b><br>' + escapeHtml(driveFolderUrl) + '<br>' : '')
    + (preview.html ? '<div style="margin-top:12px;">' + preview.html + '</div>' : '')
    + '</div>';

  var mailOptions = { attachments: tooBigForEmail ? [] : (attachments || []), htmlBody: htmlBody };
  if (preview.html) {
    mailOptions.inlineImages = preview.inlineImages;
  }

  MailApp.sendEmail(to, subject, body, mailOptions);
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
