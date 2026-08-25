import { AwsClient } from 'aws4fetch';

// Presigned URLs bleiben lange genug gültig für langsame Mobilfunk-Uploads
// grosser Dateien (PUT) und für die Zeit zwischen Upload-Abschluss und
// Formular-Absenden, bis Apps Script die Datei per GET abholt.
var URL_EXPIRY_SECONDS = 3600;

// Einfache Missbrauchsbremse: nur Bild-/Video-Uploads erlauben. Ersetzt keine
// echte Bot-Abwehr (siehe Plan-Einschränkungen), reicht aber für das aktuelle
// Volumen.
var ALLOWED_MIME_PREFIXES = ['image/', 'video/'];

function sanitizeFilename(name) {
  var safe = String(name || 'upload').replace(/[^a-zA-Z0-9._-]/g, '_');
  return safe.slice(-80) || 'upload';
}

async function handlePresign(request, env) {
  var url = new URL(request.url);
  var filename = url.searchParams.get('filename');
  var mimeType = url.searchParams.get('mimeType') || 'application/octet-stream';

  var allowed = ALLOWED_MIME_PREFIXES.some(function (prefix) {
    return mimeType.indexOf(prefix) === 0;
  });
  if (!allowed) {
    return Response.json({ error: 'mimeType nicht erlaubt' }, { status: 400 });
  }

  var key = 'uploads/' + crypto.randomUUID() + '-' + sanitizeFilename(filename);
  var objectUrl = 'https://' + env.R2_ACCOUNT_ID + '.r2.cloudflarestorage.com/'
    + 'autoankauf-schweiz-uploads/' + key;

  var client = new AwsClient({
    accessKeyId: env.R2_ACCESS_KEY_ID,
    secretAccessKey: env.R2_SECRET_ACCESS_KEY,
    service: 's3',
    region: 'auto'
  });

  // Keine zusätzlichen Header mitsignieren — sonst muss der Browser beim PUT
  // exakt dieselben Header mitschicken, sonst schlägt die Signatur fehl.
  var putUrl = new URL(objectUrl);
  putUrl.searchParams.set('X-Amz-Expires', String(URL_EXPIRY_SECONDS));
  var signedPut = await client.sign(new Request(putUrl, { method: 'PUT' }), {
    aws: { signQuery: true }
  });

  var getUrl = new URL(objectUrl);
  getUrl.searchParams.set('X-Amz-Expires', String(URL_EXPIRY_SECONDS));
  var signedGet = await client.sign(new Request(getUrl, { method: 'GET' }), {
    aws: { signQuery: true }
  });

  return Response.json({
    uploadUrl: signedPut.url.toString(),
    downloadUrl: signedGet.url.toString()
  });
}

export default {
  async fetch(request, env) {
    var url = new URL(request.url);
    if (url.pathname === '/api/presign' && request.method === 'GET') {
      return handlePresign(request, env);
    }
    return new Response('Not found', { status: 404 });
  }
};
