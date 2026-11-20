const RELAY_VERSION = '1.1.0'; // LINE 1 = the only version line in this file. Bump on every change.
/* EcoOils Walkthrough - Mail Relay. Deploy this under ptecooilsmentari@gmail.com ONLY.
   It does one job: send an email that the main Walkthrough script hands to it, so the mail goes out
   from this account and uses this account's daily quota. Nothing else is stored or exposed.
   Setup: 1) run setup() once and allow permissions; 2) copy the secret it logs;
   Optional: add Script Property REPLY_TO (your company email) so every reply is forced there, whatever the main script sends.
   3) Deploy > New deployment > Web app: Execute as Me, access Anyone; 4) copy the /exec URL. */

function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('RELAY_SECRET')) p.setProperty('RELAY_SECRET', Utilities.getUuid() + Utilities.getUuid());
  MailApp.getRemainingDailyQuota(); // forces the mail permission prompt
  Logger.log('RELAY_SECRET = ' + p.getProperty('RELAY_SECRET'));
}

function doGet() { return out_({ ok: true, relay: RELAY_VERSION, remaining: MailApp.getRemainingDailyQuota(), from: Session.getEffectiveUser().getEmail() }); }

function doPost(e) {
  try {
    const r = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const secret = PropertiesService.getScriptProperties().getProperty('RELAY_SECRET');
    if (!secret || r.secret !== secret) return out_({ ok: false, error: 'DENIED' });
    if (!r.to || !r.subject) return out_({ ok: false, error: 'BAD' });
    if (MailApp.getRemainingDailyQuota() < 3) return out_({ ok: false, error: 'QUOTA' });
    const m = { to: String(r.to), subject: String(r.subject).slice(0, 250), body: String(r.body || ''), name: String(r.name || 'EcoOils Project HSSE') };
    if (r.cc) m.cc = String(r.cc);
    const rt = PropertiesService.getScriptProperties().getProperty('REPLY_TO') || r.replyTo;
    if (rt) m.replyTo = String(rt);
    if (r.attachments && r.attachments.length) {
      m.attachments = r.attachments.map(function (a) { return Utilities.newBlob(Utilities.base64Decode(a.b64), a.type || 'image/jpeg', a.name || 'photo.jpg'); });
    }
    MailApp.sendEmail(m);
    return out_({ ok: true, remaining: MailApp.getRemainingDailyQuota() });
  } catch (err) { return out_({ ok: false, error: String(err).slice(0, 200) }); }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
