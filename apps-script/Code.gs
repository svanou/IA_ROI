/**
 * Webinar Galadrim — IA & ROI
 * Backend d'inscription (Google Apps Script, déployé en « Web App »).
 *
 * Reçoit les inscriptions du formulaire de la landing page, puis :
 *   1. ajoute l'inscrit comme invité d'UN événement Google Agenda unique
 *      (liste d'invités masquée -> les participants ne se voient pas entre eux),
 *   2. déclenche l'invitation calendrier native de Google (sendUpdates: 'all'),
 *   3. consigne l'inscription dans un Google Sheet (optionnel).
 *
 * L'événement est créé automatiquement au premier appel, puis réutilisé
 * (son id est mémorisé dans les propriétés du script).
 *
 * --- DÉPLOIEMENT (à faire une fois) ---
 *  1. script.google.com (connecté en sylvain@galadrim.ch) -> Nouveau projet.
 *  2. Colle ce fichier dans Code.gs.
 *  3. Ajuste le bloc CONFIG ci-dessous (titre, lien Zoom, description).
 *  4. Services (+) -> ajoute « Calendar API » (service avancé, identifiant: Calendar).
 *  5. (option) crée une Google Sheet, copie son ID, et mets-le dans SHEET_ID
 *     ci-dessous — ou laisse vide pour ne pas logguer.
 *  6. Déployer -> Nouveau déploiement -> type « Application web ».
 *       - Exécuter en tant que : moi (sylvain@galadrim.ch)
 *       - Qui a accès : Tout le monde
 *     Autorise les accès demandés (Agenda + envoi de mail).
 *  7. Copie l'URL .../exec et colle-la dans WEBHOOK_URL dans index.html.
 *
 * Pour tester sans la page : lance la fonction testInscription() une fois.
 */

const CONFIG = {
  CALENDAR_ID: 'primary',                         // agenda principal de sylvain@galadrim.ch
  EVENT_TITLE: 'Webinar Galadrim — IA & ROI',
  EVENT_START: '2026-07-13T12:30:00+02:00',       // doit rester aligné avec WEBINAR_DATE de la page
  EVENT_END:   '2026-07-13T13:30:00+02:00',       // +60 min
  TIMEZONE:    'Europe/Zurich',
  EVENT_LOCATION: 'En ligne (Zoom)',
  EVENT_DESCRIPTION:
    "IA & ROI : de l'identification du projet au calcul de la rentabilité.\n\n" +
    "Le lien de connexion Zoom vous sera envoyé 24h avant la session.\n\n" +
    "Galadrim Suisse — Genève & Lausanne",
  SHEET_ID: '',                                   // id d'une Google Sheet pour logguer (vide = pas de log)
};

/** Point d'entrée appelé par le formulaire (POST). */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // sérialise les écritures concurrentes sur la liste d'invités
  try {
    const data = parsePayload_(e);
    if (!data.email || !isEmail_(data.email)) {
      return json_({ ok: false, error: 'email_invalide' });
    }
    const eventId = getOrCreateEvent_();
    const status = addGuest_(eventId, data.email, data.prenom);
    logToSheet_(data, status);
    return json_({ ok: true, status: status });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/** Permet un test « santé » dans le navigateur. */
function doGet() {
  return json_({ ok: true, service: 'webinar-ia-roi' });
}

/** Crée l'événement une seule fois, puis renvoie son id mémorisé. */
function getOrCreateEvent_() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('EVENT_ID');
  if (cached) {
    try {
      Calendar.Events.get(CONFIG.CALENDAR_ID, cached); // existe toujours ?
      return cached;
    } catch (err) {
      props.deleteProperty('EVENT_ID'); // événement supprimé -> on en recrée un
    }
  }
  const created = Calendar.Events.insert({
    summary: CONFIG.EVENT_TITLE,
    location: CONFIG.EVENT_LOCATION,
    description: CONFIG.EVENT_DESCRIPTION,
    start: { dateTime: CONFIG.EVENT_START, timeZone: CONFIG.TIMEZONE },
    end:   { dateTime: CONFIG.EVENT_END,   timeZone: CONFIG.TIMEZONE },
    guestsCanSeeOtherGuests: false,  // <- participants masqués les uns aux autres
    guestsCanInviteOthers: false,
    guestsCanModify: false,
    reminders: { useDefault: false, overrides: [{ method: 'email', minutes: 24 * 60 }, { method: 'popup', minutes: 15 }] },
  }, CONFIG.CALENDAR_ID);
  props.setProperty('EVENT_ID', created.id);
  return created.id;
}

/** Ajoute l'invité à l'événement et envoie l'invitation. Renvoie 'added' ou 'already'. */
function addGuest_(eventId, email, displayName) {
  const event = Calendar.Events.get(CONFIG.CALENDAR_ID, eventId);
  const attendees = event.attendees || [];
  const exists = attendees.some(a => (a.email || '').toLowerCase() === email.toLowerCase());
  if (exists) return 'already';
  attendees.push({ email: email, displayName: displayName || undefined });
  Calendar.Events.patch(
    { attendees: attendees, guestsCanSeeOtherGuests: false },
    CONFIG.CALENDAR_ID,
    eventId,
    { sendUpdates: 'all' } // <- déclenche l'email d'invitation Google au nouvel inscrit
  );
  return 'added';
}

/** Log optionnel dans une Google Sheet. */
function logToSheet_(data, status) {
  if (!CONFIG.SHEET_ID) return;
  const sh = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Date', 'Prénom', 'Email', 'Entreprise', 'Source', 'Statut']);
  }
  sh.appendRow([new Date(), data.prenom, data.email, data.entreprise, data.source, status]);
}

/* ---------- helpers ---------- */

function parsePayload_(e) {
  let body = {};
  if (e && e.postData && e.postData.contents) {
    try { body = JSON.parse(e.postData.contents); } catch (_) { body = e.parameter || {}; }
  } else if (e && e.parameter) {
    body = e.parameter;
  }
  return {
    prenom: String(body.prenom || '').trim(),
    email: String(body.email || '').trim(),
    entreprise: String(body.entreprise || '').trim(),
    source: String(body.source || '').trim(),
  };
}

function isEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** À lancer une fois manuellement pour vérifier le flux (remplace l'email). */
function testInscription() {
  const out = doPost({ postData: { contents: JSON.stringify({
    prenom: 'Test', email: 'sylvain@galadrim.ch', entreprise: 'Galadrim', source: 'test',
  }) } });
  console.log(out.getContent());
}
