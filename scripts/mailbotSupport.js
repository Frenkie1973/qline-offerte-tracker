// Q-Line support-mailbox bot
// Leest nieuwe mails uit support@q-line.com via Microsoft Graph (app-only) en
// zet ze automatisch als supportmeldingen ("storingen") in Firestore, in
// hetzelfde formaat als het dashboard zelf gebruikt (index.html).
//
// Dit is een BEWUST APARTE bot van scripts/mailbot.js (die blijft uitsluitend
// voor de sales-mailbox store@q-line.com en wordt door dit script niet
// aangeraakt). Support-meldingen leven in een andere database (Firestore)
// dan sales-leads (Firebase Realtime Database), dus dit script praat met een
// andere opslag en gebruikt een ander inlog-mechanisme (zie hieronder).
//
// Draait via GitHub Actions (.github/workflows/mailbot-support.yml), elke
// 15 min.
// Benodigde secrets:
//   MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET  (dezelfde Microsoft-app als
//     bij de sales-bot — vereist WEL dat die app ook leesrechten heeft op de
//     support@q-line.com-mailbox; zie het bericht aan Frank hierover)
//   FIREBASE_TEAM_EMAIL, FIREBASE_TEAM_PASSWORD    (hetzelfde teamaccount
//     waarmee ook op de website wordt ingelogd — de bot logt hiermee net zo
//     in als een gewone gebruiker, en krijgt zo schrijfrechten in Firestore)
//   FIREBASE_DB_SECRET  (dezelfde die de sales-bot ook al gebruikt — hier
//     alleen gebruikt om bij te houden welke mails al verwerkt zijn, in de
//     Realtime Database, net als bij de sales-bot; de supportmeldingen zelf
//     gaan naar Firestore, zie hierboven)

const MAILBOX = 'support@q-line.com';
// Frank (en eventuele collega's) beantwoorden klanten soms vanuit de gedeelde
// support-mailbox, soms vanuit zijn eigen mailbox — daarom worden de
// "Verzonden items" van BEIDE mailboxen gecheckt, zodat een reactie altijd
// automatisch in de geschiedenis van de melding terechtkomt.
const SENT_MAILBOXES = ['support@q-line.com', 'f.timmerhuis@q-line.com'];

const TENANT_ID = process.env.MS_TENANT_ID;
const CLIENT_ID = process.env.MS_CLIENT_ID;
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
const FIREBASE_TEAM_EMAIL = process.env.FIREBASE_TEAM_EMAIL;
const FIREBASE_TEAM_PASSWORD = process.env.FIREBASE_TEAM_PASSWORD;

const FB_API_KEY = 'AIzaSyC24CbCjNFSvTR6sbslJZT6_NRQQpgN_ys';
const FS_BASE = 'https://firestore.googleapis.com/v1/projects/q-line-tracker/databases/(default)/documents';

// Statusbijhouding (welke mails al verwerkt zijn) gaat naar de Realtime
// Database, net als bij de sales-bot — zo hoeft er geen nieuwe Firestore-
// collectie (en dus geen wijziging van de Firestore-beveiligingsregels) bij
// te komen enkel voor dit boekhoudkundige stukje.
const RTDB_BASE = 'https://q-line-tracker-default-rtdb.europe-west1.firebasedatabase.app';
const DB_SECRET = process.env.FIREBASE_DB_SECRET;
const DB_AUTH_SUFFIX = DB_SECRET ? `?auth=${DB_SECRET}` : '';
const STATE_URL = `${RTDB_BASE}/supportMailbotState.json${DB_AUTH_SUFFIX}`;
async function getState() {
  try {
    const res = await fetch(STATE_URL);
    const data = await res.json();
    return data || {};
  } catch (e) { return {}; }
}
async function putState(data) {
  const res = await fetch(STATE_URL, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
  if (!res.ok) throw new Error(`Status opslaan in Realtime Database mislukt: ${res.status} ${await res.text()}`);
}

// ===== Firestore REST helpers (zelfde logica als in index.html) =====
function fsValue(v) {
  if (v === null || v === undefined || v === '') return { nullValue: null };
  if (typeof v === 'number') return Number.isInteger(v) ? { integerValue: String(v) } : { doubleValue: v };
  if (typeof v === 'boolean') return { booleanValue: v };
  if (Array.isArray(v)) return { arrayValue: { values: v.map(fsValue) } };
  if (typeof v === 'object') return { mapValue: { fields: fsEncode(v) } };
  return { stringValue: String(v) };
}
function fsEncode(obj) {
  const fields = {};
  for (const k in obj) { if (obj[k] !== undefined) fields[k] = fsValue(obj[k]); }
  return fields;
}
function fsDecodeValue(v) {
  if (!v) return null;
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('timestampValue' in v) return v.timestampValue;
  if ('nullValue' in v) return null;
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(fsDecodeValue);
  if ('mapValue' in v) return fsDecode(v.mapValue.fields || {});
  return null;
}
function fsDecode(fields) {
  const obj = {};
  for (const k in fields) obj[k] = fsDecodeValue(fields[k]);
  return obj;
}
async function fsList(token, col) {
  const res = await fetch(`${FS_BASE}/${col}?pageSize=300`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Firestore lezen (${col}) mislukt: HTTP ${res.status} ${await res.text()}`);
  const data = await res.json();
  if (!data.documents) return [];
  return data.documents.map((d) => ({ id: d.name.split('/').pop(), ...fsDecode(d.fields || {}) }));
}
async function fsCreate(token, col, obj) {
  const res = await fetch(`${FS_BASE}/${col}`, { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: fsEncode(obj) }) });
  if (!res.ok) throw new Error(`Firestore aanmaken (${col}) mislukt: HTTP ${res.status} ${await res.text()}`);
  return res.json();
}
async function fsUpdate(token, col, id, obj) {
  const mask = Object.keys(obj).map((k) => `updateMask.fieldPaths=${encodeURIComponent(k)}`).join('&');
  const res = await fetch(`${FS_BASE}/${col}/${id}?${mask}`, { method: 'PATCH', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: fsEncode(obj) }) });
  if (!res.ok) throw new Error(`Firestore bijwerken (${col}/${id}) mislukt: HTTP ${res.status} ${await res.text()}`);
  return res.json();
}

// Logt in met het gedeelde teamaccount (zelfde account als op de website) om
// een geldig token te krijgen. Geen refresh-logica nodig: elke bot-run start
// een verse container en heeft maar één inlogpoging per run nodig.
async function signIn() {
  if (!FIREBASE_TEAM_EMAIL || !FIREBASE_TEAM_PASSWORD) {
    throw new Error('Ontbrekende secrets FIREBASE_TEAM_EMAIL / FIREBASE_TEAM_PASSWORD.');
  }
  const res = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${FB_API_KEY}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email: FIREBASE_TEAM_EMAIL, password: FIREBASE_TEAM_PASSWORD, returnSecureToken: true }),
  });
  if (!res.ok) throw new Error(`Inloggen bij Firebase mislukt: HTTP ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.idToken;
}

// ===== Mail-parsing helpers (zelfde stijl als scripts/mailbot.js) =====
function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|li)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
function extractPhone(text) {
  if (!text) return '';
  const matches = text.match(/(\+?\d[\d\s\-\/\(\)]{7,}\d)/g) || [];
  for (const m of matches) {
    const digits = m.replace(/[^\d]/g, '');
    if (digits.length >= 9 && digits.length <= 14) return m.trim();
  }
  return '';
}
function extractField(text, labelPatterns) {
  for (const label of labelPatterns) {
    const re = new RegExp(`^[\\s>*-]*${label}[\\s:\\-]+(.+)$`, 'im');
    const m = text.match(re);
    if (m) {
      const v = m[1].trim().replace(/\s{2,}/g, ' ');
      if (v) return v;
    }
  }
  return '';
}
function extractSignatureName(text) {
  const m = text.match(/(?:met\s+vriendelijke\s+groet(?:en)?|groet(?:en)?|mvg|regards|kind\s+regards)[,.]?[ \t]*\n+[ \t]*([A-Z][A-Za-zÀ-ÿ'\-]+(?:[ \t]+[A-Z][A-Za-zÀ-ÿ'\-]+){0,3})/i);
  return m ? m[1].trim() : '';
}

// Zoekt structuurvelden voor een SUPPORT-melding (i.p.v. een sales-aanvraag).
function extractSupportData(bodyText, fromName, fromEmail) {
  const cleanFromName = (fromName || '').replace(/^namens\s+/i, '').trim();
  const klant = extractField(bodyText, ['naam', 'name', 'klant', 'contactpersoon'])
    || extractSignatureName(bodyText)
    || cleanFromName || '';
  const email = extractField(bodyText, ['e-?mail(?:adres)?', 'email']) || fromEmail || '';
  const telefoon = extractField(bodyText, ['telefoon(?:nummer)?', 'tel(?:efoonnr)?', 'phone\\s*number', 'phone', 'mobiel']) || extractPhone(bodyText);
  const product = extractField(bodyText, ['product', 'machine', 'type\\s*machine']);
  const model = extractField(bodyText, ['model', 'type']);
  const serienummer = extractField(bodyText, ['serienummer', 'serie\\s*nr\\.?', 'serial\\s*number', 'serienr\\.?']);
  const mkgNummer = extractField(bodyText, ['mkg\\s*nr\\.?', 'mkg\\s*nummer', 'mkgnr', 'mkgnummer']);
  const foutcode = extractField(bodyText, ['foutcode', 'fout\\s*code', 'error\\s*code', 'storing(?:scode)?']);
  const omschrijvingVeld = extractField(bodyText, ['omschrijving', 'probleem', 'klacht']);
  return { klant, email, telefoon, product, model, serienummer, mkgNummer, foutcode, omschrijvingVeld };
}

async function getAppToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const res = await fetch(url, { method: 'POST', body });
  if (!res.ok) throw new Error(`Microsoft-token ophalen mislukt: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.access_token;
}
async function fetchNewMessages(msToken, sinceISO) {
  const filter = encodeURIComponent(`receivedDateTime gt ${sinceISO}`);
  const select = 'id,subject,from,receivedDateTime,bodyPreview,body';
  const url = `https://graph.microsoft.com/v1.0/users/${MAILBOX}/mailFolders/inbox/messages?$filter=${filter}&$select=${select}&$orderby=receivedDateTime asc&$top=100`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${msToken}` } });
  if (!res.ok) throw new Error(`Support-mail ophalen mislukt: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.value || [];
}
async function fetchSentMessages(msToken, sinceISO) {
  const filter = encodeURIComponent(`sentDateTime gt ${sinceISO}`);
  const select = 'id,subject,toRecipients,sentDateTime,bodyPreview,body';
  const alle = [];
  for (const mailbox of SENT_MAILBOXES) {
    const url = `https://graph.microsoft.com/v1.0/users/${mailbox}/mailFolders/sentitems/messages?$filter=${filter}&$select=${select}&$orderby=sentDateTime asc&$top=100`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${msToken}` } });
    if (!res.ok) {
      console.error(`Verzonden mail ophalen mislukt voor ${mailbox}: ${res.status} ${await res.text()}`);
      continue;
    }
    const data = await res.json();
    alle.push(...(data.value || []));
  }
  const gezien = new Set();
  return alle.filter((m) => {
    if (gezien.has(m.id)) return false;
    gezien.add(m.id);
    return true;
  });
}
function normTel(t) { return (t || '').replace(/[^\d+]/g, ''); }

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    console.error('Ontbrekende secrets (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET).');
    process.exit(1);
  }
  if (!DB_SECRET) {
    console.error('Ontbrekend secret FIREBASE_DB_SECRET.');
    process.exit(1);
  }

  const fbToken = await signIn();
  console.log('Ingelogd bij Firebase met teamaccount.');

  // Status (welke mails al verwerkt zijn) staat in de Realtime Database,
  // los van de supportmeldingen zelf (die in Firestore staan).
  const state = await getState();

  // Zelfde principe als de sales-bot: altijd minstens 7 dagen teruglezen,
  // dubbele verwerking wordt voorkomen via processedIds, niet via dit venster.
  const zevenDagenTerug = new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString();
  const sinceISO = (state.lastCheck && state.lastCheck < zevenDagenTerug) ? state.lastCheck : zevenDagenTerug;
  const processedIds = new Set(state.processedIds || []);

  console.log(`Ophalen support-mails sinds ${sinceISO}...`);
  const msToken = await getAppToken();
  const messages = await fetchNewMessages(msToken, sinceISO);
  console.log(`${messages.length} mail(s) gevonden.`);

  const storingen = await fsList(fbToken, 'storingen');
  const GESLOTEN_STATUSSEN = ['opgelost', 'gesloten'];
  const EIGEN_ADRESSEN = ['support@q-line.com', 'f.timmerhuis@q-line.com', 'store@q-line.com'];

  let latest = sinceISO;
  let nieuweTeller = 0, uniekTeller = 0, samengevoegdTeller = 0;

  for (const m of messages) {
    if (processedIds.has(m.id)) continue;
    processedIds.add(m.id);
    if (m.receivedDateTime > latest) latest = m.receivedDateTime;

    const fromName = m.from?.emailAddress?.name || '';
    const fromEmail = m.from?.emailAddress?.address || '';
    const bodyText = m.body?.contentType === 'html' ? stripHtml(m.body.content) : (m.body?.content || m.bodyPreview || '');
    const ex = extractSupportData(bodyText, fromName, fromEmail);

    let emailKey = (ex.email || '').toLowerCase().trim();
    if (EIGEN_ADRESSEN.includes(emailKey)) emailKey = '';
    const telKey = normTel(ex.telefoon);
    const onderwerp = m.subject || 'Support-mail zonder onderwerp';

    // Bestaande OPEN melding van dezelfde klant zoeken (net als matchOpenStoring
    // in het dashboard) — op e-mail of telefoonnummer.
    const bestaande = storingen.find((s) => {
      if (GESLOTEN_STATUSSEN.includes(s.status)) return false;
      const sEmail = (s.email || '').toLowerCase().trim();
      const sTel = normTel(s.telefoon);
      if (emailKey && sEmail && sEmail === emailKey) return true;
      if (telKey && sTel && sTel === telKey) return true;
      return false;
    });

    const ontvangen = new Date(m.receivedDateTime);
    const historieRegel = {
      tijdstip: m.receivedDateTime,
      tekst: `🤖 Nieuwe mail: ${onderwerp}\n\n${bodyText.slice(0, 1500)}`,
      bron: 'Automatisch (support-mailbox)',
      richting: 'van_klant',
    };

    if (bestaande) {
      const historie = [...(bestaande.historie || []), historieRegel];
      const wijziging = { historie };
      if (!bestaande.telefoon && ex.telefoon) wijziging.telefoon = ex.telefoon;
      if (!bestaande.mkgNummer || bestaande.mkgNummer === 'Nog opvragen') { if (ex.mkgNummer) wijziging.mkgNummer = ex.mkgNummer; }
      if (!bestaande.product && ex.product) wijziging.product = ex.product;
      try {
        await fsUpdate(fbToken, 'storingen', bestaande.id, wijziging);
        bestaande.historie = historie;
        Object.assign(bestaande, wijziging);
        samengevoegdTeller++;
      } catch (e) {
        console.error(`Bijwerken melding ${bestaande.id} mislukt:`, e.message);
      }
    } else {
      const nieuw = {
        klant: ex.klant || fromEmail || 'Onbekende afzender',
        contactpersoon: '',
        telefoon: ex.telefoon || '',
        email: ex.email || fromEmail || '',
        product: ex.product || '',
        model: ex.model || '',
        serienummer: ex.serienummer || 'Nog opvragen',
        mkgNummer: ex.mkgNummer || 'Nog opvragen',
        omschrijving: ex.omschrijvingVeld || bodyText.slice(0, 1500),
        foutcode: ex.foutcode || '',
        urgentie: 'normaal',
        status: 'nieuw',
        bron: '🤖 Automatisch verwerkt uit support@-mailbox — controleer en vul aan.',
        aangemaakt: ontvangen.toISOString(),
        historie: [historieRegel],
      };
      try {
        const created = await fsCreate(fbToken, 'storingen', nieuw);
        nieuw.id = created.name.split('/').pop();
        storingen.push(nieuw);
        uniekTeller++;
      } catch (e) {
        console.error('Aanmaken nieuwe melding mislukt:', e.message);
      }
    }
    nieuweTeller++;
  }

  // === Verzonden mail (Frank/support -> klant) automatisch in de geschiedenis loggen ===
  console.log('Ophalen verzonden support-mail...');
  const sentMessages = await fetchSentMessages(msToken, sinceISO);
  console.log(`${sentMessages.length} verzonden mail(s) gevonden.`);
  let verzondenGelogd = 0;

  for (const sm of sentMessages) {
    if (processedIds.has(sm.id)) continue;
    processedIds.add(sm.id);
    if (sm.sentDateTime > latest) latest = sm.sentDateTime;

    const ontvangers = (sm.toRecipients || []).map((r) => (r.emailAddress?.address || '').toLowerCase().trim());
    if (!ontvangers.length) continue;

    const doelMelding = storingen.find((s) => ontvangers.includes((s.email || '').toLowerCase().trim()));
    if (!doelMelding) continue;

    const sentBodyText = sm.body?.contentType === 'html' ? stripHtml(sm.body.content) : (sm.body?.content || sm.bodyPreview || '');
    const verzonden = new Date(sm.sentDateTime);
    const historie = [...(doelMelding.historie || []), {
      tijdstip: sm.sentDateTime,
      tekst: sentBodyText.slice(0, 1500) || sm.subject || 'Mail verstuurd',
      bron: 'Automatisch (support-mailbox)',
      richting: 'naar_klant',
    }];
    try {
      await fsUpdate(fbToken, 'storingen', doelMelding.id, { historie });
      doelMelding.historie = historie;
      verzondenGelogd++;
      nieuweTeller++;
    } catch (e) {
      console.error(`Loggen verzonden mail bij melding ${doelMelding.id} mislukt:`, e.message);
    }
  }

  if (nieuweTeller) {
    console.log(`${nieuweTeller} mail(s) verwerkt: ${uniekTeller} nieuwe melding(en), ${samengevoegdTeller} samengevoegd met bestaande melding, ${verzondenGelogd} verzonden mail(s) gelogd.`);
  } else {
    console.log('Geen nieuwe support-mail om te verwerken.');
  }

  const trimmedIds = Array.from(processedIds).slice(-500);
  try {
    await putState({ lastCheck: latest, processedIds: trimmedIds });
  } catch (e) {
    console.error('Status opslaan mislukt:', e.message);
  }
  console.log('Klaar.');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
