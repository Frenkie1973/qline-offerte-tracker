// Q-Line mailbox bot
// Leest nieuwe mails uit store@q-line.com via Microsoft Graph (app-only) en
// zet ze als concept-leads (status "nieuw", tag auto:true) in de Firebase
// Realtime Database, in hetzelfde formaat als het dashboard zelf gebruikt.
//
// Draait via GitHub Actions (.github/workflows/mailbot.yml), elke 15 min.
// Benodigde secrets: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET

const MAILBOX = 'store@q-line.com';
const FB_BASE = 'https://q-line-tracker-default-rtdb.europe-west1.firebasedatabase.app';
const LEADS_URL = `${FB_BASE}/leads.json`;
const STATE_URL = `${FB_BASE}/mailbotState.json`;

const TENANT_ID = process.env.MS_TENANT_ID;
const CLIENT_ID = process.env.MS_CLIENT_ID;
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;

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
  // NL/DE/BE stijl telefoonnummers, met of zonder landcode
  const matches = text.match(/(\+?\d[\d\s\-\/\(\)]{7,}\d)/g) || [];
  for (const m of matches) {
    const digits = m.replace(/[^\d]/g, '');
    if (digits.length >= 9 && digits.length <= 14) return m.trim();
  }
  return '';
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
  if (!res.ok) throw new Error(`Token ophalen mislukt: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.access_token;
}

async function fetchNewMessages(token, sinceISO) {
  const filter = encodeURIComponent(`receivedDateTime gt ${sinceISO}`);
  const select = 'id,subject,from,receivedDateTime,bodyPreview,body';
  const url = `https://graph.microsoft.com/v1.0/users/${MAILBOX}/messages?$filter=${filter}&$select=${select}&$orderby=receivedDateTime asc&$top=50`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Mail ophalen mislukt: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.value || [];
}

async function getJson(url, fallback) {
  try {
    const res = await fetch(url);
    const data = await res.json();
    return data === null ? fallback : data;
  } catch (e) {
    return fallback;
  }
}

async function putJson(url, data) {
  const res = await fetch(url, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  if (!res.ok) throw new Error(`Firebase schrijven mislukt: ${res.status} ${await res.text()}`);
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    console.error('Ontbrekende secrets (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET).');
    process.exit(1);
  }

  const state = await getJson(STATE_URL, {});
  const sinceISO = state.lastCheck || new Date(Date.now() - 24 * 3600 * 1000).toISOString();
  const processedIds = new Set(state.processedIds || []);

  console.log(`Ophalen mails sinds ${sinceISO}...`);
  const token = await getAppToken();
  const messages = await fetchNewMessages(token, sinceISO);
  console.log(`${messages.length} mail(s) gevonden.`);

  const nieuwLeads = [];
  let latest = sinceISO;

  for (const m of messages) {
    if (processedIds.has(m.id)) continue;
    processedIds.add(m.id);
    if (m.receivedDateTime > latest) latest = m.receivedDateTime;

    const fromName = m.from?.emailAddress?.name || '';
    const fromEmail = m.from?.emailAddress?.address || '';
    const bodyText = m.body?.contentType === 'html' ? stripHtml(m.body.content) : (m.body?.content || m.bodyPreview || '');
    const telefoon = extractPhone(bodyText) || extractPhone(m.bodyPreview || '');

    nieuwLeads.push({
      id: `mail_${m.id}`.replace(/[^a-zA-Z0-9_]/g, '').slice(0, 40) + '_' + Date.now().toString(36),
      klant: fromName || fromEmail || 'Onbekende afzender',
      email: fromEmail,
      telefoon,
      onderwerp: m.subject || 'Mail zonder onderwerp',
      datum: todayISO(),
      status: 'nieuw',
      notitie: `🤖 Automatisch verwerkt uit mailbox — controleer en vul aan.\n\nOnderwerp: ${m.subject || ''}\nOntvangen: ${new Date(m.receivedDateTime).toLocaleString('nl-NL')}\n\n${bodyText.slice(0, 800)}`,
      contacts: [],
      auto: true,
    });
  }

  if (nieuwLeads.length) {
    const leads = await getJson(LEADS_URL, []);
    const updated = Array.isArray(leads) ? leads.concat(nieuwLeads) : nieuwLeads;
    await putJson(LEADS_URL, updated);
    console.log(`${nieuwLeads.length} nieuwe lead(s) toegevoegd aan Firebase.`);
  } else {
    console.log('Geen nieuwe leads om toe te voegen.');
  }

  const trimmedIds = Array.from(processedIds).slice(-500);
  await putJson(STATE_URL, { lastCheck: latest, processedIds: trimmedIds });
  console.log('Klaar.');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
