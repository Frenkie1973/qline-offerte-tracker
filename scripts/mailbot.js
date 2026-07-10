// Q-Line mailbox bot
// Leest nieuwe mails uit store@q-line.com via Microsoft Graph (app-only) en
// zet ze als concept-leads (status "nieuw", tag auto:true) in de Firebase
// Realtime Database, in hetzelfde formaat als het dashboard zelf gebruikt.
//
// Draait via GitHub Actions (.github/workflows/mailbot.yml), elke 15 min.
// Benodigde secrets: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET

const MAILBOX = 'store@q-line.com';
// Verwerkte aanvragen worden door Frank verplaatst naar de map "Q-Line store
// aanvragen" (archief). De bot moet dus ALLEEN de Postvak IN uitlezen, nooit
// die archiefmap, anders komen al afgehandelde leads opnieuw binnen.
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

// Zoekt naar "Label: waarde" of "Label - waarde" op een eigen regel.
// Dekt zowel gestructureerde (webshop/formulier-)mails als losse tekst waarin
// de klant toevallig zulke labels gebruikt.
function extractField(text, labelPatterns) {
  for (const label of labelPatterns) {
    const re = new RegExp(`^[\\s>*-]*${label}\\s*[:\\-]\\s*(.+)$`, 'im');
    const m = text.match(re);
    if (m) {
      const v = m[1].trim().replace(/\s{2,}/g, ' ');
      if (v) return v;
    }
  }
  return '';
}

// Probeert een naam te vinden in de afsluiting van een losse mail
// ("Met vriendelijke groet, Jan Jansen") als er geen expliciet naam-label is.
// Let op: matcht bewust alleen binnen ÉÉN regel na de groet, anders wordt een
// bedrijfsnaam op de regel eronder per ongeluk aan de naam vastgeplakt.
function extractSignatureName(text) {
  const m = text.match(/(?:met\s+vriendelijke\s+groet(?:en)?|groet(?:en)?|mvg|regards|kind\s+regards)[,.]?[ \t]*\n+[ \t]*([A-Z][A-Za-zÀ-ÿ'\-]+(?:[ \t]+[A-Z][A-Za-zÀ-ÿ'\-]+){0,3})/i);
  return m ? m[1].trim() : '';
}

function extractStructuredData(bodyText, fromName, fromEmail) {
  const cleanFromName = (fromName || '').replace(/^namens\s+/i, '').trim();
  const klant = extractField(bodyText, ['naam', 'name', 'klant', 'contactpersoon'])
    || extractSignatureName(bodyText)
    || cleanFromName || '';
  const email = extractField(bodyText, ['e-?mail(?:adres)?', 'email']) || fromEmail || '';
  const telefoon = extractField(bodyText, ['telefoon(?:nummer)?', 'tel(?:efoonnr)?', 'phone', 'mobiel']) || extractPhone(bodyText);
  const straat = extractField(bodyText, ['adres', 'address', 'straat']);
  const postcode = extractField(bodyText, ['postcode', 'zip(?:code)?', 'plz']);
  const plaats = extractField(bodyText, ['plaats', 'city', 'woonplaats', 'stadt']);
  const landVeld = extractField(bodyText, ['land', 'country']);
  const product = extractField(bodyText, ['product', 'interesse\\s*in', 'onderwerp']);
  const contactWay = extractField(bodyText, ['contact\\s*voorkeur', 'voorkeur\\s*contact', 'preferred\\s*contact(?:\\s*way)?', 'liever\\s*(?:gebeld|gemaild)']);
  const kooptermijn = extractField(bodyText, ['kooptermijn', 'wanneer.*(?:kopen|aanschaf)', 'when.*(?:buy|purchase)']);
  const typeKlant = extractField(bodyText, ['type\\s*(?:bedrijf|klant)', 'soort\\s*bedrijf', 'business\\s*type']);

  const adres = [straat, postcode, plaats, landVeld].filter(Boolean).join(', ');

  return { klant, email, telefoon, land: landVeld, contactWay, product, kooptermijn, typeKlant, adres };
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
  // Expliciet de Postvak IN-map, nooit de hele mailbox (die bevat ook de
  // archiefmap "Q-Line store aanvragen" met al afgehandelde mail).
  const url = `https://graph.microsoft.com/v1.0/users/${MAILBOX}/mailFolders/inbox/messages?$filter=${filter}&$select=${select}&$orderby=receivedDateTime asc&$top=50`;
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

const AI_PROXY_URL = 'https://qline-ai-proxy.f-timmerhuis.workers.dev';

// Genereert automatisch een concept-opvolgbericht zodra een klant reageert op
// een bestaande lead, zodat Frank het al klaar ziet staan als hij de lead
// opent — geen los knipwerk meer nodig.
async function genereerAiVoorstel(lead, nieuweMailTekst) {
  const contacts = lead.contacts || [];
  const contactlog = contacts.length
    ? contacts.map((c) => `- [${c.type}, ${c.datum} ${c.tijd || ''}] ${c.tekst || ''}`).join('\n')
    : '(nog geen contactmomenten gelogd)';

  let outboundStreak = 0;
  for (const c of contacts) {
    if (c.type === 'reactie') outboundStreak = 0;
    else outboundStreak++;
  }

  const prompt = `Je helpt Frank, eigenaar van Q-Line Equestrian (paardensystemen: horsewalkers, solariums, hekwerk, stalinrichting), een kort antwoord te schrijven op een mail die een klant zojuist stuurde. Doel: soepel en snel reageren, in Franks eigen stijl: zakelijk maar informeel, direct en praktisch, geen overbodige poespas.

KLANTGEGEVENS:
- Naam/bedrijf: ${lead.klant}
- Interesse/product: ${lead.onderwerp || 'onbekend'}
- Bedrag offerte: ${lead.bedrag ? '€' + lead.bedrag : 'nog niet bekend'}
- Land: ${lead.land || 'onbekend'}
- Status: ${lead.status}

CONTACTGESCHIEDENIS (chronologisch, wat er al verstuurd/besproken is):
${contactlog}

DE NIEUWE MAIL DIE DE KLANT ZOJUIST STUURDE (beantwoord deze specifiek):
${nieuweMailTekst.slice(0, 2000)}

Schrijf een kort antwoord (max ~120 woorden) dat:
- direct ingaat op wat de klant nu vraagt/zegt
- past bij het land/de taal van de klant (Duitse klant = Duits in de juiste formaliteit; Nederlandse klant = Nederlands)
- Franks toon aanhoudt: direct, vriendelijk-zakelijk, geen overdreven verkooppraat
- een concrete volgende stap voorstelt waar mogelijk

Geef ALLEEN de tekst van het bericht terug, geen aanhef-uitleg of extra commentaar erbij, geen aanhalingstekens eromheen.`;

  try {
    const res = await fetch(AI_PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 1000,
        messages: [{ role: 'user', content: prompt }],
      }),
    });
    if (!res.ok) throw new Error(`AI-proxy gaf ${res.status}`);
    const data = await res.json();
    const tekst = (data.content || []).filter((c) => c.type === 'text').map((c) => c.text).join('\n').trim();
    return tekst || null;
  } catch (e) {
    console.error('AI-opvolgvoorstel genereren mislukt:', e.message);
    return null;
  }
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

  const leads = await getJson(LEADS_URL, []);
  const leadsArr = Array.isArray(leads) ? leads : [];

  let latest = sinceISO;
  let nieuweTeller = 0;
  let samengevoegdTeller = 0;
  let uniekTeller = 0;

  for (const m of messages) {
    if (processedIds.has(m.id)) continue;
    processedIds.add(m.id);
    if (m.receivedDateTime > latest) latest = m.receivedDateTime;

    const fromName = m.from?.emailAddress?.name || '';
    const fromEmail = m.from?.emailAddress?.address || '';
    const bodyText = m.body?.contentType === 'html' ? stripHtml(m.body.content) : (m.body?.content || m.bodyPreview || '');

    const ex = extractStructuredData(bodyText, fromName, fromEmail);

    const extraRegels = [
      ex.kooptermijn ? `Kooptermijn: ${ex.kooptermijn}` : '',
      ex.typeKlant ? `Type klant: ${ex.typeKlant}` : '',
      ex.adres ? `Adres: ${ex.adres}` : '',
      ex.contactWay ? `Voorkeur contact: ${ex.contactWay}` : '',
    ].filter(Boolean).join('\n');

    const notitie = [
      '🤖 Automatisch verwerkt uit mailbox — controleer en vul aan.',
      `Ontvangen: ${new Date(m.receivedDateTime).toLocaleString('nl-NL')}`,
      extraRegels,
      '',
      'Oorspronkelijke tekst:',
      bodyText.slice(0, 800),
    ].filter(Boolean).join('\n');

    const onderwerp = ex.product || m.subject || 'Mail zonder onderwerp';
    const emailKey = (ex.email || '').toLowerCase().trim();
    const klantKey = (ex.klant || fromEmail || '').toLowerCase().trim();
    const GESLOTEN_STATUSSEN = ['gewonnen', 'verloren', 'deadend'];

    // Zoek een nog openstaande lead (elke actieve status: nieuw, contact,
    // offerte, dealer — alleen gewonnen/verloren/deadend telt niet mee) van
    // dezelfde klant, op e-mail óf op naam, zodat een klant waar Frank al
    // mee bezig is niet als nieuwe losse lead wordt aangemaakt.
    const bestaande = leadsArr.find((d) => {
      if (GESLOTEN_STATUSSEN.includes(d.status)) return false;
      const dEmail = (d.email || '').toLowerCase().trim();
      const dKlant = (d.klant || '').toLowerCase().trim();
      if (emailKey && dEmail && dEmail === emailKey) return true;
      if (klantKey && dKlant && dKlant === klantKey) return true;
      return false;
    });

    if (bestaande) {
      if (onderwerp && !bestaande.onderwerp.includes(onderwerp)) {
        bestaande.onderwerp = `${bestaande.onderwerp} + ${onderwerp}`;
      }
      bestaande.notitie = `${bestaande.notitie}\n\n--- Extra aanvraag van dezelfde klant ---\n${notitie}`;
      bestaande.contacts = bestaande.contacts || [];
      const ontvangen = new Date(m.receivedDateTime);
      bestaande.contacts.push({
        type: 'reactie',
        datum: ontvangen.toLocaleDateString('nl-NL', { day: 'numeric', month: 'short' }),
        tijd: ontvangen.toLocaleTimeString('nl-NL', { hour: '2-digit', minute: '2-digit' }),
        tekst: `🤖 Nieuwe mail binnengekomen: ${onderwerp}`,
      });
      // Alleen vooruit schuiven, nooit terug (bij de eenmalige inbox-inhaalslag
      // kunnen mails niet-chronologisch verwerkt worden).
      if (m.receivedDateTime.slice(0, 10) > bestaande.datum) {
        bestaande.datum = m.receivedDateTime.slice(0, 10);
      }
      if (!bestaande.telefoon && ex.telefoon) bestaande.telefoon = ex.telefoon;
      if (!bestaande.land && ex.land) bestaande.land = ex.land;
      if (!bestaande.contactWay && ex.contactWay) bestaande.contactWay = ex.contactWay;

      // Klant reageerde op een bestaande lead -> alvast een concept-antwoord
      // klaarzetten zodat Frank het meteen ziet staan bij het openen.
      console.log(`Genereer AI-opvolgvoorstel voor ${bestaande.klant}...`);
      const voorstel = await genereerAiVoorstel(bestaande, bodyText);
      if (voorstel) {
        bestaande.aiVoorstel = voorstel;
        bestaande.aiVoorstelDatum = m.receivedDateTime;
      }
      samengevoegdTeller++;
    } else {
      uniekTeller++;
      leadsArr.push({
        // Gegarandeerd uniek: hash van het volledige Graph-ID + teller + random,
        // in plaats van het bericht-ID af te kappen (dat gaf collisions bij
        // berichten uit dezelfde mailthread met een gedeeld ID-prefix).
        id: `mail_${Date.now().toString(36)}_${uniekTeller}_${Math.random().toString(36).slice(2, 8)}`,
        klant: ex.klant || fromEmail || 'Onbekende afzender',
        email: ex.email,
        telefoon: ex.telefoon,
        land: ex.land,
        contactWay: ex.contactWay,
        onderwerp,
        datum: m.receivedDateTime.slice(0, 10),
        status: 'nieuw',
        notitie,
        contacts: [],
        auto: true,
      });
    }
    nieuweTeller++;
  }

  if (nieuweTeller) {
    await putJson(LEADS_URL, leadsArr);
    console.log(`${nieuweTeller} mail(s) verwerkt: ${uniekTeller} nieuwe lead(s), ${samengevoegdTeller} samengevoegd met bestaande lead.`);
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
