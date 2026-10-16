// Q-Line mailbox bot
// Leest nieuwe mails uit store@q-line.com via Microsoft Graph (app-only) en
// zet ze als concept-leads (status "nieuw", tag auto:true) in de Firebase
// Realtime Database, in hetzelfde formaat als het dashboard zelf gebruikt.
//
// Draait via GitHub Actions (.github/workflows/mailbot.yml), elke 15 min.
//
// Werking: ALLES wat in het Postvak IN van store@ staat (binnengekomen of
// erin gesleept, laatste 14 dagen) wordt gecontroleerd. Staat de mail nog
// niet in de leads, dan wordt hij een nieuwe lead, of een reactie bij een
// openstaande lead van dezelfde klant. Verdwenen leads komen zo vanzelf terug.
// Benodigde secrets: MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET

const MAILBOX = 'store@q-line.com';
// Verwerkte aanvragen worden door Frank verplaatst naar de map "Q-Line store
// aanvragen" (archief). De bot moet dus ALLEEN de Postvak IN uitlezen, nooit
// die archiefmap, anders komen al afgehandelde leads opnieuw binnen.
//
// Frank beantwoordt klanten vanuit zijn EIGEN mailbox, niet vanuit
// store@q-line.com — daarom apart uitgelezen voor "verzonden mail".
// Frank beantwoordt klanten soms vanuit zijn eigen mailbox, soms vanuit de
// gedeelde store-mailbox — daarom worden de "Verzonden items" van BEIDE
// mailboxen gecheckt.
const SENT_MAILBOXES = ['store@q-line.com', 'f.timmerhuis@q-line.com'];
const FB_BASE = 'https://q-line-tracker-default-rtdb.europe-west1.firebasedatabase.app';
// Sinds de Firebase-regels zijn aangescherpt naar "auth != null" (30-7-2026)
// moet ELK verzoek geauthenticeerd zijn — ook vanuit deze bot, die geen
// gebruiker heeft om mee in te loggen. Daarom wordt het legacy database-secret
// als ?auth=... query-param meegestuurd; dat geeft admin-toegang die de
// beveiligingsregels omzeilt, precies zoals het dashboard dat niet nodig
// heeft omdat een ingelogde gebruiker daar wél aan "auth != null" voldoet.
const FB_SECRET = process.env.FIREBASE_DB_SECRET;
const FB_AUTH_SUFFIX = FB_SECRET ? `?auth=${FB_SECRET}` : '';
const LEADS_URL = `${FB_BASE}/leads.json${FB_AUTH_SUFFIX}`;
const STATE_URL = `${FB_BASE}/mailbotState.json${FB_AUTH_SUFFIX}`;

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

// Zoekt naar "Label: waarde", "Label - waarde" OF gewoon "Label waarde" (zonder
// scheidingsteken) op een eigen regel. De Q-Line webshop-formuliermails blijken
// namelijk GEEN dubbele punt te gebruiken (bv. "Referentie 24863", "Naam Jan
// Jansen") — vandaar dat het scheidingsteken optioneel is, met minimaal een
// spatie ertussen zodat een label niet per ongeluk vastplakt aan een woord.
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

// Probeert een naam te vinden in de afsluiting van een losse mail
// ("Met vriendelijke groet, Jan Jansen") als er geen expliciet naam-label is.
// Let op: matcht bewust alleen binnen ÉÉN regel na de groet, anders wordt een
// bedrijfsnaam op de regel eronder per ongeluk aan de naam vastgeplakt.
function extractSignatureName(text) {
  const m = text.match(/(?:met\s+vriendelijke\s+groet(?:en)?|groet(?:en)?|mvg|regards|kind\s+regards)[,.]?[ \t]*\n+[ \t]*([A-Z][A-Za-zÀ-ÿ'\-]+(?:[ \t]+[A-Z][A-Za-zÀ-ÿ'\-]+){0,3})/i);
  return m ? m[1].trim() : '';
}

// Zoekt het webshop aanvraag-/bestelnummer in de mailtekst. Webshopmails
// gebruiken hiervoor uiteenlopende labels (aanvraagnr, bestelnummer,
// ordernummer, referentie, request id...) — we proberen ze allemaal.
const AANVRAAGNR_LABELS = [
  'aanvraag\\s*nr\\.?',
  'aanvraag\\s*nummer',
  'aanvraagnr\\.?',
  'aanvraagnummer',
  'bestel\\s*nr\\.?',
  'bestelnummer',
  'order\\s*nr\\.?',
  'ordernummer',
  'order\\s*id',
  'order\\s*number',
  'request\\s*id',
  'referentie\\s*nr\\.?',
  'referentie\\s*nummer',
  'referentienummer',
  'referentie',
  'reference\\s*nr\\.?',
  'reference\\s*number',
  'reference',
  'referenz\\s*nr\\.?',
  'referenz\\s*nummer',
  'referenznummer',
  'referenz',
];

function extractAanvraagnr(text) {
  // extractField dekt nu ook het scheidingsteken-loze webshopformaat
  // ("Referentie 24863"). Alleen accepteren als de waarde er ook echt als een
  // nummer/code uitziet — anders vangt een algemeen "Reference"-veld per
  // ongeluk een bedrijfs- of plaatsnaam.
  const val = extractField(text, AANVRAAGNR_LABELS);
  return (val && /\d/.test(val) && val.length <= 20) ? val : '';
}

function extractStructuredData(bodyText, fromName, fromEmail) {
  const cleanFromName = (fromName || '').replace(/^namens\s+/i, '').trim();
  const klant = extractField(bodyText, ['naam', 'name', 'klant', 'contactpersoon'])
    || extractSignatureName(bodyText)
    || cleanFromName || '';
  const email = extractField(bodyText, ['e-?mail(?:adres)?', 'email']) || fromEmail || '';
  // "phone\s*number" MOET vóór het losse "phone" staan, anders vangt "phone"
  // alleen "Phone" en komt "number 0612345678" als waarde mee.
  const telefoon = extractField(bodyText, ['telefoon(?:nummer)?', 'tel(?:efoonnr)?', 'phone\\s*number', 'phone', 'mobiel']) || extractPhone(bodyText);
  // Idem: de samengestelde labels moeten vóór de losse "straat"/"address" staan.
  const straat = extractField(bodyText, ['straat\\s*en\\s*huisnummer', 'street\\s*and\\s*house\\s*number', 'adres', 'address', 'straat']);
  const postcode = extractField(bodyText, ['postcode', 'zip(?:code)?', 'plz']);
  const plaats = extractField(bodyText, ['plaats', 'city', 'woonplaats', 'stadt']);
  const landVeld = extractField(bodyText, ['land', 'country']);
  const product = extractField(bodyText, ['product', 'interesse\\s*in', 'onderwerp']);
  const contactWay = extractField(bodyText, ['gewenste\\s*manier\\s*van\\s*contact', 'contact\\s*voorkeur', 'voorkeur\\s*contact', 'preferred\\s*contact\\s*(?:way|method)?', 'liever\\s*(?:gebeld|gemaild)']);
  const kooptermijn = extractField(bodyText, ['kooptermijn', 'wanneer.*(?:kopen|aanschaf)', 'when.*(?:buy|purchase)']);
  const typeKlant = extractField(bodyText, ['type\\s*(?:bedrijf|klant)', 'soort\\s*bedrijf', 'business\\s*type']);
  const aanvraagnr = extractAanvraagnr(bodyText);

  const adres = [straat, postcode, plaats, landVeld].filter(Boolean).join(', ');

  return { klant, email, telefoon, land: landVeld, contactWay, product, kooptermijn, typeKlant, adres, aanvraagnr };
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

// Haalt alle pagina's op van een Graph-lijst (volgt @odata.nextLink), zodat
// er op drukke dagen niets wegvalt achter de eerste 100 resultaten.
async function graphGetAll(url, token, maxPaginas = 20) {
  const alle = [];
  let next = url;
  let paginas = 0;
  while (next && paginas < maxPaginas) {
    const res = await fetch(next, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`Graph-verzoek mislukt: ${res.status} ${await res.text()}`);
    const data = await res.json();
    alle.push(...(data.value || []));
    next = data['@odata.nextLink'] || null;
    paginas++;
  }
  return alle;
}

// Postvak IN van store@. Er wordt gefilterd op lastModifiedDateTime in plaats
// van receivedDateTime: een mail die Frank vanuit zijn eigen mailbox naar
// store@ SLEEPT behoudt namelijk zijn oorspronkelijke ontvangstdatum (kan
// weken oud zijn), maar krijgt bij het verplaatsen wél een nieuwe
// lastModifiedDateTime. Zo wordt ook een oude, versleepte mail opgepakt.
async function fetchNewMessages(token, sinceISO) {
  const filter = encodeURIComponent(`lastModifiedDateTime ge ${sinceISO}`);
  const select = 'id,internetMessageId,subject,from,receivedDateTime,lastModifiedDateTime,bodyPreview,body';
  // Expliciet de Postvak IN-map, nooit de hele mailbox (die bevat ook de
  // archiefmap "Q-Line store aanvragen" met al afgehandelde mail).
  const url = `https://graph.microsoft.com/v1.0/users/${MAILBOX}/mailFolders/inbox/messages?$filter=${filter}&$select=${select}&$orderby=lastModifiedDateTime asc&$top=100`;
  return graphGetAll(url, token);
}

// Zelfde principe, maar dan voor mail die Frank zelf heeft verstuurd, zodat
// die ook automatisch als contactmoment gelogd wordt.
async function fetchSentMessages(token, sinceISO) {
  const filter = encodeURIComponent(`sentDateTime ge ${sinceISO}`);
  const select = 'id,internetMessageId,subject,toRecipients,sentDateTime,bodyPreview,body';
  const alle = [];
  for (const mailbox of SENT_MAILBOXES) {
    const url = `https://graph.microsoft.com/v1.0/users/${mailbox}/mailFolders/sentitems/messages?$filter=${filter}&$select=${select}&$orderby=sentDateTime asc&$top=100`;
    try {
      alle.push(...(await graphGetAll(url, token)));
    } catch (e) {
      console.error(`Verzonden mail ophalen mislukt voor ${mailbox}: ${e.message}`);
    }
  }
  // Dezelfde mail kan in beide mailboxen als "verzonden" staan -> ontdubbelen.
  const gezien = new Set();
  return alle.filter((m) => {
    const k = m.internetMessageId || m.id;
    if (gezien.has(k)) return false;
    gezien.add(k);
    return true;
  });
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


// ===================================================================
// ZELFHERSTELLENDE VERWERKING
// -------------------------------------------------------------------
// Vroeger hield de bot een lijst "processedIds" bij: een mail die één keer
// verwerkt was, werd nooit meer bekeken. Probleem: het dashboard slaat bij
// elke wijziging de VOLLEDIGE leadslijst op. Stond het dashboard al een tijd
// open (met een oude lijst) en klikte Frank daarna iets aan, dan werden
// leads die de bot in de tussentijd had toegevoegd overschreven en waren ze
// weg — terwijl de bot dacht dat ze al verwerkt waren. Zo verdwenen o.a.
// Kreuzer, Hofmeister en Kühnhauser.
//
// Nu kijkt de bot bij ELKE run of een mail ook echt in de leads terug te
// vinden is (via mailIds, of via de oude "Ontvangen: ..."-vingerafdruk in
// de notitie). Is hij weg, dan wordt hij gewoon opnieuw verwerkt.
// ===================================================================

const GESLOTEN_STATUSSEN = ['gewonnen', 'verloren', 'deadend'];
// Q-Line's eigen mailboxen mogen NOOIT als "klant-e-mail" gebruikt worden om
// leads te matchen (de webshop toont soms store@ zelf als afzender).
const EIGEN_ADRESSEN = ['store@q-line.com', 'f.timmerhuis@q-line.com', 'info@q-line.com', 'support@q-line.com'];

function mailKey(m) {
  return m.internetMessageId || m.id;
}

// Mails waarvan Frank de lead in het dashboard bewust heeft VERWIJDERD
// (spam, niet relevant). Het dashboard zet die in /mailbotIgnore, zodat de
// bot ze niet steeds opnieuw aanmaakt.
const IGNORE_URL = `${FB_BASE}/mailbotIgnore.json${FB_AUTH_SUFFIX}`;
function ignoreSleutel(s) {
  return encodeURIComponent(String(s)).replace(/\./g, '%2E');
}
function isGenegeerd(ignore, m) {
  const vingerafdruk = `Ontvangen: ${new Date(m.receivedDateTime).toLocaleString('nl-NL')}`;
  return [mailKey(m), m.id, vingerafdruk].some((k) => k && ignore[ignoreSleutel(k)]);
}

function nlDatum(d) {
  return d.toLocaleDateString('nl-NL', { day: 'numeric', month: 'short' });
}
function nlTijd(d) {
  return d.toLocaleTimeString('nl-NL', { hour: '2-digit', minute: '2-digit' });
}

function contactSig(c) {
  return `${c.type || ''}|${c.datum || ''}|${c.tijd || ''}|${String(c.tekst || '').slice(0, 60)}`;
}

// Is deze inkomende mail al ergens in de leads terug te vinden?
function inboxMailAlVerwerkt(leadsArr, m) {
  const key = mailKey(m);
  const vingerafdruk = `Ontvangen: ${new Date(m.receivedDateTime).toLocaleString('nl-NL')}`;
  return leadsArr.some((d) => {
    if (!d) return false;
    if (Array.isArray(d.mailIds) && (d.mailIds.includes(key) || d.mailIds.includes(m.id))) return true;
    // Oude leads (van vóór mailIds) herkennen aan de tekst die de bot zelf
    // in de notitie zette.
    if (typeof d.notitie === 'string' && d.notitie.includes(vingerafdruk)) return true;
    return false;
  });
}

// Zoekt dezelfde lead terug in een (verse) leadslijst.
function vindLead(arr, lead) {
  if (lead.id) return arr.find((x) => x && x.id === lead.id);
  return arr.find((x) => x && !x.id && x.klant === lead.klant && x.email === lead.email && x.datum === lead.datum);
}

// Houdt per lead bij wat de bot veranderd heeft, zodat die wijzigingen vlak
// voor het opslaan op een VERSE kopie van de leads gezet kunnen worden — dan
// overschrijft de bot op zijn beurt ook geen wijzigingen die Frank in de
// tussentijd in het dashboard deed.
const patches = new Map(); // lead-object -> { set:{}, del:Set, contacts:[], mailIds:[] }
function patchVoor(lead) {
  if (!patches.has(lead)) patches.set(lead, { set: {}, del: new Set(), contacts: [], mailIds: [] });
  return patches.get(lead);
}
function zetVeld(lead, veld, waarde) {
  lead[veld] = waarde;
  const p = patchVoor(lead);
  p.set[veld] = waarde;
  p.del.delete(veld);
}
function wisVeld(lead, veld) {
  delete lead[veld];
  const p = patchVoor(lead);
  delete p.set[veld];
  p.del.add(veld);
}
function voegContactToe(lead, contact) {
  lead.contacts = lead.contacts || [];
  lead.contacts.push(contact);
  patchVoor(lead).contacts.push(contact);
}
function voegMailIdToe(lead, key) {
  lead.mailIds = lead.mailIds || [];
  if (!lead.mailIds.includes(key)) lead.mailIds.push(key);
  patchVoor(lead).mailIds.push(key);
}

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    console.error('Ontbrekende secrets (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET).');
    process.exit(1);
  }
  if (!FB_SECRET) {
    console.error('Ontbrekend secret FIREBASE_DB_SECRET — sinds de Firebase-regels "auth != null" vereisen kan de bot niet meer anoniem lezen/schrijven.');
    process.exit(1);
  }

  // Postvak IN: alles wat de afgelopen 14 dagen binnenkwam OF erin gesleept
  // is. Verzonden items: afgelopen 7 dagen.
  const inboxSinds = new Date(Date.now() - 14 * 24 * 3600 * 1000).toISOString();
  const sentSinds = new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString();

  const token = await getAppToken();
  console.log(`Ophalen Postvak IN (gewijzigd/binnengekomen sinds ${inboxSinds})...`);
  const messages = await fetchNewMessages(token, inboxSinds);
  console.log(`${messages.length} mail(s) in het venster.`);

  const leadsRuw = await getJson(LEADS_URL, null);
  if (leadsRuw !== null && !Array.isArray(leadsRuw)) {
    throw new Error('Leads in Firebase zijn geen lijst — afgebroken om niets te overschrijven.');
  }
  const leadsArr = Array.isArray(leadsRuw) ? leadsRuw : [];
  const nieuweLeads = [];
  const ignore = (await getJson(IGNORE_URL, {})) || {};

  let samengevoegdTeller = 0;
  let uniekTeller = 0;
  let overgeslagen = 0;

  for (const m of messages) {
    if (inboxMailAlVerwerkt(leadsArr, m) || isGenegeerd(ignore, m)) {
      overgeslagen++;
      continue;
    }
    const key = mailKey(m);

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
    let emailKey = (ex.email || '').toLowerCase().trim();
    if (EIGEN_ADRESSEN.includes(emailKey)) emailKey = '';
    // Als het e-mailveld uit de tekst een eigen adres was, val terug op de
    // echte afzender (bijv. bij een doorgestuurde of versleepte mail).
    if (!emailKey && fromEmail && !EIGEN_ADRESSEN.includes(fromEmail.toLowerCase())) {
      emailKey = fromEmail.toLowerCase().trim();
    }
    const klantKey = (ex.klant || '').toLowerCase().trim();

    // Openstaande lead van dezelfde klant (op e-mail of naam)? Dan wordt de
    // mail daar als reactie aan gehangen. Verschillend aanvraagnummer = andere
    // aanvraag = altijd een nieuwe lead.
    const bestaande = leadsArr.find((d) => {
      if (!d) return false;
      if (GESLOTEN_STATUSSEN.includes(d.status)) return false;
      if (ex.aanvraagnr && d.aanvraagnr && d.aanvraagnr !== ex.aanvraagnr) return false;
      const dEmail = (d.email || '').toLowerCase().trim();
      const dKlant = (d.klant || '').toLowerCase().trim();
      if (emailKey && dEmail && dEmail === emailKey) return true;
      if (klantKey && dKlant && dKlant === klantKey) return true;
      return false;
    });

    if (bestaande) {
      const huidigOnderwerp = bestaande.onderwerp || '';
      if (onderwerp && !huidigOnderwerp.includes(onderwerp)) {
        zetVeld(bestaande, 'onderwerp', huidigOnderwerp ? `${huidigOnderwerp} + ${onderwerp}` : onderwerp);
      }
      zetVeld(bestaande, 'notitie', `${bestaande.notitie || ''}\n\n--- Extra aanvraag / reactie van dezelfde klant ---\n${notitie}`);
      const ontvangen = new Date(m.receivedDateTime);
      voegContactToe(bestaande, {
        type: 'reactie',
        datum: nlDatum(ontvangen),
        tijd: nlTijd(ontvangen),
        tekst: `🤖 Nieuwe mail binnengekomen: ${onderwerp}`,
      });
      voegMailIdToe(bestaande, key);
      if (m.receivedDateTime.slice(0, 10) > (bestaande.datum || '')) {
        zetVeld(bestaande, 'datum', m.receivedDateTime.slice(0, 10));
      }
      if (!bestaande.telefoon && ex.telefoon) zetVeld(bestaande, 'telefoon', ex.telefoon);
      if (!bestaande.land && ex.land) zetVeld(bestaande, 'land', ex.land);
      if (!bestaande.contactWay && ex.contactWay) zetVeld(bestaande, 'contactWay', ex.contactWay);
      if (!bestaande.aanvraagnr && ex.aanvraagnr) zetVeld(bestaande, 'aanvraagnr', ex.aanvraagnr);

      console.log(`Reactie gekoppeld aan bestaande lead: ${bestaande.klant}. AI-opvolgvoorstel maken...`);
      const voorstel = await genereerAiVoorstel(bestaande, bodyText);
      if (voorstel) {
        zetVeld(bestaande, 'aiVoorstel', voorstel);
        zetVeld(bestaande, 'aiVoorstelDatum', m.receivedDateTime);
      }
      samengevoegdTeller++;
    } else {
      uniekTeller++;
      const lead = {
        id: `mail_${Date.now().toString(36)}_${uniekTeller}_${Math.random().toString(36).slice(2, 8)}`,
        klant: ex.klant || fromEmail || 'Onbekende afzender',
        email: emailKey || ex.email || '',
        telefoon: ex.telefoon,
        land: ex.land,
        contactWay: ex.contactWay,
        aanvraagnr: ex.aanvraagnr,
        onderwerp,
        datum: m.receivedDateTime.slice(0, 10),
        status: 'nieuw',
        notitie,
        contacts: [],
        mailIds: [key],
        auto: true,
      };
      leadsArr.push(lead);
      nieuweLeads.push(lead);
      console.log(`Nieuwe lead: ${lead.klant} — ${onderwerp}`);
    }
  }

  // === Verzonden mail (Frank -> klant) automatisch loggen ===
  console.log('Ophalen verzonden mail...');
  const sentMessages = await fetchSentMessages(token, sentSinds);
  console.log(`${sentMessages.length} verzonden mail(s) gevonden.`);
  let verzondenGelogd = 0;

  for (const sm of sentMessages) {
    const ontvangers = (sm.toRecipients || []).map((r) => (r.emailAddress?.address || '').toLowerCase().trim());
    if (!ontvangers.length) continue;

    const doelLead = leadsArr.find((d) => d && d.email && ontvangers.includes(d.email.toLowerCase().trim()));
    if (!doelLead) continue;

    const key = mailKey(sm);
    const verzonden = new Date(sm.sentDateTime);
    const datum = nlDatum(verzonden);
    const tijd = nlTijd(verzonden);
    // Al gelogd? (via mailIds, of een mail-contact op exact dezelfde datum/tijd)
    if (Array.isArray(doelLead.mailIds) && (doelLead.mailIds.includes(key) || doelLead.mailIds.includes(sm.id))) continue;
    if ((doelLead.contacts || []).some((c) => c.type === 'mail' && c.datum === datum && c.tijd === tijd)) continue;

    const sentBodyText = sm.body?.contentType === 'html' ? stripHtml(sm.body.content) : (sm.body?.content || sm.bodyPreview || '');
    voegContactToe(doelLead, {
      type: 'mail',
      datum,
      tijd,
      tekst: sentBodyText.slice(0, 800) || sm.subject || 'Mail verstuurd',
    });
    voegMailIdToe(doelLead, key);
    // Frank heeft zelf gereageerd -> klaarstaand AI-voorstel is niet meer actueel,
    // tenzij dat voorstel van NA deze verzonden mail is.
    if (doelLead.aiVoorstel && (!doelLead.aiVoorstelDatum || doelLead.aiVoorstelDatum < sm.sentDateTime)) {
      wisVeld(doelLead, 'aiVoorstel');
      wisVeld(doelLead, 'aiVoorstelDatum');
    }
    verzondenGelogd++;
  }

  console.log(`${overgeslagen} mail(s) waren al verwerkt (staan in de leads of zijn bewust verwijderd).`);

  if (!nieuweLeads.length && !patches.size) {
    console.log('Niets nieuws om op te slaan.');
  } else {
    // Vlak voor het opslaan de leads OPNIEUW ophalen en alleen onze eigen
    // wijzigingen daarop zetten — zo gaat er niets verloren van wat Frank
    // ondertussen in het dashboard heeft gedaan.
    const versRuw = await getJson(LEADS_URL, null);
    if (versRuw !== null && !Array.isArray(versRuw)) {
      throw new Error('Leads in Firebase zijn geen lijst — afgebroken om niets te overschrijven.');
    }
    if (versRuw === null && leadsArr.length > nieuweLeads.length) {
      throw new Error('Kon de actuele leads niet opnieuw ophalen — afgebroken om niets te overschrijven.');
    }
    const vers = Array.isArray(versRuw) ? versRuw : [];

    for (const [lead, p] of patches) {
      if (nieuweLeads.includes(lead)) continue; // nieuwe leads gaan hieronder in z'n geheel mee
      const doel = vindLead(vers, lead);
      if (!doel) continue; // lead is in de tussentijd in het dashboard verwijderd
      Object.assign(doel, p.set);
      for (const veld of p.del) delete doel[veld];
      if (p.contacts.length) {
        doel.contacts = doel.contacts || [];
        const sigs = new Set(doel.contacts.map(contactSig));
        for (const c of p.contacts) {
          if (!sigs.has(contactSig(c))) doel.contacts.push(c);
        }
      }
      if (p.mailIds.length) {
        doel.mailIds = Array.from(new Set([...(doel.mailIds || []), ...p.mailIds]));
      }
    }
    for (const lead of nieuweLeads) {
      if (!vers.some((x) => x && x.id === lead.id)) vers.push(lead);
    }

    await putJson(LEADS_URL, vers);
    console.log(`Opgeslagen: ${uniekTeller} nieuwe lead(s), ${samengevoegdTeller} reactie(s) gekoppeld aan bestaande lead, ${verzondenGelogd} verzonden mail(s) gelogd.`);
  }

  await putJson(STATE_URL, { lastCheck: new Date().toISOString() });
  console.log('Klaar.');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
