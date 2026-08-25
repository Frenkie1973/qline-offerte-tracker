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

// Zelfde principe, maar dan voor mail die Frank zelf heeft verstuurd, zodat
// die ook automatisch als contactmoment gelogd wordt — hij hoeft dat dan
// niet meer met de hand te doen via "Gebeld/Geappt/Gemaild".
async function fetchSentMessages(token, sinceISO) {
  const filter = encodeURIComponent(`sentDateTime gt ${sinceISO}`);
  const select = 'id,subject,toRecipients,sentDateTime,bodyPreview,body';
  const alle = [];
  for (const mailbox of SENT_MAILBOXES) {
    const url = `https://graph.microsoft.com/v1.0/users/${mailbox}/mailFolders/sentitems/messages?$filter=${filter}&$select=${select}&$orderby=sentDateTime asc&$top=50`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      console.error(`Verzonden mail ophalen mislukt voor ${mailbox}: ${res.status} ${await res.text()}`);
      continue;
    }
    const data = await res.json();
    alle.push(...(data.value || []));
  }
  // Dedupe: dezelfde mail kan soms in meerdere mailboxen als "verzonden"
  // staan (bijv. gedeelde mailbox + eigen postvak); op id ontdubbelen.
  const gezien = new Set();
  return alle.filter((m) => {
    if (gezien.has(m.id)) return false;
    gezien.add(m.id);
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

async function main() {
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    console.error('Ontbrekende secrets (MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET).');
    process.exit(1);
  }
  if (!FB_SECRET) {
    console.error('Ontbrekend secret FIREBASE_DB_SECRET — sinds de Firebase-regels "auth != null" vereisen kan de bot niet meer anoniem lezen/schrijven.');
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
    // Q-Line's eigen mailboxen mogen NOOIT als "klant-e-mail" gebruikt worden om
    // leads te matchen. De webshop toont soms store@q-line.com zelf als
    // afzender van een ordermelding — zonder deze uitsluiting werden daardoor
    // compleet verschillende, losse klantaanvragen allemaal per ongeluk
    // samengevoegd tot dezelfde (soms al foute) bestaande lead.
    const EIGEN_ADRESSEN = ['store@q-line.com', 'f.timmerhuis@q-line.com'];
    let emailKey = (ex.email || '').toLowerCase().trim();
    if (EIGEN_ADRESSEN.includes(emailKey)) emailKey = '';
    const klantKey = (ex.klant || fromEmail || '').toLowerCase().trim();
    const GESLOTEN_STATUSSEN = ['gewonnen', 'verloren', 'deadend'];

    // Zoek een nog openstaande lead (elke actieve status: nieuw, contact,
    // offerte, dealer — alleen gewonnen/verloren/deadend telt niet mee) van
    // dezelfde klant, op e-mail óf op naam, zodat een klant waar Frank al
    // mee bezig is niet als nieuwe losse lead wordt aangemaakt.
    // Uitzondering: als beide een aanvraagnr hebben en die verschillen, gaat
    // het gegarandeerd om een andere bestelling — dan nooit samenvoegen, ook
    // niet als naam/e-mail toevallig overeenkomen.
    const bestaande = leadsArr.find((d) => {
      if (GESLOTEN_STATUSSEN.includes(d.status)) return false;
      if (ex.aanvraagnr && d.aanvraagnr && d.aanvraagnr !== ex.aanvraagnr) return false;
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
      if (!bestaande.aanvraagnr && ex.aanvraagnr) bestaande.aanvraagnr = ex.aanvraagnr;

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
        aanvraagnr: ex.aanvraagnr,
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

  // === Verzonden mail (Frank -> klant) automatisch loggen ===
  console.log('Ophalen verzonden mail...');
  const sentMessages = await fetchSentMessages(token, sinceISO);
  console.log(`${sentMessages.length} verzonden mail(s) gevonden.`);
  let verzondenGelogd = 0;

  for (const sm of sentMessages) {
    if (processedIds.has(sm.id)) continue;
    processedIds.add(sm.id);
    if (sm.sentDateTime > latest) latest = sm.sentDateTime;

    const ontvangers = (sm.toRecipients || []).map((r) => (r.emailAddress?.address || '').toLowerCase().trim());
    if (!ontvangers.length) continue;

    // Zoek een lead (ongeacht status, ook al gesloten deals mogen nazorg-mail
    // gelogd krijgen) waarvan het e-mailadres overeenkomt met een ontvanger.
    const doelLead = leadsArr.find((d) => ontvangers.includes((d.email || '').toLowerCase().trim()));
    if (!doelLead) continue;

    const sentBodyText = sm.body?.contentType === 'html' ? stripHtml(sm.body.content) : (sm.body?.content || sm.bodyPreview || '');
    const verzonden = new Date(sm.sentDateTime);
    doelLead.contacts = doelLead.contacts || [];
    doelLead.contacts.push({
      type: 'mail',
      datum: verzonden.toLocaleDateString('nl-NL', { day: 'numeric', month: 'short' }),
      tijd: verzonden.toLocaleTimeString('nl-NL', { hour: '2-digit', minute: '2-digit' }),
      tekst: sentBodyText.slice(0, 800) || sm.subject || 'Mail verstuurd',
    });
    // Frank heeft zelf gereageerd -> een eventueel klaarstaand AI-voorstel is
    // niet meer actueel.
    if (doelLead.aiVoorstel) {
      delete doelLead.aiVoorstel;
      delete doelLead.aiVoorstelDatum;
    }
    verzondenGelogd++;
    nieuweTeller++;
  }

  if (verzondenGelogd) {
    console.log(`${verzondenGelogd} verzonden mail(s) automatisch gelogd bij bestaande leads.`);
  }

  if (nieuweTeller) {
    await putJson(LEADS_URL, leadsArr);
    console.log(`${nieuweTeller} mail(s) verwerkt: ${uniekTeller} nieuwe lead(s), ${samengevoegdTeller} samengevoegd met bestaande lead, ${verzondenGelogd} verzonden mail(s) gelogd.`);
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
