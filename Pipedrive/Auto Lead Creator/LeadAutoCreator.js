/***********************
 * GSADUS → Pipedrive Lead Auto-Creator (cache-first + idempotency)
 * Runs as: Sales@gsadus.com
 ***********************/

// ===== CONFIG =====
const SALES_INBOX = 'Sales@gsadus.com';
const LABEL_ACCEPT = '1 - SALES/NEW LEADS (AUTO)';
const LABEL_FILTERED = 'SPAM (AUTO)';             // spam/test submissions: skipped, not deleted (audit trail). Top-level, not under "1 - SALES".
const LABEL_FORCE = '1 - SALES/FORCE (MANUAL)';  // manual override: apply to a filtered msg to force lead creation
const LEAD_ADDRESS_FIELD_KEY = 'e76ad51def930fd350324b8057577be5bde93023'; // Lead custom field

// Pipedrive API token is read from Script Properties (Project Settings → Script Properties → PD_API_TOKEN).
// Do NOT hardcode the token in source; this file is tracked in git.

// Processing controls
const MAX_PER_RUN = 5;              // smooth bursts
const SEARCH_WINDOW = '3d';         // Gmail search window

// Idempotency
const DEDUPE_KEY = 'processed_msg_ids';  // JSON: { "<msgId>": <epoch_ms>, ... }
const DEDUPE_TTL_DAYS = 60;

// Pipedrive rate limit backoff
const PD_BACKOFF_PROP = 'pd_backoff_until_ms';
const DEFAULT_PD_BACKOFF_MS = 60 * 60 * 1000; // 1 hour fallback if reset header not provided

// Toggle: skip expensive person search and always create person (accept duplicates)
const SKIP_PERSON_SEARCH =
  (PropertiesService.getScriptProperties().getProperty('PD_SKIP_PERSON_SEARCH') || 'false')
    .toLowerCase() === 'true';

// ===== PERSONS CACHE (persistent mirror) =====
// Stored in Script Properties as a JSON map: { "email_lower": { id, name, updated } }
const PERSONS_CACHE_KEY = 'pd_persons_cache_v1'; // Script Properties JSON
const PERSONS_CACHE_MAX = 50000;                 // safety cap

function loadPersonsCache() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PERSONS_CACHE_KEY) || '{}'); }
  catch (_) { return {}; }
}
function savePersonsCache(map) {
  let m = map || {};
  const keys = Object.keys(m);
  if (keys.length > PERSONS_CACHE_MAX) {
    // keep newest ~90%
    const pruned = {};
    keys.sort((a,b) => (m[b].updated||0) - (m[a].updated||0))
        .slice(0, Math.floor(PERSONS_CACHE_MAX * 0.9))
        .forEach(k => pruned[k] = m[k]);
    m = pruned;
  }
  PropertiesService.getScriptProperties().setProperty(PERSONS_CACHE_KEY, JSON.stringify(m));
}
function cacheSetPerson(email, id, name) {
  if (!email || !id) return;
  const e = email.toLowerCase();
  const store = loadPersonsCache();
  store[e] = { id: Number(id), name: name || '', updated: Date.now() };
  savePersonsCache(store);
  // also 6h Script Cache for fast reads
  try { CacheService.getScriptCache().put(`pd_person_${e}`, String(id), 6 * 60 * 60); } catch(_){}
}
function cacheGetPersonId(email) {
  if (!email) return null;
  const e = email.toLowerCase();
  const run = CacheService.getScriptCache().get(`pd_person_${e}`);
  if (run) return Number(run);
  const store = loadPersonsCache();
  return store[e]?.id || null;
}

// === Phone cache (persistent)
const PERSONS_PHONE_CACHE_KEY = 'pd_persons_phone_cache_v1';

function loadPhoneCache() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PERSONS_PHONE_CACHE_KEY) || '{}'); }
  catch (_) { return {}; }
}
function savePhoneCache(map) {
  PropertiesService.getScriptProperties().setProperty(PERSONS_PHONE_CACHE_KEY, JSON.stringify(map || {}));
}
function cacheSetPhone(phone, id) {
  const p = normalizePhone(phone || '');
  if (!p || !id) return;
  const store = loadPhoneCache();
  store[p] = Number(id);
  savePhoneCache(store);
  try { CacheService.getScriptCache().put(`pd_person_phone_${p}`, String(id), 6*60*60); } catch(_) {}
}
function cacheGetPersonIdByPhone(phone) {
  const p = normalizePhone(phone || '');
  if (!p) return null;
  const run = CacheService.getScriptCache().get(`pd_person_phone_${p}`);
  if (run) return Number(run);
  const store = loadPhoneCache();
  return store[p] || null;
}

// ===== MAIN: Gmail-driven run (schedule ~5–30 min) =====
function processInbox() {
  // Respect backoff window if we recently hit Pipedrive rate limit
  const untilStr = PropertiesService.getScriptProperties().getProperty(PD_BACKOFF_PROP) || '0';
  const until = Number(untilStr);
  if (until && Date.now() < until) {
    console.log(`Pipedrive backoff active; resumes at ${new Date(until).toISOString()}`);
    return;
  }

  // One-time rollout guard: mark test/spam that was already sitting in the inbox when the
  // spam filter was deployed (labeled SPAM during rollout) as processed, so recalibration
  // does not backfill it into Pipedrive. No-op after the first successful run.
  maybeBackfillSkip_();

  // Gmail-only phase. Exclude both the accept label (already-created leads) and the
  // spam label (already-classified junk) so neither is reconsidered.
  const q = `deliveredto:${SALES_INBOX} newer_than:${SEARCH_WINDOW} -label:"${LABEL_ACCEPT}" -label:"${LABEL_FILTERED}"`;
  const threads = GmailApp.search(q, 0, 50);
  console.log(`processInbox: search q="${q}" threads=${threads.length}`);
  if (!threads.length) return;

  const acceptLabel = getOrCreateNestedLabel(LABEL_ACCEPT);
  const dedupe = loadDedupe();

  // Collect new, valid candidates using only Gmail
  const candidatesAll = collectCandidates(threads, dedupe);
  console.log(`processInbox: collected candidatesAll=${candidatesAll.length}`);
  // collectCandidates may have marked spam/test messages as processed; persist that.
  if (!candidatesAll.length) { saveDedupe(dedupe); return; }

  // Cap per run to smooth burst/budget
  const candidates = candidatesAll.slice(0, MAX_PER_RUN);
  console.log(`processInbox: processing ${candidates.length} candidates (max ${MAX_PER_RUN})`);

  // Pipedrive operations only if we have work
  const pd = new PipedriveClient();
  const runPersonCache = {}; // in-run cache email -> personId

  for (const item of candidates) {
    const { thread, msg, parsed } = item;
    try {
      const personId = ensurePersonSmart(pd, parsed, runPersonCache);
      // If the parsed name is a known placeholder like "[username]", prefer email for title.
      const isPlaceholderName = (n) => {
        if (!n) return false; const t = n.trim().toLowerCase(); return t === '[username]';
      };
      const effectiveName = (parsed.name && !isPlaceholderName(parsed.name)) ? parsed.name : '';
      const leadTitle = effectiveName || parsed.email || parsed.phone || 'GSADUS Web Lead';

      const payload = { title: leadTitle, person_id: personId };
      if (parsed.address) payload[LEAD_ADDRESS_FIELD_KEY] = parsed.address;

      const lead = pd.createLead(payload);
      pd.addNoteToLead(lead.id, buildNote(parsed, msg));

      console.log(`processInbox: lead created leadId=${lead.id} personId=${personId} msgId=${msg.getId()} title="${leadTitle}" email="${parsed.email||''}" phone="${parsed.phone||''}"`);

      msg.markRead();
      thread.addLabel(acceptLabel);
      // Clear override/spam labels so a FORCE rescue fires only once and the thread's
      // final state is just the accept label (no-op for a normal first-time lead).
      clearOverrideLabels_(thread);

      // record idempotency
      dedupe[msg.getId()] = Date.now();

      // write-through cache on success (redundant if ensurePersonSmart already set it)
      if (parsed.email) cacheSetPerson(parsed.email, personId, parsed.name);
    } catch (e) {
      if (isRateLimitError(e)) {
        console.log(`Pipedrive rate limit reached; pausing run. ${e && e.message}`);
        saveDedupe(dedupe);
        return; // exit early, try later
      }
      if (isAuthError(e)) {
        console.log(`Pipedrive auth error — check PD_API_TOKEN and PD_API_BASE. ${e && e.message}`);
        return; // abort run; configuration issue
      }
      console.log(`Error processing message ${msg.getId()}: ${e && e.message} stack=${e && e.stack}`);
      // continue with next candidate
    }
  }

  saveDedupe(dedupe);
  console.log(`processInbox: run complete processed=${candidates.length}`);
}

// Collect at most one candidate per thread, Gmail-only
function collectCandidates(threads, dedupe){
  const out = [];
  let filteredLabel = null; // lazily created only if something needs filtering
  for (const thread of threads) {
    // Newest to oldest. Evaluate EVERY message: the redesigned site threads all form
    // submissions into one Gmail conversation (identical subject), so a thread holds many
    // distinct leads, not one. Per-message dedupe keeps it idempotent.
    const messages = thread.getMessages().slice().reverse();
    for (const msg of messages) {
      try {
        if (msg.isInTrash()) continue;
        const msgId = msg.getId();

        // Manual override: a human applied the FORCE label to rescue a false positive.
        // Evaluated BEFORE the dedupe skip so a previously-filtered message can be re-admitted.
        const forced = threadHasLabel(thread, LABEL_FORCE);

        // Idempotency guard. NOTE: the -label: clauses in the Gmail query do NOT match these
        // parenthesized label names, so THIS dedupe map — not the labels — is the authoritative
        // "already handled" record for both created leads and filtered junk.
        if (dedupe[msgId] && !forced) continue;

        // Use profile-based parser from EmailProfiles.js
        const profParsed = parseLeadFromMessage(msg);
        if (!profParsed) continue; // not a recognized lead email

        // Map to existing shape expected by downstream logic
        const parsed = {
          name: (profParsed.fullName || '').trim(),
          email: (profParsed.email || '').trim(),
          phone: (profParsed.mobilePhone || '').trim(),
          address: (profParsed.address || '').trim(),
          message: (profParsed.note || '').trim()
        };

        // Minimal validity: need at least email or phone
        if (!parsed.email && !parsed.phone) continue;

        // Spam / test backstop (bypassed when forced). Filtered mail is labeled (audit trail)
        // and marked processed so it is not re-evaluated every run; it is never deleted, and
        // applying the FORCE label re-admits it via the check above.
        if (!forced) {
          const verdict = classifyLead(parsed);
          if (verdict.isSpam || verdict.isTest) {
            const kind = verdict.isSpam ? 'SPAM' : 'TEST';
            console.log(`collectCandidates: filtered ${kind} msgId=${msgId} email="${parsed.email}" phone="${parsed.phone}" score=${verdict.score} reasons=[${verdict.reasons.join(',')}]`);
            dedupe[msgId] = Date.now(); // authoritative skip on future runs
            try {
              if (!filteredLabel) filteredLabel = getOrCreateNestedLabel(LABEL_FILTERED);
              thread.addLabel(filteredLabel);
              msg.markRead();
            } catch (le) {
              console.log(`collectCandidates: could not label filtered msgId=${msgId}: ${le && le.message}`);
            }
            continue; // check older messages in this thread, if any
          }
        } else {
          console.log(`collectCandidates: FORCE override msgId=${msgId} — spam filter bypassed`);
        }

        out.push({ thread, msg, parsed });
        // No break: with all submissions in one thread, keep collecting every distinct lead
        // (MAX_PER_RUN caps how many are created per run).
      } catch (e) {
        console.log(`Error while collecting candidate for message ${msg.getId()}: ${e && e.message}`);
      }
    }
  }
  return out;
}

// ===== Idempotency helpers =====
function loadDedupe() {
  const raw = PropertiesService.getScriptProperties().getProperty(DEDUPE_KEY);
  let store = {};
  try { store = raw ? JSON.parse(raw) : {}; } catch (e) { store = {}; }
  // prune old
  const cutoff = Date.now() - DEDUPE_TTL_DAYS * 24 * 3600 * 1000;
  for (const k in store) if (store[k] < cutoff) delete store[k];
  return store;
}
function saveDedupe(store) {
  PropertiesService.getScriptProperties().setProperty(DEDUPE_KEY, JSON.stringify(store));
}

function isRateLimitError(e){
  const msg = (e && e.message) ? e.message : String(e || '');
  const s = msg.toLowerCase();
  return s.includes('http 429') || s.includes('rate limit') || s.includes('budget exceeded') || (e && e.code === 429);
}

function isAuthError(e){
  const msg = (e && e.message) ? e.message : String(e || '');
  const s = msg.toLowerCase();
  return s.includes('http 401') || s.includes('unauthorized access') || (e && e.code === 401);
}

// ===== INTEGRATION NOTE =====
// Candidate detection now relies solely on EmailProfiles.js via parseLeadFromMessage(msg).
// Subject/sender heuristics have been removed to avoid brittleness.

// ===== LABELS =====
function getOrCreateNestedLabel(path) {
  const parts = path.split('/');
  let cur = '';
  let label = null;
  for (const p of parts) {
    cur = cur ? `${cur}/${p}` : p;
    label = GmailApp.getUserLabelByName(cur) || GmailApp.createLabel(cur);
  }
  return label;
}

// True if the thread currently carries a user label with the exact given name.
function threadHasLabel(thread, name) {
  try {
    const labels = thread.getLabels() || [];
    for (let i = 0; i < labels.length; i++) {
      if (labels[i].getName() === name) return true;
    }
  } catch (_) {}
  return false;
}

// Remove the FORCE and SPAM labels from a thread (used after a successful lead creation
// so a manual rescue is one-shot and the thread ends in a clean accept-only state).
function clearOverrideLabels_(thread) {
  try { const fl = GmailApp.getUserLabelByName(LABEL_FORCE);    if (fl) thread.removeLabel(fl); } catch (_) {}
  try { const sl = GmailApp.getUserLabelByName(LABEL_FILTERED); if (sl) thread.removeLabel(sl); } catch (_) {}
}

// One-time rollout guard. When the spam filter is first deployed the inbox already holds
// test/spam submissions (labeled SPAM during rollout). Mark every currently SPAM-labeled
// message as processed so processInbox will not create leads for them. Gated by a Script
// Property so it runs exactly once.
function maybeBackfillSkip_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('spam_backfill_done')) return;
  try {
    const label = GmailApp.getUserLabelByName(LABEL_FILTERED);
    if (label) {
      const dedupe = loadDedupe();
      let start = 0, n = 0;
      while (true) {
        const threads = label.getThreads(start, 100);
        if (!threads.length) break;
        for (const t of threads) {
          const msgs = t.getMessages();
          for (let i = 0; i < msgs.length; i++) { dedupe[msgs[i].getId()] = Date.now(); n++; }
        }
        start += threads.length;
        if (threads.length < 100) break;
      }
      saveDedupe(dedupe);
      console.log(`maybeBackfillSkip_: marked ${n} SPAM-labeled messages as processed.`);
    }
  } catch (e) {
    console.log(`maybeBackfillSkip_: error (will retry next run) ${e && e.message}`);
    return; // do not set the flag; allow retry
  }
  props.setProperty('spam_backfill_done', '1');
}

// ===== UTILS =====
function normalizePhone(s){
  if(!s) return '';
  const t=s.trim();
  const plus=t.startsWith('+')?'+':'';
  const d=t.replace(/[^\d]/g,'');
  return d ? (plus && !d.startsWith('1') ? plus+d : d) : '';
}

// ===== PIPEDRIVE CLIENT =====
class PipedriveClient {
  constructor(){
    const props = PropertiesService.getScriptProperties();
    this.base = props.getProperty('PD_API_BASE') || 'https://api.pipedrive.com/v1';
    this.token = (props.getProperty('PD_API_TOKEN') || '').trim();
    if (!this.token) {
      throw new Error('Missing PD_API_TOKEN. Set it in Project Settings → Script Properties.');
    }
    console.log(`PD token loaded from Script Properties (ends with ${this.token.slice(-4)}).`);
  }
  call_(path, method='get', payload){
    const url = `${this.base}${path}${path.includes('?')?'&':'?'}api_token=${encodeURIComponent(this.token)}`;
    const resp = UrlFetchApp.fetch(url, {
      method,
      contentType:'application/json',
      muteHttpExceptions:true,
      payload: payload ? JSON.stringify(payload) : undefined
    });
    const code = resp.getResponseCode();
    const text = resp.getContentText();
    if (code>=200 && code<300) return text ? JSON.parse(text) : {};
    if (code === 429) {
      // Attempt to respect reset header if provided
      let headers = {};
      try { headers = resp.getAllHeaders ? resp.getAllHeaders() : {}; } catch (_) {}
      const resetSec = Number(headers['X-RateLimit-Reset'] || headers['x-ratelimit-reset'] || 0);
      const untilMs = resetSec ? (resetSec * 1000) : (Date.now() + DEFAULT_PD_BACKOFF_MS);
      PropertiesService.getScriptProperties().setProperty(PD_BACKOFF_PROP, String(untilMs));
      const err = new Error(`Pipedrive ${path} HTTP 429: ${text}`);
      err.code = 429;
      throw err;
    }
    throw new Error(`Pipedrive ${path} HTTP ${code}: ${text}`);
  }
  searchPersonByEmail(email){
    if (!email) return null;
    const res = this.call_(`/persons/search?term=${encodeURIComponent(email)}&fields=email&exact_match=true`);
    const item = res?.data?.items?.[0]?.item;
    return item ? { id: item.id, name: item.name } : null;
  }
  listPersonsPage(start, limit){
    return this.call_(`/persons?start=${start}&limit=${limit}`);
  }
  createPerson(payload){ return this.call_('/persons','post',payload).data; }
  createLead(payload){
    const clean = {};
    Object.keys(payload).forEach(k => {
      const v = payload[k];
      if (v === null || v === undefined || v === '') return;
      clean[k] = v;
    });
    return this.call_('/leads','post',clean).data;
  }
  addNoteToLead(lead_id, content){ return this.call_('/notes','post',{ lead_id, content }).data; }
}

// ===== PERSON RESOLUTION (cache-first) =====
function ensurePersonSmart(pd, parsed, runCache) {
  const email = (parsed.email || '').toLowerCase();
  const phone = normalizePhone(parsed.phone || '');

  // in-run memo
  if (email && runCache[email]) return runCache[email];
  if (!email && phone && runCache[`tel:${phone}`]) return runCache[`tel:${phone}`];

  // cache read path
  if (email) {
    const idFromEmail = cacheGetPersonId(email);
    if (idFromEmail) { runCache[email] = idFromEmail; return idFromEmail; }
  } else if (phone) {
    const idFromPhone = cacheGetPersonIdByPhone(phone);
    if (idFromPhone) { runCache[`tel:${phone}`] = idFromPhone; return idFromPhone; }
  }

  // last resort: query or create
  let personId = null;
  if (!SKIP_PERSON_SEARCH && email) {
    const found = pd.searchPersonByEmail(email);
    if (found?.id) personId = found.id;
  }
  if (!personId) {
    const payload = { name: parsed.name || parsed.email || parsed.phone || 'Website Lead' };
    if (email) payload.email = [{ value: email, primary: true, label: 'work' }];
    if (phone) payload.phone = [{ value: phone, primary: true, label: 'work' }];
    personId = pd.createPerson(payload).id;
  }

  // write-through caches
  if (email) { cacheSetPerson(email, personId, parsed.name); runCache[email] = personId; }
  if (phone) { cacheSetPhone(phone, personId); runCache[`tel:${phone}`] = personId; }

  return personId;
}

// ===== NOTE BUILDER =====
function buildNote(parsed, msg){
  const dt = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  return [
    `Source: Website contact form`,
    `Received: ${dt}`,
    `Subject: ${msg.getSubject() || ''}`,
    '',
    `Full Name: ${parsed.name || ''}`,
    `Email: ${parsed.email || ''}`,
    `Phone: ${parsed.phone || ''}`,
    `Address: ${parsed.address || ''}`,
    '',
    `Message:`,
    parsed.message || ''
  ].join('\n');
}

// ===== SPARSE CACHE SYNC (production: time-driven or manual) =====
const SYNC_CURSOR_PROP = 'pd_sync_cursor';
const SYNC_LIMIT = 200;            // persons per page
const SYNC_PAGES_PER_RUN = 3;      // cap work per invocation to conserve tokens

function sparseSyncPersonsCache() {
  const props = PropertiesService.getScriptProperties();
  const untilStr = props.getProperty(PD_BACKOFF_PROP) || '0';
  const until = Number(untilStr);
  if (until && Date.now() < until) {
    console.log(`Pipedrive backoff active; resumes at ${new Date(until).toISOString()}`);
    return;
  }

  const pd = new PipedriveClient();
  let start = Number(props.getProperty(SYNC_CURSOR_PROP) || '0');
  let processed = 0;

  try {
    for (let i = 0; i < SYNC_PAGES_PER_RUN; i++) {
      const res = pd.listPersonsPage(start, SYNC_LIMIT);
      const items = res?.data || [];
      if (!items.length) { start = 0; break; }

      for (const p of items) {
        const emails = Array.isArray(p.email) ? p.email : [];
        const phones = Array.isArray(p.phone) ? p.phone : [];
        for (const e of emails) {
          const val = (e && e.value) ? String(e.value).trim().toLowerCase() : '';
          if (val) cacheSetPerson(val, p.id, p.name);
        }
        for (const ph of phones) {
          const val = (ph && ph.value) ? String(ph.value).trim() : '';
          if (val) cacheSetPhone(val, p.id);
        }
      }

      processed += items.length;
      const more = res?.additional_data?.pagination?.more_items_in_collection;
      if (more) start = res.additional_data.pagination.next_start; else { start = 0; break; }
    }
  } catch (e) {
    if (isRateLimitError(e)) {
      console.log(`sparseSyncPersonsCache: halted due to rate limit. ${e && e.message}`);
      // PD_BACKOFF_PROP already set by client on 429
    } else {
      console.log(`sparseSyncPersonsCache: error ${e && e.message}`);
    }
  } finally {
    props.setProperty(SYNC_CURSOR_PROP, String(start));
    console.log(`sparseSyncPersonsCache: processed=${processed} next_start=${start}`);
  }
}

// ===== DEBUG UTILITIES (non-destructive) =====
function debugListInboxCandidates(limit) {
  const q = `deliveredto:${SALES_INBOX} newer_than:${SEARCH_WINDOW} -label:"${LABEL_ACCEPT}" -label:"${LABEL_FILTERED}"`;
  const threads = GmailApp.search(q, 0, limit || 20);
  const dedupe = loadDedupe();
  console.log(`Query: ${q} threads=${threads.length}`);
  let count = 0;
  for (const thread of threads) {
    const messages = thread.getMessages().slice().reverse();
    for (const msg of messages) {
      const msgId = msg.getId();
      const subject = msg.getSubject();
      if (dedupe[msgId]) { console.log(`skip msgId=${msgId} subject="${subject}" reason=deduped`); continue; }
      const parsed = parseLeadFromMessage(msg);
      if (!parsed) { console.log(`skip msgId=${msgId} subject="${subject}" reason=no-profile-match`); continue; }
      console.log(`CANDIDATE msgId=${msgId} subject="${subject}" parsed=${JSON.stringify(parsed)}`);
      count++;
      break; // one per thread
    }
  }
  console.log(`Total candidates=${count}`);
}

// Quick auth check helper to validate PD_API_TOKEN
function debugCheckPipedriveAuth(){
  try {
    const pd = new PipedriveClient();
    const res = pd.call_('/users/me', 'get');
    const user = res && res.data ? res.data : null;
    const name = user ? (user.name || '(unknown)') : '(no data)';
    const company = user ? (user.company_name || '(unknown company)') : '';
    console.log(`Pipedrive auth OK: user=${name} company=${company}`);
  } catch (e) {
    console.log(`Pipedrive auth FAIL: ${e && e.message}`);
  }
}
