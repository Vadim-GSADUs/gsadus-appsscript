/**
 * SpamFilter.gs
 *
 * Backstop spam / test classification for inbound website leads.
 *
 *   const v = classifyLead({ name, email, phone, message });
 *   if (v.isSpam || v.isTest) { skip }
 *
 * Design goals:
 *  - HIGH PRECISION: never drop a genuine ADU lead. Only reject on strong signals,
 *    or when two weaker signals agree. When in doubt, let it through.
 *  - Transparent: returns `reasons` so filtered items can be audited (they are
 *    labeled in Gmail, not deleted).
 *
 * IMPORTANT: This is a backstop, not the primary defense. Structurally-clean bot
 * spam (plausible name, free-mail address, on-topic-looking text) has no reliable
 * signal here and will pass. The durable fix is at the form: reCAPTCHA v3 or a
 * honeypot field on the Elementor form (gsadus.com). Keep this filter conservative.
 */

// Casino / SEO / pharma bot keywords (substring match, lowercased).
var SPAM_KEYWORDS = [
  'mostbet', '1xbet', '888starz', 'pin-up', 'pinup', 'vulkan', 'casino',
  'gate of olympus', 'crypto', 'bitcoin', 'forex', 'binary option',
  'viagra', 'cialis', 'porn', 'escort', 'payday loan', 'backlink',
  'seo service', 'rank your', 'buy followers', 'rosgvard'
];

// TLDs essentially never used by genuine California ADU customers.
var SPAM_TLDS = ['ru', 'su', 'cn', 'top', 'xyz', 'click', 'space', 'sbs', 'icu', 'rest', 'buzz', 'monster'];

// Known internal / vendor test markers.
var TEST_EMAILS = ['sample@only.com', 'test@test.test', 'kova@sampletest.com'];
var TEST_DOMAINS = ['kova.team'];
var TEST_NAME_RE = /^(sample only|sample|testing|test|kova test|marketing test)\b/i;

var SPAM_SCORE_THRESHOLD = 3;

/**
 * @param {{name:string, email:string, phone:string, message:string}} input
 * @return {{isSpam:boolean, isTest:boolean, score:number, reasons:string[]}}
 */
function classifyLead(input) {
  input = input || {};
  var name  = String(input.name || '');
  var email = String(input.email || '').toLowerCase().trim();
  var phone = String(input.phone || '');
  var msg   = String(input.message || '');
  var reasons = [];
  var score = 0;

  var nameAndMsg = name + ' ' + msg;
  var hay = (name + ' ' + msg + ' ' + email).toLowerCase();
  var domain = email.indexOf('@') !== -1 ? email.split('@')[1] : '';
  var tld = domain ? domain.split('.').pop() : '';

  // ---- Strong signals (any one is enough on its own) ----
  // Cyrillic / CJK / Hiragana-Katakana => not a local ADU lead.
  if (/[Ѐ-ӿ一-鿿぀-ヿ가-힯]/.test(nameAndMsg)) {
    score += 3; reasons.push('non-latin-script');
  }
  // Blocklisted keyword.
  for (var i = 0; i < SPAM_KEYWORDS.length; i++) {
    if (hay.indexOf(SPAM_KEYWORDS[i]) !== -1) { score += 3; reasons.push('keyword:' + SPAM_KEYWORDS[i]); break; }
  }

  // ---- Medium signals (need one more to reject) ----
  // URL / anchor in name or message. Weighted so a lone link from a real
  // homeowner does not, by itself, get dropped.
  if (/(https?:\/\/|www\.|\[url|<a\s|\bhref=)/i.test(nameAndMsg)) {
    score += 2; reasons.push('contains-url');
  }
  // Junk email TLD.
  if (tld && SPAM_TLDS.indexOf(tld) !== -1) { score += 2; reasons.push('junk-tld:' + tld); }
  // Bot name pattern: word + underscore + 2-4 mixed letters (mostbet_tnEr, kolca_wooa).
  if (/[A-Za-z]{3,}_[A-Za-z]{2,4}\b/.test(name)) { score += 2; reasons.push('bot-name-suffix'); }

  // ---- Weak supporting signals ----
  // Phone present but clearly invalid (e.g. "0", "505050"); empty phone is fine (email-only lead).
  var digits = phone.replace(/[^\d]/g, '');
  if (digits.length > 0 && digits.length < 7) { score += 1; reasons.push('invalid-phone'); }

  var isSpam = score >= SPAM_SCORE_THRESHOLD;

  // ---- Test / vendor bucket (skip but track separately) ----
  var isTest = false;
  if (TEST_EMAILS.indexOf(email) !== -1) { isTest = true; reasons.push('test-email'); }
  if (domain && TEST_DOMAINS.indexOf(domain) !== -1) { isTest = true; reasons.push('test-domain'); }
  if (TEST_NAME_RE.test(name.trim())) { isTest = true; reasons.push('test-name'); }

  return { isSpam: isSpam, isTest: isTest, score: score, reasons: reasons };
}
