/**
 * EmailProfiles.gs
 *
 * Single responsibility:
 *  - Inspect a GmailMessage.
 *  - Match against configured “profiles”.
 *  - Parse and return normalized lead data.
 *
 * Normalized result shape:
 * {
 *   profileKey: 'PROFILE_REQUEST_FORM' | 'PROFILE_CITY_QUERY' | ...,
 *   subject:    string,
 *   date:       Date,
 *   fullName:   string|null,
 *   email:      string|null,
 *   mobilePhone:string|null,
 *   address:    string|null,
 *   note:       string|null
 * }
 *
 * If message does not match any profile, returns null.
 */


/**
 * MAIN ENTRY POINT for other scripts.
 * Call from your automation like:
 *
 *   const parsed = parseLeadFromMessage(msg);
 *   if (!parsed) continue;
 */
function parseLeadFromMessage(msg) {
  if (!msg || typeof msg.getSubject !== 'function' || typeof msg.getBody !== 'function' || typeof msg.getDate !== 'function') {
    return null;
  }
  const subject = msg.getSubject() || '';
  const date    = msg.getDate();
  const bodyRaw = msg.getBody() || '';
  const body    = stripHtmlToText_(bodyRaw);

  for (var i = 0; i < EMAIL_PROFILES.length; i++) {
    var profile = EMAIL_PROFILES[i];

    if (!subjectMatchesProfile_(subject, profile)) continue;

    var parsed = parseWithProfile_(body, profile);
    if (!parsed) continue;

    // Attach common fields
    parsed.profileKey = profile.key;
    parsed.subject    = subject;
    parsed.date       = date;
    return parsed;
  }

  return null;
}

/**
 * CONFIG: list of profiles. Add new profiles here as you meet new templates.
 *
 * Each profile:
 *  - key: unique identifier.
 *  - subjectIncludes / subjectRegex: how to recognize by subject.
 *  - patterns: regex for fields in the plain-text body.
 */
var EMAIL_PROFILES = [
  // ---------------------------------------------------------
  // PROFILE 1: "Request form Query" (original GSADUS form)
  // ---------------------------------------------------------
  {
    key: 'PROFILE_REQUEST_FORM',
    subjectIncludes: 'Request form Query',

    patterns: {
      // Full Name : shirley
      fullName: {
        regex: /Full Name\s*:\s*(.+?)(?=\s*(Mobile\s*:|Email Id\s*:|Full Address\s*:|Message\s*:|--|$))/i
      },

      // Email Id / Email Address / Email: email@domain.com
      email: {
        regex: /(?:Email\s*Id|Email\s*Address|Email)\s*:\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i
      },

      // Mobile / Phone / Phone Number : 9167705270
      mobilePhone: {
        regex: /(?:Mobile|Phone|Phone\s*Number)\s*:\s*([0-9+()\-\s]+)/i
      },

      // Full Address / Address : ...
      address: {
        regex: /(?:Full\s*Address|Address)\s*:\s*(.+?)(?=\s*(?:Message\s*:|Comments\s*:|--|$))/is
      },

      // Message / Comments : ...
      note: {
        regex: /(?:Message|Comments)\s*:\s*([\s\S]+?)(?=\s*--\s*$|$)/i
      }
    }
  },

  // ---------------------------------------------------------
  // PROFILE 2: "City Query Form" (WordPress form)
  // ---------------------------------------------------------
  {
    key: 'PROFILE_CITY_QUERY',
    subjectIncludes: 'City Query Form',

    patterns: {
      // User Name / Name : [username]
      fullName: {
        regex: /(?:User\s*Name|Name)\s*:\s*(.+?)(?=\s*(?:Phone\s*:|Phone\s*Number\s*:|Email\s*Address\s*:|Email\s*:|Message\s*:|--|$))/i
      },

      // Email Address / Email : email@domain.com
      email: {
        regex: /(?:Email\s*Address|Email)\s*:\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i
      },

      // Phone / Phone Number : 9168060809
      mobilePhone: {
        regex: /(?:Phone|Phone\s*Number)\s*:\s*([0-9+()\-\s]+)/i
      },

      // No explicit address in this template; leave null.
      // Add later if they change the form.

      // Message / Comments : ...
      note: {
        regex: /(?:Message|Comments)\s*:\s*([\s\S]+?)(?=\s*--\s*$|$)/i
      }
    }
  },

  // ---------------------------------------------------------
  // PROFILE 3: "Elementor Contact Form" (redesigned gsadus.com website, added 2026-07-22)
  // Subject: New message from "Golden State ADUs"
  // From:    Golden State ADUs <email@gsadus.com>
  // Body labels: Full Name / Phone / Email / Address / Type of Visit / Message
  // A "---" separator line precedes an Elementor footer (Date / Time / Page URL /
  // User Agent / Remote IP / Powered by), which must be excluded from parsed fields.
  // ---------------------------------------------------------
  {
    key: 'PROFILE_WEB_CONTACT_ELEMENTOR',
    subjectIncludes: 'New message from "Golden State ADUs"',

    patterns: {
      // Full Name: John Doe  (value stays on its own line; empty -> no value)
      fullName: {
        regex: /Full\s*Name\s*:[ \t]*(.*?)(?=\s*(?:Phone\s*:|Email\s*:|Address\s*:|Type\s*of\s*Visit\s*:|Message\s*:|-{3,}|$))/i
      },

      // Email: john@example.com
      email: {
        regex: /Email\s*:\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i
      },

      // Phone: 5551234567
      mobilePhone: {
        regex: /Phone\s*:\s*([0-9+()\-\s]+)/i
      },

      // Address: 123 Main St, City  (value stays on its own line; empty -> no value)
      address: {
        regex: /Address\s*:[ \t]*(.*?)(?=\s*(?:Type\s*of\s*Visit\s*:|Message\s*:|-{3,}|$))/i
      },

      // Message: ... (stop at the "---" footer separator or a footer label line)
      note: {
        regex: /Message\s*:\s*([\s\S]*?)(?=\s*-{3,}\s*(?:\r?\n|$)|\n\s*(?:Date|Time|Page\s*URL|User\s*Agent|Remote\s*IP|Powered\s*by)\s*:|$)/i
      }
    }
  },

  // ---------------------------------------------------------
  // PROFILE 4: "Drop Us a Note" quote form (gsadus.com, first seen 2026-08-03; added 2026-09-01)
  // Subject: New ADU quote request from {name}
  // From:    sales@gsadus.com  (the site sends it AS the sales inbox; Reply-To = requester)
  // Body:    an HTML table where each label sits in its own row and the value in the NEXT
  //          row — so after HTML stripping every value is on the line(s) after its label:
  //            Full name / Phone / Email / Address / Drop Us a Note
  //          Regexes therefore anchor on "label, newline, first non-blank line", and a
  //          negative lookahead stops an EMPTY field from swallowing the next label.
  // ---------------------------------------------------------
  {
    key: 'PROFILE_QUOTE_REQUEST',
    subjectIncludes: 'New ADU quote request from',

    patterns: {
      fullName: {
        regex: /(?:^|\n)[ \t]*Full\s*name[ \t]*\n\s*(?!(?:Phone|Email|Address|Drop\s*Us\s*a\s*Note)\s*\n)([^\n]*\S[^\n]*)/i
      },
      mobilePhone: {
        regex: /(?:^|\n)[ \t]*Phone[ \t]*\n\s*(?!(?:Email|Address|Drop\s*Us\s*a\s*Note)\s*\n)([0-9+()\-.\s]*\d[0-9+()\-.\s]*)/i
      },
      email: {
        regex: /(?:^|\n)[ \t]*Email[ \t]*\n\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i
      },
      address: {
        regex: /(?:^|\n)[ \t]*Address[ \t]*\n\s*(?!(?:Drop\s*Us\s*a\s*Note)\s*\n)([^\n]*\S[^\n]*)/i
      },
      // Everything after the "Drop Us a Note" label up to the form-plugin footer.
      note: {
        regex: /(?:^|\n)[ \t]*Drop\s*Us\s*a\s*Note[ \t]*\n\s*([\s\S]*?)(?=\n\s*(?:This\s+e-?mail\s+was\s+sent|Sent\s+from|Powered\s+by|-{3,})|\s*$)/i
      }
    }
  }
  // Add more profiles here as needed.
];


/* ===== INTERNAL HELPERS ===== */

/**
 * Decide if a subject matches the profile.
 * Supports simple includes or full regex.
 */
function subjectMatchesProfile_(subject, profile) {
  subject = subject || '';
  if (profile.subjectRegex) {
    try {
      return profile.subjectRegex.test(subject);
    } catch (e) {
      return false;
    }
  }
  if (profile.subjectIncludes) {
    return subject.toLowerCase().indexOf(String(profile.subjectIncludes).toLowerCase()) !== -1;
  }
  return false;
}

/**
 * Parses body text using profile.patterns config.
 * Returns normalized object or null if not enough data.
 */
function parseWithProfile_(body, profile) {
  var patterns = profile.patterns || {};
  var result = {
    fullName:    null,
    email:       null,
    mobilePhone: null,
    address:     null,
    note:        null
  };

  // Helper: extract first group of regex (or full match), trimmed.
  function extract(pattern) {
    if (!pattern || !pattern.regex) return null;
    var m = body.match(pattern.regex);
    if (!m) return null;
    var raw = (m[1] !== undefined && m[1] !== null) ? m[1] : m[0];
    if (!raw) return null;
    return String(raw).toString().trim();
  }

  result.fullName    = extract(patterns.fullName);
  result.email       = extract(patterns.email);
  result.mobilePhone = normalizePhone_(extract(patterns.mobilePhone));
  result.address     = extract(patterns.address);
  result.note        = extract(patterns.note);

  // Minimal validity: must have at least email OR phone.
  if (!result.email && !result.mobilePhone) return null;

  return result;
}

/**
 * Very simple phone normalizer.
 * You can replace this with your existing normalizePhone()
 * if you already have one in another file.
 */
function normalizePhone_(raw) {
  if (!raw) return null;
  var digits = String(raw).replace(/[^\d+]/g, '');
  return digits || null;
}

/**
 * Strips HTML into plaintext. Keeps line breaks somewhat sane.
 */
function stripHtmlToText_(html) {
  if (!html) return '';
  var text = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, '');
  return text
    .replace(/\r/g, '')
    .replace(/\u00A0/g, ' ')
    // Decode a few common HTML entities that appear in emails
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'")
    .trim();
}

/*
============================================================
AI Guidance: Generating New Email Profiles
============================================================
Reusable instruction text for the AI — paste this when asking an AI
to create a new EMAIL_PROFILES entry based on sample emails.

You can store this as a note and paste it whenever you need a new profile:

I have a Gmail parsing system that uses “profiles” to normalize lead data.
Each profile is a JavaScript object inside an EMAIL_PROFILES array, used by a function parseLeadFromMessage.

I will paste one or more sample emails of the same template.
Based on those samples, you must produce exactly one JavaScript object that matches this schema:

{
  key: 'PROFILE_SOMETHING',
  subjectIncludes: '...',      // or use subjectRegex instead
  // subjectRegex: /.../i,     // only if needed
  patterns: {
    fullName:    { regex: /.../i } || null or omit,
    email:       { regex: /.../i } || null or omit,
    mobilePhone: { regex: /.../i } || null or omit,
    address:     { regex: /.../i } || null or omit,
    note:        { regex: /.../is } || null or omit
  }
}


Rules:

key must be unique and ALL_CAPS with a PROFILE_ prefix.

Prefer subjectIncludes: 'Exact subject string' if possible. Use subjectRegex only if strictly necessary.

In each regex, capture the value in group 1. Use lazy quantifiers where needed and stop at the next label, --, or end of text.

Use flags i for case-insensitive and is when matching multi-line body sections (like note).

The body has already been converted to plain text with line breaks; don’t write regexes that depend on HTML tags.

At minimum, the profile must reliably capture either an email or mobile phone number, or both.

Output only the JavaScript object, no explanation, no backticks.

After I paste the sample email(s), infer the best regex patterns and respond with the single profile object ready to append to my EMAIL_PROFILES array.
*/

function testParseLatest() {
  var t = GmailApp.search('newer_than:7d deliveredto:Sales@gsadus.com', 0, 1)[0];
  if (!t) { Logger.log('No thread'); return; }
  var msg = t.getMessages().slice(-1)[0]; // newest
  var parsed = parseLeadFromMessage(msg);
  Logger.log(JSON.stringify(parsed));
}

/**
 * Diagnostics: iterate recent messages and explain profile matching and extraction.
 */
function diagnoseRecentLeads(limit) {
  limit = limit || 10;
  var threads = GmailApp.search('newer_than:7d deliveredto:Sales@gsadus.com', 0, limit);
  if (!threads || !threads.length) { Logger.log('No threads'); return; }
  for (var i = 0; i < threads.length; i++) {
    var msg = threads[i].getMessages().slice(-1)[0];
    if (!msg) continue;
    var subject = msg.getSubject() || '';
    var body = stripHtmlToText_(msg.getBody() || '');
    Logger.log('--- Thread #' + (i+1) + ' ---');
    Logger.log('Subject: ' + subject);
    var matched = false;
    for (var p = 0; p < EMAIL_PROFILES.length; p++) {
      var profile = EMAIL_PROFILES[p];
      var subMatch = subjectMatchesProfile_(subject, profile);
      if (!subMatch) continue;
      matched = true;
      var r = parseWithProfile_(body, profile);
      Logger.log('Profile: ' + profile.key + ' => ' + JSON.stringify(r));
    }
    if (!matched) Logger.log('No profile subject match for this message.');
  }
}
