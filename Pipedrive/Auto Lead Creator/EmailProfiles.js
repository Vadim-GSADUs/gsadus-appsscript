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

      // Email Id: email@domain.com
      email: {
        regex: /Email Id\s*:\s*([^\s]+@[^\s]+)/i
      },

      // Mobile : 9167705270
      mobilePhone: {
        regex: /Mobile\s*:\s*([0-9+()\-\s]+)/i
      },

      // Full Address : ...
      address: {
        regex: /Full Address\s*:\s*(.+?)(?=\s*(Message\s*:|--|$))/is
      },

      // Message : ...
      note: {
        regex: /Message\s*:\s*([\s\S]+?)(?=\s*--\s*$|$)/i
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
      // User Name : [username]
      fullName: {
        regex: /User Name\s*:\s*(.+?)(?=\s*(Phone\s*:|Email Address\s*:|Message\s*:|--|$))/i
      },

      // Email Address : email@domain.com
      email: {
        regex: /Email Address\s*:\s*([^\s]+@[^\s]+)/i
      },

      // Phone : 9168060809
      mobilePhone: {
        regex: /Phone\s*:\s*([0-9+()\-\s]+)/i
      },

      // No explicit address in this template; leave null.
      // Add later if they change the form.

      // Message : ...
      note: {
        regex: /Message\s*:\s*([\s\S]+?)(?=\s*--\s*$|$)/i
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
  if (profile.subjectRegex) {
    return profile.subjectRegex.test(subject);
  }
  if (profile.subjectIncludes) {
    return subject.indexOf(profile.subjectIncludes) !== -1;
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

  // Helper: extract first group of regex, trimmed.
  function extract(pattern) {
    if (!pattern || !pattern.regex) return null;
    var m = body.match(pattern.regex);
    if (!m || !m[1]) return null;
    return m[1].toString().trim();
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
