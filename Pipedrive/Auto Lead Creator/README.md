# GSADUS → Pipedrive Lead Automation

This Apps Script project ingests lead emails from Gmail, normalizes them via profile-based parsing, and creates Pipedrive leads with notes. It is optimized for idempotency, caching, and Pipedrive rate limits.

## Files
- `LeadAutoCreator.js`: Main runtime — Gmail search, candidate collection, person/lead creation, notes, labels, idempotency, and PD backoff.
- `EmailProfiles.js`: Profile-based parser for diverse email templates. Provides `parseLeadFromMessage(msg)` which returns normalized lead data. Four profiles: `Request form Query`, `City Query Form`, Elementor `New message from "Golden State ADUs"` (also covers the email-only footer form), and the `New ADU quote request from …` "Drop Us a Note" form.
- `SpamFilter.js`: Backstop `classifyLead()` — strong/medium/weak scoring, threshold 4 (one strong or two medium signals reject; weak signals never reject on their own). Rejected mail is labeled `SPAM (AUTO)`, never deleted; `1 - SALES/FORCE (MANUAL)` overrides.
- `maintenance.js`: Support tasks (e.g., sparse persons cache sync).

## Flow
1. `processInbox()` searches Gmail via `searchLeadThreads_()`: `deliveredto:SALES_INBOX newer_than:SEARCH_WINDOW` plus a subject clause built from the profiles, paged 50 threads at a time, **and the same query scoped to `in:spam`** (`SEARCH_SPAM_TOO`). Gmail's own filter spam-folders roughly half of the website-form mail because the Hostinger web server sends it with an unaligned envelope (DMARC fails for gsadus.com), so the Inbox alone misses real leads.
2. `collectCandidates()` uses `parseLeadFromMessage(msg)` to recognize and parse lead emails into a normalized shape, then `classifyLead()` as the spam backstop.
3. `ensurePersonSmart()` resolves/creates a Pipedrive person using caches and optional search.
4. Creates a Pipedrive lead with optional address custom field and adds a detailed note.
5. Marks the message read, labels the thread, moves a spam-foldered thread back to the Inbox (`rescueFromSpam_`), and writes dedupe state.

Already-handled messages are skipped by the `processed_msg_ids` dedupe map. Gmail `-label:` exclusions do not match these label names and are not used.

### Backfilling after an outage
To recover older mail (as on 2026-09-01, when ~30 August leads were pulled out of Gmail's Spam), temporarily widen `SEARCH_WINDOW` (e.g. `35d`) and raise `MAX_PER_RUN`, let the scheduled trigger drain the backlog over a few runs, then restore `3d` / `5`. The dedupe map keeps re-runs idempotent. The root cause of that outage is outside this script: the website's mail must authenticate for gsadus.com (SPF/DKIM alignment), or Gmail must be told to trust the web host (admin allowlist / "never send to Spam" filter).

## Configuration
Set Script Properties:
- `PD_API_TOKEN` (required): Pipedrive API token.
- `PD_API_BASE` (optional): Defaults to `https://api.pipedrive.com/v1`.
- `PD_SKIP_PERSON_SEARCH` (optional): `'true'` to skip PD search and always create persons.

Key constants in `LeadAutoCreator.js`:
- `SALES_INBOX`: Target Gmail inbox.
- `LABEL_ACCEPT`: Label applied after successful processing.
- `LEAD_ADDRESS_FIELD_KEY`: Pipedrive Lead custom field key for address.
- `MAX_PER_RUN`, `SEARCH_WINDOW`: Throughput and Gmail window.
- Dedupe: `DEDUPE_KEY`, `DEDUPE_TTL_DAYS`.
- PD backoff: `PD_BACKOFF_PROP`, `DEFAULT_PD_BACKOFF_MS`.

## Adding Profiles
Profiles live in `EmailProfiles.js` under `EMAIL_PROFILES`. Each profile describes subject recognition and body regexes.
- Use `subjectIncludes` when possible; `subjectRegex` only if necessary.
- Regex must capture values in group 1; use lazy quantifiers and stop at the next label/`--`/end.
- Body is already plaintext with line breaks.

A reusable AI guidance block is appended to `EmailProfiles.js` to help generate new profiles. Paste sample emails to an AI and request a single JavaScript object ready to append.

## Notes on Refactor
- Legacy subject/sender heuristics removed; detection relies on `parseLeadFromMessage(msg)` for robustness and scalability.
- Keep `normalizePhone` utility for consistent phone normalization across modules.

## Running
- Set a time-driven trigger for `processInbox()` (e.g., every 5–30 minutes).
- Manual tests: run `processInbox()` with recent sample emails.

## Rate Limits & Backoff
- On PD 429, a backoff window is stored in `PD_BACKOFF_PROP`. Processing halts until reset.

## Caching
- Persons cache mirrors PD email→id mapping in Script Properties and Script Cache.
- Phone cache provides phone→person id.
- `sparseSyncPersonsCache()` paginates PD persons to refresh caches opportunistically.

## Troubleshooting
- Ensure `PD_API_TOKEN` is set and valid.
- Validate `LEAD_ADDRESS_FIELD_KEY` matches your PD custom field key.
- Check Gmail label path correctness.
- Review Logs for backoff messages and parsing outcomes.