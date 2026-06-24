# Design Bundles — Source-of-Truth & Provenance Contract

> **Purpose:** This is the authoritative record of *where every piece of Design Bundle
> data comes from*, how it is extracted, and how it is validated. It is the contract
> that `Code.js`, `MoodBoard.js`, and any future enrichment tooling must honor.
>
> If you are about to add a column, change an extraction step, or wire up a new
> consumer, update this file in the same change. A field with no row here has no
> defined source — that is a bug, not a feature.
>
> Companion Vault page: `Vault/wiki/curated/design-bundles.md` (the *what it is*).
> This doc is the *where the data comes from*.

---

## 1. The Source-of-Truth Model

There is exactly **one canonical store** and everything else is either an *input to*
it or a *derived output from* it.

| Role | Artifact | Notes |
|------|----------|-------|
| **CANONICAL** | `GSADUs Materials.gsheet` → **`Supplier` tab** | One row per `(Design_Bundle, Category)`. The only place data is curated. |
| Input (dirty) | External **Order Template** gsheet | Supplier-owned. We do not control its structure. Read-only to us. |
| Input (dirty) | Supplier product **websites** | Specs, real-world sizes, swatch images. Unstructured. |
| Input (dirty) | **Manual entry** by operator | Fallback for anything not yet auto-extracted. |
| Derived output | `bundles_library.json` | Machine contract for downstream tools. Regenerated, never hand-edited. |
| Derived output | `Design Bundles - Mood Board.gslides` | Human-facing deck. Regenerated. |

**Rule:** Never edit a derived output by hand. Fix the `Supplier` tab and re-run the
export. Never treat the Order Template as canonical — it is an upstream *input* we
mirror selectively.

---

## 2. Fixed Identifiers

These are hardcoded in `Code.js` and reproduced here so they live in one searchable place.

| Thing | ID / value |
|-------|-----------|
| External Order Template (file) | `1oGLgK-aCvKVh1EIhADQsqeqWQLlUaCTo4AkmtAY9dU4` |
| Order Template source tab | `Bundles` |
| Materials Drive folder | `1hc2moJgK51YPqYxcmm_Zgry5YxbsbGAs` |
| Interior Design Bundles folder | `1v7vLPjvPdMA42wGA9XqC_29DNtZP21Gk` |
| Exported JSON name | `bundles_library.json` |

### The grid (Code.js Step 1)

- **Bundles:** `Subway`, `Harbor`, `Navy`, `Olive`, `Antique`, `Villa`
- **Categories:** `Flooring`, `Bathroom Floor Tile`, `Shower Wall Tile`, `Shower Pan Tile`, `Kitchen Backsplash`, `Cabinet Color`

`Code.js` ensures one row exists for every `(bundle × category)` combination (6 × 6 = 36 baseline rows).

---

## 3. Per-Column Provenance Contract

This is the heart of the doc. Each column in the `Supplier` tab, by source.

| Column | Source | Extraction method | Validation rule | Status |
|--------|--------|-------------------|-----------------|--------|
| `Design_Bundle` | Fixed grid | `Code.js` Step 1 seeds the 6 bundles | Must be one of the 6 known bundles | ✅ deterministic |
| `Category` | Fixed grid | `Code.js` Step 1 seeds the 6 categories | Must be one of the 6 known categories | ✅ deterministic |
| `Supplier_URL` | **Order Template** | `Code.js` Step 1 — reads `=HYPERLINK()` / rich-text link from the template cell, writes `=HYPERLINK(url, displayText)` | URL must resolve (http/https); display text non-empty | ✅ automated |
| `Supplier` | **Website / manual** | ⚠️ **No automation today** — hand-typed | Non-empty; canonical brand name (drives filename) | ❌ **GAP — enrichment target** |
| `Product_Name` | **Website / manual** | ⚠️ Hand-typed (often = `Supplier_URL` display text, but not enforced) | Non-empty | ❌ **GAP — enrichment target** |
| `Product_Size` | **Website / manual** | ⚠️ Hand-typed | Real-world dims, e.g. `12x24 in`, `9.33x75 in` | ❌ **GAP — enrichment target** |
| `File_ID` | Computed | `Code.js` Step 2 — resolves file, returns Drive ID | Valid Drive file ID | ✅ automated |
| `Drive_URL` | Computed | `Code.js` Step 2 — canonical Drive URL after move/rename | Points into the Materials folder | ✅ automated |
| `Filename` | Computed | `Code.js` Step 2 — `Supplier_Product-Name.ext` (see `canonicalBasename`) | Matches canonical grammar | ✅ automated |
| `Sync_Status` | Computed | `Code.js` Step 2 — run result/timestamp per row | Internal bookkeeping; not exported | ✅ automated |
| `VScale` | Computed (partial) | `Compute Missing Scales` — derives from image aspect ratio + the *other* scale | Inches; needs `Product_Size` understanding to seed | ⚠️ partial |
| `HScale` | Computed (partial) | Same as `VScale` | Inches | ⚠️ partial |

### Provenance legend

- **Fixed grid** — deterministic, defined in `Code.js`. No external dependency.
- **Order Template** — mirrored from the external supplier-owned sheet. We never write back to it.
- **Website / manual** — currently human-entered; the primary hardening target (see §4).
- **Computed** — derived by `Code.js` from other columns + Drive/image state. Safe to regenerate.

### The core gap

The only auto-populated input is `Supplier_URL` (the *link*). The three fields that
describe what's *behind* the link — `Supplier`, `Product_Name`, `Product_Size` — are
hand-typed, and `Product_Size` also gates the scale math. Everything downstream is
already deterministic once those three exist. **This is the single seam where data
provenance is undefined**, and the target for the enrichment step in §4.

---

## 4. Proposed: Step 1.5 — Agentic Enrichment

> Status: **PROPOSED**, not yet built. Documented here so the contract is agreed
> before implementation.

A step between Step 1 (pull links) and Step 2 (sync assets) that converts the dirty
`Supplier_URL` into structured fields:

**Input:** rows where `Supplier_URL` is set but `Supplier` / `Product_Name` /
`Product_Size` are blank.

**Action (per row):**
1. Fetch the supplier product page (WebFetch / Chrome MCP).
2. Extract `Supplier`, `Product_Name`, `Product_Size` (and, where available, the swatch image → stage into the `Drive_URL` slot for Step 2 to canonicalize).
3. Write back with a **confidence flag**.

**Proposed new columns (additive — keep `Code.js` column-safe design):**

| Column | Purpose |
|--------|---------|
| `Provenance` | How the row's fields were filled: `template` / `extracted` / `manual` |
| `Needs_Review` | `TRUE` when extraction confidence is low — operator reviews only the exceptions |
| `Last_Enriched` | ISO timestamp of the last enrichment pass |

**Principle:** the agent fills what it can with high confidence and flags the rest.
The human reviews exceptions, not the whole sheet. No silent guessing into canonical
fields.

---

## 5. Downstream Contract — `bundles_library.json`

Written by `Code.js` Step 3 (`exportToJson`). Schema `1.1`:

```
{ _meta: { last_sync, source, schema_version },
  hardware: [ { name, image_url } ],
  bundles: [ { name, materials: [ {
    category, supplier, product_name, product_size,
    product_url, drive_file_id, drive_url, filename
  } ] } ] }
```

- `product_url` is parsed out of the `=HYPERLINK()` formula in `Supplier_URL`.
- `Sync_Status` is **intentionally omitted** — internal bookkeeping only.
- **Proposed extension:** carry `vscale` / `hscale` (non-null only) so Revit/Architextures
  and PNGTools can consume real-world scale without re-deriving it. See
  `Moodboard/Material-Scale-plan.md` §5.

**Consumers:** PNGTools (render content, live) · Revit Design Bundles via Architextures
tiling (the VScale/HScale path — GAP-05 / INT-02, open) · WebCatalog/Supabase (proposed).

---

## 6. Legacy — Already Retired (verified 2026-06-24)

The material-data pipeline is already converged on `bundles_library.json`. There is
**no legacy code or data left to remove** — recorded here so the history is clear and
nobody re-introduces a parallel path (Rule #6).

| Artifact | Status |
|----------|--------|
| `gsadus_materials.db` (SQLite) | **Gone.** No longer on the Drive; no live code reads it. Only a historical note remains in `DigitalDarkroom/docs/archive/`. |
| `db_export.json` | **Not legacy.** Active PNGTools prompt-config store — `exterior_styles` / `interior_styles` / `environments` only. Never held the canonical bundle data after migration. |
| Bundle/material consumption | **Migrated.** `PostProcess/PNGTools/core/darkroom/prompt_engine.py` loads bundles + hardware from `bundles_library.json` (`load_bundles_library()`). |

`bundles_library.json` is therefore the sole source of bundle/material data for
downstream consumers today.

---

## 7. Cross-References

- `Vault/wiki/curated/design-bundles.md` — what Design Bundles are + how Revit applies them
- `Vault/wiki/auto/database-layer.md` — all databases & Apps Script projects
- `Vault/wiki/curated/planning.md` — GAP-05 (materials undefined), INT-02 (sheet → Revit sync)
- `Vault/wiki/curated/architextures-material-sync.md` — texture path setup (VScale/HScale consumer)
- `Moodboard/Material-Scale-plan.md` — VScale/HScale derivation + JSON extension
- `Code.js` — the 3-step ETL implementing this contract
