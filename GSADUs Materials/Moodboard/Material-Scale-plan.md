# Material Scale Plan — Mood Board Image Sizing

## Problem

Material images are not captured at a consistent real-world scale. A hardwood
plank sample showing 5 planks stacked represents 46.65 in of real material
height (5 × 9.33 in plank width), while a 12×24 tile image may represent only
24 in. Without normalization, the mood board looks visually inconsistent —
smaller tiles appear the same size as wider planks.

## Goal

Scale each image on the slide so materials appear at a consistent real-world
proportion relative to one another. Layout positions are **not** enforced by
the script — the user freely arranges images on the slide; the script only
updates image content and size on re-sync.

---

## 1. Spreadsheet Schema — New Columns

Add two optional columns to the **Supplier** tab:

| Col | Name   | Type   | Description |
|-----|--------|--------|-------------|
| K   | VScale | number | Real-world height (in) the image represents |
| L   | HScale | number | Real-world width (in) the image represents |

**Rules:**
- Both are optional. If neither is set, image inserts at default card size.
- Typically only one is needed per material type (see examples below).
- Values are in **inches**.

### Example calculations

| Material | Product_Size | Image contents | VScale calc | HScale calc |
|---|---|---|---|---|
| Hardwood plank | 9.33×75 in | 5 planks stacked | 5 × 9.33 = **46.65** | — |
| 12×24 floor tile | 12×24 in | 2×2 tile array | 2 × 24 = **48** | 2 × 12 = **24** |
| 4×12 wall tile (3×5 array) | 4×12 in | 3 cols × 5 rows | 5 × 12 = **60** | 3 × 4 = **12** |
| Penny round mosaic (sheet) | 1 in dia | sheet ~12×12 in | **12** | **12** |
| Cabinet door swatch | n/a | single door | — | — (no scale) |

---

## 2. Scale Math

Define a **reference scale** `REF_INCHES` (default: **24 in**).
This is the real-world size that maps to the base card dimension on the slide.

```
base_card_w = MB_CELL_W  (240 pt)
base_card_h = MB_CELL_H  (248 pt)

If VScale defined:
  img_h = base_card_h × (REF_INCHES / VScale)
  img_w = img_h × (native_px_w / native_px_h)   ← preserve aspect ratio

If HScale defined (and VScale not defined):
  img_w = base_card_w × (REF_INCHES / HScale)
  img_h = img_w × (native_px_h / native_px_w)   ← preserve aspect ratio

If both defined:
  img_w = base_card_w × (REF_INCHES / HScale)
  img_h = base_card_h × (REF_INCHES / VScale)
  ← may intentionally distort for repeating tile patterns

If neither defined:
  img_w = base_card_w
  img_h = base_card_h  ← fills card, current behavior
```

### Example — Hardwood plank (VScale = 46.65)

```
img_h = 248 × (24 / 46.65) = 248 × 0.515 = 127.6 pt  ≈ 128 pt
img_w = 128 × (native_w / native_h)                   ← aspect-corrected
```

The plank image will appear ~half the card height, correctly smaller than a
tile sheet that represents only 12 in (which would fill the full card height).

### Reference scale tuning

`REF_INCHES = 24` is a starting point. The user can adjust this constant in
`MoodBoard.js` to zoom all materials in or out uniformly:

- Lower value (e.g. 12) → all images appear larger (zoomed in)
- Higher value (e.g. 48) → all images appear smaller (zoomed out, more context)

---

## 3. Slide Update Strategy — Tag-Based, Position-Preserving

### Problem with current approach
The current script clears ALL slide elements and rebuilds from scratch on every
sync. This destroys any manual layout the user has set up.

### New approach

**On first insert** (image not yet on slide):
- Insert image at a default staggered position (not a rigid grid)
- Tag the image with the category name via `image.setTitle(category)`
  e.g. `"Flooring"`, `"Bathroom Floor Tile"`, etc.
- Also tag with bundle name via `image.setDescription(bundleName)`

**On re-sync** (image already on slide):
- Scan `slide.getImages()` for `img.getTitle() === category`
- Found → replace blob in-place, resize per scale math, keep `(left, top)` unchanged
- Not found → insert at default position with tag

**Text elements** (category label, product name, size):
- Tagged similarly via shape title: e.g. `"label:Flooring"`
- On re-sync: find by title and update text content only
- Position preserved

### Default insert positions (first run, no prior layout)
Instead of a strict 3×2 grid, stagger positions slightly so elements are
individually movable without overlapping:

```
Col 0: left = 10
Col 1: left = 255
Col 2: left = 500
Row 0: top  = MB_GRID_TOP + 10
Row 1: top  = MB_GRID_TOP + MB_CELL_H + 10
```

These are starting positions only — user may freely reposition after first run.

---

## 4. Code Changes Required

### MoodBoard.js

1. **New constants:**
   ```javascript
   const REF_INCHES  = 24;   // reference real-world inches → base card dimension
   const MB_COL_VSCALE = 10; // col K (0-based)
   const MB_COL_HSCALE = 11; // col L (0-based)
   const MB_NUM_COLS   = 12; // expand from 10 → 12
   ```

2. **`syncMoodBoard()`** — read K & L columns alongside existing data

3. **`buildBundleSlide_()`** → replace with **`updateBundleSlide_()`**:
   - Does NOT call `slide.getPageElements().forEach(el => el.remove())`
   - Instead finds existing images/text by title tag
   - Calls `computeImageSize_(vscale, hscale, nativeW, nativeH)` for dimensions
   - Inserts new or updates existing

4. **New helper `computeImageSize_(vscale, hscale, nativeW, nativeH)`:**
   - Implements the scale math above
   - Returns `{ w, h }`

5. **New helper `findByTitle_(elements, title)`:**
   - Returns first element where `.getTitle() === title`, or null

### Supplier sheet

- Add col K header: `VScale`
- Add col L header: `HScale`
- Update `NUM_COLS = 12` in `Code.js` constants (only affects Step 3 JSON export
  if we want to include scale in `bundles_library.json` — optional)

---

## 5. bundles_library.json (optional extension)

If scale data should be available to PNGTools / PromptBuilder, add to each
material entry:

```json
{
  "category": "Flooring",
  "vscale": 46.65,
  "hscale": null,
  ...
}
```

Only export non-null values. `exportToJson()` in `Code.js` reads cols K & L
the same way.

---

## 6. Decisions (resolved)

1. **REF_INCHES = 24** — constant in `MoodBoard.js`. HTML sidebar deferred.
2. **VScale primary** — if both defined, VScale wins and width is derived from
   native aspect ratio. HScale used only when VScale is absent. Distortion mode
   deferred until a real skewed-image case arises.
3. **Floating text boxes** — overlay bar removed. One text box per image
   (cat + product name + size in a single multi-line box), co-located below the
   image. Independent (not grouped) so user can reposition separately.
4. **Standard 3×2 grid** for first-run insert positions (col × MB_CELL_W,
   MB_GRID_TOP + row × MB_CELL_H). Fits slide bounds exactly; user repositions freely.
5. **Placeholders tagged** — gray placeholder shapes get the same title tag as
   real images so position is preserved and they are replaced correctly on re-sync.
