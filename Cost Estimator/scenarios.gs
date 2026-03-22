function runCostSweep() {
  const ss    = SpreadsheetApp.getActive();
  const calcs = ss.getSheetByName('Calcs');
  const scen  = ss.getSheetByName('Scenarios');

  // Named ranges (each should be a single cell)
  const rngLivable = ss.getRangeByName('Livable');
  const rngBeds    = ss.getRangeByName('Beds');
  const rngBaths   = ss.getRangeByName('Baths');

  // Colors
  const COLOR_BG_INVALID      = '#bdbdbd'; // medium grey
  const COLOR_BG_UNLIKELY     = '#e0e0e0'; // light grey
  const COLOR_BG_QUESTIONABLE = '#fff9c4'; // light yellow
  const COLOR_BG_NORMAL       = '#ffffff'; // white

  const COLOR_FONT_DARKGREY   = '#424242';
  const COLOR_FONT_BLACK      = '#000000';

  // SF values in A4:A (down to last SF)
  const lastRow = scen.getLastRow();

  // Beds in B1:L1, Baths in B2:L2
  const bedRow  = scen.getRange('B1:M1').getValues()[0];
  const bathRow = scen.getRange('B2:M2').getValues()[0];
  const numScenarios = bedRow.length;

  // SF column (A4:A...)
  const sfValues = scen.getRange(4, 1, lastRow - 3, 1).getValues();

  // Clear old scenario values & formatting in B4:...
  if (lastRow >= 4) {
    scen.getRange(4, 2, lastRow - 3, numScenarios).clearContent().clearFormat();
  }

  // ---- Classification logic ----
  function classifyScenario(sf, beds, baths) {
    // If either beds or baths is blank, don't try to classify.
    // We are letting the Calcs sheet auto-assume in those cases.
    if (beds === '' || beds === null || baths === '' || baths === null) {
      return 'VALID';
    }

    sf    = Number(sf);
    beds  = Number(beds);
    baths = Number(baths);

    if (isNaN(sf) || isNaN(beds) || isNaN(baths)) return 'INVALID';

    // Hard impossible cases
    if (sf < 250) return 'INVALID';
    if (beds === 0 && baths > 1) return 'INVALID';
    if (beds > 0 && baths > beds + 1) return 'INVALID';

    // Minimum SF by bedroom count
    let minStrict, minPlausible;

    switch (beds) {
      case 0:
        minStrict    = 250;
        minPlausible = 300;
        break;
      case 1:
        minStrict    = 350;
        minPlausible = 400;
        break;
      case 2:
        minStrict    = 600;
        minPlausible = 650;
        break;
      case 3:
        minStrict    = 900;
        minPlausible = 950;
        break;
      case 4:
        minStrict    = 1050;
        minPlausible = 1100;
        break;
      default:
        minStrict    = 1100;
        minPlausible = 1150;
        break;
    }

    if (sf < minStrict)    return 'INVALID';
    if (sf < minPlausible) return 'UNLIKELY';

    // Too many baths for small units
    if (sf < 650 && baths >= 2) return 'UNLIKELY';

    // Oversized studios / 1-bed
    if (beds === 0 && sf > 800)   return 'QUESTIONABLE';
    if (beds === 1 && sf > 1150)  return 'QUESTIONABLE';

    // Large units with only 1 bath
    if (sf > 900 && baths === 1) return 'QUESTIONABLE';

    return 'VALID';
  }

  // ---- Main loop ----
  for (let r = 0; r < sfValues.length; r++) {
    const sf = sfValues[r][0];
    if (!sf) continue;

    for (let c = 0; c < numScenarios; c++) {
      const bedsRaw  = bedRow[c];
      const bathsRaw = bathRow[c];

      const classification = classifyScenario(sf, bedsRaw, bathsRaw);
      const targetCell = scen.getRange(4 + r, 2 + c); // row 4+r, col B(2)+c

      if (classification === 'INVALID') {
        targetCell
          .clearContent()
          .setBackground(COLOR_BG_INVALID)
          .setFontColor(COLOR_FONT_DARKGREY);
        continue;
      }

      // ---- Push values into Calcs (named ranges) ----
      rngLivable.setValue(sf);

      // If scenario cell is blank, clear the named range to let Calcs auto-assume.
      if (bedsRaw === '' || bedsRaw === null) {
        rngBeds.clearContent();
      } else {
        rngBeds.setValue(bedsRaw);
      }

      if (bathsRaw === '' || bathsRaw === null) {
        rngBaths.clearContent();
      } else {
        rngBaths.setValue(bathsRaw);
      }

      SpreadsheetApp.flush();

      const cost = calcs.getRange('K3').getValue(); // or a named range if you set one
      targetCell.setValue(cost);

      // ---- Formatting based on classification ----
      if (classification === 'UNLIKELY') {
        targetCell
          .setBackground(COLOR_BG_UNLIKELY)
          .setFontColor(COLOR_FONT_DARKGREY);
      } else if (classification === 'QUESTIONABLE') {
        targetCell
          .setBackground(COLOR_BG_QUESTIONABLE)
          .setFontColor(COLOR_FONT_BLACK);
      } else {
        // VALID (including auto-assumed cases)
        targetCell
          .setBackground(COLOR_BG_NORMAL)
          .setFontColor(COLOR_FONT_BLACK);
      }
    }
  }
}
