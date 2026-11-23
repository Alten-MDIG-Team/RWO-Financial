/**
 * Consolidates row data from two separate Google Spreadsheet sources 
 * ('RWO Dashboard' tabs) into a single master 'DashBoard' tab in this workbook.
 * Preserves both real HYPERLINK formulas and manually-applied rich-text links,
 * then forces proper blue/underline styling so all links render correctly.
 */
function combineDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = ss.getSheetByName("DashBoard");
  const id1 = '1hvrnM1uHMAdHqb0YONetIOtl0_PWUxWJCmVs20GJO2s';
  const id2 = '1teV6A-8Z1pnfp5eWdPKSknoGpOMGKMxn8p9EFcdi7hc';

  const sheet1 = SpreadsheetApp.openById(id1).getSheetByName("RWO Dashboard");
  const sheet2 = SpreadsheetApp.openById(id2).getSheetByName("RWO Dashboard");

  // Pulls every row, keeping links (whether formula-based or manually-inserted) intact
  function getData(sheet) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const formulas = range.getFormulas();
    const richText = range.getRichTextValues();

    return values.map((row, r) => row.map((cell, c) => {
      if (formulas[r][c]) return formulas[r][c]; // real formula, keep it
      const url = richText[r][c] ? richText[r][c].getLinkUrl() : null;
      if (url) return `=HYPERLINK("${url}","${cell}")`; // manual link, rebuild it
      return cell; // plain value
    }));
  }

  const data1 = getData(sheet1);
  const data2 = getData(sheet2);

  // Combine: header from sheet1, all rows from both (skip sheet2's header)
  const combined = data2.concat(data1.slice(1));

  // Wipe everything and dump the combined data
  destSheet.clear();
  const range = destSheet.getRange(1, 1, combined.length, combined[0].length);
  range.setValues(combined);

  // 1. Make every HYPERLINK cell look/act like a link FIRST
  const isLink = (f) => typeof f === "string" && f.toUpperCase().startsWith("=HYPERLINK");
  range.setFontColors(combined.map(row => row.map(c => isLink(c) ? "#1155cc" : null)));
  range.setFontLines(combined.map(row => row.map(c => isLink(c) ? "underline" : null)));

  // 2. THEN copy header formatting across different spreadsheets (so it doesn't get overwritten)
  const headerRange1 = sheet1.getRange(1, 1, 1, data1[0].length);
  const destHeaderRange = destSheet.getRange(1, 1, 1, data1[0].length);

  destHeaderRange.setBackgrounds(headerRange1.getBackgrounds());
  destHeaderRange.setFontColors(headerRange1.getFontColors());
  destHeaderRange.setFontWeights(headerRange1.getFontWeights());
  destHeaderRange.setFontFamilies(headerRange1.getFontFamilies());
  destHeaderRange.setFontSizes(headerRange1.getFontSizes());
  destHeaderRange.setHorizontalAlignments(headerRange1.getHorizontalAlignments());
  destHeaderRange.setVerticalAlignments(headerRange1.getVerticalAlignments());
}


/**
 * ✅ FIXED VERSION — avoids "Service Spreadsheets timed out" error
 *
 * Root cause: Sheets.Spreadsheets.Values.get() (Advanced Sheets API) can
 * time out on large ranges, especially with valueRenderOption: "FORMULA".
 *
 * Fix:
 *  1. Read using the native SpreadsheetApp service (getFormulas()) instead
 *     of the Advanced API — much more reliable for large ranges.
 *  2. Wrap the read in a retry-with-backoff helper as a safety net.
 */
function combineDashboardsAPI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destSheetName = "DashBoard"; // Change if different

  const id1 = '1hvrnM1uHMAdHqb0YONetIOtl0_PWUxWJCmVs20GJO2s';
  const id2 = '1teV6A-8Z1pnfp5eWdPKSknoGpOMGKMxn8p9EFcdi7hc';

  try {
    console.log("═".repeat(60));
    console.log("🚀 FIXED VERSION — Native Reads + Retry");
    console.log("═".repeat(60));

    // ========== STEP 1: OPEN SOURCE SHEETS ==========
    console.log("\n📍 STEP 1: Opening source sheets...");
    const sheet1 = SpreadsheetApp.openById(id1).getSheetByName("RWO Dashboard");
    if (!sheet1) throw new Error("Could not find 'RWO Dashboard' in Source 1");

    const sheet2 = SpreadsheetApp.openById(id2).getSheetByName("RWO Dashboard");
    if (!sheet2) throw new Error("Could not find 'RWO Dashboard' in Source 2");

    // getLastRow()/getLastColumn() can be inflated by leftover formatting,
    // deleted-but-not-trimmed cells, merges, etc. far below the real data —
    // that phantom range is what was causing both the blank rows past 445
    // AND the repeated timeouts (we were moving ~143k mostly-empty cells).
    // Find the TRUE last row by scanning column A (Unique Id — always
    // populated for real rows) instead of trusting the reported range.
    function getRealLastRow(sheet, label) {
      const reportedLastRow = sheet.getLastRow();
      if (reportedLastRow === 0) return 0;
      const colA = sheet.getRange(1, 1, reportedLastRow, 1).getValues();
      for (let r = colA.length - 1; r >= 0; r--) {
        const cellStr = colA[r][0] === null || colA[r][0] === undefined ? "" : String(colA[r][0]).trim();
        if (cellStr !== "") {
          const realLastRow = r + 1;
          if (realLastRow < reportedLastRow) {
            console.log("  ⚠️  " + label + ": getLastRow() reported " + reportedLastRow +
              " but real data ends at row " + realLastRow + " — trimming phantom range " +
              "(likely stray whitespace/invisible chars further down)");
          }
          return realLastRow;
        }
      }
      return 0;
    }

    const realLastRow1 = getRealLastRow(sheet1, "Source 1");
    const realLastRow2 = getRealLastRow(sheet2, "Source 2");

    console.log("  ✓ Source 1: " + realLastRow1 + " real rows × " + sheet1.getLastColumn() + " cols");
    console.log("  ✓ Source 2: " + realLastRow2 + " real rows × " + sheet2.getLastColumn() + " cols");

    const maxCols = Math.max(sheet1.getLastColumn(), sheet2.getLastColumn());
    console.log("  Using max: " + maxCols + " cols for all rows");

    // ========== STEP 2: FETCH DATA (with retry) ==========
    console.log("\n📥 STEP 2: Fetching data (native SpreadsheetApp, with retry)...");

    // Mimics valueRenderOption: "FORMULA" — use the formula text if a cell
    // has one, otherwise fall back to the cell's plain value. getFormulas()
    // ALONE returns "" for any non-formula cell, which is why plain typed
    // data (everything except things like HYPERLINK cells) was disappearing.
    function fetchFormulasWithRetry(sheet, realLastRow, label, maxRetries) {
      maxRetries = maxRetries || 4;
      for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
          const lastCol = sheet.getLastColumn();
          const range = sheet.getRange(1, 1, realLastRow, lastCol);
          const formulas = range.getFormulas();
          const values = range.getValues();

          const merged = formulas.map((row, r) =>
            row.map((cell, c) => (cell !== "" ? cell : values[r][c]))
          );

          console.log("  ✓ " + label + ": " + merged.length + " rows fetched (attempt " + attempt + ")");
          return merged;
        } catch (e) {
          console.warn("  ⚠️  " + label + " attempt " + attempt + " failed: " + e.toString());
          if (attempt === maxRetries) throw e;
          const waitMs = 1000 * Math.pow(2, attempt); // 2s, 4s, 8s, 16s...
          console.log("     Retrying in " + (waitMs / 1000) + "s...");
          Utilities.sleep(waitMs);
        }
      }
    }

    const data1 = fetchFormulasWithRetry(sheet1, realLastRow1, "Source 1");
    if (!data1 || data1.length === 0) throw new Error("Source 1 returned no data");

    Utilities.sleep(500);

    const data2 = fetchFormulasWithRetry(sheet2, realLastRow2, "Source 2");
    if (!data2 || data2.length === 0) throw new Error("Source 2 returned no data");

    // ========== STEP 3: PAD ROWS TO SAME LENGTH ==========
    console.log("\n🔧 STEP 3: Padding rows to match column count...");

    function padRows(dataArray, targetCols) {
      return dataArray.map(row => {
        const r = row.slice(0, targetCols);
        while (r.length < targetCols) r.push("");
        return r;
      });
    }

    const data1Padded = padRows(data1, maxCols);
    const data2Padded = padRows(data2, maxCols);
    console.log("  ✓ All rows padded to " + maxCols + " columns");

    // ========== STEP 4: COMBINE DATA ==========
    console.log("\n🔄 STEP 4: Combining data...");
    const data2NoHeader = data2Padded.slice(1);
    const combined = data1Padded.concat(data2NoHeader);
    const totalRows = combined.length;
    const totalCols = maxCols;

    console.log("  ✓ Source 1: " + data1Padded.length + " rows");
    console.log("  ✓ Source 2: " + data2NoHeader.length + " rows (minus header)");
    console.log("  ✓ TOTAL: " + totalRows + " rows × " + totalCols + " columns");

    // ========== STEP 5: GET DESTINATION SHEET ==========
    console.log("\n📋 STEP 5: Getting destination sheet...");
    const destSheet = ss.getSheetByName(destSheetName);
    if (!destSheet) throw new Error("Could not find sheet: " + destSheetName);
    console.log("  ✓ Destination: " + destSheet.getName());

    // Generic retry wrapper for any operation on this heavy, multi-tab file
    function withRetry(fn, label, maxRetries) {
      maxRetries = maxRetries || 4;
      for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
          const result = fn();
          if (attempt > 1) console.log("  ✓ " + label + " succeeded on attempt " + attempt);
          return result;
        } catch (e) {
          console.warn("  ⚠️  " + label + " attempt " + attempt + " failed: " + e.toString());
          if (attempt === maxRetries) throw e;
          const waitMs = 1000 * Math.pow(2, attempt);
          console.log("     Retrying " + label + " in " + (waitMs / 1000) + "s...");
          Utilities.sleep(waitMs);
        }
      }
    }

    // ========== STEP 6: CLEAR DESTINATION ==========
    console.log("\n🗑️  STEP 6: Clearing destination sheet...");
    withRetry(() => destSheet.clear(), "clear destination");
    Utilities.sleep(1000);
    console.log("  ✓ Cleared");

    // ========== STEP 7: WRITE DATA IN CHUNKS ==========
    console.log("\n📤 STEP 7: Writing data...");
    const CHUNK_SIZE = 400; // smaller chunks — lighter per call on this large file
    let chunkNum = 0;

    for (let i = 0; i < combined.length; i += CHUNK_SIZE) {
      const chunk = combined.slice(i, i + CHUNK_SIZE);
      const startRow = i + 1;
      const endRow = Math.min(i + CHUNK_SIZE, combined.length);
      chunkNum++;

      console.log("  [Chunk " + chunkNum + "] Writing rows " + startRow + "-" + endRow + "...");
      withRetry(
        () => destSheet.getRange(startRow, 1, chunk.length, totalCols).setValues(chunk),
        "write chunk " + chunkNum
      );
      console.log("    ✓ Chunk written");
      Utilities.sleep(500);
    }

    // ========== STEP 8: SUCCESS ==========
    console.log("\n" + "═".repeat(60));
    console.log("✅ SUCCESS! ALL DATA WRITTEN!");
    console.log("═".repeat(60));

    try {
      SpreadsheetApp.getUi().alert(
        "✅ SUCCESS!\n\n" +
        "Source 1: " + data1Padded.length + " rows\n" +
        "Source 2: " + data2NoHeader.length + " rows\n\n" +
        "🎯 TOTAL: " + totalRows + " rows × " + totalCols + " columns"
      );
    } catch (alertError) {
      // Data was already written successfully at this point — a UI timeout
      // showing the popup on this large file is cosmetic, not a data failure.
      console.warn("  ⚠️  Success alert failed to display (data was already written OK): " + alertError.toString());
    }

  } catch (error) {
    console.error("\n" + "❌".repeat(30));
    console.error("ERROR:", error.toString());
    console.error("Stack:", error.stack);
    console.error("❌".repeat(30));
    SpreadsheetApp.getUi().alert("❌ ERROR:\n\n" + error.toString());
  }
}
