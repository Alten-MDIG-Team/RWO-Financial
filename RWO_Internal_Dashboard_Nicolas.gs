
function combineDashboardsAPI2() {
  const ID_SOURCE_1 = '1hvrnM1uHMAdHqb0YONetIOtl0_PWUxWJCmVs20GJO2s';
  const ID_SOURCE_2 = '1teV6A-8Z1pnfp5eWdPKSknoGpOMGKMxn8p9EFcdi7hc';
  const TAB_SOURCE = "RWO Dashboard";
  const TAB_DEST = "DashBoard";

  // 1. Lecture API instantanée (1 seul appel par fichier)
  const data1 = fetchSheetAPI(ID_SOURCE_1, TAB_SOURCE);
  const data2 = fetchSheetAPI(ID_SOURCE_2, TAB_SOURCE);

  if (!data1.length && !data2.length) return;

  // 2. Fusion : Source 1 (avec en-tête) + Source 2 (sans en-tête)
  const combined = [...data2, ...data1.slice(1)];

  // 3. Normalisation des longueurs de tableaux (padding)
  const maxCols = Math.max(...combined.map(row => row.length));
  const normalizedData = combined.map(row => {
    while (row.length < maxCols) row.push("");
    return row;
  });

  // 4. Nettoyage de l'onglet cible
  const destSs = SpreadsheetApp.getActiveSpreadsheet();
  Sheets.Spreadsheets.Values.clear({}, destSs.getId(), "DashBoard!A:ZZ");  

  // 5. Écriture globale en 1 appel API
  const destRange = `'${TAB_DEST}'!A1:${getColLetter(maxCols)}${normalizedData.length}`;
  
  Sheets.Spreadsheets.Values.update(
    { values: normalizedData },
    destSs.getId(),
    destRange,
    { valueInputOption: 'USER_ENTERED' } // Active les formules (=HYPERLINK, etc.)
  );

  destSs.toast(`${normalizedData.length} lignes écrites en ~2 secondes.`, "✅ Consolidation terminée !", 5);
}

/**
 * Extrait les formules et valeurs brutes en un seul appel via Google Sheets API.
 */
function fetchSheetAPI(spreadsheetId, sheetName) {
  const response = Sheets.Spreadsheets.Values.get(spreadsheetId, `'${sheetName}'!A1:ZZ`, {
    valueRenderOption: 'FORMULA'
  });
  return response.values || [];
}

/**
 * Convertit un nombre de colonnes en lettres A1 (ex: 28 -> AB).
 */
function getColLetter(colIndex) {
  let temp, letter = '';
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = Math.floor((colIndex - temp) / 26);
  }
  return letter;
}
