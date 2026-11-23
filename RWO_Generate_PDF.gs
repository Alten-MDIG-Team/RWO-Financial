
const CONFIG2 = {
  FOLDER_ID: "1W3BBBoUPZHkn67KgJMwwtpry7rnh6jHa", // ID du dossier Google Drive cible
  // FOLDER_ID: "17bv1_qge_Xsbc5D2jBpVmCV4ImPrGql8", // ID du dossier Google Drive cible
  START_ROW: 19,
  SHEET_RAPPORT: "Delivery Acceptance Note",
  SHEET_DATA: "FINANCIAL",
  CELL_TYPE: "H7",
  CELL_PERIOD: "J2"
};


function downloadPDFOneClick() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG2.SHEET_RAPPORT);
  
  // --- DÉTECTION AUTOMATIQUE DE LA ZONE ---
  // On part de A1 jusqu'à la dernière cellule contenant des données
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const rangeNotation = `A1:${sheet.getRange(lastRow, lastCol).getA1Notation()}`;
  
  // Coordonnées pour l'URL (0-indexed)
  // const r1 = 1; // Ligne 1
  // const c1 = 6; // Colonne A
  const r1 = 0; // Row 1
  const c1 = 0; // Column A
  const r2 = lastRow+5;
  const c2 = lastCol;

  // // --- PRÉPARATION DU NOM ---
  // const period = sheet.getRange(CONFIG.CELL_PERIOD).getValue();
  // const type = sheet.getRange(CONFIG.CELL_TYPE).getValue();
  let fileName = sheet.getRange("N4").getValue()+".pdf";
  // const fileName = `AI_NT_WPZINML_DAN_Preparation_${type}_${period}.pdf`;

  // --- URL D'EXPORT ---
  const urlParams = [
    'export?format=pdf',
    `gid=${sheet.getSheetId()}`,
    `r1=${r1}`, `c1=${c1}`, `r2=${r2}`, `c2=${c2}`,
    'size=A4',
    'portrait=false', // Mode Paysage
    'fitw=true',      // Ajuster à la largeur
    'gridlines=false',
    'printtitle=false',
    'sheetnames=false',
    'pagenumbers=false'
  ].join('&');

  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/${urlParams}`;
  const token = ScriptApp.getOAuthToken();
  
  try {
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    const htmlContent = `
      <html>
        <body onload="document.getElementById('dlink').click(); setTimeout(()=>google.script.host.close(), 2000)">
          <a id="dlink" href="data:application/pdf;base64,${base64}" download="${fileName}"></a>
          <div style="font-family:sans-serif; text-align:center; padding:10px; color: #444;">
            <b>Génération du PDF...</b><br>Zone détectée : A1 jusqu'à ${sheet.getRange(lastRow, lastCol).getA1Notation()}
          </div>
        </body>
      </html>`;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(350).setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export en cours');

  } catch (e) {
    SpreadsheetApp.getUi().alert("Erreur : " + e.toString());
  }
}
