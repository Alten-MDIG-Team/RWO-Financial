// function onOpen() {
//   const ui = SpreadsheetApp.getUi();
  
//   // Crée un nouveau menu "⚙️ My Tools"
//   ui.createMenu('⚙️ My Tools')
//       .addItem('Refresh Dashboard', 'combineDashboards')
//       .addItem('Refresh Dashboard API', 'combineDashboardsAPI')
//       .addItem('Generate PDF', 'downloadPDFOneClick')
//       .addToUi();
// }


function onOpen() {  
  const ui = SpreadsheetApp.getUi();    
  ui.createMenu('⚙️ My Tools')      
      // .addItem('Refresh Dashboard', 'combineDashboards')
      // .addItem('Refresh Dashboard API', 'combineDashboardsAPI')
      .addItem('Refresh Dashboard API', 'combineDashboardsAPI2') 
      .addItem('Generate PDF', 'downloadPDFOneClick')   
      .addToUi();
}
