Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    document.getElementById("create-table").onclick = createTable;
  }
});

async function createTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values =
    [["Nome", "Sexo"]];

  expensesTable.rows.add(null, [
      ["José", "Masculino"]
  ]);

  expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
  expensesTable.getRange().format.autofitColumns();
  expensesTable.getRange().format.autofitRows();

    await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      } 
  });
}