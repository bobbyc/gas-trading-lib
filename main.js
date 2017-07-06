/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
var _Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var _Dashboard = _Spreadsheet.getSheetByName("Dashboard");
var _Taiwan50 = _Spreadsheet.getSheetByName("0050");

function onOpen() {
  var entries = [];
  entries.push({ name: "Provisioning", functionName: "OnProvisionDashboard" });
  entries.push({ name: "Provisioning 0050", functionName: "OnProvision0050" });
  entries.push({ name: "Provisioning FUND", functionName: "OnProvisionFund" });
  entries.push({ name: "Provisioning Symbol", functionName: "OnProvisionActiveCell" });
  _Spreadsheet.addMenu("[Scripts]", entries);

  // reset sheet index
  //var template = spreadsheet.getSheetByName("Template");
  //spreadsheet.setActiveSheet(template)
  //spreadsheet.moveActiveSheet(1);

  _Spreadsheet.setActiveSheet(_Dashbaord)
  _Spreadsheet.moveActiveSheet(1);

};

function OnProvisionDashboard() {
  resetSheet("Dashboard");
  provisionSheets("Dashboard");
};

function OnProvisionFund() {
  resetSheet("FUND");
  provisionSheets("FUND");
};

function OnProvision0050() {
  resetSheet("0050");
  provisionSheets("0050");
};

function OnProvisionActiveCell() {
  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var Cell = Sheet.getActiveCell();
  var Symbol = Sheet.getRange(Cell.getRow(), 1, 1, 1);

  provisionEx(Symbol.getValue());
}
