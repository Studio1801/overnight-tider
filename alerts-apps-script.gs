// Weekly text list collector for Overnight Tider.
// 1. Create a Google Sheet, Extensions > Apps Script, paste this in.
// 2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
// 3. Paste the web app URL into SHOP.alertsEndpoint in index.html.
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(["Phone", "Signed up", "Source"]);
  let data = {};
  try { data = JSON.parse(e.postData.contents); } catch (err) {}
  const phone = String(data.phone || "").replace(/[^\d+]/g, "");
  if (phone.length >= 10) {
    const existing = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues().flat().map(String);
    if (existing.indexOf(phone) === -1) sheet.appendRow([phone, new Date(), "website"]);
  }
  return ContentService.createTextOutput("ok");
}
