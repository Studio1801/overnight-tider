// Overnight Tider: orders, wholesale quotes and text-list signups, all into one Google Sheet.
//
// Setup (about 5 minutes):
// 1. Create a Google Sheet. Extensions > Apps Script. Paste this whole file in.
// 2. Put the owner's email in OWNER_EMAIL below (new orders and quotes get emailed there).
// 3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Authorize.
// 4. Copy the web app URL into SHOP.endpoint in index.html and push.
//
// Tabs "Orders", "Wholesale quotes", "Text list" and "Email list" are created automatically on the first submission.
// The owner works orders by changing the Status column (New, Priced, Confirmed, Done).

const OWNER_EMAIL = "";

const TABS = {
  order: { name: "Orders", head: ["Received", "Order #", "Name", "Phone", "Items", "How", "Date needed", "ZIP", "Notes", "Status"] },
  quote: { name: "Wholesale quotes", head: ["Received", "Name", "Business", "Buying for", "Phone", "Need", "How often", "First delivery", "Status"] },
  alert: { name: "Text list", head: ["Signed up", "Phone"] },
  email: { name: "Email list", head: ["Signed up", "Email", "Code given"] }
};

function clean(v) {
  v = String(v == null ? "" : v).slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // stop spreadsheet formula injection
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out("bad request"); }
  const t = TABS[d.type];
  if (!t) return out("unknown type");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(t.head);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, t.head.length).setFontWeight("bold");
    }
    let row;
    if (d.type === "order") row = [d.id, d.name, d.phone, d.items, d.how, d.when, d.zip, d.notes, "New"];
    if (d.type === "quote") row = [d.name, d.biz, d.who, d.phone, d.need, d.freq, d.when, "New"];
    if (d.type === "alert") {
      const phone = String(d.phone || "").replace(/[^\d+]/g, "");
      if (phone.replace(/\D/g, "").length < 10) return out("bad phone");
      const n = sh.getLastRow() - 1;
      const have = n > 0 ? sh.getRange(2, 2, n, 1).getValues().flat().map(x => String(x).replace(/^'/, "")) : [];
      if (have.indexOf(phone) > -1) return out("ok");
      row = [phone];
    }
    if (d.type === "email") {
      const email = String(d.email || "").trim().toLowerCase();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return out("bad email");
      const n = sh.getLastRow() - 1;
      const have = n > 0 ? sh.getRange(2, 2, n, 1).getValues().flat().map(x => String(x).toLowerCase()) : [];
      if (have.indexOf(email) > -1) return out("ok");
      row = [email, d.code];
    }
    sh.appendRow([new Date()].concat(row.map(clean)));
  } finally {
    lock.releaseLock();
  }

  if (OWNER_EMAIL && d.type !== "alert" && d.type !== "email") {
    const subject = d.type === "order"
      ? "New order " + d.id + " from " + (d.name || "a customer")
      : "Wholesale quote request from " + (d.biz || d.name);
    const skip = ["type", "page"];
    const body = Object.keys(d).filter(k => skip.indexOf(k) === -1).map(k => k + ": " + d[k]).join("\n");
    MailApp.sendEmail(OWNER_EMAIL, subject, body + "\n\nEverything is also in the Google Sheet.");
  }
  return out("ok");
}

function out(t) { return ContentService.createTextOutput(t); }
