// Overnight Tider backend: one Google Sheet runs the website and the owner dashboard.
//
// What it does
// - Website orders, wholesale quotes, text-list and email-list signups land in tabs of this Sheet.
// - Website chat: customers message from the site, the owner replies from the dashboard (Chats tab).
// - The owner dashboard (/admin) reads orders, prices them, moves them along, and publishes
//   stock, prices, the special, text and photos to the website. Everything is PIN protected.
//
// Setup (about 5 minutes)
// 1. Create a Google Sheet. Extensions > Apps Script. Paste this whole file in.
// 2. Put the owner's email in OWNER_EMAIL and a 4-digit PIN in FIRST_PIN below.
// 3. Run the function "setup" once (Run button at the top) and allow the permissions.
// 4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Deploy.
// 5. Copy the web app URL into SHOP.endpoint in index.html and push.
//    In the dashboard: Settings > Connection, paste the same URL.
//
// The owner can also work straight in the Sheet. Status values: New, Price sent, Confirmed, Paid, Delivered, Cancelled.

const OWNER_EMAIL = "";
const FIRST_PIN = "";        // e.g. "4821". Used once by setup(). Change it later from the dashboard
const PHOTO_FOLDER = "Overnight Tider website photos";

const TABS = {
  order: { name: "Orders", text: "B:L", text2: "P:S", head: ["Received", "Order #", "Name", "Phone", "Items", "How", "Date needed", "ZIP", "Notes", "Status", "Code", "Lines", "Fee", "Discount %", "Total", "Paid by", "History", "Business", "Channel"] },
  quote: { name: "Wholesale quotes", text: "B:I", head: ["Received", "Name", "Business", "Buying for", "Phone", "Need", "How often", "First delivery", "Status"] },
  alert: { name: "Text list", text: "B:B", head: ["Signed up", "Phone"] },
  email: { name: "Email list", text: "B:C", head: ["Signed up", "Email", "Code given"] },
  chat: { name: "Chats", text: "B:H", head: ["Time", "Chat", "From", "Name", "Phone", "Message", "Read", "Msg id"] }
};
const STATUS = { new: "New", sent: "Price sent", confirmed: "Confirmed", paid: "Paid", done: "Delivered", cancelled: "Cancelled" };
const STATUS_BACK = Object.keys(STATUS).reduce((m, k) => (m[STATUS[k].toLowerCase()] = k, m), { priced: "sent", done: "done" });

function setup() {
  if (FIRST_PIN && /^\d{4}$/.test(FIRST_PIN)) PropertiesService.getScriptProperties().setProperty("ADMIN_PIN", FIRST_PIN);
  Object.keys(TABS).forEach(k => textCols_(tab(k), TABS[k]));
  site_();
}
// Keep ZIPs, phones and dates exactly as typed (no 08104 -> 8104, no "Nov 20" -> a date)
function textCols_(sh, t) { [t.text, t.text2].forEach(a => a && sh.getRange(a).setNumberFormat("@")); }

function clean(v) {
  v = String(v == null ? "" : v).slice(0, 5000);
  return v === "" ? v : "'" + v; // store as typed: no formulas, no 08104 -> 8104, no "Nov 20" -> a date
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function out(t) { return ContentService.createTextOutput(t); }
function tab(type) {
  const t = TABS[type], ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
  if (sh.getLastRow() === 0) {
    textCols_(sh, t);
    sh.appendRow(t.head); sh.setFrozenRows(1); sh.getRange(1, 1, 1, t.head.length).setFontWeight("bold");
  } else if (sh.getLastColumn() < t.head.length) {
    sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight("bold"); // older sheet, add new columns
  }
  return sh;
}
function site_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Site") || ss.insertSheet("Site");
  if (!sh.getRange("A1").getValue()) sh.getRange("A1").setValue(JSON.stringify({ v: 1 }));
  sh.getRange("B1").setValue("Website settings written by the dashboard. Don't edit by hand.");
  return sh;
}

/* ---------- public: the website reads its published settings ---------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  if (action === "chat") {
    const id = String((e.parameter && e.parameter.id) || "");
    if (!/^[a-z0-9]{12,40}$/i.test(id)) return json([]);
    return json(chatMsgs_(id).map(m => ({ id: m.id, from: m.from, text: m.text, at: m.at })));
  }
  if (action === "site") {
    const cache = CacheService.getScriptCache(), hit = cache.get("site");
    if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    const txt = String(site_().getRange("A1").getValue() || "{}");
    try { cache.put("site", txt, 60); } catch (err) {}
    return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
  }
  return out("ok");
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out("bad request"); }
  if (d.type === "admin") return admin(d);
  if (d.type === "chat") return chatIn_(d);
  const t = TABS[d.type];
  if (!t) return out("unknown type");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = tab(d.type);
    let row;
    if (d.type === "order") {
      const lines = Array.isArray(d.lines) ? d.lines.slice(0, 40).map(l => ({ id: String(l.id).slice(0, 20), unit: String(l.unit).slice(0, 30), qty: +l.qty || 1 })) : [];
      row = [d.id, d.name, d.phone, d.items, d.how, d.when, d.zip, d.notes, "New", d.code || "", JSON.stringify(lines), 0, d.code ? 10 : 0, "", "", JSON.stringify([{ s: "new", at: new Date().toISOString() }]), "", "retail"];
    }
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

  if (OWNER_EMAIL && (d.type === "order" || d.type === "quote")) {
    const subject = d.type === "order" ? "New order " + d.id + " from " + (d.name || "a customer") : "Wholesale quote request from " + (d.biz || d.name);
    const skip = ["type", "page", "lines"];
    const body = Object.keys(d).filter(k => skip.indexOf(k) === -1).map(k => k + ": " + d[k]).join("\n");
    MailApp.sendEmail(OWNER_EMAIL, subject, body + "\n\nOpen the dashboard to price it.");
  }
  return out("ok");
}

/* ---------- owner dashboard API (PIN protected) ---------- */
function admin(d) {
  const props = PropertiesService.getScriptProperties(), cache = CacheService.getScriptCache();
  const fails = +(cache.get("fails") || 0);
  if (fails >= 10) return json({ error: "Locked. Too many wrong PINs, try again in 15 minutes." });
  const pin = props.getProperty("ADMIN_PIN");
  if (!pin) return json({ error: "No PIN set yet. Put one in FIRST_PIN and run setup." });
  if (String(d.pin) !== pin) { cache.put("fails", String(fails + 1), 900); Utilities.sleep(600); return json({ error: "Wrong PIN" }); }
  cache.remove("fails");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (d.op === "chats") return json({ chats: readChats_() });
    if (d.op === "chatReply") {
      const id = String(d.chat || ""), text = String(d.text || "").trim().slice(0, 2000);
      if (!/^[a-z0-9]{12,40}$/i.test(id) || !text) return json({ error: "Empty message" });
      const m = { id: String(d.id || "o" + Date.now()), chat: id, from: "o", name: "", phone: "", text: text, at: new Date().toISOString() };
      tab("chat").appendRow(chatRow_(m)); chatCache_(id, m);
      return json({ ok: true, msg: m });
    }
    if (d.op === "chatRead") {
      const sh = tab("chat"), n = sh.getLastRow() - 1;
      if (n > 0) { const v = sh.getRange(2, 2, n, 6).getValues(); v.forEach((r, i) => { if (String(r[0]).replace(/^'/, "") === String(d.chat) && r[1] === "Customer" && r[5] !== "yes") sh.getRange(i + 2, 7).setValue("yes"); }); }
      return json({ ok: true });
    }
    if (d.op === "load") return json({ chats: readChats_(), orders: readOrders(), quotes: readQuotes(), emails: readList("email"), texts: readList("alert"), site: JSON.parse(String(site_().getRange("A1").getValue() || "{}")) });
    if (d.op === "saveOrder") { writeOrder(d.order); return json({ ok: true }); }
    if (d.op === "saveQuote") {
      const sh = tab("quote"), r = +String(d.quote.id).replace(/\D/g, "");
      if (r > 1 && r <= sh.getLastRow()) sh.getRange(r, 9).setValue(QSTATUS[d.quote.status] || "New");
      return json({ ok: true });
    }
    if (d.op === "saveSite") {
      const txt = JSON.stringify(d.site || { v: 1 });
      if (txt.length > 48000) return json({ error: "Too much in one save. Use uploaded photos, not pasted ones." });
      site_().getRange("A1").setValue(txt); cache.remove("site");
      return json({ ok: true });
    }
    if (d.op === "upload") {
      const mime = /^image\/(webp|jpeg|png)$/.test(d.mime) ? d.mime : "image/jpeg";
      const bytes = Utilities.base64Decode(String(d.data || ""));
      if (bytes.length > 4 * 1024 * 1024) return json({ error: "Photo too large" });
      const it = DriveApp.getFoldersByName(PHOTO_FOLDER), folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
      const f = folder.createFile(Utilities.newBlob(bytes, mime, String(d.name || "photo").replace(/[^\w.-]/g, "-")));
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return json({ url: "https://lh3.googleusercontent.com/d/" + f.getId() });
    }
    if (d.op === "replaceAll") { // Settings: clear the Sheet, or load sample orders for a demo
      const D = d.data || {};
      const put = (k, list) => {
        const sh = tab(k), n = sh.getLastRow() - 1;
        if (n > 0) sh.getRange(2, 1, n, TABS[k].head.length).clearContent();
        if (list.length) sh.getRange(2, 1, list.length, list[0].length).setValues(list);
      };
      put("order", (D.orders || []).slice(0, 3000).map(orderRow_));
      put("quote", (D.quotes || []).slice(0, 500).map(q => [new Date(q.at)].concat([q.name, q.biz, q.who, q.phone, q.need, q.freq, q.when, QSTATUS[q.status] || "New"].map(clean))));
      put("email", (D.emails || []).slice(0, 5000).map(x => [new Date(x.at), clean(x.email), clean(x.code)]));
      put("alert", (D.texts || []).slice(0, 5000).map(x => [new Date(x.at), clean(x.phone)]));
      const oldChat = rows("chat").map(r => "chat:" + String(r[1]).replace(/^'/, ""));
      put("chat", (D.chats || []).slice(0, 5000).map(chatRow_));
      CacheService.getScriptCache().removeAll(oldChat.concat((D.chats || []).map(m => "chat:" + m.chat)).filter((k, i, a) => a.indexOf(k) === i));
      return json({ ok: true });
    }
    if (d.op === "setPin") {
      if (!/^\d{4}$/.test(String(d.newPin))) return json({ error: "Use 4 numbers" });
      props.setProperty("ADMIN_PIN", String(d.newPin)); return json({ ok: true });
    }
    return json({ error: "Unknown request" });
  } finally {
    lock.releaseLock();
  }
}

function rows(type) { const sh = tab(type), n = sh.getLastRow() - 1; return n > 0 ? sh.getRange(2, 1, n, TABS[type].head.length).getValues() : []; }
function iso(v) { return v instanceof Date ? v.toISOString() : (v ? new Date(v).toISOString() : new Date().toISOString()); }
function parse(v, d) { try { return v ? JSON.parse(v) : d; } catch (e) { return d; } }
function readOrders() {
  return rows("order").map(r => ({
    at: iso(r[0]), id: String(r[1]).replace(/^'/, ""), name: String(r[2]), phone: String(r[3]).replace(/^'/, ""), items: String(r[4]), how: String(r[5]), when: String(r[6]), zip: String(r[7]), notes: String(r[8]),
    status: STATUS_BACK[String(r[9]).toLowerCase()] || "new", code: String(r[10] || ""), lines: parse(r[11], []), fee: +r[12] || 0, disc: +r[13] || 0,
    paid: String(r[15] || ""), log: parse(r[16], null) || [{ s: "new", at: iso(r[0]) }], biz: String(r[17] || ""), channel: String(r[18] || "retail")
  })).filter(o => o.id);
}
function readQuotes() {
  return rows("quote").map((r, i) => ({ id: "Q" + (i + 2), at: iso(r[0]), name: String(r[1]), biz: String(r[2]), who: String(r[3]), phone: String(r[4]).replace(/^'/, ""), need: String(r[5]), freq: String(r[6]), when: String(r[7]),
    status: ({ "quote sent": "sent", won: "done", "didn't work out": "lost" })[String(r[8]).toLowerCase()] || "new" }));
}
function readList(type) { return rows(type).map(r => type === "email" ? { at: iso(r[0]), email: String(r[1]), code: String(r[2] || "") } : { at: iso(r[0]), phone: String(r[1]).replace(/^'/, "") }); }
const QSTATUS = { new: "New", sent: "Quote sent", done: "Won", lost: "Didn't work out" };
function orderVals_(o) {
  const lines = (o.lines || []).map(l => ({ id: l.id, unit: l.unit, qty: +l.qty || 1, price: l.price == null ? null : +l.price }));
  const sub = lines.reduce((t, l) => t + (l.price || 0) * l.qty, 0), total = Math.max(0, sub - sub * (+o.disc || 0) / 100 + (+o.fee || 0));
  return [STATUS[o.status] || "New", o.code || "", JSON.stringify(lines), +o.fee || 0, +o.disc || 0, Math.round(total * 100) / 100, o.paid || "", JSON.stringify(o.log || []), o.biz || "", o.channel || "retail"].map(v => typeof v === "number" ? v : clean(v));
}
function orderRow_(o) { return [new Date(o.at || Date.now())].concat([o.id, o.name, o.phone, o.items || "", o.how, o.when, o.zip, o.notes].map(clean), orderVals_(o)); }
function writeOrder(o) {
  const sh = tab("order"), n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, 2, n, 1).getValues().flat().map(x => String(x).replace(/^'/, "")) : [];
  const r = ids.indexOf(String(o.id));
  if (r === -1) sh.appendRow(orderRow_(o));
  else sh.getRange(r + 2, 10, 1, 10).setValues([orderVals_(o)]);
}

/* ---------- website chat ---------- */
function chatRow_(m) { return [new Date(m.at || Date.now())].concat([m.chat, m.from === "o" ? "You" : "Customer", m.name || "", m.phone || "", m.text, m.read ? "yes" : "", m.id || ""].map(clean)); }
function chatParse_(r) { return { at: iso(r[0]), chat: String(r[1]).replace(/^'/, ""), from: r[2] === "You" ? "o" : "c", name: String(r[3] || ""), phone: String(r[4] || "").replace(/^'/, ""), text: String(r[5] || "").replace(/^'/, ""), read: r[6] === "yes", id: String(r[7] || "") }; }
function readChats_() { const since = Date.now() - 120 * 864e5; return rows("chat").map(chatParse_).filter(m => m.chat && new Date(m.at) > since); }
function chatMsgs_(id) {
  const c = CacheService.getScriptCache(), hit = c.get("chat:" + id);
  if (hit) try { return JSON.parse(hit); } catch (e) {}
  const list = rows("chat").map(chatParse_).filter(m => m.chat === id).map(m => ({ id: m.id, from: m.from, text: m.text, at: m.at }));
  try { c.put("chat:" + id, JSON.stringify(list.slice(-80)), 21600); } catch (e) {}
  return list;
}
function chatCache_(id, m) {
  const c = CacheService.getScriptCache(), list = chatMsgs_(id);
  if (!list.some(x => x.id === m.id)) list.push({ id: m.id, from: m.from, text: m.text, at: m.at });
  try { c.put("chat:" + id, JSON.stringify(list.slice(-80)), 21600); } catch (e) {}
}
function chatIn_(d) {
  const id = String(d.chat || ""), text = String(d.text || "").trim().slice(0, 1000);
  if (!/^[a-z0-9]{12,40}$/i.test(id) || !text) return out("bad chat");
  const c = CacheService.getScriptCache(), k = "chatn:" + id, n = +(c.get(k) || 0);
  if (n > 40) return out("slow down");
  c.put(k, String(n + 1), 600);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  let first = false;
  try {
    first = !chatMsgs_(id).some(m => m.from === "c");
    const m = { id: String(d.id || "c" + Date.now()).slice(0, 40), chat: id, from: "c", name: String(d.name || "").slice(0, 60), phone: String(d.phone || "").slice(0, 30), text: text, at: new Date().toISOString() };
    tab("chat").appendRow(chatRow_(m)); chatCache_(id, m);
  } finally { lock.releaseLock(); }
  const quiet = !c.get("chatmail:" + id);
  if (OWNER_EMAIL && (first || quiet)) {
    c.put("chatmail:" + id, "1", 7200);
    MailApp.sendEmail(OWNER_EMAIL, "Website chat from " + (d.name || "a visitor"), text + "\n\n" + (d.phone ? "Mobile: " + d.phone + "\n" : "") + "Reply from the dashboard, under Messages.");
  }
  return out("ok");
}
