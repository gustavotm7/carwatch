/** @OnlyCurrentDoc  (the script can only touch the spreadsheet it lives in) */
/**
 * CarWatch group: ads shared from sites CarWatch can't read (OLX, Mercado Livre, Webmotors, Facebook Marketplace...).
 * Lives in a Google Sheet (Extensions > Apps Script) that stores the ads.
 *
 * Setup (once, see apps-script/README.md):
 *   1. Run setup() from the editor. It prepares the "Ads" tab and logs the group code.
 *   2. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *   3. In CarWatch: Menu > Group, paste the web app address (ends in /exec) and the code.
 *
 * GET  ?country=BR             -> {ok, ads: [...]}  active ads of that store, shared or updated in the last DAYS days
 * POST {code, action: "add", ad: {...}}   adds an ad, or updates it when the same link was shared before
 * POST {code, action: "gone", id}         marks an ad as sold or removed
 * POST {code, action: "ping"}             checks the code
 * Only public listing data is stored: link, site, model, year, km, price, version, gearbox, body, town, title.
 */
var SHEET = "Ads";
var HEAD = ["id", "added", "updated", "country", "site", "url", "name", "family", "year", "km", "price", "version", "gear",
            "body", "city", "uf", "lat", "lon", "title", "status"];
var DAYS = 60;            // an ad nobody shares or updates for this long drops out of the list
var MAX_ROWS = 5000;      // a full sheet refuses new ads instead of slowing down

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET)) {                // a new, empty spreadsheet: its first tab becomes "Ads"
    var first = ss.getSheets()[0];
    var sh = ss.getSheets().length === 1 && first.getLastRow() === 0 ? first.setName(SHEET) : ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    sh.setFrozenRows(1);
  }
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("GROUP_CODE")) {
    props.setProperty("GROUP_CODE", Utilities.getUuid().replace(/-/g, "").slice(0, 8).toUpperCase());
  }
  Logger.log("Group code: " + props.getProperty("GROUP_CODE"));
}

/** Run this to give the group a new code: old invites stop being able to add ads. */
function newCode() {
  PropertiesService.getScriptProperties().deleteProperty("GROUP_CODE");
  setup();
}

function doGet(e) {
  var country = String((e && e.parameter && e.parameter.country) || "").toUpperCase();
  var since = Date.now() - DAYS * 864e5;
  var ads = rows_(sheet_()).filter(function (a) {
    return a.id && a.status !== "gone" && (!country || a.country === country) && Date.parse(a.updated) >= since;
  }).map(function (a) {
    delete a._row;
    return a;
  });
  return out_({ok: true, ads: ads});
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var code = PropertiesService.getScriptProperties().getProperty("GROUP_CODE");
    if (!code || String(b.code || "").trim().toUpperCase() !== code) return out_({ok: false, error: "code"});
    if (b.action === "ping") return out_({ok: true});
    lock.waitLock(10000);
    locked = true;
    var sh = sheet_(), rows = rows_(sh), now = new Date().toISOString();
    if (b.action === "gone") {
      var hit = find_(rows, String(b.id || ""));
      if (!hit) return out_({ok: false, error: "missing"});
      sh.getRange(hit._row, HEAD.indexOf("updated") + 1, 1, 1).setValue(cell_(now));
      sh.getRange(hit._row, HEAD.indexOf("status") + 1, 1, 1).setValue("gone");
      return out_({ok: true, id: hit.id});
    }
    if (b.action !== "add") return out_({ok: false, error: "action"});
    var ad = clean_(b.ad || {});
    if (!ad) return out_({ok: false, error: "invalid"});
    var old = find_(rows, ad.id);
    if (!old && rows.length >= MAX_ROWS) return out_({ok: false, error: "full"});
    ad.added = old ? old.added : now;
    ad.updated = now;
    ad.status = "active";
    var row = HEAD.map(function (h) { return cell_(ad[h] == null ? "" : ad[h]); });
    if (old) sh.getRange(old._row, 1, 1, HEAD.length).setValues([row]);
    else sh.appendRow(row);
    return out_({ok: true, id: ad.id, updated: !!old});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  } finally {
    if (locked) lock.releaseLock();
  }
}

function sheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET);
  if (!sh) throw new Error("Run setup() first");
  return sh;
}

function rows_(sh) {
  var v = sh.getDataRange().getValues();
  var head = v.shift() || [];
  return v.map(function (r, i) {
    var o = {_row: i + 2};
    head.forEach(function (h, j) {
      var x = r[j];
      o[h] = x instanceof Date ? x.toISOString() : typeof x === "string" && x.charAt(0) === "'" ? x.slice(1) : x;
    });
    return o;
  });
}

function find_(rows, id) {
  for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i];
  return null;
}

/** Text that could start a formula, and dates, are stored as plain text (a leading apostrophe). */
function cell_(v) {
  if (typeof v !== "string") return v;
  return /^[=+\-@']/.test(v) || /^\d{4}-\d\d-\d\d/.test(v) ? "'" + v : v;
}

function txt_(v, n) {
  return String(v == null ? "" : v).replace(/[\u0000-\u001f\u007f]/g, " ").trim().slice(0, n);
}

/** The ad as the app sends it, checked and trimmed; null when something essential is wrong. */
function clean_(a) {
  var url = txt_(a.url, 600);
  if (!/^https?:\/\/[^\s"'<>]+$/i.test(url)) return null;
  var id = String(a.id || "");
  if (!/^sh[a-z0-9]{4,14}$/.test(id)) return null;
  var country = String(a.country || "").toUpperCase();
  if (country !== "BR" && country !== "CA") return null;
  var year = Math.round(+a.year), km = Math.round(+a.km || 0), price = Math.round(+a.price);
  if (!(year >= 1950 && year <= 2100) || !(price >= 500 && price <= 1e8) || !(km >= 0 && km <= 2e6)) return null;
  var name = txt_(a.name, 40);
  if (!name) return null;
  var lat = a.lat === "" || a.lat == null ? "" : +a.lat, lon = a.lon === "" || a.lon == null ? "" : +a.lon;
  if (lat === "" || lon === "" || !(Math.abs(lat) <= 90 && Math.abs(lon) <= 180)) { lat = ""; lon = ""; }
  return {
    id: id, country: country, site: txt_(a.site, 40), url: url, name: name, family: txt_(a.family, 20),
    year: year, km: km || "", price: price, version: txt_(a.version, 30), gear: a.gear === "a" ? "a" : "m",
    body: txt_(a.body, 10), city: txt_(a.city, 60), uf: txt_(a.uf, 4), lat: lat, lon: lon, title: txt_(a.title, 200)
  };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
