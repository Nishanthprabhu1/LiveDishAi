// ── CHEF NISHANTH — MENU SHEET API ──────────────────────────────────────────
// Paste this entire file into Google Apps Script (script.google.com)
// Link it to your Google Sheet, then Deploy → New deployment → Web App
// Execute as: Me  |  Who has access: Anyone
// Copy the Web App URL and paste into your HTML (see SHEET_API_URL in the HTML)
// ─────────────────────────────────────────────────────────────────────────────

const SHEET_NAME = "Menu"; // Change if your sheet tab has a different name

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows  = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => h.toString().trim().toLowerCase());

  const menuData = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    // Skip rows where name is empty or available = FALSE
    const name = getCellStr(row, headers, "name");
    if (!name) continue;

    const available = getCellStr(row, headers, "available").toLowerCase();
    if (available === "false" || available === "no" || available === "0") continue;

    // ── Core fields ──────────────────────────────────────────────────────────
    const item = {
      cat:        getCellStr(row, headers, "category"),
      name:       name,
      type:       getCellStr(row, headers, "type") || "veg",      // veg / nonveg
      price:      getNum(row, headers, "price"),
      bestseller: getBool(row, headers, "bestseller"),
      emoji:      getCellStr(row, headers, "emoji") || "🍽️",
      liked:      getNum(row, headers, "liked"),
      desc:       getCellStr(row, headers, "desc"),
      upsells:    parseList(getCellStr(row, headers, "upsells")),  // comma-separated names
      addons:     parseAddons(getCellStr(row, headers, "addons")), // "Name:price,Name:price"
      groups:     parseGroups(getCellStr(row, headers, "groups")), // JSON or shorthand
    };

    // ── AR / 3D fields (optional) ─────────────────────────────────────────────
    const folder = getCellStr(row, headers, "folder");
    if (folder) {
      item.folder = folder;
      item.file   = getCellStr(row, headers, "file");
      item.scale  = getNum(row, headers, "scale") || 1.0;
      item.orbit  = getCellStr(row, headers, "orbit") || "0deg 75deg 105%";
    }

    menuData.push(item);
  }

  const output = ContentService
    .createTextOutput(JSON.stringify(menuData))
    .setMimeType(ContentService.MimeType.JSON);

  // Allow the HTML page (any origin) to fetch this
  return output;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function getCellStr(row, headers, col) {
  const idx = headers.indexOf(col);
  if (idx === -1) return "";
  const v = row[idx];
  return v == null ? "" : v.toString().trim();
}

function getNum(row, headers, col) {
  const v = getCellStr(row, headers, col);
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function getBool(row, headers, col) {
  const v = getCellStr(row, headers, col).toLowerCase();
  return v === "true" || v === "yes" || v === "1";
}

// "Mango Lassi, Butter Naan" → ["Mango Lassi", "Butter Naan"]
function parseList(str) {
  if (!str) return [];
  return str.split(",").map(s => s.trim()).filter(Boolean);
}

// "Extra Sauce:30, Double Portion:180, Add Fries:90"
// → [{n:"Extra Sauce", e:30}, ...]
function parseAddons(str) {
  if (!str) return [];
  return str.split(",").map(part => {
    const [n, e] = part.split(":").map(s => s.trim());
    return { n: n || "", e: parseFloat(e) || 0 };
  }).filter(a => a.n);
}

// Supports two formats:
//
// Simple shorthand (staff-friendly):
//   "Choose your heat|required|radio|Mild:0,Medium:0,Hot:0,Extra Hot:10"
//   Multiple groups separated by "||"
//   e.g. "Choose heat|required|radio|Mild:0,Hot:0||Portion|required|radio|Half:0,Full:180"
//
// Raw JSON (power users):
//   [{"title":"Choose heat","sub":"required","type":"radio","opts":[{"n":"Mild","e":0}]}]
//
function parseGroups(str) {
  if (!str) return [];

  // Try JSON first
  if (str.startsWith("[")) {
    try { return JSON.parse(str); } catch(e) {}
  }

  // Shorthand: multiple groups split by "||"
  return str.split("||").map(groupStr => {
    const parts = groupStr.split("|").map(s => s.trim());
    // parts[0]=title, parts[1]=sub, parts[2]=type, parts[3]=opts
    const title = parts[0] || "";
    const sub   = parts[1] || "required";
    const type  = parts[2] || "radio";
    const opts  = (parts[3] || "").split(",").map(opt => {
      const [n, e] = opt.split(":").map(s => s.trim());
      return { n: n || "", e: parseFloat(e) || 0 };
    }).filter(o => o.n);
    return { title, sub, type, opts };
  }).filter(g => g.title);
}
