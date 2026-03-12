// ── CONFIG ───────────────────────────────────────────────────
var TEMPLATE_ID = "DAHDupJXvrA";
var BASE = "https://api.canva.com/rest/v1";

// Row indices (0-based) in the xlsx for each cost line
// Adjust these if your sheet has different row positions
var ROWS = {
  cost1: 0,  // Allgemeine Betriebs- und Nebenkosten
  cost2: 1,  // Instandhaltungsrücklage & Wartungskosten
  cost3: 2,  // Betriebs- und Nebenkosten WEG
  cost4: 3,  // Kosten Co-Ownership-Struktur
  usage1: 4, // Verbrauchsabh. Betriebskosten (Strom/Wasser)
  usage2: 5, // Professionelle Reinigung
  usage3: 6, // Wäschepaket
  reserve1: 7, // Kassenbestand laufende Kosten
  reserve2: 8, // Rücklage Instandhaltung
};

// Column indices (0-based): C=2, D=3
var COL_MONTHLY = 2;  // Column C
var COL_ANNUAL  = 3;  // Column D

var ELEMENTS = {
  immoNameCover:     "PBxwy6ZcMmJm4Dx4-LBX2QHSGd0nnLdLw",
  locationLineCover: "PBxwy6ZcMmJm4Dx4-LBRYWyxDVHTh7z63",
  locationP3:        "PBlsh0ChHt5yLYQC-LBTyPQwmJjlkvrn6",
  monthlyCosts:      "PBlsh0ChHt5yLYQC-LBtz45nrBJ9B7nfq",
  annualCosts:       "PBlsh0ChHt5yLYQC-LBtJNnRjGwwyzKr0",
  totalMonthly:      "PBlsh0ChHt5yLYQC-LBj0kz1trZBpQSnc",
  totalAnnual:       "PBlsh0ChHt5yLYQC-LB0Kg054slD9nCSf",
  usageCosts:        "PBlsh0ChHt5yLYQC-LB3vZlwnMjsb4ldw",
  reserves:          "PBKdMl5XXwHkcJtj-LB7L2N91d7dB0C8J",
};

// ── STATE ────────────────────────────────────────────────────
var token = localStorage.getItem("canva_token") || "";
var propertyName = "";
var propertyLocation = "";
var sheetData = [];

// ── STEP 1: Token ────────────────────────────────────────────
function saveToken() {
  var val = document.getElementById("tokenInput").value.trim();
  if (!val) return;
  token = val;
  localStorage.setItem("canva_token", val);
  document.getElementById("tokenStatus").innerHTML = "<span class='success'>Token saved!</span>";
}

window.addEventListener("DOMContentLoaded", function() {
  if (token) {
    document.getElementById("tokenInput").value = token;
    document.getElementById("tokenStatus").innerHTML = "<span class='success'>Token loaded</span>";
  }
});

// ── STEP 2: Load property name from Canva ────────────────────
function loadPropertyName() {
  var raw = document.getElementById("canvaIdInput").value.trim();
  if (!raw) return;

  // Extract design ID from URL or use directly
  var designId = raw;
  var match = raw.match(/\/design\/(DA[a-zA-Z0-9_-]+)/);
  if (match) designId = match[1];

  document.getElementById("propertyStatus").innerHTML = "<span class='spinner'></span>Loading...";

  // Read text content from page 1 of the property design
  fetch(BASE + "/designs/" + designId + "/content?content_types=richtexts&pages=1", {
    headers: { "Authorization": "Bearer " + token }
  })
  .then(function(r) { return r.json(); })
  .then(function(j) {
    // Extract all text strings from richtexts
    var texts = [];
    if (j.richtexts) {
      j.richtexts.forEach(function(rt) {
        if (rt.regions) {
          rt.regions.forEach(function(region) {
            if (region.text && region.text.trim()) texts.push(region.text.trim());
          });
        }
      });
    }
    // First non-empty text = property name (e.g. "Nova Vista")
    // Find the location line — contains "|" character
    propertyName = texts[0] || designId;
    propertyLocation = texts.find(function(t) { return t.indexOf("|") !== -1; }) || "";

    document.getElementById("propertyStatus").innerHTML =
      "<span class='success'>Loaded: <strong>" + propertyName + "</strong>" +
      (propertyLocation ? " &mdash; " + propertyLocation : "") + "</span>";
    tryShowGenerate();
  })
  .catch(function(err) {
    document.getElementById("propertyStatus").innerHTML =
      "<span class='error'>Could not load: " + err.message + "</span>";
  });
}

// ── STEP 3: Upload xlsx ──────────────────────────────────────
function handleFile(e) {
  var file = e.target.files[0];
  if (!file) return;
  var reader = new FileReader();
  reader.onload = function(ev) {
    var wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    // Get as array of arrays (raw rows)
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    sheetData = rows;
    document.getElementById("preview").innerHTML =
      "<span class='success'>" + rows.length + " rows loaded from sheet.</span>";
    tryShowGenerate();
  };
  reader.readAsArrayBuffer(file);
}

// ── Show Step 4 when both are ready ─────────────────────────
function tryShowGenerate() {
  if (!propertyName || !sheetData.length) return;

  var m1 = getVal(ROWS.cost1, COL_MONTHLY);
  var m2 = getVal(ROWS.cost2, COL_MONTHLY);
  var m3 = getVal(ROWS.cost3, COL_MONTHLY);
  var m4 = getVal(ROWS.cost4, COL_MONTHLY);
  var totalM = m1 + m2 + m3 + m4;

  document.getElementById("summaryBox").innerHTML =
    "<strong>Property:</strong> " + propertyName + "<br>" +
    "<strong>Cost 1 (monthly):</strong> " + fmt(m1) + " / " + fmt(getVal(ROWS.cost1, COL_ANNUAL)) + " p.a.<br>" +
    "<strong>Cost 2 (monthly):</strong> " + fmt(m2) + " / " + fmt(getVal(ROWS.cost2, COL_ANNUAL)) + " p.a.<br>" +
    "<strong>Cost 3 (monthly):</strong> " + fmt(m3) + " / " + fmt(getVal(ROWS.cost3, COL_ANNUAL)) + " p.a.<br>" +
    "<strong>Cost 4 (monthly):</strong> " + fmt(m4) + " / " + fmt(getVal(ROWS.cost4, COL_ANNUAL)) + " p.a.<br>" +
    "<strong>Total:</strong> " + fmt(totalM) + " / " + fmt(totalM * 12) + " p.a.<br>" +
    "<strong>Usage 1:</strong> " + fmt(getVal(ROWS.usage1, COL_MONTHLY)) + "<br>" +
    "<strong>Usage 2:</strong> " + fmt(getVal(ROWS.usage2, COL_MONTHLY)) + "<br>" +
    "<strong>Usage 3:</strong> " + fmt(getVal(ROWS.usage3, COL_MONTHLY)) + "<br>" +
    "<strong>Reserve (operating):</strong> " + fmt(getVal(ROWS.reserve1, COL_MONTHLY)) + "<br>" +
    "<strong>Reserve (maintenance):</strong> " + fmt(getVal(ROWS.reserve2, COL_MONTHLY));

  document.getElementById("step-generate").style.display = "block";
}

// ── STEP 4: Generate ─────────────────────────────────────────
function generate() {
  if (!token) { alert("Please enter your Canva API token first."); return; }

  var btn = document.getElementById("generateBtn");
  btn.disabled = true;
  btn.innerHTML = "<span class='spinner'></span>Creating design...";

  var m1 = getVal(ROWS.cost1, COL_MONTHLY);
  var m2 = getVal(ROWS.cost2, COL_MONTHLY);
  var m3 = getVal(ROWS.cost3, COL_MONTHLY);
  var m4 = getVal(ROWS.cost4, COL_MONTHLY);
  var a1 = getVal(ROWS.cost1, COL_ANNUAL);
  var a2 = getVal(ROWS.cost2, COL_ANNUAL);
  var a3 = getVal(ROWS.cost3, COL_ANNUAL);
  var a4 = getVal(ROWS.cost4, COL_ANNUAL);
  var totalM = m1 + m2 + m3 + m4;
  var totalA = a1 + a2 + a3 + a4;

  var ops = [
    op(ELEMENTS.immoNameCover,     propertyName),
    op(ELEMENTS.locationLineCover, propertyLocation || propertyName),
    op(ELEMENTS.locationP3,        propertyLocation || propertyName),
    op(ELEMENTS.monthlyCosts,     fmt(m1)+"\n\n"+fmt(m2)+"\n\n"+fmt(m3)+"\n\n"+fmt(m4)),
    op(ELEMENTS.annualCosts,      fmt(a1)+"\n\n"+fmt(a2)+"\n\n"+fmt(a3)+"\n\n"+fmt(a4)),
    op(ELEMENTS.totalMonthly,     fmt(totalM)),
    op(ELEMENTS.totalAnnual,      fmt(totalA)),
    op(ELEMENTS.usageCosts,
      fmt(getVal(ROWS.usage1, COL_MONTHLY)) + "\n\n" +
      fmt(getVal(ROWS.usage2, COL_MONTHLY)) + "\n\n" +
      fmt(getVal(ROWS.usage3, COL_MONTHLY))),
    op(ELEMENTS.reserves,
      fmtN(getVal(ROWS.reserve1, COL_MONTHLY)) + " EUR\n\n" +
      fmtN(getVal(ROWS.reserve2, COL_MONTHLY)) + " EUR"),
  ];

  dupDesign(TEMPLATE_ID, propertyName + " - Wirtschaftsplan")
    .then(function(newId) {
      return startTx(newId).then(function(txId) {
        return applyOps(newId, txId, ops).then(function() {
          return commitTx(newId, txId).then(function() {
            var url = "https://www.canva.com/design/" + newId + "/edit";
            document.getElementById("result").innerHTML =
              "<p class='success' style='margin-bottom:8px'>Design created for <strong>" + propertyName + "</strong>!</p>" +
              "<a class='result-link' href='" + url + "' target='_blank'>Open in Canva</a>";
            btn.innerHTML = "Done";
          });
        });
      });
    })
    .catch(function(err) {
      document.getElementById("result").innerHTML = "<span class='error'>" + err.message + "</span>";
      btn.disabled = false;
      btn.innerHTML = "Retry";
    });
}

// ── CANVA API ─────────────────────────────────────────────────
function H() { return { "Authorization": "Bearer " + token, "Content-Type": "application/json" }; }

function dupDesign(id, title) {
  return fetch(BASE + "/designs/" + id + "/copies", {
    method: "POST", headers: H(), body: JSON.stringify({ title: title })
  }).then(function(r) { return r.json(); })
    .then(function(j) {
      if (!j.design || !j.design.id) throw new Error("Duplicate failed: " + JSON.stringify(j));
      return j.design.id;
    });
}
function startTx(id) {
  return fetch(BASE + "/designs/" + id + "/editing_sessions", {
    method: "POST", headers: H()
  }).then(function(r) { return r.json(); })
    .then(function(j) {
      if (!j.editing_session || !j.editing_session.id) throw new Error("Session failed");
      return j.editing_session.id;
    });
}
function applyOps(id, txId, operations) {
  return fetch(BASE + "/designs/" + id + "/editing_sessions/" + txId + "/operations", {
    method: "POST", headers: H(), body: JSON.stringify({ operations: operations })
  }).then(function(r) {
    if (!r.ok) return r.text().then(function(t) { throw new Error("Apply failed: " + t); });
  });
}
function commitTx(id, txId) {
  return fetch(BASE + "/designs/" + id + "/editing_sessions/" + txId + "/commit", {
    method: "POST", headers: H()
  }).then(function(r) {
    if (!r.ok) return r.text().then(function(t) { throw new Error("Commit failed: " + t); });
  });
}

// ── HELPERS ───────────────────────────────────────────────────
function getVal(rowIdx, colIdx) {
  if (!sheetData[rowIdx]) return 0;
  var val = sheetData[rowIdx][colIdx];
  return parseFloat(String(val).replace(/[^0-9.,]/g, "").replace(",", ".")) || 0;
}
function fmt(v)  { return (parseFloat(v)||0).toLocaleString("de-DE") + " EUR"; }
function fmtN(v) { return (parseFloat(v)||0).toLocaleString("de-DE"); }
function op(element_id, text) { return { type: "replace_text", element_id: element_id, text: String(text || "") }; }
