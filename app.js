// ── CONFIG ───────────────────────────────────────────────────
var TEMPLATE_ID = "DAHDupJXvrA";
var BASE = "https://api.canva.com/rest/v1";

// Row indices (0-based): Excel row 15 = index 14
var ROWS = {
  cost1:    14, // Row 15 — Allgemeine Betriebs- und Nebenkosten
  cost2:    15, // Row 16 — Instandhaltungsrücklage
  cost3:    16, // Row 17 — Betriebs- und Nebenkosten WEG
  cost4:    17, // Row 18 — Kosten Co-Ownership
  usage1:   24, // Row 25 — Verbrauchsabh. (Strom/Wasser)
  usage2:   25, // Row 26 — Professionelle Reinigung
  usage3:   26, // Row 27 — Wäschepaket
  reserve1: 38, // Row 39 — Kassenbestand laufende Kosten
  reserve2: 39  // Row 40 — Rücklage Instandhaltung
};

var COL_MONTHLY  = 2; // Column C (index 2)
var COL_ANNUAL   = 3; // Column D (index 3)
var COL_USAGE    = 2; // Column C (index 2) — usage costs in column C
var COL_RESERVE  = 3; // Column D (index 3) — reserves in column D

var ELEMENTS = {
  immoNameCover:     "PBxwy6ZcMmJm4Dx4-LBX2QHSGd0nnLdLw",
  locationLineCover: "PBxwy6ZcMmJm4Dx4-LBRYWyxDVHTh7z63",
  locationP3:        "PBlsh0ChHt5yLYQC-LBTyPQwmJjlkvrn6",
  monthlyCosts:      "PBlsh0ChHt5yLYQC-LBtz45nrBJ9B7nfq",
  annualCosts:       "PBlsh0ChHt5yLYQC-LBtJNnRjGwwyzKr0",
  totalMonthly:      "PBlsh0ChHt5yLYQC-LBj0kz1trZBpQSnc",
  totalAnnual:       "PBlsh0ChHt5yLYQC-LB0Kg054slD9nCSf",
  usageCosts:        "PBlsh0ChHt5yLYQC-LB3vZlwnMjsb4ldw",
  reserves:          "PBKdMl5XXwHkcJtj-LB7L2N91d7dB0C8J"
};

// ── STATE ────────────────────────────────────────────────────
var token        = localStorage.getItem("canva_token") || "";
var propertyName = "";
var propertyLoc  = "";
var sheetData    = [];

// ── ON LOAD ──────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", function() {
  if (token) {
    document.getElementById("tokenInput").value = token;
    document.getElementById("tokenStatus").innerHTML = "<span class='success'>Token loaded</span>";
  }
});

// ── STEP 1: Save token ───────────────────────────────────────
function saveToken() {
  var val = document.getElementById("tokenInput").value.trim();
  if (!val) { alert("Please paste your token first."); return; }
  token = val;
  localStorage.setItem("canva_token", val);
  document.getElementById("tokenStatus").innerHTML = "<span class='success'>Token saved!</span>";
}

// ── STEP 2: Save property name ───────────────────────────────
function savePropertyName() {
  var name = document.getElementById("propertyNameInput").value.trim();
  var loc  = document.getElementById("propertyLocationInput").value.trim();
  if (!name) { alert("Please enter a property name."); return; }
  propertyName = name;
  propertyLoc  = loc;
  document.getElementById("propertyStatus").innerHTML =
    "<span class='success'>Confirmed: <strong>" + name + "</strong>" +
    (loc ? " &mdash; " + loc : "") + "</span>";
  tryShowGenerate();
}

// ── STEP 3: Upload xlsx ──────────────────────────────────────
function handleFile(e) {
  var file = e.target.files[0];
  if (!file) return;
  var reader = new FileReader();
  reader.onload = function(ev) {
    var wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
    var sheetName = wb.SheetNames.find(function(n) {
      return n.trim() === "M&S Output";
    }) || wb.SheetNames[0];
    var ws = wb.Sheets[sheetName];
    sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    document.getElementById("preview").innerHTML =
      "<span class='success'>" + sheetData.length + " rows loaded from tab: <strong>" + sheetName + "</strong></span>";
    tryShowGenerate();
  };
  reader.readAsArrayBuffer(file);
}

// ── Show Step 4 when both name + sheet are ready ─────────────
function tryShowGenerate() {
  if (!propertyName || !sheetData.length) return;

  var m1 = getVal(ROWS.cost1,    COL_MONTHLY);
  var m2 = getVal(ROWS.cost2,    COL_MONTHLY);
  var m3 = getVal(ROWS.cost3,    COL_MONTHLY);
  var m4 = getVal(ROWS.cost4,    COL_MONTHLY);
  var a1 = getVal(ROWS.cost1,    COL_ANNUAL);
  var a2 = getVal(ROWS.cost2,    COL_ANNUAL);
  var a3 = getVal(ROWS.cost3,    COL_ANNUAL);
  var a4 = getVal(ROWS.cost4,    COL_ANNUAL);
  var u1 = getVal(ROWS.usage1,   COL_USAGE);
  var u2 = getVal(ROWS.usage2,   COL_USAGE);
  var u3 = getVal(ROWS.usage3,   COL_USAGE);
  var r1 = getVal(ROWS.reserve1, COL_RESERVE);
  var r2 = getVal(ROWS.reserve2, COL_RESERVE);

  document.getElementById("summaryBox").innerHTML =
    "<strong>Property:</strong> " + propertyName + "<br>" +
    "<strong>Location:</strong> " + (propertyLoc || "—") + "<br><br>" +
    "<strong>Cost 1:</strong> " + fmt(m1) + " / " + fmt(a1) + " p.a.<br>" +
    "<strong>Cost 2:</strong> " + fmt(m2) + " / " + fmt(a2) + " p.a.<br>" +
    "<strong>Cost 3:</strong> " + fmt(m3) + " / " + fmt(a3) + " p.a.<br>" +
    "<strong>Cost 4:</strong> " + fmt(m4) + " / " + fmt(a4) + " p.a.<br>" +
    "<strong>Total:</strong> " + fmt(m1+m2+m3+m4) + " / " + fmt(a1+a2+a3+a4) + " p.a.<br><br>" +
    "<strong>Usage 1:</strong> " + fmt(u1) + "<br>" +
    "<strong>Usage 2:</strong> " + fmt(u2) + "<br>" +
    "<strong>Usage 3:</strong> " + fmt(u3) + "<br><br>" +
    "<strong>Reserve (operating):</strong> "   + fmt(r1) + "<br>" +
    "<strong>Reserve (maintenance):</strong> " + fmt(r2);

  document.getElementById("step-generate").style.display = "block";
  document.getElementById("generateBtn").disabled = false;
  document.getElementById("generateBtn").innerHTML = "&#9654; Create Canva Design";
  document.getElementById("result").innerHTML = "";
}

// ── STEP 4: Generate ─────────────────────────────────────────
function generate() {
  if (!token)        { alert("Please save your Canva API token first."); return; }
  if (!propertyName) { alert("Please confirm the property name first."); return; }
  if (!sheetData.length) { alert("Please upload the cost sheet first."); return; }

  var btn = document.getElementById("generateBtn");
  btn.disabled = true;
  btn.innerHTML = "<span class='spinner'></span>Creating design...";

  var m1 = getVal(ROWS.cost1,    COL_MONTHLY);
  var m2 = getVal(ROWS.cost2,    COL_MONTHLY);
  var m3 = getVal(ROWS.cost3,    COL_MONTHLY);
  var m4 = getVal(ROWS.cost4,    COL_MONTHLY);
  var a1 = getVal(ROWS.cost1,    COL_ANNUAL);
  var a2 = getVal(ROWS.cost2,    COL_ANNUAL);
  var a3 = getVal(ROWS.cost3,    COL_ANNUAL);
  var a4 = getVal(ROWS.cost4,    COL_ANNUAL);
  var u1 = getVal(ROWS.usage1,   COL_USAGE);
  var u2 = getVal(ROWS.usage2,   COL_USAGE);
  var u3 = getVal(ROWS.usage3,   COL_USAGE);
  var r1 = getVal(ROWS.reserve1, COL_RESERVE);
  var r2 = getVal(ROWS.reserve2, COL_RESERVE);

  var ops = [
    op(ELEMENTS.immoNameCover,     propertyName),
    op(ELEMENTS.locationLineCover, propertyLoc || propertyName),
    op(ELEMENTS.locationP3,        propertyLoc || propertyName),
    op(ELEMENTS.monthlyCosts,      fmt(m1)+"\n\n"+fmt(m2)+"\n\n"+fmt(m3)+"\n\n"+fmt(m4)),
    op(ELEMENTS.annualCosts,       fmt(a1)+"\n\n"+fmt(a2)+"\n\n"+fmt(a3)+"\n\n"+fmt(a4)),
    op(ELEMENTS.totalMonthly,      fmt(m1+m2+m3+m4)),
    op(ELEMENTS.totalAnnual,       fmt(a1+a2+a3+a4)),
    op(ELEMENTS.usageCosts,        fmt(u1)+"\n\n"+fmt(u2)+"\n\n"+fmt(u3)),
    op(ELEMENTS.reserves,          fmtN(r1)+" EUR\n\n"+fmtN(r2)+" EUR")
  ];

  dupDesign(TEMPLATE_ID, propertyName + " - Wirtschaftsplan")
    .then(function(newId) {
      return startTx(newId).then(function(txId) {
        return applyOps(newId, txId, ops).then(function() {
          return commitTx(newId, txId).then(function() {
            var url = "https://www.canva.com/design/" + newId + "/edit";
            document.getElementById("result").innerHTML =
              "<p class='success' style='margin-bottom:10px'>Design created for <strong>" + propertyName + "</strong>!</p>" +
              "<a class='result-link' href='" + url + "' target='_blank'>&#127912; Open in Canva</a>";
            btn.innerHTML = "Done!";
          });
        });
      });
    })
    .catch(function(err) {
      document.getElementById("result").innerHTML =
        "<span class='error'>Error: " + err.message + "</span>";
      btn.disabled = false;
      btn.innerHTML = "&#9654; Retry";
    });
}

// ── CANVA API ─────────────────────────────────────────────────
function H() {
  return { "Authorization": "Bearer " + token, "Content-Type": "application/json" };
}
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
      if (!j.editing_session || !j.editing_session.id) throw new Error("Session failed: " + JSON.stringify(j));
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
  var raw = sheetData[rowIdx][colIdx];
  if (raw === undefined || raw === null || raw === "") return 0;
  var cleaned = String(raw).replace(/[^0-9.,]/g, "").replace(",", ".");
  return parseFloat(cleaned) || 0;
}
function fmt(v)  { return (parseFloat(v) || 0).toLocaleString("de-DE") + " EUR"; }
function fmtN(v) { return (parseFloat(v) || 0).toLocaleString("de-DE"); }
function op(element_id, text) {
  return { type: "replace_text", element_id: element_id, text: String(text || "") };
}
