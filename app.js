// CONFIG — update column names to match your sheet headers exactly
var TEMPLATE_ID = "DAHDupJXvrA";

var COL = {
  immoName:           "Immo Name",
  location:           "Location",
  locationLine:       "Location Line",
  cost1Monthly:       "Betriebskosten mtl",
  cost2Monthly:       "Instandhaltung mtl",
  cost3Monthly:       "WEG Kosten mtl",
  cost4Monthly:       "Objektgesellschaft mtl",
  usageCost1:         "Verbrauch pro Nacht",
  usageCost2:         "Reinigung pro Aufenthalt",
  usageCost3:         "Waesche pro Aufenthalt",
  reserveOperating:   "Kassenbestand Betrieb",
  reserveMaintenance: "Kassenbestand Instandhaltung"
};

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

var token = localStorage.getItem("canva_token") || "";
var allRows = [];

function saveToken() {
  var val = document.getElementById("tokenInput").value.trim();
  if (!val) return;
  token = val;
  localStorage.setItem("canva_token", val);
  document.getElementById("tokenStatus").innerHTML = "<span class=\"success\">Token saved!</span>";
}

window.addEventListener("DOMContentLoaded", function() {
  if (token) {
    document.getElementById("tokenInput").value = token;
    document.getElementById("tokenStatus").innerHTML = "<span class=\"success\">Token loaded</span>";
  }
});

function handleFile(e) {
  var file = e.target.files[0];
  if (!file) return;
  var reader = new FileReader();
  reader.onload = function(ev) {
    var wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
    if (!rows.length) {
      document.getElementById("preview").innerHTML = "<span class=\"error\">No data found.</span>";
      return;
    }
    allRows = rows;
    document.getElementById("preview").innerHTML = "<span class=\"success\">" + rows.length + " row(s) loaded.</span>";
    if (rows.length === 1) { showSingle(rows[0]); } else { showTable(rows); }
  };
  reader.readAsArrayBuffer(file);
}

function showSingle(row) {
  document.getElementById("propertyName").textContent = row[COL.immoName] || "Unknown";
  document.getElementById("step-generate").style.display = "block";
  document.getElementById("step-table").style.display = "none";
  document.getElementById("step-generate").dataset.rowIndex = 0;
}

function showTable(rows) {
  document.getElementById("step-generate").style.display = "none";
  document.getElementById("step-table").style.display = "block";
  var html = "<table><thead><tr><th>#</th><th>Property</th><th>Location</th><th>Status</th><th></th></tr></thead><tbody>";
  for (var i = 0; i < rows.length; i++) {
    html += "<tr><td>" + (i+1) + "</td><td>" + (rows[i][COL.immoName]||"—") + "</td><td>" + (rows[i][COL.location]||"—") + "</td>";
    html += "<td><span class=\"status-badge status-pending\" id=\"status-" + i + "\">Pending</span></td>";
    html += "<td><button class=\"gen-btn\" id=\"btn-" + i + "\" onclick=\"generateRow(" + i + ")\">Generate</button></td></tr>";
  }
  html += "</tbody></table>";
  document.getElementById("tableContainer").innerHTML = html;
}

function generate() {
  var idx = parseInt(document.getElementById("step-generate").dataset.rowIndex || 0);
  run(allRows[idx], "result", "generateBtn", null);
}

function generateRow(idx) {
  run(allRows[idx], null, "btn-" + idx, "status-" + idx);
}

function run(row, resultId, btnId, statusId) {
  if (!token) { alert("Please enter your Canva API token first."); return; }
  var btn = document.getElementById(btnId);
  btn.disabled = true;
  btn.innerHTML = "<span class=\"spinner\"></span>Working...";
  if (statusId) setStatus(statusId, "loading", "Generating...");
  var m1 = num(row[COL.cost1Monthly]), m2 = num(row[COL.cost2Monthly]);
  var m3 = num(row[COL.cost3Monthly]), m4 = num(row[COL.cost4Monthly]);
  var totalM = m1 + m2 + m3 + m4;
  var ops = [
    op(ELEMENTS.immoNameCover,    row[COL.immoName]),
    op(ELEMENTS.locationLineCover, row[COL.locationLine] || row[COL.location]),
    op(ELEMENTS.locationP3,       row[COL.location]),
    op(ELEMENTS.monthlyCosts,     fmt(m1)+"\n\n"+fmt(m2)+"\n\n"+fmt(m3)+"\n\n"+fmt(m4)),
    op(ELEMENTS.annualCosts,      fmt(m1*12)+"\n\n"+fmt(m2*12)+"\n\n"+fmt(m3*12)+"\n\n"+fmt(m4*12)),
    op(ELEMENTS.totalMonthly,     fmt(totalM)),
    op(ELEMENTS.totalAnnual,      fmt(totalM*12)),
    op(ELEMENTS.usageCosts,       fmt(row[COL.usageCost1])+"\n\n"+fmt(row[COL.usageCost2])+"\n\n"+fmt(row[COL.usageCost3])),
    op(ELEMENTS.reserves,         fmtN(row[COL.reserveOperating])+" EUR\n\n"+fmtN(row[COL.reserveMaintenance])+" EUR")
  ];
  dupDesign(TEMPLATE_ID, row[COL.immoName] || "MYNE Design")
    .then(function(newId) {
      return startTx(newId).then(function(txId) {
        return applyOps(newId, txId, ops).then(function() {
          return commitTx(newId, txId).then(function() {
            var url = "https://www.canva.com/design/" + newId + "/edit";
            if (resultId) document.getElementById(resultId).innerHTML =
              "<p class=\"success\" style=\"margin-bottom:8px\">Design created!</p>" +
              "<a class=\"result-link\" href=\"" + url + "\" target=\"_blank\">Open in Canva</a>";
            if (statusId) setStatus(statusId, "done", "<a href=\"" + url + "\" target=\"_blank\" style=\"color:#155724\">Open in Canva</a>");
            btn.innerHTML = "Done";
          });
        });
      });
    })
    .catch(function(err) {
      console.error(err);
      if (statusId) setStatus(statusId, "error", "Error");
      if (resultId) document.getElementById(resultId).innerHTML = "<span class=\"error\">" + err.message + "</span>";
      btn.disabled = false;
      btn.innerHTML = "Retry";
    });
}

var BASE = "https://api.canva.com/rest/v1";
function H() { return { "Authorization": "Bearer " + token, "Content-Type": "application/json" }; }

function dupDesign(id, title) {
  return fetch(BASE + "/designs/" + id + "/copies", { method: "POST", headers: H(), body: JSON.stringify({ title: title }) })
    .then(function(r) { return r.json(); })
    .then(function(j) { if (!j.design || !j.design.id) throw new Error("Duplicate failed: " + JSON.stringify(j)); return j.design.id; });
}
function startTx(id) {
  return fetch(BASE + "/designs/" + id + "/editing_sessions", { method: "POST", headers: H() })
    .then(function(r) { return r.json(); })
    .then(function(j) { if (!j.editing_session || !j.editing_session.id) throw new Error("Session failed"); return j.editing_session.id; });
}
function applyOps(id, txId, operations) {
  return fetch(BASE + "/designs/" + id + "/editing_sessions/" + txId + "/operations", { method: "POST", headers: H(), body: JSON.stringify({ operations: operations }) })
    .then(function(r) { if (!r.ok) return r.text().then(function(t) { throw new Error("Apply failed: " + t); }); });
}
function commitTx(id, txId) {
  return fetch(BASE + "/designs/" + id + "/editing_sessions/" + txId + "/commit", { method: "POST", headers: H() })
    .then(function(r) { if (!r.ok) return r.text().then(function(t) { throw new Error("Commit failed: " + t); }); });
}

function num(v) { return parseFloat(v) || 0; }
function fmt(v) { return num(v).toLocaleString("de-DE") + " EUR"; }
function fmtN(v) { return num(v).toLocaleString("de-DE"); }
function op(element_id, text) { return { type: "replace_text", element_id: element_id, text: String(text || "") }; }
function setStatus(id, type, html) { var el = document.getElementById(id); if (el) { el.className = "status-badge status-" + type; el.innerHTML = html; } }