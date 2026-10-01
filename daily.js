let cumulativeAOA = [];
function runDailySummary() {
const fileInput = document.getElementById("dailyFile");
const file = fileInput.files[0];
if (!file) {
alert("Please upload a file.");
return;
}
const reader = new FileReader();
reader.onload = function (e) {
const data = new Uint8Array(e.target.result);
const workbook = XLSX.read(data, { type: "array" });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const newAOA = XLSX.utils.sheet_to_json(sheet, { header: 1 });
if (cumulativeAOA.length === 0) {
cumulativeAOA = newAOA;
} else {
const nextDayRows = newAOA.slice(8);
cumulativeAOA = cumulativeAOA.concat(nextDayRows);
}
const latestDOS = getLatestDOS(cumulativeAOA);
document.getElementById("dateHeader").innerText = formatDate(latestDOS);
processDailyData(cumulativeAOA, latestDOS);
};
reader.readAsArrayBuffer(file);
}
function resetDashboardData() {
cumulativeAOA = [];
document.getElementById("dateHeader").innerText = "";
document.getElementById("leftColumn").innerHTML = "";
document.getElementById("rightColumn").innerHTML = "";
alert("Dashboard data cleared.");
}
/* ---------------- PDF AUTO-FIT ---------------- */
function downloadPDF() {
html2canvas(document.body, { scale: 2 }).then(canvas => {
const pdf = new jspdf.jsPDF("p", "mm", "letter");
const pageWidth = pdf.internal.pageSize.getWidth();
const pageHeight = pdf.internal.pageSize.getHeight();
const imgData = canvas.toDataURL("image/png");
let imgWidth = pageWidth;
let imgHeight = canvas.height * (imgWidth / canvas.width);
if (imgHeight > pageHeight) {
const scale = pageHeight / imgHeight;
imgWidth *= scale;
imgHeight *= scale;
}
pdf.addImage(imgData, "PNG", (pageWidth - imgWidth) / 2, 0, imgWidth, imgHeight);
pdf.save("DailySummary.pdf");
});
}
/* ---------------- DATE FIX ---------------- */
function fixDate(value) {
if (!value) return "";
if (typeof value === "number") {
const date = XLSX.SSF.parse_date_code(value);
return ${String(date.m).padStart(2, "0")}/${String(date.d).padStart(2, "0")}/${date.y};
}
if (typeof value === "string") {
const cleaned = value.trim().replace(/\s+/g, "");
const d = new Date(cleaned);
if (!isNaN(d.getTime())) {
return d.toLocaleDateString("en-US");
}
}
return "";
}
function getLatestDOS(aoa) {
let latest = null;
for (let r = 8; r < aoa.length; r++) {
if (!aoa[r]) continue;
const raw = aoa[r][5];
const dos = fixDate(raw);
if (!dos) continue;
const d = new Date(dos);
if (!latest || d > latest) latest = d;
}
return latest ? latest.toLocaleDateString("en-US") : "";
}
function formatDate(dateString) {
if (!dateString) return "N/A";
const d = new Date(dateString);
return d.toLocaleDateString("en-US", {
year: "numeric",
month: "long",
day: "numeric",
weekday: "long"
});
}
function computeDaysBehind(dos, endDate) {
const [m1, d1, y1] = dos.split("/");
const [m2, d2, y2] = endDate.split("/");
const dosDate = new Date(${y1}-${m1}-${d1});
const end = new Date(${y2}-${m2}-${d2});
const diff = end - dosDate;
return Math.floor(diff / 86400000);
}
/* ---------------- MAIN PROCESSING ---------------- */
function processDailyData(aoa, latestDOS) {
let locCounts = {};
let modCounts = {};
let statusCounts = { Reported: 0, Pending: 0 };
let backlog = {};
let historical = {};
let noShowCount = 0;
let modalityLocation = {};
let noShowLocation = {};
const locationMap = {
"Astrana Breast Center": "ABC",
"Diagnostic Medical Group Arcadia": "AR",
"Diagnostic Medical Group City of Industry": "CI",
"Diagnostic Medical Group Monterey Park": "MP",
"Diagnostic Medical Group San Gabriel": "SG",
"Synergy San Gabriel": "SSG",
"A-Scheduling": "A-Scheduling"
};
const historicalStatuses = [
"Completed WO Report",
"Reported",
"TechComplete"
];
for (let r = 8; r < aoa.length; r++) {
if (!aoa[r]) continue;
const modality = String(aoa[r][0] || "").trim();
const locationFull = String(aoa[r][1] || "").trim();
const statusRaw = String(aoa[r][24] || "").trim();
const dosRaw = aoa[r][5];
const apptID = String(aoa[r][6] || "").trim();
const dos = fixDate(dosRaw);
if (!dos || !apptID) continue;
const statusClean = statusRaw.replace(/\s+/g, "").toLowerCase();
const location = locationMap[locationFull] || locationFull;
const d = new Date(dos);
const latest = new Date(latestDOS);
if (d.getTime() === latest.getTime()) {
if (
statusClean === "completedworeport" ||
statusClean === "reported" ||
statusClean === "techcomplete"
) {
locCounts[location] = (locCounts[location] || 0) + 1;
modCounts[modality] = (modCounts[modality] || 0) + 1;
if (!modalityLocation[modality]) modalityLocation[modality] = {};
modalityLocation[modality][location] =
(modalityLocation[modality][location] || 0) + 1;
}
if (statusClean === "reported" || statusClean === "completedworeport") {
statusCounts.Reported++;
} else if (statusClean === "techcomplete") {
statusCounts.Pending++;
}
if (statusClean === "noshow") {
noShowCount++;
if (!noShowLocation[modality]) noShowLocation[modality] = {};
noShowLocation[modality][location] =
(noShowLocation[modality][location] || 0) + 1;
}
}
if (d < latest && statusClean === "techcomplete") {
const daysBehind = computeDaysBehind(dos, latestDOS);
if (!backlog[dos]) {
backlog[dos] = { count: 0, daysBehind: daysBehind };
}
backlog[dos].count++;
}
if (d < latest && historicalStatuses.includes(statusRaw)) {
if (!historical[dos]) historical[dos] = 0;
historical[dos]++;
}
}
renderTables(locCounts, modCounts, statusCounts, backlog, historical, noShowCount);
renderNoShowSummary(noShowLocation);
renderModalityPerLocation(modalityLocation);
renderNoShowPerLocation(noShowLocation);
autoPlaceTables();
}
/* ---------------- RENDER TABLES ---------------- */
function renderTables(locCounts, modCounts, statusCounts, backlog, historical, noShowCount) {
let locHTML = "LocationProcedures%";
let totalLoc = Object.values(locCounts).reduce((a, b) => a + b, 0);
Object.keys(locCounts)
.sort()
.forEach(loc => {
const count = locCounts[loc];
const pct = totalLoc > 0 ? ((count / totalLoc) * 100).toFixed(2) : "0.00";
locHTML += <tr><td>${loc}</td><td>${count}</td><td>${pct}%</td></tr>;
});
locHTML += <tr><td>Total</td><td>${totalLoc}</td><td>100%</td></tr>;
document.getElementById("locTable").innerHTML = locHTML;
let modHTML = "ModalityProcedures%";
let totalMod = Object.values(modCounts).reduce((a, b) => a + b, 0);
Object.keys(modCounts)
.sort()
.forEach(mod => {
const count = modCounts[mod];
const pct = totalMod > 0 ? ((count / totalMod) * 100).toFixed(2) : "0.00";
modHTML += <tr><td>${mod}</td><td>${count}</td><td>${pct}%</td></tr>;
});
modHTML += <tr><td>Total</td><td>${totalMod}</td><td>100%</td></tr>;
document.getElementById("modTable").innerHTML = modHTML;
let statusHTML = "StatusCount";
statusHTML += <tr><td>Reported</td><td>${statusCounts.Reported}</td></tr>;
statusHTML += <tr><td>Pending</td><td>${statusCounts.Pending}</td></tr>;
document.getElementById("statusTable").innerHTML = statusHTML;
let backlogHTML = "DateExams Not ReadDays Behind";
Object.keys(backlog)
.sort((a, b) => new Date(a) - new Date(b))
.forEach(dos => {
backlogHTML += <tr><td>${dos}</td><td>${backlog[dos].count}</td><td>${backlog[dos].daysBehind}</td></tr>;
});
document.getElementById("backlogTable").innerHTML = backlogHTML;
renderHistoricalSummaryTable(historical);
}
function renderHistoricalSummaryTable(historical) {
const container = document.getElementById("historicalSummaryTable");
const latestDateText = document.getElementById("dateHeader").innerText;
let cleanDateStr = latestDateText.replace(/[A-Za-z\s]+,/, "").trim();
const latestDate = new Date(cleanDateStr);
const currentYear = !isNaN(latestDate.getTime()) ? latestDate.getFullYear() : 2026;
const currentMonth = !isNaN(latestDate.getTime()) ? latestDate.getMonth() : 8;
const currentQuarter = Math.ceil((currentMonth + 1) / 3);
const formatMonthKey = (m, y) => ${new Date(y, m).toLocaleDateString("en-US", { month: "short" })}-${y};
const formatQuarterKey = (q, y) => Q${q}-${y};
let reportStructure = [
{ key: formatMonthKey(currentMonth, currentYear), typeLabel: "MTD", matchFn: (m, y, q) => m === currentMonth && y === currentYear },
{ key: formatMonthKey((currentMonth - 1 + 12) % 12, currentMonth - 1 < 0 ? currentYear - 1 : currentYear), typeLabel: "Full Month", matchFn: (m, y, q) => m === ((currentMonth - 1 + 12) % 12) && y === (currentMonth - 1 < 0 ? currentYear - 1 : currentYear) },
{ key: formatMonthKey((currentMonth - 2 + 12) % 12, currentMonth - 2 < 0 ? currentYear - 1 : currentYear), typeLabel: "Full Month", matchFn: (m, y, q) => m === ((currentMonth - 2 + 12) % 12) && y === (currentMonth - 2 < 0 ? currentYear - 1 : currentYear) },
{ key: formatQuarterKey(currentQuarter, currentYear), typeLabel: "QTD", matchFn: (m, y, q) => q === currentQuarter && y === currentYear }
];
let lookbackQuarters = [
{ q: currentQuarter - 1 === 0 ? 4 : currentQuarter - 1, y: currentQuarter - 1 === 0 ? currentYear - 1 : currentYear },
{ q: currentQuarter - 2 <= 0 ? currentQuarter - 2 + 4 : currentQuarter - 2, y: currentQuarter - 2 <= 0 ? currentYear - 1 : currentYear }
];
lookbackQuarters.forEach(bq => {
if (bq.y < 2026 || (bq.y === 2026 && bq.q < 2)) return;
reportStructure.push({
key: formatQuarterKey(bq.q, bq.y),
typeLabel: "Quarter",
matchFn: (m, y, q) => q === bq.q && y === bq.y
});
});
let html = "PeriodDMGPeriod Type";
reportStructure.forEach(bucket => {
let dmgTotal = 0;
Object.keys(historical).forEach(dateStr => {
const date = new Date(dateStr);
const m = date.getMonth();
const y = date.getFullYear();
const q = Math.ceil((m + 1) / 3);
if (bucket.matchFn(m, y, q)) {
dmgTotal += historical[dateStr] || 0;
}
});
html += `
${bucket.key}
${dmgTotal.toLocaleString()}
${bucket.typeLabel}
`;
});
container.innerHTML = html;
}
function renderModalityPerLocation(modalityLocation) {
const container = document.getElementById("modalityPerLocationTable");
const locations = ["AR", "CI", "MP", "SG", "SSG"];
let html = "Modality";
locations.forEach(loc => {
html += <th>${loc}</th>;
});
html += "Total";
Object.keys(modalityLocation)
.sort()
.forEach(mod => {
let rowTotal = 0;
html += <tr><td>${mod}</td>;
locations.forEach(loc => {
const val = modalityLocation[mod][loc] || 0;
rowTotal += val;
html += <td>${val}</td>;
});
html += <td>${rowTotal}</td></tr>;
});
let colTotals = {};
locations.forEach(loc => colTotals[loc] = 0);
let grandTotal = 0;
Object.keys(modalityLocation).forEach(mod => {
locations.forEach(loc => {
const val = modalityLocation[mod][loc] || 0;
colTotals[loc] += val;
grandTotal += val;
});
});
html += "Total";
locations.forEach(loc => {
html += <td>${colTotals[loc]}</td>;
});
html += <td>${grandTotal}</td></tr>;
container.innerHTML = html;
}
function renderNoShowPerLocation(noShowLocation) {
const container = document.getElementById("noShowPerLocationTable");
const locations = ["AR", "CI", "MP", "SG", "SSG"];
let html = "Modality";
locations.forEach(loc => {
html += <th>${loc}</th>;
});
html += "Total";
Object.keys(noShowLocation)
.sort()
.forEach(mod => {
let rowTotal = 0;
html += <tr><td>${mod}</td>;
locations.forEach(loc => {
const val = noShowLocation[mod][loc] || 0;
rowTotal += val;
html += <td>${val}</td>;
});
html += <td>${rowTotal}</td></tr>;
});
let colTotals = {};
locations.forEach(loc => colTotals[loc] = 0);
let grandTotal = 0;
Object.keys(noShowLocation).forEach(mod => {
locations.forEach(loc => {
const val = noShowLocation[mod][loc] || 0;
colTotals[loc] += val;
grandTotal += val;
});
});
html += "Total";
locations.forEach(loc => {
html += <td>${colTotals[loc]}</td>;
});
html += <td>${grandTotal}</td></tr>;
container.innerHTML = html;
}
function renderNoShowSummary(noShowLocation) {
const container = document.getElementById("noShowSummaryTable");
let html = "LocationProcedures with No Show";
const locations = ["AR", "CI", "MP", "SG", "SSG"];
let totalNoShow = 0;
locations.forEach(loc => {
let locTotal = 0;
Object.keys(noShowLocation).forEach(mod => {
locTotal += noShowLocation[mod][loc] || 0;
});
totalNoShow += locTotal;
html += <tr><td>${loc}</td><td>${locTotal}</td></tr>;
});
html += <tr><td>Total No Show</td><td>${totalNoShow}</td></tr>;
container.innerHTML = html;
}
function autoPlaceTables() {
const left = document.getElementById("leftColumn");
const right = document.getElementById("rightColumn");
left.innerHTML = "";
right.innerHTML = "";
const forcedLeftTables = [
["Summary by Location", "locTable"],
["Summary by Modality", "modTable"],
["# No Show Summary", "noShowSummaryTable"]
];
const balancedTables = [
["Total Reported / Pending Read", "statusTable"],
["Backlog (All Dates Before Latest DOS)", "backlogTable"],
["# Historical Exam Data", "historicalSummaryTable"],
["Summary by Modality per Location", "modalityPerLocationTable"],
["Summary No Show by Modality per Location", "noShowPerLocationTable"]
];
let leftHeight = 0;
let rightHeight = 0;
function createBlock(title, id) {
const wrapper = document.createElement("div");
wrapper.className = "report-block";
wrapper.style.marginBottom = "25px";
wrapper.innerHTML = <h2>${title}</h2>;
const originalTable = document.getElementById(id);
const tableClone = originalTable.cloneNode(true);
wrapper.appendChild(tableClone);
document.body.appendChild(wrapper);
const height = wrapper.offsetHeight;
wrapper.remove();
return { element: wrapper, height: height };
}
forcedLeftTables.forEach(([title, id]) => {
const block = createBlock(title, id);
left.appendChild(block.element);
leftHeight += block.height;
});
balancedTables.forEach(([title, id]) => {
const block = createBlock(title, id);
if (leftHeight <= rightHeight) {
left.appendChild(block.element);
leftHeight += block.height;
} else {
right.appendChild(block.element);
rightHeight += block.height;
}
});
}

