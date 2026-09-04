
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
        const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const endDateLabel = extractEndDate(aoa);   // Row 5 label
        const latestDOS = getLatestDOS(aoa);        // Actual latest DOS in column

        document.getElementById("dateHeader").innerText = formatDate(latestDOS);

        processDailyData(aoa, latestDOS);
    };

    reader.readAsArrayBuffer(file);
}

/* ----------------------------------------------------------
   Extract END DATE from row 5 label
----------------------------------------------------------- */
function extractEndDate(aoa) {
    const periodRow = aoa[4][0]; // row 5 = index 4

    if (!periodRow || typeof periodRow !== "string") return "";

    const match = periodRow.match(/-\s*([A-Za-z]{3}\s+\d{2}\s+\d{4})/);
    return match ? match[1] : "";
}

/* ----------------------------------------------------------
   Find the latest Date of Service in the column
----------------------------------------------------------- */
function getLatestDOS(aoa) {
    let latest = null;

    for (let r = 8; r < aoa.length; r++) {
        const raw = aoa[r][5];
        const dos = fixDate(raw);
        if (!dos) continue;

        const d = new Date(dos);
        if (!latest || d > latest) latest = d;
    }

    return latest ? latest.toLocaleDateString("en-US") : "";
}

/* ----------------------------------------------------------
   Format date nicely
----------------------------------------------------------- */
function formatDate(dateString) {
    const d = new Date(dateString);
    return d.toLocaleDateString("en-US", {
        year: "numeric",
        month: "long",
        day: "numeric",
        weekday: "long"
    });
}

/* ----------------------------------------------------------
   Fix Excel serial dates
----------------------------------------------------------- */
function fixDate(value) {
    if (!value) return "";
    if (typeof value === "number") {
        const date = XLSX.SSF.parse_date_code(value);
        return `${String(date.m).padStart(2, "0")}/${String(date.d).padStart(2, "0")}/${date.y}`;
    }
    return value;
}

/* ----------------------------------------------------------
   Main processing logic
----------------------------------------------------------- */
function processDailyData(aoa, latestDOS) {
    let locCounts = {};
    let modCounts = {};
    let statusCounts = { Reported: 0, Pending: 0 };
    let backlog = {};
    let historical = {};

    const locationMap = {
        "Astrana Breast Center": "ABC",
        "Diagnostic Medical Group Arcadia": "AR",
        "Diagnostic Medical Group City of Industry": "CI",
        "Diagnostic Medical Group Monterey Park": "MP",
        "Diagnostic Medical Group San Gabriel": "SG",
        "Synergy San Gabriel": "SSG",
        "A-Scheduling": "A-Scheduling"
    };

    const techStatuses = [
        "TechComplete",
        "Completed",
        "ExamComplete",
        "ExamDone",
        "StudyComplete",
        "StudyDone",
        "Final",
        "Dictated",
        "Signed"
    ];

    for (let r = 8; r < aoa.length; r++) {
        const modality = String(aoa[r][0] || "").trim();
        const locationFull = String(aoa[r][1] || "").trim();
        const status = String(aoa[r][4] || "").trim();
        const dosRaw = aoa[r][5];
        const apptID = String(aoa[r][6] || "").trim();

        const dos = fixDate(dosRaw);
        if (!apptID || !dos) continue;

        const location = locationMap[locationFull] || locationFull;

        /* DAILY TABLES — ONLY latest DOS */
        if (dos === latestDOS) {
            locCounts[location] = (locCounts[location] || 0) + 1;
            modCounts[modality] = (modCounts[modality] || 0) + 1;

            if (status === "Reported") statusCounts.Reported++;
            else statusCounts.Pending++;

            if (techStatuses.includes(status)) {
                if (!backlog[dos]) backlog[dos] = new Set();
                backlog[dos].add(apptID);
            }
        }

        /* HISTORICAL — ONLY dates BEFORE latest DOS */
        const d = new Date(dos);
        const latest = new Date(latestDOS);

        if (d < latest) {
            if (!historical[dos]) historical[dos] = 0;
            historical[dos]++;
        }
    }

    renderTables(locCounts, modCounts, statusCounts, backlog, historical);
}

/* ----------------------------------------------------------
   Render tables
----------------------------------------------------------- */
function renderTables(locCounts, modCounts, statusCounts, backlog, historical) {

    /* LOCATION TABLE */
    let locHTML = "<tr><th>Location</th><th>Procedures</th><th>%</th></tr>";
    let totalLoc = Object.values(locCounts).reduce((a, b) => a + b, 0);

    Object.keys(locCounts).forEach(loc => {
        const count = locCounts[loc];
        const pct = ((count / totalLoc) * 100).toFixed(2);
        locHTML += `<tr><td>${loc}</td><td>${count}</td><td>${pct}%</td></tr>`;
    });

    locHTML += `<tr><td>Total</td><td>${totalLoc}</td><td>100%</td></tr>`;
    document.getElementById("locTable").innerHTML = locHTML;

    /* MODALITY TABLE */
    let modHTML = "<tr><th>Modality</th><th>Procedures</th><th>%</th></tr>";
    let totalMod = Object.values(modCounts).reduce((a, b) => a + b, 0);

    Object.keys(modCounts).forEach(mod => {
        const count = modCounts[mod];
        const pct = ((count / totalMod) * 100).toFixed(2);
        modHTML += `<tr><td>${mod}</td><td>${count}</td><td>${pct}%</td></tr>`;
    });

    modHTML += `<tr><td>Total</td><td>${totalMod}</td><td>100%</td></tr>`;
    document.getElementById("modTable").innerHTML = modHTML;

    /* STATUS TABLE */
    let statusHTML = "<tr><th>Status</th><th>Count</th></tr>";
    statusHTML += `<tr><td>Reported</td><td>${statusCounts.Reported}</td></tr>`;
    statusHTML += `<tr><td>Pending</td><td>${statusCounts.Pending}</td></tr>`;
    document.getElementById("statusTable").innerHTML = statusHTML;

    /* BACKLOG TABLE — ONLY latest DOS */
    let backlogHTML = "<tr><th>Date</th><th>Exams</th></tr>";

    Object.keys(backlog).forEach(dos => {
        const count = backlog[dos].size;
        backlogHTML += `<tr><td>${dos}</td><td>${count}</td></tr>`;
    });

    document.getElementById("backlogTable").innerHTML = backlogHTML;

    /* HISTORICAL TABLE — 7/1/2026 → 9/1/2026 */
    let histHTML = "<tr><th>Date</th><th>Exams</th></tr>";

    Object.keys(historical)
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(dos => {
            histHTML += `<tr><td>${dos}</td><td>${historical[dos]}</td></tr>`;
        });

    document.getElementById("historicalTable").innerHTML = histHTML;
}
