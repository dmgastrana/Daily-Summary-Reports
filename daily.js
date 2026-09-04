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

        extractEndDate(aoa);   // NEW: pull date from row 5
        processDailyData(aoa); // process data starting row 8
    };

    reader.readAsArrayBuffer(file);
}

/* ----------------------------------------------------------
   EXTRACT END DATE FROM ROW 5
   Example row:
   "Report ran for the period: Jul 01 2026 - Sep 02 2026"
----------------------------------------------------------- */
function extractEndDate(aoa) {
    const periodRow = aoa[4][0]; // row 5 = index 4

    let endDate = "";

    if (periodRow && typeof periodRow === "string") {
        const match = periodRow.match(/-\s*([A-Za-z]{3}\s+\d{2}\s+\d{4})/);
        if (match) {
            endDate = match[1]; // "Sep 02 2026"
        }
    }

    let formattedEndDate = "";
    if (endDate) {
        const d = new Date(endDate);
        formattedEndDate = d.toLocaleDateString("en-US", {
            year: "numeric",
            month: "long",
            day: "numeric",
            weekday: "long"
        });
    }

    document.getElementById("dateHeader").innerText = formattedEndDate;
}

/* ----------------------------------------------------------
   FIX DATE FORMAT (Excel serial → MM/DD/YYYY)
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
   DAYS BEHIND CALCULATION
----------------------------------------------------------- */
function computeDaysBehind(dos) {
    const referenceDate = new Date();
    referenceDate.setHours(0, 0, 0, 0);

    const [month, day, year] = dos.split("/");
    const dosDate = new Date(`${year}-${month}-${day}`);

    const diff = referenceDate - dosDate;
    return Math.floor(diff / 86400000);
}

/* ----------------------------------------------------------
   MAIN PROCESSING LOGIC
----------------------------------------------------------- */
function processDailyData(aoa) {
    let locCounts = {};
    let modCounts = {};
    let statusCounts = { Reported: 0, Pending: 0 };
    let backlog = {};

    // NEW: short location code mapping
    const locationMap = {
        "Astrana Breast Center": "ABC",
        "Diagnostic Medical Group Arcadia": "AR",
        "Diagnostic Medical Group City of Industry": "CI",
        "Diagnostic Medical Group Monterey Park": "MP",
        "Diagnostic Medical Group San Gabriel": "SG",
        "Synergy San Gabriel": "SSG",
        "A-Scheduling": "A-Scheduling"
    };

    // NEW: multi-status backlog support
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

    for (let r = 0; r < aoa.length; r++) {

        // NEW: skip rows 0–7 (header row is row 8)
        if (r < 8) continue;

        const modality = String(aoa[r][0] || "").trim();
        const locationFull = String(aoa[r][1] || "").trim();
        const status = String(aoa[r][4] || "").trim();
        const dosRaw = aoa[r][5];
        const apptID = String(aoa[r][6] || "").trim();

        const dos = fixDate(dosRaw);

        if (!apptID || !dos) continue;

        // Convert full location → short code
        const location = locationMap[locationFull] || locationFull;

        // Count locations
        locCounts[location] = (locCounts[location] || 0) + 1;

        // Count modalities
        modCounts[modality] = (modCounts[modality] || 0) + 1;

        // Count statuses
        if (status === "Reported") statusCounts.Reported++;
        else statusCounts.Pending++;

        // Backlog logic
        if (techStatuses.includes(status)) {
            if (!backlog[dos]) backlog[dos] = new Set();
            backlog[dos].add(apptID);
        }
    }

    renderTables(locCounts, modCounts, statusCounts, backlog);
}

/* ----------------------------------------------------------
   RENDER TABLES
----------------------------------------------------------- */
function renderTables(locCounts, modCounts, statusCounts, backlog) {

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

    /* BACKLOG TABLE */
    let backlogHTML = "<tr><th>Date</th><th>Exams</th><th>Days Behind</th></tr>";

    Object.keys(backlog)
        .filter(dos => backlog[dos].size > 0)
        .filter(dos => /^\d{2}\/\d{2}\/\d{4}$/.test(dos))
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(dos => {
            const count = backlog[dos].size;
            const daysBehind = computeDaysBehind(dos);
            backlogHTML += `<tr><td>${dos}</td><td>${count}</td><td>${daysBehind}</td></tr>`;
        });

    document.getElementById("backlogTable").innerHTML = backlogHTML;
}
