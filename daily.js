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

        const latestDOS = getLatestDOS(aoa);
        document.getElementById("dateHeader").innerText = formatDate(latestDOS);

        processDailyData(aoa, latestDOS);
    };

    reader.readAsArrayBuffer(file);
}

/* ----------------------------------------------------------
   Bulletproof date parser
----------------------------------------------------------- */
function fixDate(value) {
    if (!value) return "";

    if (typeof value === "number") {
        const date = XLSX.SSF.parse_date_code(value);
        return `${String(date.m).padStart(2, "0")}/${String(date.d).padStart(2, "0")}/${date.y}`;
    }

    if (typeof value === "string") {
        const cleaned = value.trim().replace(/\s+/g, "");
        const d = new Date(cleaned);
        if (!isNaN(d)) {
            return d.toLocaleDateString("en-US");
        }
    }

    return "";
}

/* ----------------------------------------------------------
   Find latest DOS in column F
----------------------------------------------------------- */
function getLatestDOS(aoa) {
    let latest = null;

    for (let r = 8; r < aoa.length; r++) {
        const raw = aoa[r][5]; // Column F
        const dos = fixDate(raw);
        if (!dos) continue;

        const d = new Date(dos);
        if (!latest || d > latest) latest = d;
    }

    return latest.toLocaleDateString("en-US");
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
   Days Behind = End Date − DOS
----------------------------------------------------------- */
function computeDaysBehind(dos, endDate) {
    const [m1, d1, y1] = dos.split("/");
    const [m2, d2, y2] = endDate.split("/");

    const dosDate = new Date(`${y1}-${m1}-${d1}`);
    const end = new Date(`${y2}-${m2}-${d2}`);

    const diff = end - dosDate;
    return Math.floor(diff / 86400000);
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
    let noShowCount = 0;

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

        const modality = String(aoa[r][0] || "").trim();   // Column A
        const locationFull = String(aoa[r][1] || "").trim(); // Column B
        const statusRaw = String(aoa[r][24] || "").trim();   // Column Y
        const dosRaw = aoa[r][5];                            // Column F
        const apptID = String(aoa[r][6] || "").trim();       // Column G

        const dos = fixDate(dosRaw);
        if (!dos || !apptID) continue;

        const statusClean = statusRaw.replace(/\s+/g, "").toLowerCase();
        const location = locationMap[locationFull] || locationFull;

        const d = new Date(dos);
        const latest = new Date(latestDOS);

        /* DAILY — ONLY latest DOS AND ONLY correct statuses */
        if (d.getTime() === latest.getTime()) {

            // Count ONLY these statuses as procedures:
            if (
                statusClean === "completedworeport" ||
                statusClean === "reported" ||
                statusClean === "techcomplete"
            ) {
                locCounts[location] = (locCounts[location] || 0) + 1;
                modCounts[modality] = (modCounts[modality] || 0) + 1;
            }

            // Reported / Pending Read
            if (statusClean === "reported" || statusClean === "completedworeport") {
                statusCounts.Reported++;
            } else if (statusClean === "techcomplete") {
                statusCounts.Pending++;
            }

            // No Show
            if (statusClean === "noshow") {
                noShowCount++;
            }
        }

        /* BACKLOG — ONLY TechComplete */
        if (d < latest && statusClean === "techcomplete") {
            const daysBehind = computeDaysBehind(dos, latestDOS);

            if (!backlog[dos]) {
                backlog[dos] = { count: 0, daysBehind: daysBehind };
            }

            backlog[dos].count++;
        }

        /* HISTORICAL — old dates only */
        if (d < latest && historicalStatuses.includes(statusRaw)) {
            if (!historical[dos]) historical[dos] = 0;
            historical[dos]++;
        }
    }

    renderTables(locCounts, modCounts, statusCounts, backlog, historical, noShowCount);
}

/* ----------------------------------------------------------
   Render tables
----------------------------------------------------------- */
function renderTables(locCounts, modCounts, statusCounts, backlog, historical, noShowCount) {

    /* LOCATION TABLE — SORT A → Z */
    let locHTML = "<tr><th>Location</th><th>Procedures</th><th>%</th></tr>";
    let totalLoc = Object.values(locCounts).reduce((a, b) => a + b, 0);

    Object.keys(locCounts)
        .sort()  // ⭐ alphabetical order
        .forEach(loc => {
            const count = locCounts[loc];
            const pct = ((count / totalLoc) * 100).toFixed(2);
            locHTML += `<tr><td>${loc}</td><td>${count}</td><td>${pct}%</td></tr>`;
        });

    locHTML += `<tr><td>Total</td><td>${totalLoc}</td><td>100%</td></tr>`;
    document.getElementById("locTable").innerHTML = locHTML;

    /* MODALITY TABLE — SORT A → Z */
    let modHTML = "<tr><th>Modality</th><th>Procedures</th><th>%</th></tr>";
    let totalMod = Object.values(modCounts).reduce((a, b) => a + b, 0);

    Object.keys(modCounts)
        .sort()  // ⭐ alphabetical order
        .forEach(mod => {
            const count = modCounts[mod];
            const pct = ((count / totalMod) * 100).toFixed(2);
            modHTML += `<tr><td>${mod}</td><td>${count}</td><td>${pct}%</td></tr>`;
        });

    modHTML += `<tr><td>Total</td><td>${totalMod}</td><td>100%</td></tr>`;
    document.getElementById("modTable").innerHTML = modHTML;

    /* TOTAL NO SHOW */
    document.getElementById("modTable").insertAdjacentHTML(
        "afterend",
        `<div style="margin-top:10px;">Total No Show: ${noShowCount}</div>`
    );

    /* STATUS TABLE */
    let statusHTML = "<tr><th>Status</th><th>Count</th></tr>";
    statusHTML += `<tr><td>Reported</td><td>${statusCounts.Reported}</td></tr>`;
    statusHTML += `<tr><td>Pending</td><td>${statusCounts.Pending}</td></tr>`;
    document.getElementById("statusTable").innerHTML = statusHTML;

    /* BACKLOG TABLE */
    let backlogHTML = "<tr><th>Date</th><th>Exams Not Read</th><th>Days Behind</th></tr>";

    Object.keys(backlog)
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(dos => {
            backlogHTML += `<tr><td>${dos}</td><td>${backlog[dos].count}</td><td>${backlog[dos].daysBehind}</td></tr>`;
        });

    document.getElementById("backlogTable").innerHTML = backlogHTML;

    /* HISTORICAL TABLE */
    let histHTML = "<tr><th>Date</th><th>Exams</th></tr>";

    Object.keys(historical)
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(dos => {
            histHTML += `<tr><td>${dos}</td><td>${historical[dos]}</td></tr>`;
        });

    document.getElementById("historicalTable").innerHTML = histHTML;
}
