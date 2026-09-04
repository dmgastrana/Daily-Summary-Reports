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

        processDailyData(aoa);
    };

    reader.readAsArrayBuffer(file);
}

function fixDate(value) {
    if (!value) return "";
    if (typeof value === "number") {
        const date = XLSX.SSF.parse_date_code(value);
        return `${String(date.m).padStart(2, "0")}/${String(date.d).padStart(2, "0")}/${date.y}`;
    }
    return value;
}

function computeDaysBehind(dos) {
    const referenceDate = new Date();
    referenceDate.setHours(0, 0, 0, 0);

    const [month, day, year] = dos.split("/");
    const dosDate = new Date(`${year}-${month}-${day}`);

    const diff = referenceDate - dosDate;
    return Math.floor(diff / 86400000);
}

function processDailyData(aoa) {
    let locCounts = {};
    let modCounts = {};
    let statusCounts = { Reported: 0, Pending: 0 };
    let backlog = {};

    for (let r = 0; r < aoa.length; r++) {

        if (r === 0) continue; // skip header row

        const modality = String(aoa[r][0] || "").trim();
        const location = String(aoa[r][1] || "").trim();
        const status = String(aoa[r][4] || "").trim();
        const dosRaw = aoa[r][5];
        const apptID = String(aoa[r][6] || "").trim();

        const dos = fixDate(dosRaw);

        if (!apptID || !dos) continue;

        locCounts[location] = (locCounts[location] || 0) + 1;
        modCounts[modality] = (modCounts[modality] || 0) + 1;

        if (status === "Reported") statusCounts.Reported++;
        else statusCounts.Pending++;

        if (status === "TechComplete") {
            if (!backlog[dos]) backlog[dos] = new Set();
            backlog[dos].add(apptID);
        }
    }

    renderTables(locCounts, modCounts, statusCounts, backlog);
}

function renderTables(locCounts, modCounts, statusCounts, backlog) {
    const today = new Date();
    const dateString = today.toLocaleDateString("en-US", {
        year: "numeric",
        month: "long",
        day: "numeric",
        weekday: "long"
    });

    document.getElementById("dateHeader").innerText = dateString;

    let locHTML = "<tr><th>Location</th><th>Procedures</th><th>%</th></tr>";
    let totalLoc = Object.values(locCounts).reduce((a, b) => a + b, 0);

    Object.keys(locCounts).forEach(loc => {
        const count = locCounts[loc];
        const pct = ((count / totalLoc) * 100).toFixed(2);
        locHTML += `<tr><td>${loc}</td><td>${count}</td><td>${pct}%</td></tr>`;
    });

    locHTML += `<tr><td>Total</td><td>${totalLoc}</td><td>100%</td></tr>`;
    document.getElementById("locTable").innerHTML = locHTML;

    let modHTML = "<tr><th>Modality</th><th>Procedures</th><th>%</th></tr>";
    let totalMod = Object.values(modCounts).reduce((a, b) => a + b, 0);

    Object.keys(modCounts).forEach(mod => {
        const count = modCounts[mod];
        const pct = ((count / totalMod) * 100).toFixed(2);
        modHTML += `<tr><td>${mod}</td><td>${count}</td><td>${pct}%</td></tr>`;
    });

    modHTML += `<tr><td>Total</td><td>${totalMod}</td><td>100%</td></tr>`;
    document.getElementById("modTable").innerHTML = modHTML;

    let statusHTML = "<tr><th>Status</th><th>Count</th></tr>";
    statusHTML += `<tr><td>Reported</td><td>${statusCounts.Reported}</td></tr>`;
    statusHTML += `<tr><td>Pending</td><td>${statusCounts.Pending}</td></tr>`;
    document.getElementById("statusTable").innerHTML = statusHTML;

    // FINAL FIX — remove phantom dates like 09/01/2026
    let backlogHTML = "<tr><th>Date</th><th>Exams</th><th>Days Behind</th></tr>";

    Object.keys(backlog)
        .filter(dos => backlog[dos].size > 0)
        .filter(dos => /^\d{2}\/\d{2}\/\d{4}$/.test(dos)) // must be real date
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(dos => {
            const count = backlog[dos].size;
            const daysBehind = computeDaysBehind(dos);
            backlogHTML += `<tr><td>${dos}</td><td>${count}</td><td>${daysBehind}</td></tr>`;
        });

    document.getElementById("backlogTable").innerHTML = backlogHTML;
}
