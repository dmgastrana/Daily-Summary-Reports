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

function getLatestDOS(aoa) {
    let latest = null;

    for (let r = 8; r < aoa.length; r++) {
        const raw = aoa[r][5];
        const dos = fixDate(raw);
        if (!dos) continue;

        const d = new Date(dos);
        if (!latest || d > latest) latest = d;
    }

    return latest.toLocaleDateString("en-US");
}

function formatDate(dateString) {
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

    const dosDate = new Date(`${y1}-${m1}-${d1}`);
    const end = new Date(`${y2}-${m2}-${d2}`);

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

    autoPlaceTables();  // ⭐ AUTO-FIT LEFT/RIGHT COLUMN
}

/* ---------------- RENDER TABLES ---------------- */

function renderTables(locCounts, modCounts, statusCounts, backlog, historical, noShowCount) {

    let locHTML = "<tr><th>Location</th><th>Procedures</th><th>%</th></tr>";
    let totalLoc = Object.values(locCounts).reduce((a, b) => a + b, 0);

    Object.keys(locCounts)
        .sort()
        .forEach(loc => {
            const count = locCounts[loc];
            const pct = ((count / totalLoc) * 100).toFixed(2);
            locHTML += `<tr><td>${loc}</td><td>${count}</td><td>${pct}%</td></tr>`;
        });

    locHTML += `<tr><td>Total</td><td>${totalLoc}</td><td>100%</td></tr>`;
    document.getElementById("locTable").innerHTML = locHTML;

    let modHTML = "<tr><th>Modality</th><th>Procedures</th><th>%</th></tr>";
    let totalMod = Object.values(modCounts).reduce((a, b) => a + b, 0);

    Object.keys(modCounts)
        .sort()
        .forEach(mod => {
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

    let backlogHTML = "<tr><th>Date</th><th>Exams Not Read</th><th>Days Behind</th></tr>";

    Object.keys(backlog)
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(dos => {
            backlogHTML += `<tr><td>${dos}</td><td>${backlog[dos].count}</td><td>${backlog[dos].daysBehind}</td></tr>`;
        });

    document.getElementById("backlogTable").innerHTML = backlogHTML;

    renderHistoricalSummaryTable(historical);
}

function renderHistoricalSummaryTable(historical) {

    const container = document.getElementById("historicalSummaryTable");

    let monthlyTotals = {};
    let quarterlyTotals = {};

    Object.keys(historical).forEach(d => {
        const date = new Date(d);
        const month = date.toLocaleDateString("en-US", { month: "short" });
        const year = date.getFullYear();
        const key = `${month}-${year}`;

        monthlyTotals[key] = (monthlyTotals[key] || 0) + historical[d];

        const q = Math.ceil((date.getMonth() + 1) / 3);
        const qKey = `Q${q}-${year}`;

        quarterlyTotals[qKey] = (quarterlyTotals[qKey] || 0) + historical[d];
    });

    let html = "<tr><th>Period</th><th>Exams</th><th>Date Range</th></tr>";

    Object.keys(monthlyTotals)
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(period => {

            const [mon, yr] = period.split("-");
            const monthIndex = new Date(`${mon} 1, ${yr}`).getMonth();

            const datesInMonth = Object.keys(historical)
                .filter(d => {
                    const dt = new Date(d);
                    return dt.getMonth() === monthIndex && dt.getFullYear() === Number(yr);
                })
                .sort((a, b) => new Date(a) - new Date(b));

            const firstDOS = datesInMonth[0];
            const lastDOS = datesInMonth[datesInMonth.length - 1];

            const periodText = `${firstDOS} to ${lastDOS}`;

            html += `<tr><td>${period}</td><td>${monthlyTotals[period]}</td><td>${periodText}</td></tr>`;
        });

    Object.keys(quarterlyTotals)
        .sort()
        .forEach(period => {

            const [qLabel, yr] = period.split("-");
            const qNum = Number(qLabel.replace("Q", ""));

            const startMonth = (qNum - 1) * 3 + 1;

            const datesInQuarter = Object.keys(historical)
                .filter(d => {
                    const dt = new Date(d);
                    const m = dt.getMonth() + 1;
                    return dt.getFullYear() === Number(yr) &&
                           m >= startMonth &&
                           m < startMonth + 3;
                })
                .sort((a, b) => new Date(a) - new Date(b));

            const firstDOS = datesInQuarter[0];
            const lastDOS = datesInQuarter[datesInQuarter.length - 1];

            const periodText = `${firstDOS} to ${lastDOS}`;

            html += `<tr><td>${period}</td><td>${quarterlyTotals[period]}</td><td>${periodText}</td></tr>`;
        });

    container.innerHTML = html;
}

function renderModalityPerLocation(modalityLocation) {

    const container = document.getElementById("modalityPerLocationTable");

    const locations = ["AR", "CI", "MP", "SG", "SSG"];

    let html = "<tr><th>Modality</th>";

    locations.forEach(loc => {
        html += `<th>${loc}</th>`;
    });

    html += "<th>Total</th></tr>";

    Object.keys(modalityLocation)
        .sort()
        .forEach(mod => {
            let rowTotal = 0;
            html += `<tr><td>${mod}</td>`;

            locations.forEach(loc => {
                const val = modalityLocation[mod][loc] || 0;
                rowTotal += val;
                html += `<td>${val}</td>`;
            });

            html += `<td>${rowTotal}</td></tr>`;
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

    html += "<tr><td>Total</td>";

    locations.forEach(loc => {
        html += `<td>${colTotals[loc]}</td>`;
    });

    html += `<td>${grandTotal}</td></tr>`;

    container.innerHTML = html;
}

function renderNoShowPerLocation(noShowLocation) {

    const container = document.getElementById("noShowPerLocationTable");

    const locations = ["AR", "CI", "MP", "SG", "SSG"];

    let html = "<tr><th>Modality</th>";

    locations.forEach(loc => {
        html += `<th>${loc}</th>`;
    });

    html += "<th>Total</th></tr>";

    Object.keys(noShowLocation)
        .sort()
        .forEach(mod => {
            let rowTotal = 0;
            html += `<tr><td>${mod}</td>`;

            locations.forEach(loc => {
                const val = noShowLocation[mod][loc] || 0;
                rowTotal += val;
                html += `<td>${val}</td>`;
            });

            html += `<td>${rowTotal}</td></tr>`;
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

    html += "<tr><td>Total</td>";

    locations.forEach(loc => {
        html += `<td>${colTotals[loc]}</td>`;
    });

    html += `<td>${grandTotal}</td></tr>`;

    container.innerHTML = html;
}

function renderNoShowSummary(noShowLocation) {

    const container = document.getElementById("noShowSummaryTable");

    let html = "<tr><th>Location</th><th>Procedures with No Show</th></tr>";

    const locations = ["AR", "CI", "MP", "SG", "SSG"];

    let totalNoShow = 0;

    locations.forEach(loc => {
        let locTotal = 0;

        Object.keys(noShowLocation).forEach(mod => {
            locTotal += noShowLocation[mod][loc] || 0;
        });

        totalNoShow += locTotal;

        html += `<tr><td>${loc}</td><td>${locTotal}</td></tr>`;
    });

    html += `<tr><td>Total No Show</td><td>${totalNoShow}</td></tr>`;

    container.innerHTML = html;
}

/* ---------------- AUTO-FIT LEFT/RIGHT COLUMN ---------------- */

function autoPlaceTables() {
    const left = document.getElementById("leftColumn");
    const right = document.getElementById("rightColumn");

    left.innerHTML = "";
    right.innerHTML = "";

    const tables = [
        ["Summary by Location", "locTable"],
        ["Summary by Modality", "modTable"],
        ["# No Show Summary", "noShowSummaryTable"],
        ["Total Reported / Pending Read", "statusTable"],
        ["Backlog (All Dates Before Latest DOS)", "backlogTable"],
        ["# Historical Exam Data", "historicalSummaryTable"],
        ["Summary by Modality per Location", "modalityPerLocationTable"],
        ["Summary No Show by Modality per Location", "noShowPerLocationTable"]
    ];

    let leftHeight = 0;
    const maxHeight = 900; // portrait page limit

    tables.forEach(([title, id]) => {
        const wrapper = document.createElement("div");
        wrapper.innerHTML = `<h2>${title}</h2>`;
        wrapper.appendChild(document.getElementById(id));

        document.body.appendChild(wrapper);
        const height = wrapper.offsetHeight;
        wrapper.remove();

        if (leftHeight + height < maxHeight) {
            left.appendChild(wrapper);
            leftHeight += height;
        } else {
            right.appendChild(wrapper);
        }
    });
}
