/* ============================================================
   FACILITY + MODALITY DEFINITIONS
   ============================================================ */

const FACILITIES = {
    "Diagnostic Medical Group Arcadia": "AR",
    "Diagnostic Medical Group City of Industry": "CI",
    "Diagnostic Medical Group Monterey Park": "MP",
    "Diagnostic Medical Group San Gabriel": "SG",
    "Synergy San Gabriel": "SSG"
};

const MODALITIES = [
    "CT",
    "Bone Density",
    "EKG",
    "MG",
    "MR",
    "US",
    "XR",
    "KS"
];

/* ============================================================
   STATUS MESSAGES
   ============================================================ */

function showStatus(id, msg) {
    document.getElementById("processing").style.display = "none";
    document.getElementById("errorMessage").style.display = "none";
    document.getElementById("successMessage").style.display = "none";

    const el = document.getElementById(id);
    el.innerText = msg;
    el.style.display = "block";
}

/* ============================================================
   MAIN ENTRY
   ============================================================ */

function runDailySummary() {
    showStatus("processing", "Processing…");

    const fileInput = document.getElementById("dailyFile");
    const file = fileInput.files[0];

    if (!file) {
        showStatus("errorMessage", "No file selected.");
        return;
    }

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    reader.onload = function(e) {
        if (fileName.endsWith(".csv")) {
            parseCSV(e.target.result);
        } else {
            parseExcel(e.target.result);
        }
    };

    if (fileName.endsWith(".csv")) {
        reader.readAsBinaryString(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

/* ============================================================
   FIX DATE (handles Excel serial numbers + real dates)
   ============================================================ */

function fixDate(value) {
    if (!value) return "";

    // Excel serial number
    const num = Number(value);
    if (!isNaN(num) && num > 20000 && num < 60000) {
        const base = new Date(1899, 11, 30);
        base.setDate(base.getDate() + num);
        return base.toLocaleDateString("en-US");
    }

    // Normal date
    const d = new Date(String(value).trim());
    return isNaN(d) ? "" : d.toLocaleDateString("en-US");
}

/* ============================================================
   PARSE CSV USING XLSX ENGINE (bulletproof)
   ============================================================ */

function parseCSV(raw) {
    const uint8 = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i++) {
        uint8[i] = raw.charCodeAt(i) & 0xFF;
    }

    const workbook = XLSX.read(uint8, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    generateDailyTables(aoa);
}

/* ============================================================
   PARSE EXCEL
   ============================================================ */

function parseExcel(buffer) {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    generateDailyTables(aoa);
}

/* ============================================================
   MAIN DAILY TABLE GENERATION
   ============================================================ */

function generateDailyTables(aoa) {
    const HEADER_ROW = 8;
    const DATA_START = HEADER_ROW + 1;

    if (!aoa[HEADER_ROW]) {
        showStatus("errorMessage", "Header row 8 not found.");
        return;
    }

    const daily = {};

    for (let r = DATA_START; r < aoa.length; r++) {
        const row = aoa[r];
        if (!row || row.length < 7) continue;

        const modality = String(row[0] || "").trim();
        const locationFull = String(row[1] || "").trim();
        const dosRaw = row[5];

        const dos = fixDate(dosRaw);
        const facility = FACILITIES[locationFull];

        if (!dos || !facility || !modality) continue;

        if (!daily[dos]) daily[dos] = {};
        if (!daily[dos][facility]) daily[dos][facility] = {};

        daily[dos][facility][modality] =
            (daily[dos][facility][modality] || 0) + 1;
    }

    displayDailyTable(daily);
    showStatus("successMessage", "Completed! Tables generated below.");
}

/* ============================================================
   DISPLAY DAILY TABLE (matches your screenshot)
   ============================================================ */

function displayDailyTable(daily) {
    const left = document.getElementById("leftColumn");
    const right = document.getElementById("rightColumn");

    left.innerHTML = "";
    right.innerHTML = "";

    const table = document.createElement("table");

    /* HEADER ROW */
    let header = "<tr><th>Date</th>";

    Object.values(FACILITIES).forEach(fac => {
        MODALITIES.forEach(mod => {
            header += `<th>${fac} ${mod}</th>`;
        });
        header += `<th>${fac} Total</th>`;
    });

    header += "<th>Grand Total</th></tr>";
    table.innerHTML = header;

    /* DATA ROWS */
    Object.keys(daily)
        .sort((a, b) => new Date(a) - new Date(b))
        .forEach(dos => {
            let row = `<tr><td>${dos}</td>`;
            let grandTotal = 0;

            Object.values(FACILITIES).forEach(fac => {
                let facTotal = 0;

                MODALITIES.forEach(mod => {
                    const val =
                        daily[dos][fac] && daily[dos][fac][mod]
                            ? daily[dos][fac][mod]
                            : "";
                    row += `<td>${val}</td>`;
                    facTotal += Number(val || 0);
                });

                row += `<td>${facTotal}</td>`;
                grandTotal += facTotal;
            });

            row += `<td>${grandTotal}</td></tr>`;
            table.innerHTML += row;
        });

    left.appendChild(table);
}

/* ============================================================
   DOWNLOAD OUTPUT
   ============================================================ */

function downloadOutput() {
    alert("Daily tables are visual only.");
}
