// PDF download function
async function downloadPDF(elementId, filename) {
  const element = document.getElementById(elementId);

  const canvas = await html2canvas(element, {
    scale: 3,
    useCORS: true
  });

  const imgData = canvas.toDataURL("image/png");

  const pdf = new jspdf.jsPDF({
    orientation: "portrait",
    unit: "pt",
    format: "a4"
  });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const imgWidth = pageWidth - 40;
  const imgHeight = (canvas.height * imgWidth) / canvas.width;

  pdf.addImage(imgData, "PNG", 20, 20, imgWidth, imgHeight);
  pdf.save(filename);
}


// Fix Excel dates
function fixDate(v) {
  if (!v) return "";
  let d;

  if (typeof v === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    d = new Date(excelEpoch.getTime() + v * 86400000);
  } else {
    d = new Date(v);
  }

  if (isNaN(d)) return "";

  const month = d.getMonth() + 1;
  const day = d.getDate();
  const year = d.getFullYear();

  return (
    String(month).padStart(2, "0") + "/" +
    String(day).padStart(2, "0") + "/" +
    year
  );
}


// Convert DOS → “September 2, 2026      Wednesday”
function formatLongDate(dosString) {
  const d = new Date(dosString + "T00:00:00");

  const months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];

  const weekdays = [
    "Sunday","Monday","Tuesday","Wednesday",
    "Thursday","Friday","Saturday"
  ];

  const monthName = months[d.getMonth()];
  const day = d.getDate();
  const year = d.getFullYear();
  const weekday = weekdays[d.getDay()];

  return `${monthName} ${day}, ${year}      ${weekday}`;
}


// Location mapping
const locationMap = {
  "Astrana Breast Center": "ABC",
  "Diagnostic Medical Group Arcadia": "AR",
  "Diagnostic Medical Group City of Industry": "CI",
  "Diagnostic Medical Group Monterey Park": "MP",
  "Diagnostic Medical Group San Gabriel": "SG",
  "Synergy San Gabriel": "SSG"
};


// Backlog reference date
const referenceDate = new Date("2026-09-02T00:00:00");

function computeDaysBehind(dosString) {
  const dosDate = new Date(dosString + "T00:00:00");
  const diff = referenceDate - dosDate;
  return Math.floor(diff / 86400000);
}


// MAIN DAILY SUMMARY
async function runDailySummary() {

  try {
    const dailyFile = document.getElementById("dailyFile").files[0];
    if (!dailyFile) {
      alert("Please upload the RIS file.");
      return;
    }

    const buf = await dailyFile.arrayBuffer();
    const wb = XLSX.read(buf);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const totalSet = new Set();
    const reportedSet = new Set();
    const pendingSet = new Set();
    const noShowSet = new Set();

    const modalityCounts = {};
    const locationCounts = {};
    const backlog = {};

    // Extract Date of Service from RIS
    let serviceDate = null;

    for (let r = 8; r < aoa.length; r++) {
      const dos = fixDate(aoa[r][5]);
      if (dos) {
        serviceDate = dos;
        break;
      }
    }

    // Write formatted date header
    if (serviceDate) {
      const header = document.getElementById("dateHeader");
      if (header) {
        header.textContent = formatLongDate(serviceDate);
      }
    }

    // Process rows
    for (let r = 8; r < aoa.length; r++) {
      const apptID = String(aoa[r][6] || "").trim();
      const status = String(aoa[r][24] || "").trim();
      const rawLocation = String(aoa[r][1] || "").trim();
      const modality = String(aoa[r][0] || "").trim();
      const dos = fixDate(aoa[r][5]);

      if (!apptID) continue;

      if (status !== "Cancel") totalSet.add(apptID);
      if (status === "Reported") reportedSet.add(apptID);
      if (status === "TechComplete") pendingSet.add(apptID);
      if (status === "NoShow") noShowSet.add(apptID);

      const loc = locationMap[rawLocation] || rawLocation;
      if (!locationCounts[loc]) locationCounts[loc] = new Set();
      locationCounts[loc].add(apptID);

      if (!modalityCounts[modality]) modalityCounts[modality] = new Set();
      modalityCounts[modality].add(apptID);

      if (status === "TechComplete" && dos) {
        if (!backlog[dos]) backlog[dos] = new Set();
        backlog[dos].add(apptID);
      }
    }

    const locTotal = totalSet.size;

    // Summary by Location
    let locHTML = "<tr><th>Location</th><th>Procedures</th><th>%</th></tr>";
    Object.keys(locationCounts).forEach(loc => {
      const count = locationCounts[loc].size;
      const pct = ((count / locTotal) * 100).toFixed(2) + "%";
      locHTML += `<tr><td>${loc}</td><td>${count}</td><td>${pct}</td></tr>`;
    });
    locHTML += `<tr><td>Total</td><td>${locTotal}</td><td>100%</td></tr>`;
    document.getElementById("locTable").innerHTML = locHTML;


    // Summary by Modality
    let modHTML = "<tr><th>Modality</th><th>Procedures</th><th>%</th></tr>";
    Object.keys(modalityCounts).forEach(mod => {
      const count = modalityCounts[mod].size;
      const pct = ((count / locTotal) * 100).toFixed(2) + "%";
      modHTML += `<tr><td>${mod}</td><td>${count}</td><td>${pct}</td></tr>`;
    });
    modHTML += `<tr><td>Total</td><td>${locTotal}</td><td>100%</td></tr>`;
    document.getElementById("modTable").innerHTML = modHTML;


    // Total Reported / Pending Read
    let statusHTML = "<tr><th>Status</th><th>Count</th><th>%</th></tr>";
    const reportedPct = ((reportedSet.size / locTotal) * 100).toFixed(2) + "%";
    const pendingPct = ((pendingSet.size / locTotal) * 100).toFixed(2) + "%";

    statusHTML += `<tr><td>Total Reported</td><td>${reportedSet.size}</td><td>${reportedPct}</td></tr>`;
    statusHTML += `<tr><td>Pending Read</td><td>${pendingSet.size}</td><td>${pendingPct}</td></tr>`;
    document.getElementById("statusTable").innerHTML = statusHTML;


    // Backlog table
    let backlogHTML = "<tr><th>Date</th><th>Exams</th><th>Days Behind</th></tr>";

    Object.keys(backlog)
      .sort((a, b) => new Date(a) - new Date(b))
      .forEach(dos => {
        const count = backlog[dos].size;
        const daysBehind = computeDaysBehind(dos);
        backlogHTML += `<tr><td>${dos}</td><td>${count}</td><td>${daysBehind}</td></tr>`;
      });

    document.getElementById("backlogTable").innerHTML = backlogHTML;


  } catch (err) {
    alert("ERROR: " + err.message);
  }
}
