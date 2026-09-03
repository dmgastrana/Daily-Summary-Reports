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

// Utility: fix Excel dates
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
  return (
    String(d.getUTCMonth() + 1).padStart(2, "0") + "/" +
    String(d.getUTCDate()).padStart(2, "0") + "/" +
    d.getUTCFullYear()
  );
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
const referenceDate = new Date("2026-09-02");

function computeDaysBehind(dosString) {
  const dosDate = new Date(dosString);
  if (isNaN(dosDate)) return "";
  const diff = referenceDate - dosDate;
  return Math.floor(diff / 86400000);
}

// MAIN DAILY SUMMARY
async function runDailySummary() {
  const summary = document.getElementById("summary");
  summary.textContent = "Processing…";

  try {
    const dailyFile = document.getElementById("dailyFile").files[0];
    if (!dailyFile) {
      summary.textContent = "ERROR: Please upload the RIS file.";
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

    let out = "";

    out += "Total Exams: " + totalSet.size + "\n";
    out += "Total Reported: " + reportedSet.size + "\n";
    out += "Pending Read: " + pendingSet.size + "\n";
    out += "No-Show: " + noShowSet.size + "\n\n";

    out += "Summary by Location:\n";
    Object.keys(locationCounts).forEach(loc => {
      out += loc + ": " + locationCounts[loc].size + "\n";
    });
    out += "\n";

    out += "Summary by Modality:\n";
    Object.keys(modalityCounts).forEach(mod => {
      out += mod + ": " + modalityCounts[mod].size + "\n";
    });
    out += "\n";

    out += "Backlog:\n";
    Object.keys(backlog).forEach(dos => {
      const count = backlog[dos].size;
      const daysBehind = computeDaysBehind(dos);
      out += dos + " → " + count + " exams → " + daysBehind + " days behind\n";
    });

    summary.textContent = out;

  } catch (err) {
    summary.textContent = "ERROR: " + err.message;
  }
}
