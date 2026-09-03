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
  const pageHeight = pdf.internal.pageSize.getHeight();

  const imgWidth = pageWidth - 40;
  const imgHeight = (canvas.height * imgWidth) / canvas.width;

  pdf.addImage(imgData, "PNG", 20, 20, imgWidth, imgHeight);
  pdf.save(filename);
}
 
