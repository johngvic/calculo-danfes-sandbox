const xlsx = require("xlsx");
const fs = require("fs");
const PDFDocument = require("pdfkit");
const { execSync } = require('child_process');

// Command to parse sheet
// xlsx2csv -d "¦" job.xlsx job.csv

function formatCNPJ(cnpj) {
  return cnpj.replace(
    /^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/,
    '$1.$2.$3/$4-$5'
  );
}

const generatePdf = ({ data, cnpj, name }) => {
  // Generate PDF with formatted table
  const doc = new PDFDocument({
    margin: 40,
    size: "A4",
  });

  doc.pipe(fs.createWriteStream(`${cnpj}.pdf`));

  doc
    .fontSize(12)
    .font("Helvetica-Bold")
    .text(`CLIENTE: `, { continued: true, align: "left" })
    .font("Helvetica")
    .text(name)
    .moveDown(0.3);

  doc
    .fontSize(12)
    .font("Helvetica-Bold")
    .text('CNPJ: ', { continued: true, align: "left" })
    .font("Helvetica")
    .text(formatCNPJ(cnpj))
    .moveDown(0.3);

  doc
    .fontSize(12)
    .font("Helvetica-Bold")
    .text('TRABALHO: ', { continued: true, align: "left" })
    .font("Helvetica")
    .text("EXCLUSÃO PIS/COFINS PRÓPRIA BASE")
    .moveDown(1);

  doc.moveDown(1);

  // Table
  const colWidths = [48, 80, 80, 80, 80, 80, 80];
  const cellHeight = 25;
  const headerHeight = 24;
  const margin = 35;
  let currentY = doc.y;

  // Header background
  doc.fillColor("#D9D9D9");
  doc.rect(margin, currentY, doc.page.width - 2 * margin, headerHeight).fill();

  // Header text
  doc.fillColor("#000000").fontSize(8).font("Helvetica-Bold");
  let currentX = margin + 5;
  data[0].forEach((header, index) => {
    doc.text(header, currentX, currentY + 8, {
      width: colWidths[index] - 10,
      align: "center",
    });
    currentX += colWidths[index];
  });

  currentY += headerHeight;

  // Rows
  doc.fillColor("#000000").font("Helvetica");
  data.slice(1).forEach((row, rowIndex) => {
    // Alternate row colors
    if (rowIndex % 2 === 0) {
      doc
        .fillColor("#F2F2F2")
        .rect(margin, currentY, doc.page.width - 2 * margin, cellHeight)
        .fill();
    }

    doc.fillColor("#000000").fontSize(8);
    currentX = margin + 5;
    row.forEach((cell, colIndex) => {
      const value = colIndex === 0 ? cell : formatCurrency(Number(cell));
      doc.text(value, currentX, currentY + 10, {
        width: colWidths[colIndex] - 10,
        align: "center",
      });
      currentX += colWidths[colIndex];
    });

    currentY += cellHeight;
  });

  // position footer a bit below the last drawn table row
  const footerY = currentY + 50;

  doc
    .fontSize(10)
    .font("Helvetica-Oblique")
    .fillColor("#666666")
    .text(`Gerado em: ${new Date().toLocaleDateString("pt-BR")}`, margin, footerY, {
      align: "left",
    });

  // Finalize
  doc.end();
};

const inspect = (file) => {
  // read the csv file
  const csvFilePath = `${file}.csv`;
  const csvData = fs.readFileSync(csvFilePath, "utf8");

  // handle csv data
  const rows = csvData.split("\n").map((row) => row.split("¦"));
  const headers = rows[0];

  const customerNameIndex = headers.indexOf("Nome do cliente/fornecedor");
  const customerCodeIndex = headers.indexOf("Código do cliente/fornecedor");
  const rpsIndex = headers.indexOf("Nº RPS");
  const nfIndex = headers.indexOf("Nº NF");
  const installmentIndex = headers.indexOf("Prestação");

  const result = [];

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    const name = rows[rowIndex][customerNameIndex];
    const code = rows[rowIndex][customerCodeIndex];
    const rps = rows[rowIndex][rpsIndex];
    const nf = rows[rowIndex][nfIndex];
    const installment = rows[rowIndex][installmentIndex];
    const [current, _, last] = installment?.split(" ") || [null, null, null];

    // if (current) {
    //   result.push({
    //     name,
    //     code,
    //     rps,
    //     nf,
    //     installment,
    //     origin: file.replace(/(\d{2})(\d{4})/, '$1/$2')
    //   })
    // }

    if (current && (current != last)) {
      result.push({
        name,
        code,
        rps,
        nf,
        installment,
        origin: file.replace(/(\d{2})(\d{4})/, '$1/$2')
      })
    }
  }

  fs.writeFile(`./payloads/${file}_PENDING.json`, JSON.stringify(result, null, 2), () => { })
}

(async () => {
  const files = ["092025", "102025", "112025", "122025", "012026"];

  for (const file of files) {
    inspect(file);
  }
})();
