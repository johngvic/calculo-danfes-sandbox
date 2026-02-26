const xlsx = require("xlsx");
const fs = require("fs");
const PDFDocument = require("pdfkit");
const { execSync } = require('child_process');

// Command to parse sheet
// xlsx2csv -d "¦" job.xlsx job.csv

const getYear = (date) => {
  return date.split("/")[2];
};

const returnIfIsNumber = (value) => {
  return isNaN(value) ? 0 : Number(value);
};

const negativateValue = (value) => {
  return value * -1;
};

const formatCurrency = (value) => {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
  }).format(value);
};

function formatCNPJ(cnpj) {
  return cnpj.replace(
    /^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/,
    '$1.$2.$3/$4-$5'
  );
}

const calculate = (csvPath) => {
  // read the csv file
  const csvFilePath = csvPath;
  const csvData = fs.readFileSync(csvFilePath, "utf8");

  // handle csv data
  const rows = csvData.split("\n").map((row) => row.split("¦"));
  const headers = rows[0];

  const periodIndex = headers.indexOf("Período");
  const cfopFaturamentoIndex = headers.indexOf("CFOP Faturamento");
  const basePisIndex = headers.indexOf("Vlr Base Cálculo PIS");
  const basePisStfIndex = headers.indexOf("Vlr Base Cálculo PIS - STF");
  const pisIndex = headers.indexOf("Vlr Diferença PIS");
  const cofinsIndex = headers.indexOf("Vlr Diferença Cofins");
  const pisSelicIndex = headers.indexOf("Vlr SELIC S/PIS");
  const cofinsSelicIndex = headers.indexOf("Vlr SELIC S/Cofins");

  const register = {};

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    const date = rows[rowIndex][periodIndex];
    const isReturn =
      rows[rowIndex][cfopFaturamentoIndex] == "Devolução Faturamento";

    if (date) {
      const year = getYear(date);
      const basePisValue = returnIfIsNumber(rows[rowIndex][basePisIndex]);
      const basePisStfValue = returnIfIsNumber(rows[rowIndex][basePisStfIndex]);

      const baseValue = isReturn
        ? negativateValue(Number((basePisValue - basePisStfValue).toFixed(2)))
        : Number((basePisValue - basePisStfValue).toFixed(2));
      const pisValue = isReturn
        ? negativateValue(returnIfIsNumber(rows[rowIndex][pisIndex]))
        : returnIfIsNumber(rows[rowIndex][pisIndex]);
      const cofinsValue = isReturn
        ? negativateValue(returnIfIsNumber(rows[rowIndex][cofinsIndex]))
        : returnIfIsNumber(rows[rowIndex][cofinsIndex]);
      const pisSelicValue = isReturn
        ? negativateValue(returnIfIsNumber(rows[rowIndex][pisSelicIndex]))
        : returnIfIsNumber(rows[rowIndex][pisSelicIndex]);
      const cofinsSelicValue = isReturn
        ? negativateValue(returnIfIsNumber(rows[rowIndex][cofinsSelicIndex]))
        : returnIfIsNumber(rows[rowIndex][cofinsSelicIndex]);

      if (!register[year]) {
        register[year] = {
          baseValue,
          pisValue,
          cofinsValue,
          pisSelicValue,
          cofinsSelicValue,
        };
      } else {
        register[year].baseValue = Number(
          (register[year].baseValue + baseValue).toFixed(2),
        );
        register[year].pisValue = Number(
          (register[year].pisValue + pisValue).toFixed(2),
        );
        register[year].cofinsValue = Number(
          (register[year].cofinsValue + cofinsValue).toFixed(2),
        );
        register[year].pisSelicValue = Number(
          (register[year].pisSelicValue + pisSelicValue).toFixed(2),
        );
        register[year].cofinsSelicValue = Number(
          (register[year].cofinsSelicValue + cofinsSelicValue).toFixed(2),
        );
      }
    }
  }

  const entries = Object.entries(register);
  const totalBaseValue = entries.map(([_, obj]) => obj.baseValue);
  const totalPisValue = entries.map(([_, obj]) => obj.pisValue);
  const totalCofinsValue = entries.map(([_, obj]) => obj.cofinsValue);
  const totalPisSelicValue = entries.map(([_, obj]) => obj.pisSelicValue);
  const totalCofinsSelicValue = entries.map(([_, obj]) => obj.cofinsSelicValue);

  register[""] = {
    baseValue: totalBaseValue.reduce((acc, vl) => {
      return acc + vl;
    }, 0),
    pisValue: totalPisValue.reduce((acc, vl) => {
      return acc + vl;
    }, 0),
    cofinsValue: totalCofinsValue.reduce((acc, vl) => {
      return acc + vl;
    }, 0),
    pisSelicValue: totalPisSelicValue.reduce((acc, vl) => {
      return acc + vl;
    }, 0),
    cofinsSelicValue: totalCofinsSelicValue.reduce((acc, vl) => {
      return acc + vl;
    }, 0),
  };

  const data = [
    [
      "PERÍODO",
      "BASE CRÉDITO",
      "PIS",
      "COFINS",
      "PIS - SELIC",
      "COFINS - SELIC",
      "TOTAL CRÉDITO",
    ],
  ];

  for (const [year, values] of Object.entries(register)) {
    const total =
      values.pisValue +
      values.cofinsValue +
      values.pisSelicValue +
      values.cofinsSelicValue;

    data.push([
      year,
      values.baseValue.toFixed(2),
      values.pisValue.toFixed(2),
      values.cofinsValue.toFixed(2),
      values.pisSelicValue.toFixed(2),
      values.cofinsSelicValue.toFixed(2),
      total.toFixed(2),
    ]);
  }

  return { data, cnpj: rows[1][0] };
};

const generateXlsx = (data) => {
  const worksheet = xlsx.utils.aoa_to_sheet(data);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Resumo Anual");
  xlsx.writeFile(workbook, "resumo_anual.xlsx");
};

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

(async () => {
  execSync('xlsx2csv -d "¦" job.xlsx job.csv')
  const { data, cnpj } = calculate("job.csv");
  const response = await fetch(`https://api.opencnpj.org/${cnpj}`, { method: 'GET' });
  const { razao_social } = await response.json();
  // generateXlsx(dt);
  generatePdf({ data, cnpj, name: razao_social });
})();
