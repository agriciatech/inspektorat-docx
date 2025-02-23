const { PrismaClient } = require("@prisma/client");
const fs = require("fs");
const prisma = new PrismaClient();
const ExcelJS = require("exceljs");
const axios = require("axios");
const enpoint = require("../../config/url");

exports.setExcelLKE = async (req, res) => {
  let result = [];
  let average = {};
  const ALphabet = [
    "C",
    "D",
    "E",
    "F",
    "G",
    "H",
    "I",
    "J",
    "K",
    "L",
    "M",
    "N",
    "O",
    "P",
  ];
  try {
    await axios
      .get(`${enpoint}/api/v1/inspeksis`)
      .then(function (response) {
        result = response.data.data;
        average = response.data.yearData;
      })
      .catch(function (error) {
        console.log(error);
      });
    const resultTahun =
      await prisma.$queryRaw`SELECT * from "Tahun" ORDER BY tahun ASC`;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("LKE UTAMA");

    worksheet.getCell("A1").value = "Rekap Evaluasi Akip";
    worksheet.mergeCells(`A1:P1`);

    worksheet.getCell("A3").value = "No";
    worksheet.mergeCells("A3:A4");
    worksheet.getCell("B3").value = "Nama OPD";
    worksheet.mergeCells("B3:B4");

    let adding = 0;
    resultTahun.map((item, index) => {
      worksheet.getCell(`${ALphabet[index + adding]}3`).value = parseFloat(
        item.tahun
      );
      worksheet.mergeCells(
        `${ALphabet[index + adding]}3:${ALphabet[index + adding + 1]}3`
      );
      worksheet.getCell(`${ALphabet[index + adding]}4`).value = "Nilai";
      worksheet.getCell(`${ALphabet[index + adding + 1]}4`).value = "Kategori";
      adding += 1;
    });

    result.map((result, index) => {
      worksheet.addRow([index + 1, result.namaOPD]);
      let addingTahun = 0;
      const countWorksheet = worksheet.actualRowCount + 1;
      result.tahunData.map((data, index) => {
        worksheet.getCell(
          `${ALphabet[index + addingTahun]}${countWorksheet}`
        ).value = parseFloat(data.Inspeksi.nilai);
        if (data.Inspeksi.kategori == " ") {
          data.Inspeksi.kategori = "E";
        }
        worksheet.getCell(
          `${ALphabet[index + addingTahun + 1]}${countWorksheet}`
        ).value = data.Inspeksi.kategori;
        addingTahun += 1;
      });
    });

    worksheet.addRow(["Rata-Rata"]);
    const countWorksheet = worksheet.actualRowCount + 1;
    worksheet.mergeCells(`A${countWorksheet}:B${countWorksheet}`);

    let addingTahun = 0;
    average.map((data, index) => {
      worksheet.getCell(
        `${ALphabet[index + addingTahun]}${countWorksheet}`
      ).value = parseFloat(data.Average);

      worksheet.getCell(
        `${ALphabet[index + addingTahun + 1]}${countWorksheet}`
      ).value = "-";
      addingTahun += 1;
    });
    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        if (rowNumber === 3 || rowNumber === 4) {
          // Header row
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "521807" }, // Red color
          };
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // White font color
          cell.alignment = { vertical: "middle", horizontal: "center" };
        }
      });
    });

    worksheet.getColumn("A").width = 5;
    worksheet.getColumn("B").width = 50;
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        cell.alignment = { wrapText: true };
      });
    });
    worksheet.getColumn("A").alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    worksheet.getColumn("B").alignment = {
      vertical: "middle",
      horizontal: "left",
      wrapText: true,
    };

    worksheet.getCell("A1").alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };

    const borderStyle = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
      color: { argb: "000000" }, // Black color
    };

    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        cell.border = borderStyle;
      });
    });
    worksheet.getCell("A1").border = {};

    ALphabet.forEach((col) => {
      worksheet.getColumn(col).eachCell({ includeEmpty: true }, (cell) => {
        cell.alignment = { vertical: "middle", horizontal: "center" };
      });
    });

    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        if (rowNumber === 3 || rowNumber === 4 || rowNumber == countWorksheet) {
          // Header row
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "521807" }, // Red color
          };
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; // White font color
          cell.alignment = { vertical: "middle", horizontal: "center" };
        }
      });
    });
    // worksheet.addRow([index + 1, result.namaOPD]);

    const outputFile = `uploads/Rekapitulasi Evaluasi SAKIP.xlsx`;
    fs.access(outputFile, fs.constants.F_OK, (err) => {
      if (err) {
        if (err.code === "ENOENT") {
          console.log("File does not exist");
        } else {
          console.error("Error checking file existence:", err);
        }
      } else {
        // File exists, so unlink (delete) it
        fs.unlink(outputFile, (unlinkErr) => {
          if (unlinkErr) {
            console.error("Error deleting file:", unlinkErr);
          } else {
            console.log(`${outputFile} has been deleted successfully`);
          }
        });
      }
    });
    await workbook.xlsx.writeFile(outputFile);
    res.status(200).json({ data: outputFile });
  } catch (error) {
    console.error("Gagal membuat pengguna:", error);
    res.status(500).json({ error: "Gagal membuat pengguna" });
  }
};
