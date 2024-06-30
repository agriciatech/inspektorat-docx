const { PrismaClient } = require("@prisma/client");
const fs = require("fs");
const prisma = new PrismaClient();
const excelJS = require("exceljs");
const axios = require("axios"); // node

exports.setExcelLKEUtama = async (req, res) => {
  let result = [];
  let average = {};
  try {
    await axios
      .get("https://inspektorat-dev.agriciatech.com/api/v1/inspeksis")
      .then(function (response) {
        result = response.data.data;
        average = response.data.yearData;
      })
      .catch(function (error) {
        console.log(error);
      });

    var workbook = new excelJS.Workbook();
    var worksheet = workbook.getWorksheet("Rekap");

    worksheet.columns = [
      { header: "Id", key: "id", width: 10 },
      { header: "Name", key: "name", width: 32 },
      { header: "D.O.B.", key: "DOB", width: 10, outlineLevel: 1 },
    ];

    const idCol = worksheet.getColumn("id");
    const nameCol = worksheet.getColumn("B");
    const dobCol = worksheet.getColumn(3);

    dobCol.header = "Date of Birth";

    dobCol.header = ["Date of Birth", "A.K.A. D.O.B."];

    dobCol.key = "dob";

    dobCol.width = 15;

    dobCol.hidden = true;

    worksheet.getColumn(4).outlineLevel = 0;
    worksheet.getColumn(5).outlineLevel = 1;

    expect(worksheet.getColumn(4).collapsed).to.equal(false);
    expect(worksheet.getColumn(5).collapsed).to.equal(true);

    dobCol.eachCell(function (cell, rowNumber) {
      // ...
    });

    dobCol.eachCell({ includeEmpty: true }, function (cell, rowNumber) {
      // ...
    });

    // add a column of new values
    worksheet.getColumn(6).values = [1, 2, 3, 4, 5];

    // add a sparse column of values
    worksheet.getColumn(7).values = [, , 2, 3, , 5, , 7, , , , 11];

    worksheet.spliceColumns(3, 2);
    const newCol3Values = [1, 2, 3, 4, 5];
    const newCol4Values = ["one", "two", "three", "four", "five"];
    worksheet.spliceColumns(3, 1, newCol3Values, newCol4Values);
    const outputFile = `uploads/Rekapitulasi Evaluasi SAKIP.xlsx`;
    await workbook.xlsx.writeFile(outputFile);
    res.status(200).json({ data: outputFile });
  } catch (error) {
    console.error("Gagal membuat pengguna:", error);
    res.status(500).json({ error: "Gagal membuat pengguna" });
  }
};
