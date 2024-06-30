const { PrismaClient } = require("@prisma/client");
const fs = require("fs");
const prisma = new PrismaClient();
const axios = require("axios");
const xlsx = require("xlsx");

exports.setLKEExcel = async (req, res) => {
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

    const workbook = xlsx.read("../../../uploads/sakipuntuk1.xlsx");
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json(sheet);
    const outputFile = `uploads/Rekapitulasi Evaluasi SAKIP.xlsx`;
    await workbook.xlsx.writeFile(outputFile);
    res.status(200).json({ data: outputFile });
  } catch (error) {
    console.error("Gagal membuat pengguna:", error);
    res.status(500).json({ error: "Gagal membuat pengguna" });
  }
};
