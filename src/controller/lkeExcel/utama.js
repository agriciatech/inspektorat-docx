const { PrismaClient } = require("@prisma/client");
const fs = require("fs");
const prisma = new PrismaClient();
const ExcelJS = require("exceljs");
const axios = require("axios");
const { start } = require("repl");
const { response } = require("express");

exports.setExcelLKEUtama = async (req, res) => {
  const result = [];
  const resultRekomendasi = [];
  let dataInspeksi = {};
  try {
    const { id } = req.params;

    const result1 =
      await prisma.$queryRaw`SELECT * from "Inspeksi" WHERE id = ${parseInt(
        id
      )}`;

    await axios
      .get(
        `https://inspektorat-be.agriciatech.com/api/v1/components/${1}/${id}`
      )
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });

    await axios
      .get(
        `https://inspektorat-be.agriciatech.com/api/v1/components/${2}/${id}`
      )
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });
    await axios
      .get(
        `https://inspektorat-be.agriciatech.com/api/v1/components/${3}/${id}`
      )
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });
    await axios
      .get(
        `https://inspektorat-be.agriciatech.com/api/v1/components/${4}/${id}`
      )
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });

    await axios
      .get(
        `https://inspektorat-be.agriciatech.com/api/v1/rekomendasi?inspeksi=${id}`
      )
      .then(function (response) {
        resultRekomendasi.push(response.data.data);
      })
      .catch(function (error) {
        console.log(error);
      });

    await axios
      .get(
        `https://inspektorat-be.agriciatech.com/api/v1/inspeksis/${parseInt(
          result1[0].fk_user
        )}/${parseInt(result1[0].fk_tahun)}`
      )
      .then(function (response) {
        dataInspeksi = response.data;
      })
      .catch(function (error) {
        console.log(error);
      });
    const predikat = dataInspeksi.kategori;

    console.log(parseInt(result1[0].fk_tahun));

    const resultTahun =
      await prisma.$queryRaw`SELECT * from "Tahun" WHERE pk_tahun_id = ${parseInt(
        result1[0].fk_tahun
      )}`;
    const resultInspektur =
      await prisma.$queryRaw`SELECT * FROM "user" WHERE id = ${parseInt(
        result1[0].fk_user
      )}`;

    const resultComponent =
      await prisma.$queryRaw`SELECT * FROM "Component" ORDER BY nomor`;

    // const resultRekomendasi =
    //   await prisma.$queryRaw`SELECT * FROM "RInspeksiKomponen" WHERE fk_inspeksi = ${parseInt(
    //     id
    //   )} `;

    const resultRInspeksiSubKomponen =
      await prisma.$queryRaw`SELECT SUM(nilai), fk_component FROM "RInspeksiSubKomponen" as a JOIN "SubKomponen" as b ON a.fk_sub_component = b.id  WHERE fk_inspeksi = ${parseInt(
        id
      )} GROUP BY fk_component ORDER BY fk_component`;

    let total = 0;
    result.map((item, index) => {
      total += parseFloat(item.nilai);
    });
    const tahun = resultTahun[0].tahun;
    const user = resultInspektur[0].name;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("LKE UTAMA");

    worksheet.getCell("A1").value = "HASIL EVALUASI AKUNTABILITAS KINERJA";
    worksheet.getCell("A2").value = `${user} Lampung Selatan`;
    worksheet.getCell("A3").value = `Tahun ${tahun}`;

    // Merge cells A1:C1
    worksheet.mergeCells("A1:D1");
    worksheet.mergeCells("A2:D2");
    worksheet.mergeCells("A3:D3");
    worksheet.getCell("A1").alignment = { horizontal: "center" };
    worksheet.getCell("A2").alignment = { horizontal: "center" };
    worksheet.getCell("A3").alignment = { horizontal: "center" };

    worksheet.addRow();
    const headers = [
      "No",
      "Komponen/Sub Komponen/Kriteria",
      "Bobot",
      `Nilai ${tahun}`,
    ];

    // Add headers to the worksheet
    worksheet.addRow(headers);

    // Add some sample data
    const dataComponent = [];

    resultComponent.map((item, index) =>
      dataComponent.push([
        parseFloat(item.nomor),
        item.component,
        parseFloat(item.bobot),
        parseFloat(resultRInspeksiSubKomponen[index].sum),
      ])
    );

    // Add rows to the worksheet
    dataComponent.forEach((row) => {
      worksheet.addRow(row);
    });

    worksheet.addRow([
      "Nilai Akuntabilitas Kinerja",
      "",
      "",
      parseFloat(total),
    ]);
    const headerAktual = worksheet.actualRowCount + 1;
    worksheet.getCell(`A${headerAktual}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "521807" }, // Red background color
    };
    worksheet.getCell(`A${headerAktual}`).font = {
      bold: true,
      color: { argb: "FFFFFFFF" },
    };
    worksheet.mergeCells(`A${headerAktual}:C${headerAktual}`);
    worksheet.addRow(["", "", "", predikat]);
    worksheet.addRow();
    const headers2 = ["No", "Catatan"];

    // Add headers to the worksheet
    worksheet.addRow(headers2);
    const headerAktual2 = worksheet.actualRowCount + 2;
    worksheet.mergeCells(`B${headerAktual2}:D${headerAktual2}`);

    let nomorStart = 0;
    result.map((item, index) => {
      item.SubKomponen.map((subkomponen, index) => {
        subkomponen.Keriteria.map((test, index) => {
          nomorStart++;
          worksheet.addRow([nomorStart, test.Catatan.catatan]);
          const row = worksheet.actualRowCount + 1;
          // worksheet.mergeCells(`B${row}:D${row}`);
        });
      });
    });
    // const catatanComponent = [];

    // resultInspeksiKriteria.forEach((resultRekomendasi, index) => {
    //   catatanComponent.push([index + 1, resultRekomendasi.catatan]);
    // });
    // catatanComponent.forEach((row) => {
    //   worksheet.addRow(row);
    //   const test = worksheet.actualRowCount + 1;
    //   // worksheet.mergeCells(`B${test}:D${test}`);
    // });

    worksheet.addRow();
    const headers3 = ["No", "Rekomendasi"];

    // Add headers to the worksheet
    worksheet.addRow(headers3);
    const headerAktual3 = worksheet.actualRowCount + 3;
    worksheet.mergeCells(`B${headerAktual3}:D${headerAktual3}`);

    nomorStart = 0;
    resultRekomendasi[0].forEach((resultRekomendasiMap, indexfirst) => {
      resultRekomendasiMap.Rekomendasi.map((result, index) => {
        nomorStart++;
        worksheet.addRow([nomorStart, result.rekomendasi]);
        const test = worksheet.actualRowCount + 1;
        // worksheet.mergeCells(`B${test}:D${test}`);
      });
    });

    worksheet.addRow();

    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        if (
          rowNumber === 5 ||
          rowNumber === headerAktual2 ||
          rowNumber === headerAktual3
        ) {
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

    const centerColumns = ["A", "C", "D"];
    centerColumns.forEach((col) => {
      worksheet.getColumn(col).eachCell({ includeEmpty: true }, (cell) => {
        cell.alignment = { vertical: "middle", horizontal: "center" };
      });
    });
    worksheet.getColumn("A").width = 5;
    worksheet.getColumn("B").width = 50;
    worksheet.getColumn("B").alignment = {
      vertical: "middle",
      horizontal: "left",
    };
    worksheet.getCell("B1").alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getCell("B2").alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getCell("B3").alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    // Apply formatting (optional)
    // worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    //   row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
    //     cell.alignment = { vertical: "middle", horizontal: "center" };
    //     cell.border = {
    //       top: { style: "thin" },
    //       left: { style: "thin" },
    //       bottom: { style: "thin" },
    //       right: { style: "thin" },
    //     };
    //   });
    // });

    // Auto-size columns (optional)

    const outputFile = `uploads/LKE Utama Evaluasi SAKIP ${user} ${tahun}.xlsx`;

    await workbook.xlsx.writeFile(outputFile);
    res.status(200).json({ data: outputFile });
  } catch (error) {
    console.error("Gagal membuat pengguna:", error);
    res.status(500).json({ error: "Gagal membuat pengguna" });
  }
};
