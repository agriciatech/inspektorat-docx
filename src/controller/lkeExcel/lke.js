const { PrismaClient } = require("@prisma/client");
const fs = require("fs");
const prisma = new PrismaClient();
const ExcelJS = require("exceljs");
const { default: axios } = require("axios");

exports.setLKEExcel = async (req, res) => {
  let result = [];
  let average = {};
  try {
    const { id } = req.params;

    const result1 =
      await prisma.$queryRaw`SELECT * from "Inspeksi" WHERE id = ${parseInt(
        id
      )}`;
    const resultTahun =
      await prisma.$queryRaw`SELECT * from "Tahun" WHERE pk_tahun_id = ${parseInt(
        result1[0].fk_tahun
      )}`;
    const resultInspektur =
      await prisma.$queryRaw`SELECT * FROM "user" WHERE id = ${parseInt(
        result1[0].fk_user
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

    const resultRInspeksiSubKomponen =
      await prisma.$queryRaw`SELECT SUM(nilai), fk_component FROM "RInspeksiSubKomponen" as a JOIN "SubKomponen" as b ON a.fk_sub_component = b.id  WHERE fk_inspeksi = ${parseInt(
        id
      )} GROUP BY fk_component ORDER BY fk_component`;

    let total = 0;
    console.log(resultRInspeksiSubKomponen);
    resultRInspeksiSubKomponen.map((item, index) => {
      total += parseFloat(item.sum);
    });
    const tahun = resultTahun[0].tahun;
    const user = resultInspektur[0].name;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("LKE");

    worksheet.getCell("A1").value = "LEMBAR KERJA EVALUASI";

    // Merge cells A1:C1
    worksheet.mergeCells("A1:G1");
    worksheet.getCell("A1").alignment = { horizontal: "center" };
    worksheet.addRow();

    worksheet.getCell("A3").value = "No";
    worksheet.mergeCells("A3:A4");
    worksheet.getCell("B3").value = "Komponen/Sub Komponen/Kriteria";
    worksheet.mergeCells("B3:B4");
    worksheet.getCell("C3").value = "Bobot";
    worksheet.mergeCells("C3:C4");
    worksheet.getCell("D3").value = "Unit/Satker";
    worksheet.mergeCells("D3:E3");
    worksheet.getCell("D4").value = "Jawaban";
    worksheet.getCell("E4").value = "Nilai";
    worksheet.getCell("F3").value = "Status";
    worksheet.mergeCells("F3:F4");
    worksheet.getCell("G3").value = "Catatan";
    worksheet.mergeCells("G3:G4");
    const Alphabet = ["A", "B", "C", "D", "E", "F", "G"];

    for (let i = 0; i < Alphabet.length; i++) {
      for (let j = 3; j <= 4; j++) {
        worksheet.getCell(`${Alphabet[i]}${j}`).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "521807" }, // Red color
        };
        worksheet.getCell(`${Alphabet[i]}${j}`).font = {
          bold: true,
          color: { argb: "FFFFFFFF" },
        }; // White font color
        worksheet.getCell(`${Alphabet[i]}${j}`).alignment = {
          vertical: "middle",
          horizontal: "center",
        };
      }
    }
    result.map((component, index) => {
      console.log(component);
      worksheet.addRow([
        parseFloat(component.nomor),
        component.component,
        parseFloat(component.bobot),
        "",
        parseFloat(resultRInspeksiSubKomponen[index].sum),
      ]);
      let headerAktual = worksheet.actualRowCount + 1;
      for (let i = 0; i < Alphabet.length; i++) {
        worksheet.getCell(`${Alphabet[i]}${headerAktual}`).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FACFC2" }, // Red color
        };
        worksheet.getCell(`${Alphabet[i]}${headerAktual}`).font = {
          bold: true,
          color: { argb: "000000" },
        }; // White font color
        worksheet.getCell(`${Alphabet[i]}${headerAktual}`).alignment = {
          vertical: "middle",
          horizontal: "center",
        };
      }
      component.SubKomponen.map((subKomponen, subKomponenIndex) => {
        worksheet.addRow([
          subKomponen.nomor,
          subKomponen.nama,
          parseFloat(subKomponen.bobot),
          subKomponen.Penilaian.jawaban,
          parseFloat(subKomponen.Penilaian.nilai),
        ]);
        let headerAktual = worksheet.actualRowCount + 1;
        for (let i = 0; i < Alphabet.length; i++) {
          worksheet.getCell(`${Alphabet[i]}${headerAktual}`).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFCC" }, // Red color
          };
          worksheet.getCell(`${Alphabet[i]}${headerAktual}`).font = {
            bold: true,
            color: { argb: "000000" },
          }; // White font color
          worksheet.getCell(`${Alphabet[i]}${headerAktual}`).alignment = {
            vertical: "middle",
            horizontal: "center",
          };
        }
        worksheet.addRow(["Keriteria:"]);
        let nowCount = worksheet.actualRowCount + 1;
        worksheet.mergeCells(`A${nowCount}:G${nowCount}`);
        subKomponen.Keriteria.map((keriteria, indexinspeksi) => {
          worksheet.addRow([
            parseInt(keriteria.nomor),
            keriteria.keriteria,
            "",
            "",
            "",
            keriteria.Catatan.verifikasi,
            keriteria.Catatan.catatan,
          ]);
          let nowCount = worksheet.actualRowCount + 1;
          worksheet.mergeCells(`B${nowCount}:E${nowCount}`);

          // console.log(inspeksi.Catatan.catatan);
        });
      });
    });

    let headerAktual = worksheet.actualRowCount + 2;
    for (let i = 0; i < Alphabet.length; i++) {
      worksheet.getCell(`${Alphabet[i]}${headerAktual}`).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "521807" }, // Red color
      };
      worksheet.getCell(`${Alphabet[i]}${headerAktual}`).font = {
        bold: true,
        color: { argb: "000000" },
      }; // White font color
      worksheet.getCell(`${Alphabet[i]}${headerAktual}`).alignment = {
        vertical: "middle",
        horizontal: "center",
      };
    }

    worksheet.getColumn("A").width = 5;
    worksheet.getColumn("B").width = 50;
    worksheet.getColumn("G").width = 50;
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        cell.alignment = { wrapText: true };
      });
    });

    worksheet.getColumn("C").alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    worksheet.getColumn("D").alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    worksheet.getColumn("E").alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
    worksheet.getColumn("A").alignment = {
      vertical: "middle",
      horizontal: "left",
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

    const outputFile = `uploads/LKE ${user} ${tahun}.xlsx`;
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
