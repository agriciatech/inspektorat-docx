const { PrismaClient } = require("@prisma/client");
const fs = require("fs");
const prisma = new PrismaClient();
const excelJS = require("exceljs");

exports.setExcelLKE = async (req, res) => {
  try {
    const resultTahun = await prisma.$queryRaw`SELECT * from "Tahun" `;
    const result1 = await prisma.$queryRaw`SELECT * from "Inspeksi"`;
    const resultInspektur =
      await prisma.$queryRaw`SELECT * FROM "user" as a JOIN "Role" as b ON a.fk_role = b.pk_role_id WHERE role = 'OPD';`;

    const resultRInspeksiSubKomponen =
      await prisma.$queryRaw`SELECT SUM(nilai) as tot, fk_component,fk_inspeksi FROM "RInspeksiSubKomponen" as a JOIN "SubKomponen" as b ON a.fk_sub_component = b.id GROUP BY fk_component,fk_inspeksi ORDER BY fk_component`;
    console.log(result1);

    resultInspektur.map((index, number) => {
      result1.map((result, number1) => {
        resultRInspeksiSubKomponen.map((result2, number2) => {
          if (
            result.fk_tahun == 1 &&
            resultInspektur.fk_user == index.id &&
            result2.fk_inspeksi == result1.id
          ) {
            index.tahun1 = result2.tot;
          } else {
            index.tahun1 = 0;
          }
          if (
            result.fk_tahun == 2 &&
            resultInspektur.fk_user == index.id &&
            result2.fk_inspeksi == result1.id
          ) {
            index.tahun2 = result2.tot;
          } else {
            index.tahun2 = 0;
          }
          if (
            result.fk_tahun == 3 &&
            resultInspektur.fk_user == index.id &&
            result2.fk_inspeksi == result1.id
          ) {
            index.tahun3 = result2.tot;
          } else {
            index.tahun3 = 0;
          }
          if (
            result.fk_tahun == 4 &&
            resultInspektur.fk_user == index.id &&
            result2.fk_inspeksi == result1.id
          ) {
            index.tahun4 = result2.tot;
          } else {
            index.tahun4 = 0;
          }
          if (
            result.fk_tahun == 5 &&
            resultInspektur.fk_user == index.id &&
            result2.fk_inspeksi == result1.id
          ) {
            index.tahun5 = result2.tot;
          } else {
            index.tahun5 = 0;
          }
          if (
            result.fk_tahun == 6 &&
            resultInspektur.fk_user == index.id &&
            result2.fk_inspeksi == result1.id
          ) {
            index.tahun6 = result2.tot;
          } else {
            index.tahun6 = 0;
          }
          if (
            result.fk_tahun == 7 &&
            resultInspektur.fk_user == index.id &&
            result2.fk_inspeksi == result1.id
          ) {
            index.tahun7 = result2.tot;
          } else {
            index.tahun7 = 0;
          }
        });
      });
      index.no = number + 1;
    });

    const workbook = new excelJS.Workbook();
    const worksheet = workbook.addWorksheet("My Users");
    worksheet.columns = [
      { header: "No", key: "no", width: 10 },
      { header: "OPD", key: "name", width: 50 },
      { header: "2020", key: "tahun1", width: 10 },
      { header: "2021", key: "tahun2", width: 10 },
      { header: "2022", key: "tahun3", width: 10 },
      { header: "2023", key: "tahun4", width: 10 },
      { header: "2024", key: "tahun5", width: 10 },
      { header: "2025", key: "tahun6", width: 10 },
      { header: "2026", key: "tahun7", width: 10 },
    ];
    // Looping through User data
    resultInspektur.forEach((user) => {
      worksheet.addRow(user);
    });
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true };
    });

    const outputFile = `uploads/Rekapitulasi Evaluasi SAKIP.xlsx`;
    await workbook.xlsx.writeFile(outputFile);
    res.status(200).json({ data: outputFile });
  } catch (error) {
    console.error("Gagal membuat pengguna:", error);
    res.status(500).json({ error: "Gagal membuat pengguna" });
  }
};
