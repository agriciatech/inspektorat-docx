const { PrismaClient } = require("@prisma/client");
const {
  AlignmentType,
  Document,
  Footer,
  Header,
  Packer,
  PageBreak,
  PageNumber,
  SectionType,
  NumberFormat,
  Paragraph,
  TextRun,
  LevelFormat,
  HeadingLevel,
  UnderlineType,
  Tab,
  VerticalAlign,
  Table,
  TabStopType,
  TextDirection,
  TableRow,
  TableCell,
  WidthType,
  TabStopPosition,
  BorderStyle,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  PageBorderZOrder,
  convertInchesToTwip,
} = require("docx");
const fs = require("fs");
const prisma = new PrismaClient();
const axios = require("axios");
const { text } = require("express");
const enpoint = require("../../config/url");

function cmToTwip(cm) {
  return Math.round(cm * 28.35 * 20); // Convert cm to points and then to twentieths of a point
}

exports.formulir10 = async (req, res) => {
  try {
    const { id } = req.params;
    const result = [];
    const resultRekomendasi = [];
    let dataInspeksi = {};

    const result1 =
      await prisma.$queryRaw`SELECT * from "Inspeksi" WHERE id = ${parseInt(
        id
      )}`;

    await axios
      .get(`${enpoint}/api/v1/components/${1}/${id}`)
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });

    await axios
      .get(`${enpoint}/api/v1/components/${2}/${id}`)
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });
    await axios
      .get(`${enpoint}/api/v1/components/${3}/${id}`)
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });
    await axios
      .get(`${enpoint}/api/v1/components/${4}/${id}`)
      .then(function (response) {
        result.push(response.data);
      })
      .catch(function (error) {
        console.log(error);
      });

    await axios
      .get(`${enpoint}/api/v1/rekomendasi?inspeksi=${id}`)
      .then(function (response) {
        resultRekomendasi.push(response.data.data);
      })
      .catch(function (error) {
        console.log(error);
      });

    await axios
      .get(
        `${enpoint}/api/v1/inspeksis/${parseInt(result1[0].fk_user)}/${parseInt(
          result1[0].fk_tahun
        )}`
      )
      .then(function (response) {
        dataInspeksi = response.data;
      })
      .catch(function (error) {
        console.log(error);
      });
    const resultTahun =
      await prisma.$queryRaw`SELECT * from "Tahun" WHERE pk_tahun_id = ${parseInt(
        result1[0].fk_tahun
      )}`;
    const resultInspektur =
      await prisma.$queryRaw`SELECT * FROM "user" WHERE id = ${parseInt(
        result1[0].fk_user
      )}`;

    const predikat = dataInspeksi.kategori;

    const resultComponent =
      await prisma.$queryRaw`SELECT * FROM "Component" ORDER BY nomor`;
    const resultInspeksiKriteria =
      await prisma.$queryRaw`SELECT * FROM "RInspeksiKriteria" as r JOIN "Keriteria" as k ON r.fk_keriteria = k.id JOIN "SubKomponen" as s ON k.fk_komponen = s.id WHERE fk_inspeksi = ${parseInt(
        id
      )} AND NOT verifikasi like 'Sesuai' AND NOT catatan = '' AND NOT catatan = '-' ;`;

    const resultRInspeksiSubKomponen =
      await prisma.$queryRaw`SELECT SUM(nilai), fk_component FROM "RInspeksiSubKomponen" as a JOIN "SubKomponen" as b ON a.fk_sub_component = b.id  WHERE fk_inspeksi = ${parseInt(
        id
      )} GROUP BY fk_component ORDER BY fk_component`;

    let total = 0;
    resultRInspeksiSubKomponen.map((item, index) => {
      total += parseFloat(item.sum);
    });
    const tahun = resultTahun[0].tahun;
    const user = resultInspektur[0].name;
    const margintop = cmToTwip(2.5);
    const marginleft = cmToTwip(3);
    const marginright = cmToTwip(3);
    const marginbottom = cmToTwip(2.5);

    const generateRows = (allData, subKomponen) =>
      allData.map(
        (item, index) =>
          new TableRow({
            children: [
              new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    text: item.nomor.toString(),
                    alignment: AlignmentType.CENTER,
                  }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    text: item.component,
                  }),
                ],
                verticalAlign: VerticalAlign.CENTER,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    text: item.bobot.toString(),
                    alignment: AlignmentType.CENTER,
                  }),
                ],
                verticalAlign: VerticalAlign.CENTER,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    text: subKomponen[index].sum.toFixed(2).toString(),
                    alignment: AlignmentType.CENTER,
                  }),
                ],
                verticalAlign: VerticalAlign.CENTER,
              }),
            ],
          })
      );

    const doc = new Document({
      styles: {
        default: {
          heading1: {
            run: {
              size: 28,
              bold: true,
              italics: true,
              color: "FF0000",
            },
            paragraph: {
              spacing: {
                after: 120,
              },
            },
          },

          listParagraph: {
            run: {
              size: "11pt",
              font: "Bookman Old Style",
            },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
            },
          },
          document: {
            run: {
              size: "11pt",
              font: "Bookman Old Style",
            },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
            },
          },
        },
      },
      paragraphStyles: [
        {
          id: "aside",
          name: "Aside",
          basedOn: "Normal",
          next: "Normal",
          run: {
            size: "11pt",
            font: "Bookman Old Style",
          },
          paragraph: {
            alignment: AlignmentType.JUSTIFIED,
            indent: {
              left: convertInchesToTwip(0.5),
            },
            spacing: { line: 360 },
          },
        },
        {
          id: "wellSpaced",
          name: "Well Spaced",
          basedOn: "Normal",
          quickFormat: true,
          paragraph: {
            spacing: {
              line: 276,
              before: 20 * 72 * 0.1,
              after: 20 * 72 * 0.05,
            },
          },
        },
        {
          id: "strikeUnderline",
          name: "Strike Underline",
          basedOn: "Normal",
          quickFormat: true,
          run: {
            strike: true,
            underline: {
              type: UnderlineType.SINGLE,
            },
          },
        },
      ],
      characterStyles: [
        {
          id: "strikeUnderlineCharacter",
          name: "Strike Underline",
          basedOn: "Normal",
          quickFormat: true,
          run: {
            strike: true,
            underline: {
              type: UnderlineType.SINGLE,
            },
          },
        },
      ],

      numbering: {
        config: [
          {
            reference: "numbering-format",
            levels: [
              {
                level: 0,
                format: LevelFormat.DECIMAL,
                text: "%1",
                alignment: AlignmentType.START,
                style: {
                  paragraph: {
                    alignment: AlignmentType.JUSTIFIED,
                    indent: {
                      left: convertInchesToTwip(0.2),
                      hanging: convertInchesToTwip(0.18),
                    },
                  },
                },
              },
              {
                level: 1,
                format: LevelFormat.UPPER_LETTER,
                text: "%2.",
                alignment: AlignmentType.UPPER_LETTER,
                style: {
                  paragraph: {
                    indent: {
                      left: convertInchesToTwip(0.4),
                      hanging: convertInchesToTwip(0.18),
                    },
                  },
                },
              },
              {
                level: 2,
                format: LevelFormat.LOWER_LETTER,
                text: "%3.",
                alignment: AlignmentType.START,
                style: {
                  paragraph: {
                    indent: {
                      left: convertInchesToTwip(0.6),
                      hanging: convertInchesToTwip(0.2),
                    },
                  },
                },
              },
              {
                level: 3,
                format: LevelFormat.LOWER_LETTER,
                text: "%4.",
                alignment: AlignmentType.START,
                style: {
                  paragraph: {
                    indent: { left: 2880, hanging: 2420 },
                  },
                },
              },
            ],
          },
        ],
      },
      sections: [
        {
          margins: {
            top: margintop,
            right: marginright,
            bottom: marginbottom,
            left: marginleft,
          },
          properties: {
            titlePage: true,
            page: {
              pageNumbers: {
                start: 1,
                formatType: NumberFormat.DECIMAL,
              },
            },

            type: SectionType.CONTINUOUS,
          },
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      children: [PageNumber.CURRENT],
                    }),
                  ],
                }),
              ],
            }),
            first: new Header({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      children: [],
                    }),
                  ],
                }),
              ],
            }),
          },
          children: [
            // new Paragraph({
            //   border: {
            //     top: {
            //       color: "auto",
            //       space: 1,
            //       value: BorderStyle.SINGLE,
            //       size: 6,
            //     },
            //     bottom: {
            //       color: "auto",
            //       space: 1,
            //       value: BorderStyle.SINGLE,
            //       size: 6,
            //     },
            //     left: {
            //       color: "auto",
            //       space: 1,
            //       value: BorderStyle.SINGLE,
            //       size: 6,
            //     },
            //   },
            //   spacing: {
            //     before: 200, // Space before the paragraph with the border
            //     after: 200, // Space after the paragraph with the border
            //   },
            //   margin: {
            //     left: 400, // Space inside the border on the left
            //   },
            // }),
            new Paragraph({
              text: "PERNYATAAN PERJANJIAN KINERJA ",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: "TINGKAT UNIT KERJA/SKPD/SATUAN KERJA",
              alignment: AlignmentType.CENTER,
            }),

            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: "-Logo Lembaga",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: "PERJANJIAN KINERJA TAHUN ..........................",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: " ",
            }),

            new Paragraph({
              text: `Dalam rangka mewujudkan manajemen pemerintahan yang efektif, transparandan akuntabel serta berorientasi pada hasil, kami yang bertanda tangan di bawah ini:`,
            }),
            new Paragraph({
              text: " ",
            }),

            new Paragraph({
              spacing: { line: 360 },
              text: `   Nama : `,
            }),
            new Paragraph({
              spacing: { line: 360 },
              text: `   Jabatan : `,
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              spacing: { line: 360 },
              text: `selanjutnya disebut pihak pertama  : `,
            }),
            new Paragraph({
              text: "",
            }),

            new Paragraph({
              spacing: { line: 360 },
              text: `   Nama : `,
            }),

            new Paragraph({
              spacing: { line: 360 },
              text: `   Jabatan : `,
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: "selaku atasan pihak pertama, selanjutnya disebut pihak kedua",
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: `Pihak pertama berjanji akan mewujudkan target kinerja yang seharusnya
                    sesuai lampiran perjanjian ini, dalam rangka mencapai target kinerja jangka
                    menengah seperti yang telah ditetapkan dalam dokumen perencanaan.
                    Keberhasilan dan kegagalan pencapaian target kinerja tersebut menjadi
                    tanggung jawab kami.`,
            }),
            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: `Pihak kedua akan melakukan supervisi yang diperlukan serta akan melakukan
                    evaluasi terhadap capaian kinerja dari perjanjian ini dan mengambil tindakan
                    yang diperlukan dalam rangka pemberian penghargaan dan sanksi.`,
            }),

            new Paragraph({
              text: "",
            }),
            new Paragraph({
              text: "",
            }),
          ],
        },
      ],
    });
    const table = new Table({
      columnWidths: [3505, 5505],
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              width: {
                size: 5505,
                type: WidthType.DXA,
              },
              children: [
                new Paragraph({
                  text: "",
                }),
              ],
            }),
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              width: {
                size: 5505,
                type: WidthType.DXA,
              },
              children: [
                new Paragraph({
                  text: "......................,................",
                  alignment: AlignmentType.CENTER,
                }),
                new Paragraph({
                  text: "",
                }),
              ],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              width: {
                size: 5505,
                type: WidthType.DXA,
              },
              children: [
                new Paragraph({
                  text: "Pihak Kedua,",
                  alignment: AlignmentType.CENTER,
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
              ],
            }),
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              width: {
                size: 5505,
                type: WidthType.DXA,
              },
              children: [
                new Paragraph({
                  text: "Pihak Pertama,",
                  alignment: AlignmentType.CENTER,
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
                new Paragraph({
                  text: "",
                }),
              ],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              width: {
                size: 5505,
                type: WidthType.DXA,
              },
              children: [
                new Paragraph({
                  text: "...........................................",
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
            new TableCell({
              borders: {
                top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
              },
              width: {
                size: 5505,
                type: WidthType.DXA,
              },

              children: [
                new Paragraph({
                  text: "...........................................",
                  alignment: AlignmentType.CENTER,
                }),
              ],
            }),
          ],
        }),
      ],
    });

    doc.addSection({
      properties: {
        type: SectionType.CONTINUOUS,
      },
      children: [new Paragraph({ text: " " }), table],
    });
    const outputFile = `uploads/formulir10-${user}-${tahun}.docx`;

    fs.access(outputFile, fs.constants.F_OK, (err) => {
      if (err) {
        if (err.code === "ENOENT") {
          console.log("File does not exist");
        } else {
          console.error("Error checking file existence:", err);
        }
      } else {
        fs.unlink(outputFile, (unlinkErr) => {
          if (unlinkErr) {
            console.error("Error deleting file:", unlinkErr);
          } else {
            console.log(`${outputFile} has been deleted successfully`);
          }
        });
      }
    });

    Packer.toBuffer(doc).then((buffer) => {
      fs.writeFileSync(outputFile, buffer);
    });
    res.status(200).json({ data: outputFile });
  } catch (error) {
    console.error("Gagal membuat pengguna:", error);
    res.status(500).json({ error: "Gagal membuat pengguna" });
  }
};
