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
  convertInchesToTwip,
} = require("docx");
const fs = require("fs");
const prisma = new PrismaClient();

function cmToTwip(cm) {
  return Math.round(cm * 28.35 * 20); // Convert cm to points and then to twentieths of a point
}

exports.setPDFLKE = async (req, res) => {
  try {
    const { id } = req.params;

    const resultTahun =
      await prisma.$queryRaw`SELECT * from "Tahun" WHERE pk_tahun_id = ${parseInt(
        id
      )}`;
    const tahun = resultTahun[0].tahun;

    const result1 =
      await prisma.$queryRaw`SELECT * from "Inspeksi" WHERE fk_tahun = ${parseInt(
        id
      )}`;
    const resultInspektur =
      await prisma.$queryRaw`SELECT * FROM "user" as a JOIN "Role" as b ON a.fk_role = b.pk_role_id WHERE role = 'OPD';`;

    const resultRInspeksiSubKomponen =
      await prisma.$queryRaw`SELECT SUM(nilai), fk_component,fk_inspeksi FROM "RInspeksiSubKomponen" as a JOIN "SubKomponen" as b ON a.fk_sub_component = b.id GROUP BY fk_component,fk_inspeksi ORDER BY fk_component`;

    const resultPredikat =
      await prisma.$queryRaw`SELECT * FROM "Predikat" where index = 1 ORDER BY predikat `;

    let allSum = 0;
    let lengthInspektur = resultInspektur.length;
    resultInspektur.map(async (item, index) => {
      item.nilai = 0;
      result1.map(async (itemInspeksi, index) => {
        if (item.id == itemInspeksi.fk_user) {
          item.kategori = itemInspeksi.kategori;
          resultRInspeksiSubKomponen.map((subKomponen, index) => {
            if (itemInspeksi.id == subKomponen.fk_inspeksi) {
              item.nilai += parseFloat(subKomponen.sum);
            }
          });
        }
      });
      allSum += item.nilai;
    });

    const average = (allSum / resultInspektur.length).toFixed(2);
    let onceTime = true;
    let predikat = "";
    resultPredikat.map((item, index) => {
      if (average > parseInt(item.nilai) && onceTime) {
        onceTime = false;
        predikat = item.predikat;
      }
    });

    const generateRows = (allData) =>
      allData.map(
        (item, index) =>
          new TableRow({
            children: [
              new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    text: (index + 1).toString(),
                    alignment: AlignmentType.CENTER,
                  }),
                ],
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    text: item.name,
                  }),
                ],
                verticalAlign: VerticalAlign.CENTER,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    text: item.nilai.toFixed(2).toString(),
                    alignment: AlignmentType.CENTER,
                  }),
                ],
                verticalAlign: VerticalAlign.CENTER,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    text: item.kategori,
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
              size: "12pt",
              font: "Arial",
            },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
            },
          },
          document: {
            run: {
              size: "12pt",
              font: "Arial",
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
            size: "12pt",
            font: "Arial",
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
                format: LevelFormat.LOWER_LETTER,
                text: "%2.",
                alignment: AlignmentType.LOWER_LETTER,
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
                format: LevelFormat.UPPER_LETTER,
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
          properties: {
            titlePage: true,
            page: {
              pageNumbers: {
                start: 1,
                formatType: NumberFormat.DECIMAL,
              },
              margins: {
                top: cmToTwip(2.5),
                right: cmToTwip(2),
                bottom: cmToTwip(2.5),
                left: cmToTwip(3),
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
                      children: ["", PageNumber.CURRENT],
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
            new Paragraph({
              children: [
                new TextRun({
                  text: "NOTA DINAS",
                  bold: true,
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: " ",
            }),

            new Paragraph({
              children: [
                new TextRun({ text: "Yth." }),
                new TextRun("\t: Bupati Lampung Selatan."),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "" }),
                new TextRun(
                  "\t  Melalui: 1. Sekretaris Daerah Kabupaten Lampung Selatan"
                ),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "" }),
                new TextRun(
                  "\t \t      2. Asisten Administrasi Umum Sekretariat Daerah "
                ),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "" }),
                new TextRun("\t \t          Kabupaten Lampung Selatan."),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),

            new Paragraph({
              children: [
                new TextRun({ text: "Dari" }),
                new TextRun("\t: Inspektur Kabupaten Lampung Selatan"),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Tanggal" }), new TextRun("\t: ")],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [new TextRun({ text: "Nomor" }), new TextRun("\t: ")],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Sifat" }),
                new TextRun("\t: Rahasia"),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Lampiran" }),
                new TextRun("\t: -"),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Hal" }),
                new TextRun(
                  `\t: Laporan Rekapitulasi Hasil Evaluasi atas Sistem Akuntabilitas `
                ),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "" }),
                new TextRun(
                  `\t  Kinerja Instansi Pemerintah (SAKIP) Internal `
                ),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "" }),
                new TextRun(`\t  Kabupaten Lampung Selatan Tahun ${tahun}.`),
              ],
              tabStops: [
                {
                  type: TabStopType.LEFT,
                  position: 1500,
                },
              ],
            }),

            new Paragraph({
              text: "___________________________________________________________________",
            }),

            new Paragraph({
              text: "",
            }),
            new Paragraph({
              style: "aside",
              spacing: { line: 360 },
              children: [
                new TextRun({
                  text: `\t Dalam rangka pelaksanaan Peraturan Pemerintah Nomor 8 Tahun 2006 tentang Pelaporan Keuangan dan Kinerja Instansi Pemerintah, Peraturan Presiden Nomor 29 Tahun 2014 tentang Sistem Akuntabilitas Kinerja Instansi Pemerintah (SAKIP) dan Peraturan Menteri Pendayagunaan Aparatur Negara dan Reformasi Birokrasi Nomor 88 Tahun 2021 tentang Evaluasi Akuntabilitas Kinerja Instansi Pemerintah, Tim Inspektorat Kabupaten Lampung Selatan telah melakukan evaluasi atas Sistem Akuntabilitas Kinerja Instansi Pemerintah (SAKIP) internal Tahun ${tahun} kepada seluruh Perangkat Daerah dengan hasil sebagai berikut :`,
                }),
              ],
            }),
            new Paragraph({
              spacing: { line: 360 },
              text: `Berdasarkan hasil evaluasi atas Sistem Akuntabilitas Kinerja Instansi Pemerintah (SAKIP) Internal memperoleh nilai sebesar ${average} dengan kategori "${predikat}" dengan rincian capaian per masing-masing Perangkat Daerah sebagai berikut :`,
            }),
          ],
        },
      ],
    });

    doc.addSection({
      properties: {
        type: SectionType.CONTINUOUS,
      },
      children: [
        new Table({
          width: {
            size: 9020,
            type: WidthType.DXA,
          },

          rows: [
            new TableRow({
              children: [
                new TableCell({
                  rowSpan: 2,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "No",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                  rowSpan: 2,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Nama Perangkat Daerah",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                  columnSpan: 2,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Akumulasi Penilaian",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Nilai",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Kategori",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
              ],
            }),
            ...generateRows(resultInspektur),
            new TableRow({
              children: [
                new TableCell({
                  columnSpan: 2,
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Nilai Akuntabilitas Kinerja",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: (allSum / lengthInspektur)
                            .toFixed(2)
                            .toString(),
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: predikat,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
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
      children: [
        new Paragraph({
          spacing: { line: 360 },
          text: `Demikian rekapitulasi hasil evaluasi atas Sistem Akuntabilitas Kinerja Instansi Pemerintah (SAKIP) Internal ini dibuat dengan sebenarnya, atas perhatian Bapak Bupati Lampung Selatan kami ucapkan terimakasih.`,
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "",
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "INSPEKTUR,",
          indent: {
            left: convertInchesToTwip(4.2),
            hanging: convertInchesToTwip(0),
          },
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "",
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "",
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "XXXXXXXXX",
          indent: {
            left: convertInchesToTwip(4.2),
            hanging: convertInchesToTwip(0),
          },
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "XXXXXXXXX",
          indent: {
            left: convertInchesToTwip(4.2),
            hanging: convertInchesToTwip(0),
          },
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "NIP.XXXXXXX",
          indent: {
            left: convertInchesToTwip(4.2),
            hanging: convertInchesToTwip(0),
          },
        }),
      ],
    });

    const outputFile = `uploads/ND Rekapitulasi Evaluasi SAKIP-${tahun}.docx`;
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
    Packer.toBuffer(doc).then((buffer) => {
      fs.writeFileSync(outputFile, buffer);
    });

    // const inputBuffer = fs.readFileSync(outputFile);
    // const docs = new docx.Document(inputBuffer);

    // // Access the first section and remove the header from the first page
    // const firstSection = docs.sections[0];
    // if (firstSection) {
    //   const header = firstSection.headers.default;
    //   if (header) {
    //     firstSection.headers.default = new docx.Header();
    //   }
    // }

    // // Save the modified document
    // const packer = new docx.Packer();
    // packer
    //   .toBuffer(doc)
    //   .then((buffer) => {
    //     fs.writeFileSync(outputFile, buffer);
    //     console.log(`Modified DOCX saved to: ${outputFile}`);
    //   })
    //   .catch((err) => {
    //     console.error("Error saving DOCX:", err);
    //   });
    res.status(200).json({ data: outputFile });
  } catch (error) {
    console.error("Gagal membuat pengguna:", error);
    res.status(500).json({ error: "Gagal membuat pengguna" });
  }
};
