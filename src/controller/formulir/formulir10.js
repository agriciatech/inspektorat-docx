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
                      children: ["TEST", PageNumber.CURRENT],
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
                      children: ["TEST", PageNumber.CURRENT],
                    }),
                  ],
                }),
              ],
            }),
          },
          children: [
            new Paragraph({
              text: "halo",
              border: {
                top: {
                  color: "auto",
                  space: 1,
                  value: BorderStyle.SINGLE,
                  size: 6,
                },
                bottom: {
                  color: "auto",
                  space: 1,
                  value: BorderStyle.SINGLE,
                  size: 6,
                },
                left: {
                  color: "auto",
                  space: 1,
                  value: BorderStyle.SINGLE,
                  size: 6,
                },
              },
              spacing: {
                before: 200, // Space before the paragraph with the border
                after: 200, // Space after the paragraph with the border
              },
              margin: {
                left: 400, // Space inside the border on the left
              },
            }),
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
              text: "-Logo Lembaga",
            }),
            new Paragraph({
              text: "PERJANJIAN KINERJA TAHUN ..........................",
            }),
          ],
        },
      ],
    });

    const items = [
      "Memperoleh informasi tentang Implementasi SAKIP",
      "Menilai tingkat implementasi SAKIP",
      "Menilai tingkat Akuntabilitas Kinerja",
      "Memberikan saran perbaikan untuk peningkatan implementasi SAKIP.",
      "Memonitor tindak lanjut rekomendasi hasil evaluasi periode sebelumnya",
    ];

    items.forEach((item, index) => {
      doc.addSection({
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: [
          new Paragraph({
            spacing: { line: 360 },
            text: item,
            numbering: {
              reference: "numbering-format",
              level: 1,
            },
          }),
        ],
      });
    });

    doc.addSection({
      properties: {
        type: SectionType.CONTINUOUS,
      },
      children: [
        new Paragraph({
          spacing: { line: 360 },
          text: `Evaluasi dilaksanakan terhadap 4 (empat) komponen besar manajemen kinerja, yang meliputi: Perencanaan Kinerja, Pengukuran Kinerja, Pelaporan kinerja, Evaluasi Akuntabilitas Kinerja Internal. Laporan Kinerja Instansi Pemerintah (LKjIP) tahun ${
            parseInt(tahun) - 1
          } merupakan salah satu dokumen yang dievaluasi selain dokumen Rencana Strategis (Renstra), dokumen Rencana Kerja (Renja), dokumen Penetapan Kinerja (PK), dokumen Pohon Kinerja serta dokumen terkait lainnya.`,
          numbering: {
            reference: "numbering-format",
            level: 0,
          },
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: `Hasil evaluasi dituangkan dalam bentuk nilai dengan kisaran mulai dari 0 s.d. 100. Berdasarkan hasil evaluasi, tingkat akuntabilitas ${user} Kabupaten Lampung Selatan memperoleh nilai sebesar ${total.toString()} dengan kategori ${predikat}`,
          numbering: {
            reference: "numbering-format",
            level: 0,
          },
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
          text: `Nilai tersebut merupakan akumulasi penilaian terhadap seluruh komponen manajemen kinerja yang dievaluasi pada ${user} Kabupaten Lampung Selatan untuk Tahun ${tahun}, dengan rincian sebagai berikut:`,
          numbering: {
            reference: "numbering-format",
            level: 0,
          },
        }),
      ],
    });

    doc.addSection({
      properties: {
        type: SectionType.CONTINUOUS,
      },
      children: [
        new Table({
          width: {
            size: 8090,
            type: WidthType.DXA,
          },
          indent: {
            size: 320,
            type: WidthType.DXA,
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
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
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Komponen",
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
                          text: "Bobot (%)",
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
                          text: "Nilai",
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  verticalAlign: VerticalAlign.CENTER,
                }),
              ],
            }),
            ...generateRows(resultComponent, resultRInspeksiSubKomponen),
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
                          text: "100",
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
                          text: total.toFixed(2).toString(),
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
          text: `Berdasarkan hasil evaluasi, dijumpai permasalahan terkait implementasi SAKIP pada ${user} Kabupaten Lampung Selatan sebagai berikut:`,
          numbering: {
            reference: "numbering-format",
            level: 0,
          },
        }),
      ],
    });

    resultComponent.forEach((item, index) => {
      doc.addSection({
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: [
          new Paragraph({
            spacing: { line: 360 },
            text: item.component,
            numbering: {
              reference: "numbering-format",
              level: 1,
            },
          }),
        ],
      });

      result.map((result, index) => {
        result.SubKomponen.map((subkomponen, index) => {
          subkomponen.Keriteria.map((resultCatatan, index) => {
            if (
              item.id == subkomponen.FK_Component &&
              (resultCatatan.Catatan.verifikasi.includes("Sebagian") ||
                resultCatatan.Catatan.verifikasi.includes("Tidak"))
            ) {
              console.log(resultCatatan);
              doc.addSection({
                properties: {
                  type: SectionType.CONTINUOUS,
                },
                children: [
                  new Paragraph({
                    spacing: { line: 360 },
                    text: resultCatatan.Catatan.catatan,
                    numbering: {
                      reference: "numbering-format",
                      level: 2,
                    },
                  }),
                ],
              });
            }
          });
        });
      });
      resultInspeksiKriteria.forEach((resultKeriteria, index) => {});
    });

    doc.addSection({
      properties: {
        type: SectionType.CONTINUOUS,
      },
      children: [
        new Paragraph({
          spacing: { line: 360 },
          text: `Terhadap permasalahan di atas, kami rekomendasikan kepada Kepala ${user} Kabupaten Lampung Selatan beserta seluruh jajaran untuk melakukan perbaikan sebagai berikut:`,
          numbering: {
            reference: "numbering-format",
            level: 0,
          },
        }),
      ],
    });

    resultComponent.forEach((item, index) => {
      doc.addSection({
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: [
          new Paragraph({
            spacing: { line: 360 },
            text: item.component,
            numbering: {
              reference: "numbering-format",
              level: 1,
            },
          }),
        ],
      });
      resultRekomendasi[0].forEach((resultRekomendasiMap, indexfirst) => {
        resultRekomendasiMap.Rekomendasi.map((result, index) => {
          if (item.id == result.Fk_Komponen) {
            doc.addSection({
              properties: {
                type: SectionType.CONTINUOUS,
              },
              children: [
                new Paragraph({
                  spacing: { line: 360 },
                  text: result.rekomendasi,
                  numbering: {
                    reference: "numbering-format",
                    level: 2,
                  },
                }),
              ],
            });
          }
        });
      });
    });

    doc.addSection({
      properties: {
        type: SectionType.CONTINUOUS,
      },
      children: [
        new Paragraph({
          spacing: { line: 360 },
          text: `\t Demikian hasil evaluasi atas AKIP pada ${user} Kabupaten Lampung Selatan Tahun ${tahun}, dimohon kerjasama Saudara dalam melakukan perbaikan dan menindaklanjuti saran yang telah diberikan oleh Tim Evaluasi Inspektorat Kabupaten Lampung Selatan. Atas perhatian dan kerjasamanya, kami ucapkan terimakasih.`,
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
          text: "Tembusan disampaikan kepada Yth:",
        }),
        new Paragraph({
          spacing: { line: 360 },
          text: "Bupati Lampung Selatan (sebagai laporan)",
        }),
      ],
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
