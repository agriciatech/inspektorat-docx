const { Document, Packer, Paragraph, TextRun, AlignmentType, UnderlineType, TextAlignment, Table, TableRow, TableCell, VerticalAlign, WidthType, convertInchesToTwip, BorderStyle } = require('docx');
const fs = require('fs');

exports.setForm9page2 = async (tahun, data) => {
    let call1 = [];
    let call2 = [];
    let call3 = [];
    let call4 = [];

    data.data.forEach((value, index) => {
        value.Indikator.forEach((payload, index2) => {
            if (index2 === 0) {
                call1.push(new Paragraph((index + 1).toString()));
                call2.push(new Paragraph(value.SasaranPD));
            } else {
                call1.push(new Paragraph(" "));
                call2.push(new Paragraph(" "));
            }

            call3.push(new Paragraph(payload.Indikator_sasaran_pd));
            call4.push(new Paragraph(payload.Target.toString()));
        });
    });

    const programRows = [];

    data.program.forEach((value, index) => {
        programRows.push(
            new Paragraph({
                children: [
                    new TextRun(`${index + 1}.) ${value.Program} (Rp.${value.Rp},0-)`),
                ],
                alignment: AlignmentType.LEFT,
            })
        );
    });

    // Membuat array rows secara dinamis berdasarkan data
    const tableRows = [
        // Baris pertama untuk header
        new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "No.",
                                    bold: true,
                                }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    width: {
                        size: 5, // Ukuran disesuaikan dalam persen
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Sasaran Strategis",
                                    bold: true,
                                }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    width: {
                        size: 50,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Indikator Kinerja",
                                    bold: true,
                                }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    width: {
                        size: 50,
                        type: WidthType.PERCENTAGE,
                    },
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Target",
                                    bold: true,
                                }),
                            ],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    verticalAlign: VerticalAlign.CENTER,
                    width: {
                        size: 50,
                        type: WidthType.PERCENTAGE,
                    },
                }),
            ],
            tableHeader: true,
        })
    ];

    // Menambahkan baris-baris data secara dinamis
    for (let i = 0; i < call1.length; i++) {
        tableRows.push(
            new TableRow({
                children: [
                    new TableCell({
                        children: [call1[i]],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [call2[i]],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [call3[i]],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [call4[i]],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                ],
                height: {
                    value: convertInchesToTwip(0.5),
                },
            })
        );
    }

    const doc = new Document({
        sections: [
            {
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "ANAK LAMPIRAN I/3-6",
                                bold: false,
                            }),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { after: 200 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "CONTOH FORMULIR LAMPIRAN PERJANJIAN KINERJA KEMENTERIAN/LEMBAGA/PROVINSI/KABUPATEN/KOTA",
                                bold: false,
                                color: '#000000'
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        scale: 14,
                        heading: 'Heading2',

                        spacing: { after: 100 },
                    }),

                    // Title
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `PERJANJIAN KINERJA TAHUN ${tahun} KEMENTERIAN/LEMBAGA/PROVINSI/KABUPATEN/KOTA`,
                                bold: true,
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 100 },
                    }),

                    new Table({
                        alignment: 'center',
                        rows: tableRows,
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 1 },
                            bottom: { style: BorderStyle.SINGLE, size: 1 },
                            left: { style: BorderStyle.SINGLE, size: 1 },
                            right: { style: BorderStyle.SINGLE, size: 1 },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
                            insideVertical: { style: BorderStyle.SINGLE, size: 1 },
                        },
                    }),

                    // Additional paragraphs below the table
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Program",
                                bold: true,
                            }),
                        ],
                        alignment: AlignmentType.LEFT,
                        spacing: { before: 100, after: 100 },
                    }),
                    // new Paragraph({
                    //     children: [
                    //         new TextRun("1................................."),
                    //     ],
                    //     alignment: AlignmentType.LEFT,
                    // }),
                    ...programRows,

                    // Date and Signature
                    new Paragraph({
                        children: [
                            new TextRun(`..................., .........${tahun}`),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { before: 800 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Menteri/Kepala/Gubernur/Bupati/Walikota",
                                bold: true,
                            }),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { before: 800 },
                    }),
                    new Paragraph({
                        children: [
                            new TextRun("................................................"),
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { before: 200, after: 800 },
                    }),

                    // Explanation section
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Penjelasan pengisian terhadap lampiran di atas adalah sebagai berikut: \n",
                                bold: false,
                                color: '#000000'
                            }),
                        ],
                        alignment: AlignmentType.JUSTIFIED,
                        spacing: { before: 100 },
                        heading: 'Heading3'
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "1) Pada kolom (1) diisi no urut;",
                                bold: false,
                                color: '#000000'
                            }),
                        ],
                        alignment: AlignmentType.JUSTIFIED,
                        spacing: { before: 100 },
                        heading: 'Heading3'
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "2) Pada kolom (2) diisi dengan sasaran strategis K/L/Pemda atau kondisi terakhir yang seharusnya terwujud pada tahun yang bersangkutan;\n",
                                bold: false,
                                color: '#000000'
                            }),
                        ],
                        alignment: AlignmentType.JUSTIFIED,
                        spacing: { before: 100 },
                        heading: 'Heading3'
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "3) Pada kolom (3) diisi dengan indikator kinerja utama dan indikator lain dari K/L/Pemda yang relevan dengan sasaran atau kondisi yang ingin diwujudkan;\n",
                                bold: false,
                                color: '#000000'
                            }),
                        ],
                        alignment: AlignmentType.JUSTIFIED,
                        spacing: { before: 100 },
                        heading: 'Heading3'
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "4) Pada kolom (4) diisi dengan target kinerja yang akan dicapai atau seharusnya dicapai oleh K/L/Pemda;\n",
                                bold: false,
                                color: '#000000'
                            }),
                        ],
                        alignment: AlignmentType.JUSTIFIED,
                        spacing: { before: 100 },
                        heading: 'Heading3'
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "5) Pada kolom Program diisi dengan nama program K/L/Pemda yang terkait dengan sasaran yang akan dicapai;\n",
                                bold: false,
                                color: '#000000'
                            }),
                        ],
                        alignment: AlignmentType.JUSTIFIED,
                        spacing: { before: 100 },
                        heading: 'Heading3'
                    }),
                ],
            },
        ],
    });

    const buffer = await Packer.toBuffer(doc);
    const outputFile = `uploads/formulir9_page2_${tahun}.docx`;

    fs.writeFileSync(outputFile, buffer);

    return outputFile

}
