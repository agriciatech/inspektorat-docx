const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  UnderlineType,
  TextAlignment,
} = require("docx");
const fs = require("fs");

exports.setForm9 = async (tahun, data) => {
  // const today = new Date();

  // const year = today.getFullYear();
  // const month = today.getMonth(); // Months are 0-based, so we add 1
  // const day = String(today.getDate()).padStart(2, '0');

  // const namaBulan = [
  //     "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  //     "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  // ];

  // const formattedDate = `${day} ${namaBulan[month]} ${year}`;
  const doc = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: "ANAK LAMPIRAN I/1-6",
                bold: false,
              }),
            ],
            alignment: AlignmentType.RIGHT,
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: "PERNYATAAN PERJANJIAN KINERJA TINGKAT KEMENTERIAN/LEMBAGA/PEMERINTAH PROVINSI/KABUPATEN/KOTA",
                bold: false,
                color: "#000000",
              }),
            ],
            alignment: AlignmentType.CENTER,
            scale: 14,
            heading: "Heading2",

            spacing: { after: 400 },
          }),

          // Logo placeholder
          new Paragraph({
            children: [
              new TextRun({
                text: "-Logo Lembaga-",
                bold: true,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),

          // Title
          new Paragraph({
            children: [
              new TextRun({
                text: `PERJANJIAN KINERJA TAHUN ${tahun}`,
                bold: true,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // First paragraph
          new Paragraph({
            children: [
              new TextRun(
                "Dalam rangka mewujudkan manajemen pemerintahan yang efektif, transparan dan akuntabel serta berorientasi pada hasil, yang bertanda tangan di bawah ini:"
              ),
            ],
            alignment: AlignmentType.JUSTIFIED,
            // spacing: { after: 200 },
          }),

          // Name and Job Title for Pihak Pertama
          new Paragraph({
            children: [
              new TextRun({
                text: `Nama\t\t: ${data.pejabat.Nama}`,
                break: 2,
              }),
              new TextRun({
                text: `Jabatan\t\t: ${data.pejabat.Jabatan}`,
                break: 2,
              }),
            ],
            spacing: { after: 200 },
          }),
          // Pihak Pertama
          new Paragraph({
            children: [
              new TextRun({
                text: "berjanji akan mewujudkan target kinerja yang seharusnya sesuai lampiran perjanjian ini, dalam rangka mencapai target kinerja jangka menengah seperti yang telah ditetapkan dalam dokumen perencanaan.",
              }),
            ],
            spacing: { after: 400 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Keberhasilan dan kegagalan pencapaian target kinerja tersebut menjadi tanggung jawab kami.",
              }),
            ],
            spacing: { after: 600 },
          }),

          // Signature placeholders
          new Paragraph({
            children: [
              new TextRun({
                text: `........................., .......................`,
                break: 3,
              }),
              new TextRun({
                text: "Menteri/Pimpinan Lembaga/Gubernur/Bupati/Walikota",
                break: 5,
              }),
              new TextRun({
                text: "........................., .......................",
                break: 2,
              }),
            ],
            alignment: AlignmentType.RIGHT,
            spacing: { before: 400, after: 200 },
          }),
        ],
      },
    ],
  });

  const outputFile = `uploads/formulir9_page1_${tahun}.docx`;

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputFile, buffer);

  return outputFile;
};
