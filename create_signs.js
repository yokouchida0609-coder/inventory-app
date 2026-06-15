const { Document, Packer, Paragraph, TextRun, AlignmentType, PageOrientation, SectionType, VerticalAlign } = require('docx');
const fs = require('fs');

const pageProps = {
  size: {
    width: 11906,
    height: 16838,
    orientation: PageOrientation.LANDSCAPE,
  },
  margin: { top: 720, right: 720, bottom: 720, left: 720 },
};

function makeSection(text, type) {
  return {
    properties: {
      type,
      page: pageProps,
      verticalAlign: VerticalAlign.CENTER,
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [
          new TextRun({
            text,
            bold: true,
            font: "Arial",
            size: 320, // 160pt
          }),
        ],
      }),
    ],
  };
}

const doc = new Document({
  sections: [
    makeSection("STOP", SectionType.NEXT_PAGE),
    makeSection("5-Minute", SectionType.NEXT_PAGE),
    makeSection("Overtime", SectionType.NEXT_PAGE),
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("signs.docx", buffer);
  console.log("signs.docx created!");
});
