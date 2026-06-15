const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
        BorderStyle, WidthType, ShadingType, HeadingLevel } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 4, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 160 },
    children: [new TextRun({ text, bold: true, font: "Meiryo", size: 36 })],
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    children: [new TextRun({ text, bold: true, font: "Meiryo", size: 28 })],
  });
}

function p(text) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, font: "Meiryo", size: 24 })],
  });
}

function cell(text, isHeader = false) {
  return new TableCell({
    borders,
    width: { size: 4500, type: WidthType.DXA },
    shading: isHeader ? { fill: "DDDDDD", type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    children: [new Paragraph({
      children: [new TextRun({ text, bold: isHeader, font: "Meiryo", size: 22 })],
    })],
  });
}

function cell2(text, width, isHeader = false) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: isHeader ? { fill: "DDDDDD", type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    children: [new Paragraph({
      children: [new TextRun({ text, bold: isHeader, font: "Meiryo", size: 22 })],
    })],
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Meiryo", size: 24 } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children: [
      // タイトル
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: "プレゼン用プラカード作成　仕様書", bold: true, font: "Meiryo", size: 52 })],
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 600 },
        children: [new TextRun({ text: "作成日：2026年4月22日", font: "Meiryo", size: 22 })],
      }),

      // 1. 概要
      h1("1. 概要"),
      p("プレゼンテーション・会議等で使用する時間管理用プラカードのWordファイルを作成する。"),
      p("印刷後、ボードに貼り付けて手持ちプラカードとして使用する。"),

      // 2. 成果物
      h1("2. 成果物"),
      p("・Wordファイル（.docx）　1点"),
      p("・ページ数：3ページ（各ページ1枚分）"),

      // 3. 仕様
      h1("3. 仕様"),

      h2("3-1. ページ設定"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 6026],
        rows: [
          new TableRow({ children: [cell2("項目", 3000, true), cell2("内容", 6026, true)] }),
          new TableRow({ children: [cell2("用紙サイズ", 3000), cell2("A4", 6026)] }),
          new TableRow({ children: [cell2("向き", 3000), cell2("横（ランドスケープ）", 6026)] }),
          new TableRow({ children: [cell2("余白", 3000), cell2("上下左右　各 12.7mm（0.5インチ）", 6026)] }),
        ],
      }),

      h2("3-2. 各ページの内容"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1500, 3000, 4526],
        rows: [
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 1500, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, shading: { fill: "DDDDDD", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: "ページ", bold: true, font: "Meiryo", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, shading: { fill: "DDDDDD", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: "表示テキスト", bold: true, font: "Meiryo", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 4526, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, shading: { fill: "DDDDDD", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: "用途", bold: true, font: "Meiryo", size: 22 })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 1500, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "1ページ目", font: "Meiryo", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "STOP", font: "Arial", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 4526, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "発表中止・終了の合図", font: "Meiryo", size: 22 })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 1500, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "2ページ目", font: "Meiryo", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "5-Minute", font: "Arial", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 4526, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "残り5分の合図", font: "Meiryo", size: 22 })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 1500, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "3ページ目", font: "Meiryo", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "Overtime", font: "Arial", size: 22 })] })] }),
            new TableCell({ borders, width: { size: 4526, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 150, right: 150 }, children: [new Paragraph({ children: [new TextRun({ text: "時間超過の合図", font: "Meiryo", size: 22 })] })] }),
          ]}),
        ],
      }),

      h2("3-3. テキスト書式"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 6026],
        rows: [
          new TableRow({ children: [cell2("項目", 3000, true), cell2("内容", 6026, true)] }),
          new TableRow({ children: [cell2("フォント", 3000), cell2("Arial", 6026)] }),
          new TableRow({ children: [cell2("フォントサイズ", 3000), cell2("160pt（目安）　※1枚に収まるサイズに調整", 6026)] }),
          new TableRow({ children: [cell2("スタイル", 3000), cell2("太字（Bold）", 6026)] }),
          new TableRow({ children: [cell2("色", 3000), cell2("黒（#000000）", 6026)] }),
          new TableRow({ children: [cell2("配置", 3000), cell2("ページ中央（水平・垂直ともに中央揃え）", 6026)] }),
        ],
      }),

      // 4. 備考
      h1("4. 備考"),
      p("・印刷後に手持ちプラカードとして使用するため、文字は遠くから視認できるサイズであること。"),
      p("・各ページは独立した1枚として完結すること（テキストがページをまたがないこと）。"),
      p("・ファイル形式は .docx（Microsoft Word形式）とする。"),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("spec.docx", buffer);
  console.log("spec.docx created!");
});
