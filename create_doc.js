const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, width, opts) {
  opts = opts || {};
  return new TableCell({
    borders: borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: text, bold: opts.bold || false, size: opts.size || 22, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })]
    })]
  });
}

function infoTable(rows) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [2400, 6960],
    rows: rows.map(function(r) {
      return new TableRow({
        cantSplit: true,
        children: [ cell(r[0], 2400, { shade: "D5E8F0", bold: true }), cell(r[1], 6960) ]
      });
    })
  });
}

function heading(text, level) {
  return new Paragraph({
    heading: level,
    children: [new TextRun({ text: text, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })]
  });
}

function para(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { before: opts.before || 60, after: opts.after || 60 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({
      text: text,
      bold: opts.bold || false,
      size: opts.size || 24,
      color: opts.color || "000000",
      font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }
    })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: text, size: 24, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 100, after: 100 }, children: [] });
}

function storyPara(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 360 },
    indent: { firstLine: 480 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text: text, size: 24, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })]
  });
}

function storyQuote(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: text, size: 22, italics: true, color: "808080", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })]
  });
}

function restTable(rows, widths) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: widths,
    rows: rows.map(function(r, i) {
      return new TableRow({
        cantSplit: true,
        children: r.map(function(c, j) {
          var w = widths[j];
          if (i === 0) return cell(c, w, { shade: "D5E8F0", bold: true });
          return cell(c, w);
        })
      });
    })
  });
}

function summaryTable(headers, rows, widths) {
  var allRows = [new TableRow({
    cantSplit: true,
    children: headers.map(function(h, j) {
      return cell(h, widths[j], { shade: "1F4E79", bold: true, align: AlignmentType.CENTER });
    })
  })];
  rows.forEach(function(r) {
    allRows.push(new TableRow({
      cantSplit: true,
      children: r.map(function(c, j) { return cell(c, widths[j]); })
    }));
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: widths,
    rows: allRows
  });
}

var content = [];

// Title
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 400, after: 200 },
  children: [new TextRun({ text: "\u5317\u4EAC\u666F\u70B9\u6E38\u89C8\u653B\u7565", size: 48, bold: true, color: "1F4E79", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })]
}));
content.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 400 },
  children: [new TextRun({ text: "9\u67081\u65E5\u5929\u5B89\u95E8\u5E7F\u573A\u53CA\u5468\u8FB9\u666F\u70B9\u6E38\u89C8\u8BA1\u5212", size: 28, color: "808080", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })]
}));

// Summary table
content.push(heading("\u6E38\u89C8\u4FE1\u606F\u6C47\u603B\u8868", HeadingLevel.HEADING_2));
content.push(summaryTable(
  ["\u666F\u70B9", "70\u5468\u5C81\u95E8\u7968", "\u7EFF\u8272\u901A\u9053", "\u9884\u8BA1\u65F6\u957F", "\u9884\u7EA6\u65B9\u5F0F"],
  [
    ["\u5929\u5B89\u95E8\u5347\u65D7", "\u514D\u8D39", "\u975E\u5B98\u65B9\u63D0\u53CA", "2-3\u5C0F\u65F6", "\u201C\u5929\u5B89\u95E8\u5E7F\u573A\u9884\u7EA6\u53C2\u89C8\u201D\u5C0F\u7A0B\u5E8F"],
    ["\u4EBA\u6C11\u5927\u4F1A\u5802", "\u514D\u8D39", "\u65E0\u660E\u786E\u901A\u9053", "1-2\u5C0F\u65F6", "\u201C\u4EBA\u6C11\u5927\u4F1A\u5802\u53C2\u89C8\u9884\u7EA6\u201D\u5C0F\u7A0B\u5E8F"],
    ["\u56FD\u5BB6\u535A\u7269\u9986", "\u514D\u8D39", "\u5317\u95E8\u7EFF\u8272\u901A\u9053", "2-3\u5C0F\u65F6", "\u56FD\u535A\u5B98\u7F51/\u5C0F\u7A0B\u5E8F\uFF0C\u63D0\u524D7\u5929"],
    ["\u6545\u5BAB", "\u534A\u4EF7(60\u5C81+)", "\u7231\u5FC3\u901A\u9053", "3-6\u5C0F\u65F6", "\u201C\u6545\u5BAB\u535A\u7269\u9662\u201D\u5C0F\u7A0B\u5E8F\uFF0C\u63D0\u524D20\u5929"],
    ["\u4E2D\u592E\u793C\u54C1\u4E2D\u5FC3", "\u514D\u8D39", "60\u5C81+\u73B0\u573A\u9884\u7EA6", "2-3\u5C0F\u65F6", "\u516C\u4F17\u53F7\u5B9E\u540D\u9884\u7EA6"],
    ["\u56FD\u5BB6\u5927\u5267\u9662", "\u514D\u7968", "\u65E0\u969C\u788D\u8BBE\u65BD", "2-3\u5C0F\u65F6", "\u5B98\u65B9\u5E73\u53F0\u8D2D\u7968"],
    ["\u9E1F\u5DE2", "\u514D\u7968", "\u65E0\u969C\u788D\u901A\u9053", "1-2\u5C0F\u65F6", "\u5C0F\u7A0B\u5E8F/\u73B0\u573A\u8D2D\u7968"],
    ["\u6C34\u7ACB\u65B9", "\u514D\u8D39(65\u5C81+)", "\u8BC1\u4EF6\u6838\u9A8C", "\u7EA61\u5C0F\u65F6", "\u201C\u6C34\u7ACB\u65B9\u201D\u516C\u4F17\u53F7\u8D2D\u7968"],
    ["\u5965\u68EE\u516C\u56ED", "\u514D\u8D39", "\u65E0\u9700", "\u534A\u5929\u81F3\u5168\u5929", "\u65E0\u9700\u9884\u7EA6"],
    ["\u5929\u575B", "\u514D\u8D39(60\u5C81+)", "\u5168\u56ED\u65E0\u53F0\u9636", "2-3\u5C0F\u65F6", "\u201C\u5929\u575B\u201D/\u201C\u7545\u6E38\u516C\u56ED\u201D\u516C\u4F17\u53F7\uFF0C\u63D0\u524D7\u5929"],
  ],
  [1800, 1400, 1400, 1400, 3360]
));

content.push(new Paragraph({ children: [new PageBreak()] }));

// Section 1
content.push(heading("\u4E00\u3001\u6E38\u89C8\u5B89\u6392", HeadingLevel.HEADING_1));

// 1. Tiananmen
content.push(heading("1. \u5929\u5B89\u95E8\u5E7F\u573A\u5347\u65D7\u4EEA\u5F0F", HeadingLevel.HEADING_2));
content.push(para("9\u67081\u65E5\u5347\u65D7\u65F6\u95F4\uFF1A\u7EA605:42\uFF089\u67081\u65E5\u6709\u519B\u4E50\u56E2\u73B0\u573A\u6F14\u594F\uFF09", { bold: true }));
content.push(infoTable([
  ["\u6E38\u89C8\u65E5\u671F", "2026\u5E749\u67081\u65E5\uFF088\u670831\u65E5\u534A\u591C\u6392\u961F\uFF09"],
  ["\u5347\u65D7\u65F6\u95F4", "\u7EA605:42:09"],
  ["\u95E8\u7968\u4EF7\u683C", "\u514D\u8D39\uFF0C\u4F46\u987B\u63D0\u524D\u9884\u7EA6"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u5347\u65D7\u4EEA\u5F0F\u7EA610\u5206\u949F\uFF1B\u542B\u6392\u961F\u3001\u62CD\u7167\u5EFA\u8BAE\u9884\u75592-3\u5C0F\u65F6"],
]));

content.push(para("70\u5468\u5C81\u8001\u5E74\u4EBA\u7EFF\u8272\u901A\u9053", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u5929\u5B89\u95E8\u5E7F\u573A\u5347\u65D7\u9884\u7EA6\u89C4\u5219\u4E2D\uFF0C60\u5C81\u4EE5\u4E0A\u8001\u4EBA\u5C5E\u4E8E\u9700\u9884\u7EA6\u53C2\u89C8\u4EBA\u7FA4\uFF0C\u53EF\u4E0E\u672A\u6210\u5E74\u4EBA\u3001\u519B\u4EBA\u3001\u6B8B\u75BE\u4EBA\u7B49\u4E00\u8D77\u901A\u8FC7\u9884\u7EA6\u5E73\u53F0\u9884\u7EA6\u3002"));
content.push(bullet("\u90E8\u5206\u65C5\u6E38\u653B\u7565\u63D0\u5230\u201C60\u5C81\u4EE5\u4E0A\u51ED\u8EAB\u4EFD\u8BC1\u8D70\u5FEB\u901F\u901A\u9053\uFF0C\u53EF\u514D\u6392\u961F\u5B89\u68C0\u201D\uFF0C\u4F46\u6B64\u4FE1\u606F\u6765\u81EA\u975E\u5B98\u65B9\u6765\u6E90\uFF0C\u5EFA\u8BAE\u4EE5\u73B0\u573A\u5DE5\u4F5C\u4EBA\u5458\u6307\u5F15\u4E3A\u51C6\u3002"));
content.push(bullet("\u8001\u4EBA\u53EF\u7531\u540C\u884C\u6210\u5E74\u4EBA\u4EE3\u4E3A\u9884\u7EA6\uFF0C\u586B\u5199\u5BF9\u5E94\u8EAB\u4EFD\u4FE1\u606F\u5373\u53EF\uFF0C\u65E0\u9700\u8001\u4EBA\u672C\u4EBA\u64CD\u4F5C\u624B\u673A\u3002"));

content.push(para("\u5BB6\u5C5E\u966A\u540C\u653F\u7B56", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u672A\u5355\u72EC\u5217\u51FA\u201C70\u5468\u5C81\u8001\u4EBA\u96501\u540D\u5BB6\u5C5E\u966A\u540C\u201D\u7684\u5B98\u65B9\u89C4\u5219\u3002\u8001\u4EBA\u53EF\u7531\u540C\u884C\u6210\u5E74\u4EBA\u4EE3\u4E3A\u9884\u7EA6\uFF0C\u5165\u573A\u65F6\u914D\u5408\u8EAB\u4EFD\u6838\u9A8C\u5373\u53EF\u3002"));
content.push(bullet("\u5B9E\u9645\u64CD\u4F5C\u4E2D\uFF0C\u5EFA\u8BAE1\u540D\u5BB6\u5C5E\u966A\u540C\u8001\u4EBA\u524D\u5F80\uFF0C\u65B9\u4FBF\u7167\u987E\u548C\u6C9F\u901A\u3002"));

content.push(para("\u6E29\u99A8\u63D0\u793A", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9884\u7EA6\u65B9\u5F0F\uFF1A\u5FAE\u4FE1\u641C\u7D22\u201C\u5929\u5B89\u95E8\u5E7F\u573A\u9884\u7EA6\u53C2\u89C8\u201D\u5C0F\u7A0B\u5E8F\uFF0C\u63D0\u524D1-9\u5929\u9884\u7EA6\uFF0C\u4E0D\u53EF\u5F53\u5929\u9884\u7EA6\u3002\u6BCF\u65E5\u5206\u65F6\u6BB5\uFF0C\u987B\u9009\u3010\u5347\u65D7\u3011\u65F6\u6BB5\u3002"));
content.push(bullet("\u9884\u7EA6\u5355\u6700\u591A\u53EF\u9884\u7EA67\u4EBA\uFF0C\u5176\u4E2D\u6700\u591A4\u540D\u6210\u5E74\u4EBA\u3002"));
content.push(bullet("\u672A\u9884\u7EA6\u4E0D\u80FD\u8FDB\u5165\u5929\u5B89\u95E8\u5E7F\u573A\uFF0C\u73B0\u573A\u4E0D\u8865\u9884\u7EA6\u30021\u4E2A\u6708\u5185\u723D\u7EA63\u6B21\u53CA\u4EE5\u4E0A\uFF0C\u534A\u5E74\u5185\u4E0D\u652F\u6301\u518D\u9884\u7EA6\u3002"));
content.push(bullet("\u5B89\u68C0\u975E\u5E38\u4E25\u683C\uFF0C\u5EFA\u8BAE\u53EA\u5E26\u8EAB\u4EFD\u8BC1\u548C\u624B\u673A\uFF0C\u5426\u5219\u5B89\u68C0\u73AF\u8282\u4F1A\u803D\u8BEF\u65F6\u95F4\u3002"));
content.push(bullet("\u63A8\u8350\u5B89\u68C0\u8DEF\u7EBF\uFF1A\u524D\u95E8\u897F\u4FA7\u8DEF\u7EBF\u6392\u961F\u4EBA\u6570\u6700\u5C11\u3001\u7528\u65F6\u6700\u5C11\uFF0C\u5B89\u68C0\u51FA\u53E3\u5373\u662F\u516C\u5395\u3002"));
content.push(bullet("\u901A\u8FC7\u5B89\u68C0\u540E\u76F4\u884C\uFF0C\u5DE6\u62D0\u8FDB\u5730\u4E0B\u901A\u9053\uFF0C\u76F4\u884C\u540E\u53F3\u62D0\u8FC7\u9A6C\u8DEF\uFF0C\u5230\u5929\u5B89\u95E8\u5E7F\u573A\u6B63\u5BF9\u9762\u3002"));
content.push(bullet("\u4E0D\u8981\u63A5\u8DEF\u8FB9\u6709\u4EBA\u9012\u56FD\u65D7\uFF08\u6709\u4EBA\u4E13\u95E8\u8D5A\u6E38\u5BA2\u94B1\uFF09\uFF1B\u4E0D\u8981\u5750\u6469\u7684\u3002"));
content.push(bullet("\u51CC\u6668\u4E09\u56DB\u70B9\u6392\u961F\u62A2\u524D\u6392\u5BF9\u8001\u4EBA\u8EAB\u4F53\u8D1F\u62C5\u8FC7\u91CD\uFF0C\u5EFA\u8BAE\u4FDD\u8BC1\u7761\u7720\uFF0C\u9009\u62E9\u8212\u9002\u7AD9\u4F4D\u3002"));
content.push(bullet("\u5DF2\u9884\u7EA6\u6BDB\u4E3B\u5E2D\u7EAA\u5FF5\u5802\u3001\u4EBA\u6C11\u5927\u4F1A\u5802\u3001\u56FD\u5BB6\u535A\u7269\u9986\u3001\u6545\u5BAB\u7B49\u666F\u70B9\u7684\u6E38\u5BA2\uFF0C\u53EF\u51ED\u9884\u7EA6\u8BB0\u5F55\u8FDB\u5165\u5929\u5B89\u95E8\u5E7F\u573A\u53C2\u89C8\uFF0C\u4F46\u89C2\u770B\u5347\u65D7\u4EEA\u5F0F\u987B\u53E6\u884C\u9884\u7EA6\u201C\u5347\u65D7\u201D\u65F6\u6BB5\u3002"));

content.push(spacer());

// 2. Great Hall
content.push(heading("2. \u4EBA\u6C11\u5927\u4F1A\u5802", HeadingLevel.HEADING_2));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u4E1C\u57CE\u533A\u5929\u5B89\u95E8\u5E7F\u573A\u897F\u4FA7"],
  ["\u5F00\u653E\u65F6\u95F4", "9:00-15:00\uFF0C14:30\u505C\u6B62\u68C0\u7968"],
  ["\u95E8\u7968\u4EF7\u683C", "\u6210\u4EBA\u796330\u5143\uFF1B\u5B66\u751F\u796815\u5143\uFF1B60\u5468\u5C81\uFF08\u542B\uFF09\u4EE5\u4E0A\u4E2D\u56FD\u516C\u6C11\u514D\u8D39"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u5EFA\u8BAE\u9884\u75591-2\u5C0F\u65F6\uFF08\u542B\u6392\u961F\u5B89\u68C0\u3001\u68C0\u7968\u548C\u901A\u884C\u65F6\u95F4\uFF09"],
]));

content.push(para("70\u5468\u5C81\u8001\u5E74\u4EBA\u7EFF\u8272\u901A\u9053", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("60\u5C81\u4EE5\u4E0A\u8001\u4EBA\u53EF\u4EAB\u53D7\u514D\u7968\uFF0C\u4F46\u4ECD\u9700\u6309\u89C4\u5B9A\u63D0\u524D\u7F51\u4E0A\u9884\u7EA6\u3002"));
content.push(bullet("\u5B98\u65B9\u653F\u7B56\u4E2D\u672A\u5355\u72EC\u5217\u51FA\u201C70\u5468\u5C81\u7EFF\u8272\u901A\u9053\u201D\u6216\u201C\u514D\u6392\u961F\u201D\u89C4\u5219\uFF0C\u4F18\u60E0/\u514D\u7968\u4EBA\u7FA4\u540C\u6837\u9700\u8981\u63D0\u524D\u7F51\u4E0A\u9884\u7EA6\uFF0C\u5E76\u5728\u5165\u9986\u65F6\u51FA\u793A\u6709\u6548\u8BC1\u4EF6\u539F\u4EF6\u6838\u9A8C\u3002"));

content.push(para("\u5BB6\u5C5E\u966A\u540C\u653F\u7B56", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u5B98\u65B9\u660E\u786E\uFF1A\u672A\u6EE114\u5468\u5C81\u5C11\u513F\u7AE5\u3001\u6B8B\u75BE\u4EBA\u3001\u884C\u52A8\u4E0D\u4FBF\u6216\u65E0\u81EA\u4E3B\u884C\u4E3A\u80FD\u529B\u6E38\u5BA2\u53C2\u89C8\u65F6\u987B\u6709\u6210\u5E74\u4EBA\u966A\u540C\uFF0C\u966A\u540C\u4EBA\u9700\u6309\u89C4\u5B9A\u9884\u7EA6\u8D2D\u7968\u3002"));
content.push(bullet("70\u5468\u5C81\u4EE5\u4E0A\u8001\u4EBA\u867D\u65E0\u5F3A\u5236\u966A\u540C\u8981\u6C42\uFF0C\u4F46\u5EFA\u8BAE1\u540D\u5BB6\u5C5E\u966A\u540C\uFF0C\u65B9\u4FBF\u7167\u987E\u3002"));

content.push(para("\u6E29\u99A8\u63D0\u793A", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9884\u7EA6\u65B9\u5F0F\uFF1A\u5FAE\u4FE1\u5C0F\u7A0B\u5E8F\u201C\u4EBA\u6C11\u5927\u4F1A\u5802\u53C2\u89C8\u9884\u7EA6\u201D\uFF0C\u63D0\u524D1-3\u5929\u9884\u7EA6\uFF0C\u6BCF\u65E517:00\u653E\u7968\u3002\u4E0D\u8BBE\u73B0\u573A\u552E\u7968\uFF0C\u4E0D\u552E\u5F53\u65E5\u7968\u548C\u56E2\u4F53\u7968\u3002"));
content.push(bullet("\u5FC5\u987B\u643A\u5E26\u9884\u7EA6\u7684\u8EAB\u4EFD\u8BC1\u539F\u4EF6\uFF0C\u7535\u5B50\u8BC1\u4EF6\u53EF\u80FD\u4E0D\u597D\u4F7F\u3002"));
content.push(bullet("\u7981\u6B62\u643A\u5E26\u6DB2\u4F53\u3001\u98DF\u54C1\u3001\u5145\u7535\u5B9D\u3001\u706B\u79CD\u7B49\u5B58\u5728\u5B89\u5168\u9690\u60A3\u7684\u7269\u54C1\u3002"));
content.push(bullet("\u9152\u9B3C\u8005\u3001\u8863\u51A0\u4E0D\u6574\u8005\u8C22\u7EDD\u53C2\u89C8\u3002\u4E25\u7981\u8FDB\u5165\u975E\u5F00\u653E\u533A\u57DF\uFF0C\u7981\u6B62\u5927\u58F0\u55A7\u54D4\u3001\u8FFD\u9010\u6253\u95F9\u3002"));
content.push(bullet("\u54A8\u8BE2\u7535\u8BDD\uFF1A010-83084776\uFF0C\u5DE5\u4F5C\u65F6\u95F49:00-15:00\u3002"));

content.push(spacer());

// 3. National Museum
content.push(heading("3. \u56FD\u5BB6\u535A\u7269\u9986", HeadingLevel.HEADING_2));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u4E1C\u57CE\u533A\u4E1C\u957F\u5B89\u885716\u53F7\uFF08\u5929\u5B89\u95E8\u5E7F\u573A\u4E1C\u4FA7\uFF09"],
  ["\u5F00\u653E\u65F6\u95F4", "\u6BCF\u65E59:00-17:00\uFF0C16:00\u505C\u6B62\u5165\u9986\uFF1B\u5468\u4E00\u95ED\u9986\uFF08\u6CD5\u5B9A\u8282\u5047\u65E5\u9664\u5916\uFF09\u30026\u67081\u65E5\u81F310\u670831\u65E5\u5EF6\u957F\u81F317:30"],
  ["\u95E8\u7968\u4EF7\u683C", "\u514D\u8D39\u53C2\u89C8\uFF0C\u987B\u5B9E\u540D\u9884\u7EA6"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u5EFA\u8BAE\u9884\u75592-3\u5C0F\u65F6\uFF1B\u6DF1\u5EA6\u6E38\u89C8\u5EFA\u8BAE3\u5C0F\u65F6\u4EE5\u4E0A"],
]));

content.push(para("70\u5468\u5C81\u8001\u5E74\u4EBA\u7EFF\u8272\u901A\u9053", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u5B98\u65B9\u660E\u786E\uFF1A60\u5468\u5C81\uFF08\u542B\uFF09\u4EE5\u4E0A\u8001\u4EBA\uFF0C\u51ED\u672C\u4EBA\u9884\u7EA6\u8BC1\u4EF6\u539F\u4EF6\u3001\u9884\u7EA6\u6210\u529F\u4FE1\u606F\u548C\u4F18\u5F85\u8BC1\u4EF6\uFF0C\u5728\u6240\u9884\u7EA6\u65F6\u6BB5\u5185\u4ECE\u5317\u95E8\u7EFF\u8272\u901A\u9053\u6838\u9A8C\u5165\u9986\u3002"));
content.push(bullet("\u6B64\u653F\u7B56\u81EA2025\u5E748\u67081\u65E5\u8D77\u5B9E\u65BD\uFF0C70\u5468\u5C81\u8001\u4EBA\u8D70\u5317\u95E8\u7EFF\u8272\u901A\u9053\u6709\u660E\u786E\u5B98\u65B9\u4F9D\u636E\u3002"));

content.push(para("\u5BB6\u5C5E\u966A\u540C\u653F\u7B56", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u5B98\u65B9\u672A\u5355\u72EC\u5217\u51FA\u201C70\u5468\u5C81\u8001\u4EBA\u96501\u540D\u5BB6\u5C5E\u966A\u540C\u201D\u7684\u89C4\u5219\u300260\u5468\u5C81\u4EE5\u4E0A\u8001\u4EBA\u51ED\u672C\u4EBA\u9884\u7EA6\u8BC1\u4EF6\u4ECE\u5317\u95E8\u7EFF\u8272\u901A\u9053\u6838\u9A8C\u5165\u9986\u3002"));
content.push(bullet("\u672A\u6EE114\u5468\u5C81\uFF08\u542B\uFF09\u7684\u672A\u6210\u5E74\u4EBA\uFF0C\u987B\u7531\u6210\u5E74\u4EBA\u4EE3\u4E3A\u9884\u7EA6\uFF0C\u5E76\u4E0E\u9884\u7EA6\u4EBA\u540C\u6B65\u6838\u9A8C\u5165\u9986\u3002"));
content.push(bullet("\u5EFA\u8BAE1\u540D\u5BB6\u5C5E\u966A\u540C\u8001\u4EBA\u4ECE\u7EFF\u8272\u901A\u9053\u4E00\u8D77\u5165\u9986\uFF08\u5B9E\u9645\u64CD\u4F5C\u4E2D\u5BB6\u5C5E\u8D70\u540C\u884C\u901A\u9053\u5373\u53EF\uFF09\u3002"));

content.push(para("\u6E29\u99A8\u63D0\u793A", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9884\u7EA6\u65B9\u5F0F\uFF1A\u56FD\u5BB6\u535A\u7269\u9986\u5B98\u7F51\u3001\u5B98\u65B9\u9884\u7EA6\u5C0F\u7A0B\u5E8F\u3001\u5FAE\u4FE1\u5C0F\u7A0B\u5E8F\u548C\u5FAE\u4FE1\u516C\u4F17\u53F7\u9884\u7EA6\u3002\u63D0\u524D7\u65E5\u5185\u9884\u7EA6\uFF0C\u6BCF\u65E517:00\u653E\u7968\u3002"));
content.push(bullet("\u6BCF\u65E5\u5206\u4E3A9:00-11:00\u300111:00-13:30\u300113:30-16:00\u4E09\u4E2A\u9884\u7EA6\u5165\u9986\u65F6\u6BB5\u3002"));
content.push(bullet("\u6BCF\u4E2A\u8D26\u53F7\u6BCF\u5468\u6700\u591A\u9884\u7EA61\u6B21\uFF0C\u6BCF\u6B21\u6700\u591A\u9884\u7EA65\u4EBA\u3002\u540C\u4E00\u8BC1\u4EF6\u53F7\u6BCF\u6708\u6700\u591A\u9884\u7EA64\u6B21\uFF0C\u6BCF\u5929\u4EC5\u53EF\u9884\u7EA61\u6B21\u3002"));
content.push(bullet("\u9884\u7EA6\u540E\u9000\u7968\u4ECD\u8BA1\u5165\u9884\u7EA6\u6B21\u6570\u3002\u6BCF\u4E2A\u8BC1\u4EF6\u53F7\u7D2F\u79EF3\u6B21\u672A\u5C5A\u7EA6\u5C06\u88AB\u9650\u5236\u9884\u7EA630\u65E5\u3002"));
content.push(bullet("60\u5468\u5C81\u4EE5\u4E0A\u8001\u4EBA\u8D70\u5317\u95E8\u7EFF\u8272\u901A\u9053\u6838\u9A8C\u5165\u9986\u3002\u56FD\u5BB6\u535A\u7269\u9986\u8BBE\u6709\u65E0\u969C\u788D\u7535\u68AF\u3001\u4F11\u606F\u5EA7\u6905\u3001\u996E\u6C34\u5904\uFF0C\u5BF9\u8001\u5E74\u7FA4\u4F53\u8F83\u53CB\u597D\u3002"));
content.push(bullet("\u5FC5\u770B\u5C55\u54C1\u63A8\u8350\uFF1A\u53E4\u4EE3\u4E2D\u56FD\u57FA\u672C\u9648\u5217\u5C55\uFF08\u5730\u4E0B\u4E00\u5C42\uFF09\u3001\u590D\u5174\u4E4B\u8DEF\u5C55\uFF08\u5317\u4FA7\uFF09\u3001\u540E\u6BCD\u620E\u9F0E\u3001\u56DB\u7F8A\u65B9\u5C0A\u7B49\u3002"));
content.push(bullet("\u54A8\u8BE2\u7535\u8BDD\uFF1A010-65116400\u3002"));

content.push(spacer());

// 4. Forbidden City
content.push(heading("4. \u6545\u5BAB", HeadingLevel.HEADING_2));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u4E1C\u57CE\u533A\u666F\u5C71\u524D\u88574\u53F7\uFF08\u5929\u5B89\u95E8\u5E7F\u573A\u5317\u4FA7\uFF09"],
  ["\u5F00\u653E\u65F6\u95F4", "\u65FA\u5B63\uFF084\u67081\u65E5-10\u670831\u65E5\uFF098:30-17:00\uFF0816:00\u505C\u6B62\u5165\u9986\uFF09\uFF1B\u6DE1\u5B63\uFF0811\u67081\u65E5-3\u670831\u65E5\uFF098:30-16:30\u3002\u5468\u4E00\u95ED\u9986\uFF08\u6CD5\u5B9A\u8282\u5047\u65E5\u9664\u5916\uFF09"],
  ["\u95E8\u7968\u4EF7\u683C", "\u65FA\u5B6360\u5143/\u4EBA\uFF1B\u6DE1\u5B6340\u5143/\u4EBA\u3002\u73CD\u5B9D\u9986\u3001\u949F\u8868\u9986\u540410\u5143\u300260\u5468\u5C81\uFF08\u542B\uFF09\u4EE5\u4E0A\u534A\u4EF7\uFF08\u65FA\u5B6330\u5143\uFF0C\u6DE1\u5B6320\u5143\uFF09\uFF1B18\u5468\u5C81\u4EE5\u4E0B\u514D\u8D39"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u534A\u65E5\u6E383-4\u5C0F\u65F6\uFF1B\u6DF1\u5EA6\u6E38\u89C86\u5C0F\u65F6\u4EE5\u4E0A"],
]));

content.push(para("70\u5468\u5C81\u8001\u5E74\u4EBA\u7EFF\u8272\u901A\u9053", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("60\u5468\u5C81\u4EE5\u4E0A\u8001\u5E74\u89C2\u4F17\u53EF\u4ECE\u7231\u5FC3\u901A\u9053\u68C0\u7968\u901A\u884C\uFF0C\u65E0\u9700\u6392\u961F\u3002\u4E58\u5750\u8F6E\u6905\u7684\u89C2\u4F17\u53EF\u4ECE\u65E0\u969C\u788D\u901A\u9053\u68C0\u7968\u901A\u884C\u3002"));
content.push(bullet("\u6545\u5BAB\u672A\u5B9E\u884C70\u5468\u5C81\u4EE5\u4E0A\u5168\u514D\u653F\u7B56\uFF0C60\u5468\u5C81\u4EE5\u4E0A\u4EAB\u53D7\u534A\u4EF7\u4F18\u60E0\uFF0C\u4ECD\u9700\u6309\u89C4\u5B9A\u9884\u7EA6\u8D2D\u7968\u3002"));
content.push(bullet("\u7231\u5FC3\u901A\u9053\u4F4D\u4E8E\u5348\u95E8\u5B89\u68C0\u533A\uFF0C\u8001\u5E74\u89C2\u4F17\u51ED\u6709\u6548\u8BC1\u4EF6\u5373\u53EF\u4F7F\u7528\u3002"));

content.push(para("\u5BB6\u5C5E\u966A\u540C\u653F\u7B56", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u6240\u6709\u89C2\u4F17\uFF08\u542B\u4F18\u60E0\u4EBA\u7FA4\u53CA\u966A\u62A4\u4EBA\uFF09\u5747\u987B\u5B9E\u540D\u9884\u7EA6\u53C2\u89C8\u3002\u5BB6\u5C5E\u966A\u540C\u9700\u53E6\u884C\u9884\u7EA6\u8D2D\u7968\u3002"));
content.push(bullet("\u672A\u6EE114\u5468\u5C81\u672A\u6210\u5E74\u4EBA\u9700\u7531\u6210\u5E74\u4EBA\u966A\u540C\u3002\u5EFA\u8BAE1\u540D\u5BB6\u5C5E\u966A\u540C\u8001\u4EBA\u4E00\u8D77\u4ECE\u7231\u5FC3\u901A\u9053\u5165\u5BAB\u3002"));

content.push(para("\u6E29\u99A8\u63D0\u793A", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9884\u7EA6\u65B9\u5F0F\uFF1A\u201C\u6545\u5BAB\u535A\u7269\u9662\u201D\u5FAE\u4FE1\u5C0F\u7A0B\u5E8F\u5B9E\u540D\u9884\u7EA6\u3002\u6700\u65E9\u53EF\u4E8E\u53C2\u89C87\u65E5\u524D20:00\u5F00\u59CB\u9884\u8BA2\u3002\u73CD\u5B9D\u9986\u3001\u949F\u8868\u9986\u9700\u5355\u72EC\u9884\u7EA6\u3002\u73B0\u573A\u4E0D\u552E\u7968\u3002"));
content.push(bullet("\u5165\u9662\u4EC5\u9650\u5357\u9762\u5348\u95E8\uFF0C\u53C2\u89C8\u7ED3\u675F\u540E\u4ECE\u795E\u6B66\u95E8\u6216\u4E1C\u534E\u95E8\u79BB\u9662\u3002"));
content.push(bullet("\u643A\u5E26\u5927\u4EF6\u884C\u674E\u9700\u5148\u53BB\u7AEF\u95E8\u5E7F\u573A\u897F\u4FA7\u89C2\u4F17\u670D\u52A1\u4E2D\u5FC3\u5B58\u653E\uFF0C\u518D\u5230\u5348\u95E8\u5B89\u68C0\u533A\uFF0C\u4EE5\u514D\u91CD\u590D\u6392\u961F\u3002"));
content.push(bullet("\u5FC5\u770B\u770B\u70B9\uFF1A\u5348\u95E8\u3001\u592A\u548C\u6BBF\u3001\u4E2D\u548C\u6BBF\u3001\u4FDD\u548C\u6BBF\u3001\u4E7E\u6E05\u5BAB\u3001\u5764\u5B81\u5BAB\u3001\u5FA1\u82B1\u56ED\u3001\u73CD\u5B9D\u9986\u3001\u949F\u8868\u9986\u7B49\u3002"));
content.push(bullet("\u5EFA\u8BAE\u6E38\u89C8\u8DEF\u7EBF\uFF1A\u5348\u95E8\u2192\u592A\u548C\u95E8\u2192\u4E09\u5927\u6BBF\u2192\u4E7E\u6E05\u95E8\u2192\u540E\u5BAB\u2192\u5FA1\u82B1\u56ED\u2192\u795E\u6B66\u95E8\u51FA\u3002"));
content.push(bullet("\u670D\u52A1\u70ED\u7EBF\uFF1A400-950-1925\u3002"));

content.push(spacer());

// 5. Central Gift Center
content.push(heading("5. \u4E2D\u592E\u793C\u54C1\u6587\u7269\u7BA1\u7406\u4E2D\u5FC3\uFF08\u53EF\u9009\uFF09", HeadingLevel.HEADING_2));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u4E1C\u57CE\u533A\u897F\u5174\u9686\u88571\u53F7\uFF08\u8FD1\u5730\u94C12/5\u53F7\u7EBF\u5D07\u6587\u95E8\u7AD9H\u53E3\uFF0C\u6B65\u884C\u7EA612\u5206\u949F\uFF09"],
  ["\u5F00\u653E\u65F6\u95F4", "\u5468\u4E8C\u81F3\u5468\u65E5 9:00-17:00\uFF0816:00\u505C\u6B62\u5165\u573A\uFF09\uFF0C\u5468\u4E00\u95ED\u9986"],
  ["\u95E8\u7968\u4EF7\u683C", "\u514D\u8D39"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "2-3\u5C0F\u65F6"],
]));

content.push(para("70\u5468\u5C81\u8001\u5E74\u4EBA\u7EFF\u8272\u901A\u9053", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u672A\u627E\u5230\u660E\u786E\u201C70\u5468\u5C81\u7EFF\u8272\u901A\u9053\u201D\u8868\u8FF0\u3002\u4F4660\u5468\u5C81\u53CA\u4EE5\u4E0A\u8001\u5E74\u4EBA\u53EF\u51ED\u672C\u4EBA\u4F18\u5F85\u8BC1\u4EF6\u548C\u6709\u6548\u8EAB\u4EFD\u8BC1\u4EF6\u539F\u4EF6\u73B0\u573A\u9884\u7EA6\u53C2\u89C8\u3002"));
content.push(bullet("\u9AD8\u9F84\u8001\u4EBA\u3001\u884C\u52A8\u4E0D\u4FBF\u8005\u548C\u672A\u6EE114\u5468\u5C81\uFF08\u542B\uFF09\u7684\u672A\u6210\u5E74\u4EBA\u9700\u7531\u6210\u5E74\u4EBA\u966A\u62A4\u3002"));

content.push(para("\u5BB6\u5C5E\u966A\u540C\u653F\u7B56", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9AD8\u9F84\u8001\u4EBA\u3001\u884C\u52A8\u4E0D\u4FBF\u8005\u9700\u7531\u6210\u5E74\u4EBA\u966A\u62A4\u3002\u5EFA\u8BAE1\u540D\u5BB6\u5C5E\u966A\u540C\u3002"));
content.push(bullet("\u9884\u7EA6\u89C4\u5219\uFF1A\u6BCF\u4E2A\u8BC1\u4EF6\u6BCF\u65E5\u53EF\u9884\u7EA61\u6B21\uFF0C\u6BCF\u6B21\u6700\u591A\u9884\u7EA65\u4EBA\u3002"));

content.push(para("\u6E29\u99A8\u63D0\u793A", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9884\u7EA6\u65B9\u5F0F\uFF1A\u5FAE\u4FE1\u516C\u4F17\u53F7\u201C\u4E2D\u592E\u793C\u54C1\u6587\u7269\u7BA1\u7406\u4E2D\u5FC3\u201D\u5B9E\u540D\u9884\u7EA6\u3002\u9700\u63D0\u524D1\u81F37\u5929\u9884\u7EA6\uFF0C\u4E0D\u63A5\u53D7\u73B0\u573A\u9884\u7EA6\u300260\u5468\u5C81\u53CA\u4EE5\u4E0A\u8001\u5E74\u4EBA\u53EF\u51ED\u4F18\u5F85\u8BC1\u4EF6\u548C\u6709\u6548\u8EAB\u4EFD\u8BC1\u4EF6\u539F\u4EF6\u73B0\u573A\u9884\u7EA6\u3002"));
content.push(bullet("\u9884\u7EA6\u8D44\u683C1\u6B21\u6709\u6548\uFF0C\u51FA\u573A\u9986\u540E\u4E0D\u80FD\u518D\u6B21\u8FDB\u5165\u3002"));
content.push(bullet("\u5982\u65E0\u6CD5\u6309\u9884\u7EA6\u65F6\u95F4\u53C2\u89C8\uFF0C\u8BF7\u4E8E\u53C2\u89C8\u5F53\u592916:00\u524D\u53D6\u6D88\u9884\u7EA6\u3002\u7D2F\u8BA13\u6B21\u903E\u671F\u672A\u53D6\u6D88\u9884\u7EA6\uFF0C90\u5929\u5185\u4E0D\u80FD\u9884\u7EA6\u3002"));
content.push(bullet("\u7981\u6B62\u643A\u5E26\u89C4\u683C\u8D85\u8FC7\u957F51CM\u3001\u5BBD30CM\u3001\u9AD876.5CM\u7684\u5927\u578B\u884C\u674E\uFF0C\u8BF7\u52FF\u5E26\u5BA0\u7269\u5165\u573A\u3002"));
content.push(bullet("\u8BF7\u52FF\u4F7F\u7528\u5404\u7C7B\u7535\u52A8\u8F6E\u6905\u3001\u8001\u5E74\u4EE3\u6B65\u8F66\u3001\u6ED1\u677F\u8F66\u3001\u5E73\u8861\u8F66\u7B49\uFF0C\u5982\u9700\u4F7F\u7528\u8F6E\u6905\u8BF7\u4E0E\u5DE5\u4F5C\u4EBA\u5458\u8054\u7CFB\u3002"));

content.push(spacer());

// 6. NCPA
content.push(heading("6. \u56FD\u5BB6\u5927\u5267\u9662\uFF08\u53EF\u9009\uFF09", HeadingLevel.HEADING_2));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u897F\u57CE\u533A\u897F\u957F\u5B89\u88572\u53F7\uFF08\u5929\u5B89\u95E8\u5E7F\u573A\u897F\u4FA7\uFF09"],
  ["\u5F00\u653E\u65F6\u95F4", "\u5468\u4E8C\u81F3\u5468\u65E5\u53CA\u56FD\u5BB6\u6CD5\u5B9A\u8282\u5047\u65E5 9:00-17:00\uFF1B\u5468\u4E00\u95ED\u9986"],
  ["\u95E8\u7968\u4EF7\u683C", "\u53C2\u89C8\u796840\u5143/\u4EBA\uFF1B60-69\u5468\u5C81\u534A\u4EF720\u5143\uFF1B70\u5468\u5C81\u53CA\u4EE5\u4E0A\u514D\u7968"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "2-3\u5C0F\u65F6"],
]));

content.push(para("70\u5468\u5C81\u8001\u5E74\u4EBA\u7EFF\u8272\u901A\u9053", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("70\u5468\u5C81\u53CA\u4EE5\u4E0A\u8001\u5E74\u4EBA\u51ED\u6709\u6548\u8BC1\u4EF6\u514D\u7968\u300260-69\u5468\u5C81\u4EAB\u53D7\u534A\u4EF7\u4F18\u60E0\uFF0820\u5143\uFF09\u3002"));
content.push(bullet("\u5B98\u65B9\u65E0\u969C\u788D\u8BBE\u65BD\u5305\u62EC\u4E13\u95E8\u7684\u5B89\u68C0\u901A\u9053\u53CA\u4EBA\u5DE5\u5B89\u68C0\u3001\u6B8B\u75BE\u4EBA\u4E13\u7528\u536B\u751F\u95F4\u3001\u76F4\u68AF\u7B49\u8BBE\u65BD\uFF0C\u53EF\u51CF\u5C11\u7B49\u5F85\u5E76\u65B9\u4FBF\u901A\u884C\u3002"));

content.push(para("\u5BB6\u5C5E\u966A\u540C\u653F\u7B56", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("6\u5468\u5C81\u4EE5\u4E0B\u6216\u8EAB\u9AD81.2\u7C73\u4EE5\u4E0B\u513F\u7AE5\uFF08\u9700\u6709\u6210\u4EBA\u966A\u4F34\uFF09\u53EF\u514D\u8D39\u53C2\u89C8\u3002"));
content.push(bullet("70\u5468\u5C81\u4EE5\u4E0A\u8001\u4EBA\u867D\u65E0\u5F3A\u5236\u966A\u540C\u8981\u6C42\uFF0C\u4F46\u5EFA\u8BAE1\u540D\u5BB6\u5C5E\u966A\u540C\uFF0C\u65B9\u4FBF\u7167\u987E\u3002"));

content.push(para("\u6E29\u99A8\u63D0\u793A", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9884\u7EA6\u65B9\u5F0F\uFF1A\u901A\u8FC7\u56FD\u5BB6\u5927\u5267\u9662\u5B98\u65B9\u5E73\u53F0\u8D2D\u7968/\u9884\u7EA6\u3002\u6BCF\u4E2A\u8BC1\u4EF6\u5728\u6BCF\u4E2A\u5165\u9662\u65E5\u9650\u8D2D\u4E00\u5F20\u95E8\u7968\uFF0C\u6BCF\u4E2A\u8BA2\u5355\u6700\u591A\u53EF\u9884\u8BA110\u4EBA\u3002"));
content.push(bullet("\u53C2\u89C8\u65F6\u53EF\u643A\u5E26\u76F8\u673A\u5165\u573A\uFF08\u8BF7\u52FF\u643A\u5E26\u4E09\u811A\u67B6\uFF09\u3002\u4F53\u79EF\u8F83\u5C0F\u7684\u4FBF\u643A\u5305\u53EF\u5E26\u5165\uFF0C\u4F53\u79EF\u5927\u7684\u5305\u9700\u5728\u5317\u95E8\u5904\u8863\u5E3D\u95F4\u5B58\u5305\u3002"));
content.push(bullet("\u4E0D\u80FD\u643A\u5E26\u98DF\u7269\u548C\u996E\u6599\u8FDB\u5165\u5267\u9662\u3002"));
content.push(bullet("\u5317\u6C34\u4E0B\u5ECA\u9053\u6709\u514D\u8D39\u8BB2\u89E3\u670D\u52A1\u548C\u5E38\u8BBE\u514D\u8D39\u996E\u6C34\u5904\u3002"));
content.push(bullet("\u5927\u5385\u53CA\u516C\u5171\u7A7A\u95F4\u8DEF\u9762\u5766\u5766\uFF0C\u5B9A\u671F\u8FDB\u884C\u9632\u6ED1\u5904\u7406\uFF0C\u9002\u5408\u8001\u5E74\u4EBA\u884C\u8D70\u3002"));

content.push(spacer());

// 7. Olympic Park
content.push(heading("7. \u9E1F\u5DE2\u3001\u6C34\u7ACB\u65B9\u7B49\u5965\u6797\u5339\u514B\u516C\u56ED\uFF08\u53EF\u9009\uFF09", HeadingLevel.HEADING_2));

content.push(para("\u9E1F\u5DE2\uFF08\u56FD\u5BB6\u4F53\u80B2\u573A\uFF09", { bold: true, before: 160 }));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u671D\u9633\u533A\u56FD\u5BB6\u4F53\u80B2\u573A\u5357\u8DEF1\u53F7"],
  ["\u5F00\u653E\u65F6\u95F4", "10:00-22:00\uFF0821:00\u505C\u6B62\u68C0\u7968\uFF09"],
  ["\u95E8\u7968\u4EF7\u683C", "\u6210\u4EBA\u796850\u5143\uFF1B70\u5468\u5C81\u4EE5\u4E0A\u514D\u7968\uFF1B60-69\u5468\u5C81\u3001\u5B66\u751F\u7B4925\u5143\uFF08\u534A\u4EF7\uFF09"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u7EA61-2\u5C0F\u65F6"],
]));
content.push(bullet("70\u5468\u5C81\u4EE5\u4E0A\u8001\u4EBA\u514D\u7968\u3002\u9E1F\u5DE2\u63D0\u4F9B\u65E0\u969C\u788D\u4E13\u7528\u901A\u9053\u3001\u65E0\u969C\u788D\u5750\u5E2D\u53CA\u8F85\u52A9\u8BBE\u65BD\uFF0C\u5B89\u6392\u4E13\u4EBA\u5168\u7A0B\u5F15\u5BFC\u5E2E\u6276\uFF0C\u4FDD\u969C\u884C\u52A8\u4E0D\u4FBF\u3001\u8001\u5E74\u7FA4\u4F53\u7B49\u6709\u9700\u8981\u7684\u89C2\u4F17\u3002"));
content.push(bullet("1.2\u7C73\u4EE5\u4E0B\u513F\u7AE5\u514D\u7968\uFF0C\u9700\u6709\u6210\u4EBA\u966A\u540C\u5165\u573A\u3002"));
content.push(bullet("\u8D2D\u7968\u65B9\u5F0F\uFF1A\u5B98\u65B9\u8D2D\u7968\u5C0F\u7A0B\u5E8F\u6216\u73B0\u573A\u552E\u7968\uFF08\u9E1F\u5DE2\u5357DE\u53E3\u5916\u3001\u5317KL\u53E3\u5916\u53CA\u897F\u5165\u53E3\uFF09\u3002\u670D\u52A1\u70ED\u7EBF\uFF1A400-600-2008\u3002"));
content.push(bullet("\u770B\u70B9\uFF1A\u91D1\u8272\u5927\u5385\u53CC\u5965\u5C55\u533A\u3001\u7965\u4E91\u91D1\u5385\u3001\u6D41\u6C34\u94F6\u5385\u3001\u5965\u8FD0\u6587\u5316\u9057\u4EA7\u957F\u5ECA\u3001\u94A2\u7ED3\u6784\u68EE\u6797\u3001\u7A7A\u4E2D\u89C2\u5149\u8D70\u5ECA\u7B49\u3002"));

content.push(para("\u6C34\u7ACB\u65B9\uFF08\u56FD\u5BB6\u6E38\u6CF3\u4E2D\u5FC3\uFF09", { bold: true, before: 160 }));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u671D\u9633\u533A\u5965\u6797\u5339\u514B\u516C\u56ED\u6838\u5FC3\u533A"],
  ["\u5F00\u653E\u65F6\u95F4", "10:00-22:30\uFF0822:00\u505C\u6B62\u5165\u573A\uFF09\uFF0C\u5177\u4F53\u4EE5\u573A\u9986\u5F53\u65E5\u516C\u793A\u4E3A\u51C6"],
  ["\u95E8\u7968\u4EF7\u683C", "\u5168\u4EF7\u796830\u5143\uFF1B60-64\u5468\u5C81\u534A\u4EF715\u5143\uFF1B65\u5468\u5C81\u53CA\u4EE5\u4E0A\u514D\u8D39"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u7EA61\u5C0F\u65F6"],
]));
content.push(bullet("65\u5468\u5C81\u53CA\u4EE5\u4E0A\u8001\u4EBA\u514D\u8D39\u3002\u4EAB\u53D7\u4F18\u60E0\u653F\u7B56\u8005\u9700\u5728\u8D2D\u7968\u53CA\u68C0\u7968\u65F6\u51FA\u5177\u672C\u4EBA\u76F8\u5173\u8BC1\u4EF6\u539F\u4EF6\u3002"));
content.push(bullet("6\u5468\u5C81\uFF08\u542B\uFF09\u4EE5\u4E0B\u6216\u8EAB\u9AD81.2\u7C73\uFF08\u542B\uFF09\u4EE5\u4E0B\u513F\u7AE5\u9700\u6709\u6210\u5E74\u4EBA\u966A\u4F34\u5165\u5185\u514D\u8D39\u53C2\u89C8\u3002"));
content.push(bullet("\u8D2D\u7968\u65B9\u5F0F\uFF1A\u5173\u6CE8\u201C\u6C34\u7ACB\u65B9\u201D\u516C\u4F17\u53F7\u8D2D\u7968\u3002\u5168\u56FD\u7EDF\u4E00\u670D\u52A1\u70ED\u7EBF\uFF1A4006202022\u3002"));
content.push(bullet("\u5982\u9700\u53C2\u89C8\u201C\u6C34\u4E4B\u5149\u201D\u5149\u5F71\u79C0\uFF0C\u6CE8\u610F\u6D3B\u52A8\u8C03\u6574\u53EF\u80FD\u5F71\u54CD\u9879\u76EE\u5F00\u653E\u3002"));

content.push(para("\u5965\u6797\u5339\u514B\u68EE\u6797\u516C\u56ED\uFF08\u5965\u68EE\uFF09", { bold: true, before: 160 }));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u671D\u9633\u533A\u79D1\u835E\u8DEF33\u53F7"],
  ["\u5F00\u653E\u65F6\u95F4", "\u65FA\u5B63\uFF083\u670815\u65E5-11\u670815\u65E5\uFF096:00-20:00\uFF1B\u6DE1\u5B63 7:00-19:00"],
  ["\u95E8\u7968\u4EF7\u683C", "\u514D\u8D39\uFF0C\u65E0\u9700\u9884\u7EA6"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u534A\u5929\u81F3\u5168\u5929"],
]));
content.push(bullet("\u516C\u56ED\u514D\u8D39\u5F00\u653E\uFF0C60\u5C81\u4EE5\u4E0A\u8001\u4EBA\u51ED\u8EAB\u4EFD\u8BC1\u53EF\u76F4\u63A5\u5165\u56ED\uFF0C\u65E0\u9700\u989D\u5916\u95E8\u7968\u6216\u9884\u7EA6\u3002"));
content.push(bullet("\u516C\u56ED\u9762\u79EF\u8F83\u5927\uFF0C\u9002\u5408\u6563\u6B65\u3001\u6162\u8DD1\u3001\u9A91\u884C\u3002\u914D\u5907\u4F11\u606F\u5EA7\u6905\u3001\u536B\u751F\u95F4\u3001\u996E\u6C34\u5904\u7B49\u8BBE\u65BD\u3002"));
content.push(bullet("\u5EFA\u8BAE\u6839\u636E\u8001\u4EBA\u4F53\u529B\u9009\u62E9\u90E8\u5206\u533A\u57DF\u6E38\u89C8\uFF0C\u4E0D\u5FC5\u8D70\u5B8C\u5168\u7A0B\u3002"));

content.push(spacer());

// 8. Temple of Heaven
content.push(heading("8. \u5929\u575B\u516C\u56ED\uFF08\u53EF\u9009\uFF09", HeadingLevel.HEADING_2));
content.push(infoTable([
  ["\u5730\u5740", "\u5317\u4EAC\u5E02\u4E1C\u57CE\u533A\u5929\u575B\u5185\u4E1C\u91CC7\u53F7"],
  ["\u5F00\u653E\u65F6\u95F4", "\u516C\u56ED\uFF1A\u65FA\u5B63\uFF084\u67081\u65E5-10\u670831\u65E5\uFF096:00-22:00\uFF0821:00\u505C\u6B62\u5165\u56ED\uFF09\uFF1B\u6DE1\u5B63\uFF0811\u67081\u65E5-3\u670831\u65E5\uFF096:30-22:00\u3002\u6838\u5FC3\u666F\u70B9\uFF08\u7948\u5E74\u6BBF\u3001\u56DE\u97F3\u58C1\u3001\u56DE\u4E18\uFF09\uFF1A\u65FA\u5B638:00-17:30\uFF0817:00\u505C\u6B62\u8FDB\u5165\uFF09\uFF1B\u6DE1\u5B638:00-17:00\uFF0816:30\u505C\u6B62\u8FDB\u5165\uFF09\u3002\u5468\u4E00\u6838\u5FC3\u666F\u70B9\u5173\u95ED\uFF08\u6CD5\u5B9A\u8282\u5047\u65E5\u9664\u5916\uFF09"],
  ["\u95E8\u7968\u4EF7\u683C", "\u65FA\u5B63\u5927\u95E8\u796815\u5143\uFF0C\u8054\u796834\u5143\uFF1B\u6DE1\u5B63\u5927\u95E8\u796810\u5143\uFF0C\u8054\u796828\u5143\u3002\u8054\u7968\u5305\u542B\u7948\u5E74\u6BBF\u3001\u56DE\u97F3\u58C1\u3001\u56DE\u4E18\u300260\u5468\u5C81\u53CA\u4EE5\u4E0A\u514D\u8D39\uFF08\u542B\u6838\u5FC3\u666F\u70B9\uFF09"],
  ["\u9884\u8BA1\u6E38\u89C8\u65F6\u957F", "\u542B\u6838\u5FC3\u666F\u70B92-3\u5C0F\u65F6\uFF1B\u4EC5\u516C\u56ED\u6563\u6B651-2\u5C0F\u65F6"],
]));

content.push(para("70\u5468\u5C81\u8001\u5E74\u4EBA\u7EFF\u8272\u901A\u9053", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("60\u5468\u5C81\u53CA\u4EE5\u4E0A\u8001\u4EBA\u51ED\u8EAB\u4EFD\u8BC1\u514D\u8D39\u5165\u56ED\uFF0C\u7948\u5E74\u6BBF\u3001\u56DE\u97F3\u58C1\u3001\u56DE\u4E18\u7B49\u6838\u5FC3\u666F\u70B9\u5168\u90E8\u514D\u8D39\u53C2\u89C8\uFF0C\u4E0D\u9650\u6237\u7C4D\u3002"));
content.push(bullet("\u5929\u575B\u516C\u56ED\u5730\u52BF\u5766\u5766\uFF0C\u5168\u56ED\u65E0\u53F0\u9636\uFF0C\u53E4\u67CF\u6811\u6797\u906E\u9634\uFF0C\u5BF9\u8001\u5E74\u4EBA\u975E\u5E38\u53CB\u597D\u3002"));
content.push(bullet("\u516C\u56ED\u672A\u5355\u72EC\u5217\u51FA\u201C70\u5468\u5C81\u7EFF\u8272\u901A\u9053\u201D\u4E13\u9879\u89C4\u5B9A\uFF0C\u4F46\u56ED\u5185\u9053\u8DEF\u5766\u5766\uFF0C\u901A\u884C\u4FBF\u5229\uFF0C\u9002\u5408\u8001\u4EBA\u6E38\u89C8\u3002"));

content.push(para("\u5BB6\u5C5E\u966A\u540C\u653F\u7B56", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("60\u5468\u5C81\u4EE5\u4E0A\u8001\u4EBA\u514D\u8D39\u5165\u56ED\u65E0\u9700\u989D\u5916\u966A\u540C\u9884\u7EA6\uFF0C\u5BB6\u5C5E\u9700\u81EA\u884C\u9884\u7EA6\u8D2D\u7968\u5165\u56ED\u3002"));
content.push(bullet("\u5EFA\u8BAE1\u540D\u5BB6\u5C5E\u966A\u540C\u8001\u4EBA\u6E38\u89C8\uFF0C\u56ED\u5185\u9762\u79EF\u8F83\u5927\uFF0C\u65B9\u4FBF\u7167\u5E94\u3002"));

content.push(para("\u6E29\u99A8\u63D0\u793A", { bold: true, color: "C00000", before: 200 }));
content.push(bullet("\u9884\u7EA6\u65B9\u5F0F\uFF1A\u201C\u5929\u575B\u201D\u5FAE\u4FE1\u516C\u4F17\u53F7\u3001\u201C\u7545\u6E38\u516C\u56ED\u201D\u516C\u4F17\u53F7\u3001\u5929\u575B\u516C\u56ED\u5B98\u7F51\u3001\u4EAC\u901A\u5C0F\u7A0B\u5E8F\u3002\u514D\u8D39\u4EBA\u7FA4\u4E5F\u5FC5\u987B\u9884\u7EA6\uFF0C\u514D\u8D39\u4E0D\u7B49\u4E8E\u514D\u9884\u7EA6\u3002"));
content.push(bullet("\u63D0\u524D7\u5929\u53EF\u9884\u7EA6\uFF0C\u6BCF\u65E50:00\u653E\u7968\u3002\u65FA\u5B63\u5468\u672B\u53CA\u8282\u5047\u65E5\u8054\u7968\u62A2\u624B\uFF0C\u5EFA\u8BAE\u5DE5\u4F5C\u65E5\u524D\u5F80\u3002"));
content.push(bullet("\u63A8\u8350\u4E1C\u95E8\u5165\u56ED\uFF0C\u76F4\u8FBE\u7948\u5E74\u6BBF\u6838\u5FC3\u666F\u533A\uFF0C\u6E38\u89C8\u8DEF\u7EBF\u6700\u77ED\u3002"));
content.push(bullet("\u5FC5\u770B\u666F\u70B9\uFF1A\u7948\u5E74\u6BBF\u3001\u56DE\u97F3\u58C1\u3001\u56DE\u4E18\u3001\u4E39\u9646\u6865\u3001\u7687\u7A79\u5BAE\u3001\u795E\u4E50\u7F72\u7B49\u3002"));
content.push(bullet("\u5EFA\u8BAE\u6E38\u89C8\u8DEF\u7EBF\uFF1A\u4E1C\u95E8\u2192\u7948\u5E74\u6BBF\u2192\u4E39\u9646\u6865\u2192\u7687\u7A79\u5BAE\uFF08\u56DE\u97F3\u58C1\uFF09\u2192\u56DE\u4E18\u2192\u5357\u95E8\u51FA\uFF08\u6216\u539F\u8DEF\u8FD4\u4E1C\u95E8\u51FA\uFF09\u3002"));
content.push(bullet("\u516C\u56ED\u5185\u6709\u7535\u74F6\u8F66\uFF0C20\u5143/\u4EBA\uFF0C\u53EF\u4EE3\u6B65\u5F80\u8FD4\u6838\u5FC3\u666F\u70B9\uFF0C\u8001\u4EBA\u4F53\u529B\u4E0D\u652F\u53EF\u8003\u8651\u4E58\u5750\u3002"));
content.push(bullet("\u6E05\u6668\u53EF\u770B\u5230\u672C\u5730\u8001\u4EBA\u6668\u7EC3\u3001\u6253\u592A\u6781\uFF0C\u6C1B\u56F4\u5B89\u9038\u8212\u7F13\uFF0C\u9002\u5408\u6162\u6E38\u3002"));
content.push(bullet("\u54A8\u8BE2\u7535\u8BDD\uFF1A010-67028866\u3002"));

content.push(spacer());

content.push(new Paragraph({ children: [new PageBreak()] }));

// Section 2 - 北京吃食故事
content.push(heading("二、北京吃食——一口京城味，半部烟火史", HeadingLevel.HEADING_1));
content.push(storyPara("北京这地界儿，六朝古都，天子脚下。王公贵胄吃的是排场，市井百姓嚼的是日子。这一出一入之间，便生出无数吃食的故事来。今儿不荐店，不论价，单说这些吃食的前世今生，权当饭桌上多添一碟下酒的谈资。"));

content.push(spacer());

// 1. 炙子烤肉
content.push(heading("1. 炙子烤肉——铁条上的蒙古风", HeadingLevel.HEADING_2));
content.push(storyPara("要说烤肉，得从蒙古人说起。当年成吉思汗的铁骑踏遍欧亚，行军打仗，没锅没灶，头盔往火上一架，肉片往上一摊，便是烤肉。这法子传到北京，让回回们拾掇出了新花样。"));
content.push(storyPara("所谓“炙子”，是一副铁条焊成的烤盘，形如百叶窗，下面烧松枝或炭火。肉片事先用酱油、料酒、葱姜末腌过，往炙子上一摊，滋啦一声，香气便顺着烟气往上蹿。讲究的主儿，自己端着盘子站着烤，一只脚踏在板凳上，边烤边吃，这叫“武吃”。文雅的便坐着让伙计烤好端上来，这叫“文吃”。"));
content.push(storyPara("老北京烤肉有“三大家”——烤肉宛、烤肉季、烤肉刘。烤肉宛在宣武门，打咸丰年间就支摊子了；烤肉季在什刹海，是个姓季的回回先摆的浮摊，后来发了家；烤肉刘在友谊医院边上，年头最晚，口碑不差。三家各有各的绝活，但规矩一样：肉只选牛羊上脑，片得薄如纸，迎风能飘起来。"));

content.push(spacer());

// 2. 门钉肉饼
content.push(heading("2. 门钉肉饼——紫禁城门上的金疙瘩", HeadingLevel.HEADING_2));
content.push(storyPara("门钉肉饼这名儿，透着京城独有的那份讲究。"));
content.push(storyPara("紫禁城的门上钉着铜鎏金门钉，九路九排，九九八十一个，取的是“九五之尊”的意思。这肉饼做出来，圆鼓鼓、厚墩墩，正面还捏出一圈花边，往那儿一摆，活脱脱就是一个门钉。"));
content.push(storyPara("传说是这么说的：慈禧太后有一回嘴馋，传膳时御膳房犯了难——山珍海味老佛爷吃腻了。有个厨子灵机一动，拿牛肉大葱做馅，包进面皮里，烙得两面金黄，端上去。太后一看，嘿，这不就是宫门上的门钉吗？一口咬下去，皮脆馅香，汤汁四溢，凤颜大悦。"));
content.push(storyPara("这故事真假难辨，但门钉肉饼的讲究是真的。皮要半烫面，凉了不硬；馅要牛肉大葱，七分肉三分葱，多了腻少了寡。煎的时候先烙正面，锁住花边，再翻面烙底，两面金黄才出锅。趁热吃，一口下去那汁水烫嘴，但舍不得吐——这就是门钉肉饼的厉害。"));

content.push(spacer());

// 3. 南门涮肉
content.push(heading("3. 南门涮肉——铜锅里的涮法儿", HeadingLevel.HEADING_2));
content.push(storyPara("北京人涮羊肉，不叫“火锅”，叫“涮锅子”。一个铜锅，中间是烟囱，底下烧炭，水开得翻花，夹一片薄如蝉翼的羊肉，在滚水里一涮——就那么三五秒——肉色一变就捞，蘸麻酱小料吃。这涮法儿，讲究的是“快”字，涮久了肉老，涮快了不熟，差一秒都不行。"));
content.push(storyPara("铜锅也有学问。为什么非得铜的？铜导热快，水翻得急，肉片下去不降温度，出来才嫩。那烟囱不叫烟囱，叫“火筒”，烧的是硬木炭，没烟没味。要是烧了带松脂的炭，羊肉就串了味儿，老饕一吃便知。"));
content.push(storyPara("小料是灵魂。芝麻酱打底，韭菜花提鲜，酱豆腐增醇，虾油点睛，最后撒一把香菜末。各家小料配方是命根子，传男不传女。其实说白了就是比例不同，但差一点味儿就不对，你说玄不玄？"));
content.push(storyPara("涮肉的肉也讲究，只取羊后腿和上脑，俗称“黄瓜条”的部位。手工切，片要大、薄、匀。以前有名的大师傅，一刀下去，片片透光，往墙上贴能粘住。如今切肉多用机器，快是快了，但少了那份手艺活的精气神。"));

content.push(spacer());

// 4. 北京烤鸭
content.push(heading("4. 北京烤鸭——挂炉与焖炉的百年之争", HeadingLevel.HEADING_2));
content.push(storyPara("北京烤鸭，分两派：挂炉与焖炉。"));
content.push(storyPara("挂炉派以全聚德为首。炉膛半开放，果木明火烤，鸭子挂里头，火苗舔着鸭皮，油脂一滴滴落入炉底的铁盘里，噼啪作响。老师傅拿着长杆子，隔一会儿就把鸭子换个位置，务使受热均匀。烤出来的鸭子，皮脆如玻璃，肉嫩似豆腐，片的时候得连皮带肉，薄厚均匀。全聚德从同治三年（1864年）开张至今，用的是“北京填鸭”，个头大，肉厚实。"));
content.push(storyPara("焖炉派以便宜坊为代表。焖炉是封闭的，不直接见明火。炉壁先烧热，再把鸭子放进去，关上炉门，用炉壁的余温把鸭子“焖”熟。烤出来的鸭子皮肉之间多了一层油脂，口感更滋润。便宜坊比全聚德还老，永乐年间就在菜市口摆摊了，算下来六百多年了。"));
content.push(storyPara("两派争了一百多年，谁也说服不了谁。挂炉的嫌焖炉的不够脆，焖炉的嫌挂炉的太干。其实各有各的好，就看你好哪一口。但有一件事是两派共同的规矩：鸭子只能烤，不能炸。你要是在北京跟人说吃“炸鸭”，人家能跟你急。"));
content.push(storyPara("片鸭子也是绝活。一整只鸭子，片一百零八片，片片带皮带肉，不多不少。这是老师傅的功夫，如今能片到一百片的就算好手了。片完的鸭架子别扔，熬白菜汤，撒把盐，那汤鲜得能把眉毛鲜掉。"));

content.push(spacer());

// 5. 卤煮
content.push(heading("5. 卤煮——穷人嘴里的御膳", HeadingLevel.HEADING_2));
content.push(storyPara("卤煮火烧，听着俗，吃着香，是老北京南城穷苦人嘴里蹦出来的宝贝。"));
content.push(storyPara("早年间，皇宫里吃苏造肉——那是拿五花大片肉和各种香料炖出来的，香得能让人从东华门闻到西华门。可御膳的东西，老百姓哪吃得起？但御膳房的太监们有法子：把苏造肉的配方偷出来，拿老百姓吃得起的猪下水——肠子、肺头、豆腐——替了那五花肉，再加上火烧（面饼），一锅炖了。这便是卤煮的由来。"));
content.push(storyPara("小肠陈是卤煮的头一块招牌。打清末就在南城摆摊，传到如今已是第四代。老汤是命根子，那锅老汤据说从开业那天起就没断过火，每日添水加料，百年下来，汤里的味道一层叠一层，跟城墙根的土似的，年岁越长越厚实。"));
content.push(storyPara("吃卤煮有讲究。先看那汤色——得是酱红色，浓而不浑。火烧切成井字块，豆腐三角块，肠子切段，肺头切片，码在碗底，浇上滚汤，再搁蒜泥、韭菜花、腐乳汁。一口下去，肠子弹牙，肺头软烂，火烧吸饱了汤汁但不烂，嚼着有劲儿。冬天来一碗，热气从肚子里往上顶，能把一天的寒气全赶跑。"));
content.push(storyPara("有人说卤煮是“下水做的，上不得台面”。北京人听了这话不乐意：你知道什么？这卤煮里头，有宫里的方子，有民间的智慧，有百年的老汤，有一代代手艺人的命。它不是上不了台面，它就是台面。"));

content.push(spacer());

// 6. 豆汁
content.push(heading("6. 豆汁——爱之如命，恨之入骨", HeadingLevel.HEADING_2));
content.push(storyPara("豆汁这东西，是北京吃食里最能分出“自己人”和“外人”的试金石。"));
content.push(storyPara("绿豆泡涨，磨成浆，滤去豆渣，剩下的液体搁在那儿发酵。发到发酸发馊，表面起一层沫子，这就成了豆汁。灰绿色的，闻着一股酸腐味儿，头回见的人，十个有九个想吐。"));
content.push(storyPara("可北京人爱它爱得要命。老北京有句话：“没喝过豆汁儿，不算到过北京。”这话不假。豆汁喝进嘴里，酸、甜、馊、涩，什么味都有，但咽下去之后，回甘从嗓子眼儿往上冒，浑身通透。配一碗豆汁的，必须是焦圈和咸菜丝。焦圈炸得酥脆，豆汁滚烫，一凉一热，一脆一滑，对冲之下，那股子酸馊味竟变得妙不可言。"));
content.push(storyPara("豆汁的来历也有意思。原本是做绿豆淀粉时的下脚料，穷人舍不得倒，搁着搁着发酵了，一尝，嘿，还别有风味。后来连宫里都喝上了。乾隆爷好这口，专门让御膳房每日备着。据说有一次乾隆召见大臣，端出豆汁来赐饮，大臣们捏着鼻子灌下去，脸上还得堆着笑说“好喝”。这画面，想想就好笑。"));
content.push(storyPara("如今喝豆汁的地方不多了。老磁器口豆汁店、尹三豆汁，数得过来的就那几家。坐的也都是老主顾，大爷大妈们端着碗，就着焦圈，一口豆汁一口咸菜，吃的是回忆，品的是岁月。你要是头一回喝，别急着吐——忍住第一口，第二口就顺了，到第三口，你可能就上瘾了。当然，也可能直接放下碗再也不碰。这就是豆汁的脾气：它不迎合任何人。"));

content.push(spacer());

// 7. 爆肚
content.push(heading("7. 爆肚——三秒定乾坤", HeadingLevel.HEADING_2));
content.push(storyPara("爆肚，一个“爆”字，写尽了这道吃食的全部学问。"));
content.push(storyPara("牛百叶也好，羊肚领也罢，切好了往滚水里一汆——多久？三秒。多了老了，少了生了，就这三秒，一秒不能差。捞出来过凉水，装盘，蘸麻酱吃。就这么简单的东西，但做到极致就是绝活。"));
content.push(storyPara("爆肚分十三个部位，各有各的名堂：肚仁、肚领、肚板、肚丝、蘑菇头、散丹……每个部位汆的时间不同，短的几秒，长的十几秒。一个老掌柜，靠一锅滚水和一双筷子，能在不同部位之间切换自如，毫厘不差。这门手艺，不是三年五年能练出来的。"));
content.push(storyPara("蘸料也讲究。麻酱、酱油、醋、辣油、香菜、葱花，各有比例。但各家最大的不同，在于那一碟“爆肚蘸料”的秘方——据说有用虾油的，有用黄酒的，有的还搁一点点白糖提鲜。这方子，掌柜的比自家存折还看得紧。"));
content.push(storyPara("吃爆肚最高境界：端上来不蘸料，先吃一口原味。好的爆肚，本身带着一股子清甜和脆嫩，什么料都不蘸，光嚼那口脆劲儿，就是满嘴的香。蘸料是锦上添花，不是遮丑的粉。"));

content.push(spacer());

// 8. 炒肝
content.push(heading("8. 炒肝——名不副实的醇厚", HeadingLevel.HEADING_2));
content.push(storyPara("炒肝不炒，里面也不仅仅有肝。"));
content.push(storyPara("这东西其实是猪大肠和猪肝，用蒜和黄酱熬成的浓汤，勾芡而成。叫“炒”是虚的，“煮”才是实的；叫“肝”是虚的，“肠”才是主角。一碗炒肝端上来，先看见的是黑红浓稠的汤，用勺子一搅，大肠段和肝片便浮上来，蒜香扑鼻。"));
content.push(storyPara("天兴居的炒肝最有名，从咸丰年间就开张了。老北京吃炒肝讲究不用勺，端着碗转着喝——因为勾芡浓稠，碗边一吸溜就进嘴了，勺子反而碍事。配着包子吃，一个肉包子一碗炒肝，这顿早饭，老北京能从天蒙蒙亮吃到日上三竿。"));

content.push(spacer());

// 9. 驴打滚
content.push(heading("9. 驴打滚——滚出来的甜", HeadingLevel.HEADING_2));
content.push(storyPara("最后说个甜的。驴打滚，学名叫豆面糕，是老北京传统小吃。黄米面蒸熟，摊开抹上红豆沙，卷起来，在炒熟的黄豆面里一滚——那黄豆面沾满了一身，像极了驴子在土里打滚扬起一溜尘烟。这名字，俗到极点，妙到极致。"));
content.push(storyPara("还有艾窝窝，雪白的糯米团子，里头裹着芝麻、核桃、瓜子仁，顶部点一粒红点。这红点是点睛之笔，像美人眉心的花钿，小小一枚，精致得很。"));
content.push(storyPara("护国寺小吃店里，这些一口一个的小点心摆了一排。北京人管它们叫“京味儿”，其实里面装的全是光阴。过去庙会上，大人牵着孩子，花几文钱买一包驴打滚，孩子举着边走边吃，豆面掉了一身也不在乎。那光景，如今想起来，甜的不是点心，是那会儿的日子。"));

content.push(spacer());

content.push(storyQuote("——人间烟火气，最抚凡人心。"));

var doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }, size: 24 }
      }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, color: "1F4E79", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0, keepNext: false, keepLines: false } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, color: "2E75B6", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1, keepNext: false, keepLines: false } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: "404040", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
        paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2, keepNext: false, keepLines: false } },
    ]
  },
  numbering: { config: [ { reference: "bullets", levels: [ { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } } ] } ] },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: content
  }]
});

Packer.toBuffer(doc).then(function(buffer) {
  var outPath = "c:\\Users\\sucheng.li\\Desktop\\tare\u5206\u6790\\\u5317\u4EAC\u666F\u70B9\u6E38\u89C8\u653B\u7565\\\u5317\u4EAC\u666F\u70B9\u6E38\u89C8\u653B\u7565.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Document created: " + outPath);
});