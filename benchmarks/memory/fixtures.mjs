// Deterministic synthetic documents for the memory benchmark.
//
// One "unit" is roughly one page of ordinary business content: a heading, body paragraphs with
// mixed run formatting, a short bullet list, and a small bordered table on every second unit.
// Every page carries a header and a footer with a PAGE field. The same unit count always
// produces the same bytes (fixed zip timestamps), so a size label names the same document on
// every machine and every run.

import { zipSync, strToU8 } from 'fflate';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const FIXED_MTIME = new Date(Date.UTC(2020, 0, 1));
const WORDS = [
  'agreement',
  'balance',
  'contract',
  'delivery',
  'estimate',
  'forecast',
  'guarantee',
  'heading',
  'invoice',
  'journal',
  'ledger',
  'margin',
  'notice',
  'obligation',
  'payment',
  'quarter',
  'revenue',
  'schedule',
  'term',
  'update',
  'vendor',
  'warranty',
];

const words = (seed, count) =>
  Array.from({ length: count }, (_, i) => WORDS[(seed * 7 + i * 3) % WORDS.length]).join(' ');

const run = (text, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

function bodyParagraph(unit, index) {
  const seed = unit * 31 + index;
  return (
    `<w:p>` +
    run(`Clause ${unit + 1}.${index + 1}. `, '<w:b/>') +
    run(`${words(seed, 18)}. `) +
    run(`${words(seed + 1, 6)}`, '<w:i/>') +
    run(` ${words(seed + 2, 14)} `) +
    run(words(seed + 3, 4), '<w:color w:val="1F4E79"/><w:u w:val="single"/>') +
    run(`. ${words(seed + 4, 20)}.`) +
    `</w:p>`
  );
}

const bullet = (unit, index) =>
  `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${run(
    `Item ${index + 1}: ${words(unit * 13 + index, 9)}`
  )}</w:p>`;

function table(unit) {
  const cols = 4;
  const width = Math.floor(9360 / cols);
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${width}"/>`).join('');
  const rows = Array.from({ length: 5 }, (_, row) => {
    const cells = Array.from({ length: cols }, (_, col) => {
      const text =
        row === 0 ? `Column ${col + 1}` : `${words(unit + row * 5 + col, 3)} ${row * 100 + col}`;
      return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p>${run(text, row === 0 ? '<w:b/>' : '')}</w:p></w:tc>`;
    }).join('');
    return `<w:tr>${row === 0 ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${cells}</w:tr>`;
  }).join('');
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

function unitXml(unit) {
  const parts = [
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run(`Section ${unit + 1}: ${words(unit, 3)}`)}</w:p>`,
  ];
  for (let i = 0; i < 3; i += 1) parts.push(bodyParagraph(unit, i));
  for (let i = 0; i < 3; i += 1) parts.push(bullet(unit, i));
  if (unit % 2 === 0) parts.push(table(unit));
  for (let i = 3; i < 5; i += 1) parts.push(bodyParagraph(unit, i));
  return parts.join('');
}

const sectPr =
  '<w:sectPr><w:headerReference w:type="default" r:id="rId4"/><w:footerReference w:type="default" r:id="rId5"/>' +
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const parts = {
  '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`,
  '_rels/.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`,
  'word/styles.xml': `${XML}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`,
  'word/settings.xml': `${XML}<w:settings ${W}><w:defaultTabStop w:val="720"/></w:settings>`,
  'word/numbering.xml': `${XML}<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  'word/header1.xml': `${XML}<w:hdr ${W}><w:p>${run('Memory benchmark document', '<w:i/>')}</w:p></w:hdr>`,
  'word/footer1.xml': `${XML}<w:ftr ${W}><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${run('Page ')}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('1')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`,
};

/** Build the synthetic document with `units` content units, as DOCX bytes. */
export function syntheticDocx(units) {
  const body = Array.from({ length: units }, (_, unit) => unitXml(unit)).join('');
  const files = {
    ...parts,
    'word/document.xml': `${XML}<w:document ${W} ${R}><w:body>${body}${sectPr}</w:body></w:document>`,
  };
  const entries = Object.fromEntries(
    Object.entries(files).map(([name, text]) => [name, [strToU8(text), { mtime: FIXED_MTIME }]])
  );
  return zipSync(entries, { level: 6 });
}
