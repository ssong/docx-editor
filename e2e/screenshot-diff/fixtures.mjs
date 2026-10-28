// The documents the screenshot diff renders, and the synthetic ones it builds.
//
// Each entry names a document and how many of its first pages to capture. The set covers
// the cases a layout or paint change breaks first: ordinary business text, tables, inline
// and floating images, CJK, right-to-left text and justified lines. The synthetic
// documents are deterministic (fixed zip timestamps), so the same entry is the same bytes
// on every machine and every run.

import { strToU8, zipSync } from 'fflate';
import { syntheticDocx } from '../../benchmarks/memory/fixtures.mjs';

const FIXTURES_DIR = new URL('../fixtures/', import.meta.url);

/**
 * `file` is a path under `e2e/fixtures/`; `build` returns the bytes of a synthetic document.
 * `pages` is the number of leading pages to capture.
 */
export const SCREENSHOT_FIXTURES = [
  { name: 'demo', file: 'demo.docx', pages: 3 },
  { name: 'with-tables', file: 'with-tables.docx', pages: 2 },
  { name: 'repeated-table-header', file: 'repeated-table-header.docx', pages: 2 },
  { name: 'example-with-image', file: 'example-with-image.docx', pages: 2 },
  { name: 'images-wrap-sides', file: 'images-wrap-sides.docx', pages: 2 },
  { name: 'comprehensive-word-elements', file: 'comprehensive-word-element-test.docx', pages: 3 },
  { name: 'rtl-table-bidivisual', file: 'rtl-table-bidivisual.docx', pages: 1 },
  { name: 'justified-multi-section', file: 'EP_ZMVZ_MULTI_v4.docx', pages: 2 },
  { name: 'justified-lorem', file: 'lorem-caret-fix-check.docx', pages: 2 },
  { name: 'synthetic-scripts', build: scriptsDocx, pages: 2 },
  { name: 'synthetic-memory-bench', build: () => syntheticDocx(4), pages: 3 },
];

/** The bytes of one fixture entry. */
export async function fixtureBytes(entry) {
  if (entry.build) return entry.build();
  const { readFile } = await import('node:fs/promises');
  return new Uint8Array(await readFile(new URL(entry.file, FIXTURES_DIR)));
}

// ---- The synthetic scripts document: justified Latin, CJK and right-to-left paragraphs.

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const FIXED_MTIME = new Date(Date.UTC(2020, 0, 1));

const LATIN =
  'The parties agree that the delivery schedule in this agreement is an estimate. ' +
  'Each invoice states the payment terms, the ledger reference and the warranty period. ' +
  'Notices take effect on receipt, and a revised forecast replaces the previous quarter. ';
const CHINESE =
  '这是一个用于检查中文排版的段落。文字在行尾换行时，标点符号不应该出现在行首。' +
  '合同双方同意按照本协议规定的时间表交付货物，并在收到发票后三十日内付款。';
const JAPANESE =
  'これは日本語の組版を確認するための段落です。漢字、ひらがな、カタカナ、' +
  'そしてEnglish wordsと数字2026が混在しています。行末の句読点にも注意してください。';
const KOREAN =
  '이 문단은 한국어 조판을 확인하기 위한 것입니다. 한글 단어 사이에는 공백이 있으며, ' +
  '줄 끝에서 단어 단위로 줄이 바뀌어야 합니다. 계약 당사자는 일정에 동의합니다.';
const ARABIC =
  'هذه فقرة لاختبار عرض النص العربي من اليمين إلى اليسار. تحتوي الفقرة على كلمات كثيرة ' +
  'حتى تلتف على عدة أسطر، ويجب أن تبقى الحروف متصلة في كل سطر. ';
const HEBREW =
  'זוהי פסקה לבדיקת הצגת טקסט בעברית מימין לשמאל. הפסקה ארוכה מספיק כדי לגלוש לכמה שורות, ' +
  'והמילים צריכות להופיע בסדר הנכון בכל שורה. ';

const esc = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const run = (text, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const para = (runs, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;
const heading = (text) => para(run(text), '<w:pStyle w:val="Heading1"/>');

const EAST_ASIA = (font, lang) => `<w:rFonts w:eastAsia="${font}"/><w:lang w:eastAsia="${lang}"/>`;
const RTL = (lang) => `<w:rtl/><w:lang w:bidi="${lang}"/>`;

function scriptsBody() {
  return [
    heading('Justified text'),
    para(
      run(LATIN) +
        run('Bold terms, ', '<w:b/>') +
        run('italic notes, ', '<w:i/>') +
        run('larger type, ', '<w:sz w:val="28"/>') +
        run('coloured and underlined words', '<w:color w:val="1F4E79"/><w:u w:val="single"/>') +
        run(', a superscript', '') +
        run('2', '<w:vertAlign w:val="superscript"/>') +
        run(' and an unbreakable-hyphenated-compound-word-that-is-rather-long. ' + LATIN),
      '<w:jc w:val="both"/>'
    ),
    para(run(LATIN + LATIN), '<w:jc w:val="both"/><w:ind w:firstLine="720"/>'),
    para(
      run('Term\tValue\tNote'),
      '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="6480"/><w:tab w:val="left" w:pos="7200"/></w:tabs>'
    ),
    para(run(LATIN), '<w:jc w:val="distribute"/>'),
    para(run('Centred line of text'), '<w:jc w:val="center"/>'),
    para(run('Right aligned line of text'), '<w:jc w:val="right"/>'),
    heading('CJK'),
    para(run(CHINESE + CHINESE, EAST_ASIA('SimSun', 'zh-CN'))),
    para(run(CHINESE + CHINESE, EAST_ASIA('SimSun', 'zh-CN')), '<w:jc w:val="both"/>'),
    para(run(JAPANESE + JAPANESE, EAST_ASIA('MS Mincho', 'ja-JP')), '<w:jc w:val="both"/>'),
    para(run(KOREAN + KOREAN, EAST_ASIA('Malgun Gothic', 'ko-KR'))),
    heading('Right to left'),
    para(run(ARABIC + ARABIC, RTL('ar-SA')), '<w:bidi/>'),
    para(run(ARABIC + ARABIC, RTL('ar-SA')), '<w:bidi/><w:jc w:val="both"/>'),
    para(run(HEBREW + HEBREW, RTL('he-IL')), '<w:bidi/>'),
    para(
      run('الإصدار ', RTL('ar-SA')) +
        run('Version 2.5') +
        run(' صدر في عام ', RTL('ar-SA')) +
        run('2026') +
        run(' مع ', RTL('ar-SA')) +
        run('DOCX', '<w:b/>') +
        run(' و ', RTL('ar-SA')) +
        run('PDF'),
      '<w:bidi/>'
    ),
  ].join('');
}

const sectPr =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

const parts = {
  '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  '_rels/.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  'word/styles.xml': `${XML}<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style></w:styles>`,
};

/** The synthetic scripts document, as DOCX bytes. */
export function scriptsDocx() {
  const files = {
    ...parts,
    'word/document.xml': `${XML}<w:document ${W}><w:body>${scriptsBody()}${sectPr}</w:body></w:document>`,
  };
  const entries = Object.fromEntries(
    Object.entries(files).map(([name, text]) => [name, [strToU8(text), { mtime: FIXED_MTIME }]])
  );
  return zipSync(entries, { level: 6 });
}
