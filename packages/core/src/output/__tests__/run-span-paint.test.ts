import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
import { expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/index.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../../layout/index.ts';
import type { SpanGranularity } from '../../layout/index.ts';
import { paintSemanticLayout } from '../semantic-paint.ts';

// A line painted from one span per run segment shows the same text at the same places as one
// painted from one span per word, with fewer elements.

const run = (text: string, rPr = '') =>
  `<w:r><w:rPr><w:sz w:val="22"/>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const PROSE = 'agreement balance contract delivery estimate forecast guarantee heading';

function paint(spans: SpanGranularity): HTMLElement {
  const parsed = readOoxmlPart(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      `<w:p>${run(`${PROSE} `)}${run('invoice journal ', '<w:b/>')}${run('ledger.')}</w:p>` +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${run(PROSE)}</w:p>` +
      `<w:p><w:pPr><w:jc w:val="both"/></w:pPr>${run(`${PROSE} invoice journal.`)}</w:p>` +
      '<w:sectPr><w:pgSz w:w="3800" w:h="12000"/>' +
      '<w:pgMar w:left="400" w:right="400" w:top="200" w:bottom="200"/></w:sectPr>' +
      '</w:body></w:document>',
    { name: '/word/document.xml', contentType: 'app/xml' }
  );
  if (!parsed.ok) throw new Error(parsed.reason);
  const container = document.createElement('div');
  paintSemanticLayout(
    container,
    layoutSemanticDocument(parsed.part, 1, { measurer: createFixedMeasurer(6, 14), spans }),
    { scale: 1 }
  );
  return container;
}

const lines = (container: HTMLElement) => [
  ...container.querySelectorAll<HTMLElement>('.docx-line'),
];
const runs = (line: HTMLElement) => [...line.querySelectorAll<HTMLElement>('[data-start]')];

test('each line paints the same text from fewer elements', () => {
  const word = lines(paint('word'));
  const merged = lines(paint('run'));
  expect(merged.map((line) => line.textContent)).toEqual(word.map((line) => line.textContent));
  expect(merged.map((line) => line.style.left)).toEqual(word.map((line) => line.style.left));
  const count = (all: HTMLElement[]) => all.reduce((sum, line) => sum + runs(line).length, 0);
  expect(count(merged)).toBeLessThan(count(word));
  // The model offsets the elements carry still cover each line without a gap.
  for (const line of merged) {
    const elements = runs(line);
    for (const [index, element] of elements.slice(1).entries()) {
      expect(element.dataset.start).toBe(elements[index]!.dataset.end);
    }
  }
});

test('a justified line paints its stretched spaces the same way', () => {
  const justified = (container: HTMLElement) =>
    lines(container)
      .filter((line) => runs(line).some((element) => element.style.wordSpacing))
      .map((line) =>
        runs(line).map((element) => [
          element.textContent,
          element.style.wordSpacing,
          element.style.marginLeft,
        ])
      );
  const word = justified(paint('word'));
  expect(word.length).toBeGreaterThan(0);
  expect(justified(paint('run'))).toEqual(word);
});
