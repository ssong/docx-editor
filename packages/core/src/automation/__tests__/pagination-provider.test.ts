import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { openDocumentForExport } from '../../export/export-session.ts';
import { createFixedMeasurer } from '../../layout/fixed-measurer.ts';
import { headlessViewOfStore } from '../../store/headless-document-view.ts';
import { normalizeParagraphIdentity } from '../../store/package/para-id.ts';
import { readOoxmlPackage } from '../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../store/package/ooxml-tree.ts';
import { TreePackageStore } from '../../store/store/tree-package-store.ts';
import { createAutomationPaginationProvider, paginationSnapshotOf } from '../pagination.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

function liveStore(body: string): TreePackageStore {
  const loaded = readOoxmlPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rDoc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
      ),
    })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const main = loaded.package.parts.get(loaded.package.mainDocumentPart)!;
  return new TreePackageStore(loaded.package, normalizeParagraphIdentity(main));
}

test('the one-shot provider matches a retained live session and follows later revisions', async () => {
  const body = Array.from(
    { length: 80 },
    (_, index) => `<w:p><w:r><w:t>Line ${index}</w:t></w:r></w:p>`
  );
  const store = liveStore(body.join(''));
  const view = headlessViewOfStore(store);
  const measurer = createFixedMeasurer();
  const provider = createAutomationPaginationProvider({ measurer, producer: 'test-fixed' });

  const retained = openDocumentForExport(view, {
    measurer,
    producer: 'test-fixed',
    displayMode: 'proposed',
    reuseAcrossRevisions: true,
  });
  expect(retained.ok).toBe(true);
  if (!retained.ok) return;
  try {
    const before = await provider(view);
    expect(before).toEqual(paginationSnapshotOf(await retained.session.layout()));

    const firstParagraph = (node: OoxmlNode): string | null => {
      if (node.kind === 'paragraph') return node.id;
      if (node.kind === 'textValue') return null;
      for (const child of node.children) {
        const found = firstParagraph(child);
        if (found) return found;
      }
      return null;
    };
    const paragraphId = firstParagraph(store.bodyStore().part.root)!;
    const changed = store.transact({ kind: 'body' }, (context) => {
      context.apply({ op: 'insertText', paragraphId, offset: 0, text: 'word '.repeat(4000) });
    });
    expect(changed.ok).toBe(true);

    const after = await provider(view);
    expect(after.pageCount).toBeGreaterThan(before.pageCount);
    expect(after).toEqual(paginationSnapshotOf(await retained.session.layout()));
  } finally {
    retained.session.dispose();
  }
});
