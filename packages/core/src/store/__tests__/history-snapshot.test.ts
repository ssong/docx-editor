// Compact undo history: an entry below the top of its stack keeps the body's child list as
// chunks shared with its neighbours, and undo and redo still restore every state exactly.

import { describe, expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint } from '../package/ooxml-serialize.ts';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { findNode, nodeIndexTestRecorder } from '../package/ooxml-edit.ts';
import { partHistoryKey } from '../package/part-history-key.ts';
import { chunkChildren } from '../store/history-snapshot.ts';
import { TreeDocumentStore } from '../store/tree-store.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function widePart(paragraphs: number): OoxmlPart {
  const body = Array.from(
    { length: paragraphs },
    (_, at) => `<w:p><w:r><w:t>Paragraph ${at}</w:t></w:r></w:p>`
  ).join('');
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function bodyOf(part: OoxmlPart): readonly OoxmlNode[] {
  const body = part.root.children[0]!;
  if (body.kind === 'textValue') throw new Error('no body');
  return body.children;
}

function paragraphIds(part: OoxmlPart): string[] {
  return bodyOf(part)
    .filter((node) => node.kind === 'paragraph')
    .map((node) => node.id);
}

function commit(store: TreeDocumentStore, op: TreeDocOp): void {
  const result = store.transact((ctx) => ctx.apply(op));
  if (!result.ok) throw new Error(`${result.reason} ${result.detail ?? ''}`);
}

/** A mix of typing, Enter and Backspace at spread-out paragraphs. */
function editStep(store: TreeDocumentStore, step: number): void {
  const ids = paragraphIds(store.part);
  const paragraphId = ids[(step * 37) % ids.length]!;
  if (step % 7 === 3) {
    commit(store, { op: 'splitParagraph', paragraphId, offset: 2 });
  } else if (step % 11 === 5) {
    const at = ids.indexOf(paragraphId);
    const next = ids[at + 1];
    if (next) commit(store, { op: 'joinParagraphs', firstId: paragraphId, secondId: next });
  } else {
    commit(store, { op: 'insertText', paragraphId, offset: 0, text: `${step}` });
  }
}

describe('chunkChildren', () => {
  const nodes = (count: number, from = 0): OoxmlNode[] =>
    Array.from(
      { length: count },
      (_, at) =>
        ({ kind: 'textValue', id: `/word/document.xml#0.${from + at}`, value: '' }) as OoxmlNode
    );
  const flat = (chunks: readonly (readonly OoxmlNode[])[]): OoxmlNode[] => chunks.flat();

  test('chunks hold the children in order', () => {
    const children = nodes(5_000);
    const chunks = chunkChildren(children, undefined);
    expect(flat(chunks)).toEqual(children);
    expect(chunks.length).toBeGreaterThan(20);
    expect(chunks.length).toBeLessThan(250);
  });

  test('a replaced child changes one chunk', () => {
    const children = nodes(5_000);
    const chunks = chunkChildren(children, undefined);
    const edited = [...children];
    edited[2_500] = { ...children[2_500]! } as OoxmlNode;
    const next = chunkChildren(edited, chunks);
    expect(flat(next)).toEqual(edited);
    expect(next.filter((chunk) => !chunks.includes(chunk))).toHaveLength(1);
  });

  test('an inserted or removed child keeps the chunks after it aligned', () => {
    const children = nodes(5_000);
    const chunks = chunkChildren(children, undefined);
    const inserted = [...children.slice(0, 1_000), ...nodes(1, 9_000), ...children.slice(1_000)];
    const afterInsert = chunkChildren(inserted, chunks);
    expect(flat(afterInsert)).toEqual(inserted);
    expect(afterInsert.filter((chunk) => !chunks.includes(chunk)).length).toBeLessThanOrEqual(2);

    const removed = children.filter((_, at) => at !== 4_000);
    const afterRemove = chunkChildren(removed, chunks);
    expect(flat(afterRemove)).toEqual(removed);
    expect(afterRemove.filter((chunk) => !chunks.includes(chunk)).length).toBeLessThanOrEqual(2);
  });

  test('distant changes reuse every chunk between them', () => {
    const children = nodes(5_000);
    const chunks = chunkChildren(children, undefined);
    const edited = [...children];
    edited[100] = { ...children[100]! } as OoxmlNode;
    edited[4_900] = { ...children[4_900]! } as OoxmlNode;
    const next = chunkChildren(edited, chunks);
    expect(flat(next)).toEqual(edited);
    expect(next.filter((chunk) => !chunks.includes(chunk))).toHaveLength(2);
  });
});

describe('compact undo history', () => {
  test('undo and redo restore every state of a wide document', () => {
    const store = new TreeDocumentStore(widePart(600));
    const states = [canonicalOoxmlFingerprint(store.part)];
    for (let step = 0; step < 60; step += 1) {
      editStep(store, step);
      states.push(canonicalOoxmlFingerprint(store.part));
    }
    for (let at = states.length - 2; at >= 0; at -= 1) {
      expect(store.undo()).not.toBeNull();
      expect(canonicalOoxmlFingerprint(store.part)).toBe(states[at]!);
    }
    expect(store.canUndo).toBe(false);
    for (let at = 1; at < states.length; at += 1) {
      expect(store.redo()).not.toBeNull();
      expect(canonicalOoxmlFingerprint(store.part)).toBe(states[at]!);
    }
  });

  test('entries below the top keep no copy of the body child list', () => {
    const store = new TreeDocumentStore(widePart(2_000));
    for (let step = 0; step < 40; step += 1) editStep(store, step);
    const { undoStack } = store.checkpoint();
    const lists = new Set<readonly OoxmlNode[]>([bodyOf(store.part)]);
    // Only the newest entry holds a whole package; every other one holds chunks.
    for (const entry of undoStack.slice(0, -1)) {
      const held = entry.snapshot as unknown as { whole: unknown; compact: unknown };
      expect(held.whole).toBeNull();
      expect(held.compact).not.toBeNull();
    }
    lists.add(bodyOf(undoStack.at(-1)!.snapshot.package().parts.get('/word/document.xml')!));
    expect(lists.size).toBe(2);
  });

  test('a restored state keeps the node index and the id allocator lineage', () => {
    const store = new TreeDocumentStore(widePart(600));
    for (let step = 0; step < 30; step += 1) editStep(store, step);
    const recorder = nodeIndexTestRecorder();
    recorder.reset();
    for (let step = 0; step < 20; step += 1) {
      store.undo();
      expect(findNode(store.part, paragraphIds(store.part)[300]!)).not.toBeNull();
    }
    // Each undo patched the index it inherited; none walked the whole part again.
    expect(recorder.completeBuilds).toBe(0);
    // New paragraphs after a deep undo take ids that no earlier state used.
    const seen = new Set(paragraphIds(store.part));
    for (let step = 0; step < 10; step += 1) {
      const ids = paragraphIds(store.part);
      commit(store, { op: 'splitParagraph', paragraphId: ids[step * 50]!, offset: 1 });
    }
    const after = paragraphIds(store.part);
    expect(new Set(after).size).toBe(after.length);
    for (const id of after) expect(findNode(store.part, id)?.id).toBe(id);
    expect(after.filter((id) => !seen.has(id))).toHaveLength(10);
  });

  test('a part restored from a compact entry keeps its history key', () => {
    const store = new TreeDocumentStore(widePart(600));
    const keys = [partHistoryKey(store.part)];
    for (let step = 0; step < 10; step += 1) {
      editStep(store, step);
      keys.push(partHistoryKey(store.part));
    }
    for (let at = keys.length - 2; at >= 0; at -= 1) {
      store.undo();
      expect(partHistoryKey(store.part)).toBe(keys[at]!);
    }
  });
});
