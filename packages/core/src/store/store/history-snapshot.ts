import { adoptHistoryRootLineage, historyRootLineage } from '../package/ooxml-edit.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import { adoptPartHistoryKey, partHistoryKey } from '../package/part-history-key.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import type { HistoryGroup } from './history-group.ts';
import type { SelectionMark } from './tree-store.ts';

// Undo history that does not hold one copy of the body's child list per edit.
//
// A history entry is the whole package as it was. Structural sharing makes that cheap for
// everything an edit did not touch, but not for the element that holds the edited block:
// typing one character rebuilds the body, and the body's child list is one pointer per
// top-level block. At 9,500 blocks (1,270 pages) every entry held about 80 KB of its own.
//
// An entry below the top of its stack therefore keeps a WIDE element (a part root, or a
// child of one, with many children) as a list of chunks. A chunk is a frozen slice of the
// child list, and the chunks of one entry are mostly the very same arrays as the chunks of
// the entry before it, because two neighbouring states differ in a few blocks. Where a chunk
// ends depends only on the child at its end (a hash of its id), so an inserted or removed
// block changes one chunk and the chunks after it still line up. An entry then costs about
// one pointer per chunk, plus the chunks its own edit changed.
//
// Undo rebuilds the spine from the chunks: new objects for the package, the part, the root
// and the wide element, and the same objects for every block. Callers that keyed state on a
// restored part keep it through `partHistoryKey`; the node index and the id allocator follow
// through the root lineage, so a restored root patches the index instead of walking the part.

/** One entry of a store's undo or redo stack. */
export interface HistoryEntry {
  /**
   * The whole package as it was, not just the story part.
   *
   * Affordable for the same reason a part snapshot was: parts are immutable and deep-frozen, so
   * a package snapshot is a Map of references and every part the transaction did not touch is
   * object-identical to the one before it. Undo reverses every part one intent wrote rather
   * than only the story. Below the top of its stack, an entry keeps wide elements in chunks.
   */
  readonly snapshot: HistorySnapshot;
  readonly revision: number;
  readonly selectionBefore: SelectionMark | null;
  readonly selectionAfter: SelectionMark | null;
  /**
   * The gesture this entry is still collecting, while it is the newest entry and nothing
   * has closed it. Only ever set on the top of the undo stack: `closeHistoryGroup`
   * strips it, and undo and redo never copy it onto the entries they push.
   */
  readonly group?: HistoryGroup;
}

/** Average number of children in one chunk is `CHUNK_MASK + 1`. */
const CHUNK_MASK = 63;
/** An element with fewer children costs less than its chunk list would save. */
const WIDE = 256;

type Chunk = readonly OoxmlNode[];

/** A wide element without its child list, and the chunks that hold the list. */
class ChunkedElement {
  constructor(
    readonly shell: OoxmlElement,
    readonly chunks: readonly Chunk[]
  ) {}
}

/** A narrow root whose children include at least one chunked element. */
class SpineElement {
  constructor(
    readonly shell: OoxmlElement,
    readonly children: readonly (OoxmlNode | ChunkedElement)[]
  ) {}
}

type CompactRoot = ChunkedElement | SpineElement;

interface CompactPart {
  readonly shell: OoxmlPart;
  readonly root: CompactRoot;
  readonly lineage: object;
  readonly key: object;
}

interface CompactPackage {
  readonly shell: OoxmlPackage;
  /** The parts in package order. An array, because a Map costs several times as much. */
  readonly parts: readonly (OoxmlPart | CompactPart)[];
}

const NO_CHILDREN: readonly OoxmlNode[] = Object.freeze([]);
const NO_NAMES: readonly string[] = Object.freeze([]);
const NO_PARTS: ReadonlyMap<string, OoxmlPart> = new Map();

/** The element with an empty child list: its kind, id, name and attributes, and no subtree. */
function withoutChildren(element: OoxmlElement): OoxmlElement {
  return Object.freeze({ ...element, children: NO_CHILDREN }) as OoxmlElement;
}

/** Whether a chunk ends after `node`. Depends on the id alone, so equal runs chunk equally. */
function endsChunk(node: OoxmlNode): boolean {
  const id = node.id;
  let hash = 0x811c9dc5;
  for (let at = 0; at < id.length; at += 1) {
    hash ^= id.charCodeAt(at);
    hash = Math.imul(hash, 0x01000193);
  }
  // Parsed ids differ only in their last digits; mix those into the low bits.
  hash ^= hash >>> 15;
  hash = Math.imul(hash, 0x2c1b3c6d);
  hash ^= hash >>> 12;
  return (hash & CHUNK_MASK) === 0;
}

function holdsAt(children: readonly OoxmlNode[], at: number, chunk: Chunk): boolean {
  if (at < 0 || at + chunk.length > children.length) return false;
  for (let offset = 0; offset < chunk.length; offset += 1) {
    if (children[at + offset] !== chunk[offset]) return false;
  }
  return true;
}

/**
 * Split `children` into chunks, reusing the chunks of `previous` wherever the same children
 * sit between the same chunk ends.
 *
 * Neighbouring states share a long prefix and a long suffix, so those are matched first by
 * pointer comparison, and only the changed middle is hashed. A middle chunk that equals one
 * of `previous` (a state further away) is reused too.
 */
export function chunkChildren(
  children: readonly OoxmlNode[],
  previous: readonly Chunk[] | undefined
): Chunk[] {
  const out: Chunk[] = [];
  let start = 0;
  let first = 0;
  let last = (previous?.length ?? 0) - 1;
  if (previous) {
    for (; first <= last; first += 1) {
      const chunk = previous[first]!;
      const end = start + chunk.length;
      if (!holdsAt(children, start, chunk)) break;
      // The last chunk of `previous` ended with its list; here the list may go on.
      if (first === last && end < children.length && !endsChunk(children[end - 1]!)) break;
      out.push(chunk);
      start = end;
    }
  }
  const tail: Chunk[] = [];
  let stop = children.length;
  if (previous) {
    for (; last >= first; last -= 1) {
      const chunk = previous[last]!;
      const begin = stop - chunk.length;
      if (begin < start || !holdsAt(children, begin, chunk)) break;
      if (begin > 0 && !endsChunk(children[begin - 1]!)) break;
      tail.push(chunk);
      stop = begin;
    }
  }
  const byFirstChild = new Map<OoxmlNode, Chunk>();
  if (previous) {
    for (let at = first; at <= last; at += 1) {
      const chunk = previous[at]!;
      if (chunk.length > 0) byFirstChild.set(chunk[0]!, chunk);
    }
  }
  let from = start;
  for (let at = start; at < stop; at += 1) {
    if (at < stop - 1 && !endsChunk(children[at]!)) continue;
    const known = byFirstChild.get(children[from]!);
    out.push(
      known && known.length === at + 1 - from && holdsAt(children, from, known)
        ? known
        : Object.freeze(children.slice(from, at + 1))
    );
    from = at + 1;
  }
  for (let at = tail.length - 1; at >= 0; at -= 1) out.push(tail[at]!);
  // An array grown by `push` keeps spare capacity; the copy holds exactly one slot a chunk.
  return out.slice();
}

function chunkElement(element: OoxmlElement, previous: ChunkedElement | undefined): ChunkedElement {
  return new ChunkedElement(
    withoutChildren(element),
    chunkChildren(element.children, previous?.chunks)
  );
}

/** The chunked element `previous` holds for `id`, at the root or one level under it. */
function previousChunked(
  previous: CompactRoot | undefined,
  id: string
): ChunkedElement | undefined {
  if (!previous) return undefined;
  if (previous instanceof ChunkedElement) return previous.shell.id === id ? previous : undefined;
  for (const child of previous.children) {
    if (child instanceof ChunkedElement && child.shell.id === id) return child;
  }
  return undefined;
}

/** The compact form of a root, or null when nothing in its first two levels is wide. */
function compactRoot(root: OoxmlElement, previous: CompactRoot | undefined): CompactRoot | null {
  if (root.children.length >= WIDE) return chunkElement(root, previousChunked(previous, root.id));
  let wide = false;
  const children = root.children.map((child): OoxmlNode | ChunkedElement => {
    if (child.kind === 'textValue' || child.children.length < WIDE) return child;
    wide = true;
    return chunkElement(child, previousChunked(previous, child.id));
  });
  if (!wide) return null;
  return new SpineElement(withoutChildren(root), children);
}

function rebuildChunked(element: ChunkedElement): OoxmlElement {
  let length = 0;
  for (const chunk of element.chunks) length += chunk.length;
  const children = new Array<OoxmlNode>(length);
  let at = 0;
  for (const chunk of element.chunks) for (const child of chunk) children[at++] = child;
  return Object.freeze({ ...element.shell, children }) as OoxmlElement;
}

function rebuildRoot(root: CompactRoot): OoxmlElement {
  if (root instanceof ChunkedElement) return rebuildChunked(root);
  const children = root.children.map((child) =>
    child instanceof ChunkedElement ? rebuildChunked(child) : child
  );
  return Object.freeze({ ...root.shell, children }) as OoxmlElement;
}

function rebuildPart(part: CompactPart): OoxmlPart {
  const root = rebuildRoot(part.root);
  adoptHistoryRootLineage(root, part.lineage);
  const rebuilt = Object.freeze({ ...part.shell, root }) as OoxmlPart;
  adoptPartHistoryKey(rebuilt, part.key);
  return rebuilt;
}

function isCompactPart(part: OoxmlPart | CompactPart): part is CompactPart {
  return 'shell' in part;
}

/**
 * The newest compact form of each part name, so the next entry to compact can reuse its
 * chunks. One per store: the entries of one store are states of the same parts.
 */
export type HistoryChunkSource = Map<string, CompactPart>;

/**
 * One package in undo or redo history: whole while it is the newest entry of its stack,
 * and compact once another entry covers it.
 */
export class HistorySnapshot {
  private whole: OoxmlPackage | null;
  private compact: CompactPackage | null = null;
  /** Wide parts left whole because the live package held them at the last compaction. */
  private shared: readonly string[] | null = null;

  constructor(pkg: OoxmlPackage) {
    this.whole = pkg;
  }

  /** Whether `compactAgainst` can still make this entry smaller. */
  get pending(): boolean {
    return this.whole !== null
      ? this.shared === null || this.shared.length > 0
      : !!this.shared?.length;
  }

  /** The package this entry restores. A compact entry builds new spine objects each time. */
  package(): OoxmlPackage {
    if (this.whole) return this.whole;
    const compact = this.compact!;
    const parts = new Map<string, OoxmlPart>();
    for (const part of compact.parts) {
      const restored = isCompactPart(part) ? rebuildPart(part) : part;
      parts.set(restored.name, restored);
    }
    return Object.freeze({ ...compact.shell, parts }) as OoxmlPackage;
  }

  /**
   * Keep wide elements as chunks, for every part that differs from `live`.
   *
   * A part that `live` still holds costs nothing here, so it stays whole until a later call
   * finds that the live package has moved on.
   */
  compactAgainst(live: OoxmlPackage, source: HistoryChunkSource): void {
    const held = this.whole ? [...this.whole.parts.values()] : this.compact!.parts;
    const names = this.shared;
    let shared: string[] | null = null;
    let parts: (OoxmlPart | CompactPart)[] | null = null;
    for (let at = 0; at < held.length; at += 1) {
      const part = held[at]!;
      if (isCompactPart(part) || (names && !names.includes(part.name))) continue;
      if (!isWorthCompacting(part)) continue;
      if (live.parts.get(part.name) === part) {
        (shared ??= []).push(part.name);
        continue;
      }
      const compact = compactPart(part, source);
      if (!compact) continue;
      parts ??= held.slice();
      parts[at] = compact;
    }
    this.shared = shared ?? NO_NAMES;
    if (!parts) return;
    this.compact = {
      shell: this.whole
        ? (Object.freeze({ ...this.whole, parts: NO_PARTS }) as OoxmlPackage)
        : this.compact!.shell,
      parts,
    };
    this.whole = null;
  }
}

/**
 * Compact every entry that another entry now covers, against the live package.
 *
 * The newest entry of each stack stays whole: it is the one the next undo or redo restores,
 * and a gesture may still extend it. The store runs this after each commit, undo and redo;
 * an entry with nothing left to compact costs one check.
 */
export function compactCoveredEntries(
  stacks: readonly (readonly HistoryEntry[])[],
  live: OoxmlPackage,
  source: HistoryChunkSource
): void {
  for (const stack of stacks) {
    for (let at = 0; at < stack.length - 1; at += 1) {
      const snapshot = stack[at]!.snapshot;
      if (snapshot.pending) snapshot.compactAgainst(live, source);
    }
  }
}

function isWorthCompacting(part: OoxmlPart): boolean {
  const root = part.root;
  if (root.children.length >= WIDE) return true;
  return root.children.some((child) => child.kind !== 'textValue' && child.children.length >= WIDE);
}

/** One compact form per part object, so entries that share a part share its chunks too. */
const compactParts = new WeakMap<OoxmlPart, CompactPart>();

function compactPart(part: OoxmlPart, source: HistoryChunkSource): CompactPart | null {
  const known = compactParts.get(part);
  if (known) return known;
  const root = compactRoot(part.root, source.get(part.name)?.root);
  if (!root) return null;
  const compact: CompactPart = {
    shell: Object.freeze({ ...part, root: root.shell }) as OoxmlPart,
    root,
    lineage: historyRootLineage(part.root),
    key: partHistoryKey(part),
  };
  compactParts.set(part, compact);
  source.set(part.name, compact);
  return compact;
}
