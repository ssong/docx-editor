import type { OoxmlPart } from './ooxml-tree.ts';

// A key for state that belongs to one part STATE, and must come back when undo restores it.
//
// Parts are immutable, so a part object used to be that key: undo put the very same object
// back. Undo history now keeps older states in a compact form (`history-snapshot.ts`), and
// restoring one builds a new part object with the same content. The key moves with it.

const keys = new WeakMap<OoxmlPart, object>();

/** The key for `part`'s state: the same object for the part undo or redo restores. */
export function partHistoryKey(part: OoxmlPart): object {
  let key = keys.get(part);
  if (!key) {
    key = {};
    keys.set(part, key);
  }
  return key;
}

/** Give a part rebuilt from history the key of the part it stands for. @internal */
export function adoptPartHistoryKey(part: OoxmlPart, key: object): void {
  keys.set(part, key);
}
