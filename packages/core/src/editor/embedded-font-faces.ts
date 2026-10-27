// Paint-side twin of embedded-font auto-wiring, done safely (issue #78).
//
// Measurement never touches the browser's font machinery: HarfBuzz shapes the admitted
// bytes directly. But painted pages are ordinary DOM, so a font that exists only inside
// the DOCX renders in a platform substitute unless the SAME bytes are registered as a
// `FontFace`.
//
// The naive version of that — register under the family the DOCUMENT declares — is a
// zero-click UI-redressing hole, and was reverted once already. `FontFaceSet` is
// per-DOCUMENT, not per-subtree: a face added for a family name hides the installed font
// of that name for the WHOLE page. A file fully controls `w:font/@w:name`, so a document
// naming its face `Segoe UI` would repaint the host application's own chrome with
// attacker-chosen glyphs. CSS-quoting the name stops CSS injection, not shadowing —
// quoting is not namespacing.
//
// So faces register under an ENGINE-MINTED alias derived from the admitted face's content
// hash, and the painter emits `font-family: "<alias>", "<declared family>"`. Document text
// gets the embedded glyphs; the declared name keeps whatever it means to the host page.

import type { FontSource, FontSourceSubstitution } from '../contracts/editor.ts';
import {
  boundedStructuralFontValidator,
  FontResolutionError,
  type FontResourceSnapshot,
} from '../layout/font-resource.ts';
import { materializeCollectionFace } from './font-collection-face.ts';

/** The slice of `FontFace` this module needs; injectable for tests. */
export interface FontFaceLike {
  load(): Promise<unknown>;
}

/** The slice of `FontFaceSet` this module needs; injectable for tests. */
export interface FontFaceSetLike {
  add(face: FontFaceLike): unknown;
  delete(face: FontFaceLike): unknown;
}

export interface EmbeddedFontFaceEnvironment {
  readonly fontSet?: FontFaceSetLike | undefined;
  readonly createFontFace?: (
    family: string,
    bytes: Uint8Array,
    descriptors: { readonly weight: string; readonly style: 'normal' | 'italic' }
  ) => FontFaceLike;
}

/** Faces registered for one document load; disposing removes exactly those faces. */
export interface EmbeddedFontFaceRegistration {
  /** Faces actually admitted into the set (per-face failures are dropped silently). */
  readonly installed: number;
  /**
   * Declared family -> alias, for the painter. Empty when nothing registered, which makes
   * painting fall back to the declared family alone.
   */
  readonly alias: (family: string) => string | undefined;
  /**
   * Register the same faces in another document's `FontFaceSet`, once per document. A set is
   * per document, so a surface that moves to a pop-out window or an iframe paints in
   * fallback glyphs until its new document has them too.
   */
  showIn(document: Document): void;
  dispose(): void;
}

const NO_REGISTRATION: EmbeddedFontFaceRegistration = Object.freeze({
  installed: 0,
  alias: () => undefined,
  showIn() {},
  dispose() {},
});

type FaceDescriptors = { readonly weight: string; readonly style: 'normal' | 'italic' };

/**
 * The alias for one document's face set.
 *
 * Derived from content hashes, so it is engine-minted and a document cannot collide with
 * a host family name — the whole point. The `docx-embedded-` prefix keeps it recognisable
 * in devtools, and the alias is per-FAMILY (not per-face) because CSS selects weight and
 * style through the descriptors, not the name.
 */
function aliasForFamily(hashes: readonly string[]): string {
  let digest = 0;
  for (const hash of hashes) {
    for (let index = 0; index < hash.length; index += 1) {
      digest = (digest * 31 + hash.charCodeAt(index)) >>> 0;
    }
  }
  return `docx-embedded-${digest.toString(36)}`;
}

/** Match the family normalization used by measurement's fontRequestKey. */
const familyKey = (family: string): string => family.trim().toLowerCase();

/**
 * The font machinery of `doc`, or of the script's own document when none is given.
 *
 * Both halves come from the SAME window: a `FontFace` made by one window's constructor is
 * refused by another window's `FontFaceSet`.
 */
export function fontFaceEnvironmentOf(doc?: Document | null): EmbeddedFontFaceEnvironment {
  const target =
    doc ?? (typeof globalThis.document !== 'undefined' ? globalThis.document : undefined);
  const fontSet = (target as { fonts?: FontFaceSetLike } | undefined)?.fonts;
  const FontFaceCtor = (target?.defaultView as (Window & typeof globalThis) | null | undefined)
    ?.FontFace;
  if (!fontSet || typeof FontFaceCtor !== 'function') return {};
  return {
    fontSet,
    createFontFace: (family, bytes, descriptors) =>
      // A copy, not the admitted view: FontFace snapshots its source, but slicing keeps
      // the engine's buffers from ever being aliased by browser internals.
      new FontFaceCtor(family, bytes.slice().buffer, descriptors) as unknown as FontFaceLike,
  };
}

/**
 * Register admitted faces under engine-minted aliases and return the lookup the painter
 * uses. Resolves after every face it registers up front has either loaded into the set or
 * failed; per-face failure only lowers `installed`.
 *
 * Callers MUST pass faces that survived validation — these bytes reach the browser's own
 * font loader.
 *
 * `substitutions` extends the lookup to the names runs actually write: a document says
 * `Calibri` while the registered face is `Carlito`, so without the map the text measures
 * shaped and then paints in a platform substitute. Aliasing the `from` name is what makes
 * app-supplied stand-ins paintable WITHOUT registering a face called "Calibri" globally —
 * the shadowing hole this module exists to avoid applies to Word's family names just as
 * much as to a document's own.
 *
 * `usedFamilies` names the families the document asks for. Faces those families reach,
 * directly or through `substitutions`, register before this resolves, so the first paint
 * already has them. Every other family registers the first time the painter asks for its
 * alias. A browser keeps its own decoded copy of every registered face, so registering the
 * whole configuration held all of it for a document that names one family. Omitted, every
 * face registers up front.
 */
export async function registerEmbeddedFontFaces(
  sources: readonly FontSource[],
  environment: EmbeddedFontFaceEnvironment = fontFaceEnvironmentOf(),
  substitutions: readonly FontSourceSubstitution[] = [],
  usedFamilies?: readonly string[]
): Promise<EmbeddedFontFaceRegistration> {
  const { fontSet, createFontFace } = environment;
  if (!fontSet || !createFontFace || sources.length === 0) return NO_REGISTRATION;

  // One alias per declared family, covering all of that family's faces.
  const byFamily = new Map<string, FontSource[]>();
  for (const source of sources) {
    const family = familyKey(source.request.family);
    const list = byFamily.get(family);
    if (list) list.push(source);
    else byFamily.set(family, [source]);
  }
  /** A substituted-from name -> the registered family whose alias it borrows. */
  const substituted = new Map<string, string>();
  for (const substitution of substitutions) {
    const from = familyKey(substitution.from.family);
    const to = familyKey(substitution.to.family);
    if (byFamily.has(to) && !substituted.has(from)) substituted.set(from, to);
  }
  const used = usedFamilies && new Set(usedFamilies.map(familyKey));
  const eager = new Set<string>();
  for (const family of byFamily.keys()) {
    if (!used || used.has(family)) eager.add(family);
  }
  if (used) {
    for (const [from, to] of substituted) if (used.has(from)) eager.add(to);
  }

  const aliasOf = new Map<string, string>();
  for (const [family, faces] of byFamily) {
    aliasOf.set(family, aliasForFamily(faces.map((face) => `${face.hash}:face:${face.faceIndex}`)));
  }
  /** Families with at least one face the browser accepted. */
  const loaded = new Set<string>();
  /** Families that have registered, or tried to: each is attempted once. */
  const attempted = new Set<string>();
  const added: { readonly set: FontFaceSetLike; readonly face: FontFaceLike }[] = [];
  /**
   * What each accepted face was made from, so `showIn` can make it again elsewhere. The SOURCE,
   * not its bytes: a source can hand out a fresh copy on each read, and keeping one copy per
   * face for a window that may never open held every document font once more.
   */
  const accepted: { alias: string; source: FontSource; descriptors: FaceDescriptors }[] = [];
  /** Other documents' font machinery, from `showIn`; a family registered later reaches them too. */
  const shownElsewhere: EmbeddedFontFaceEnvironment[] = [];
  const browserBytesOf = (face: FontSource): Uint8Array => {
    const bytes = face.bytes;
    const isCollection =
      bytes.length >= 4 &&
      bytes[0] === 0x74 &&
      bytes[1] === 0x74 &&
      bytes[2] === 0x63 &&
      bytes[3] === 0x66;
    if (isCollection && face.faceIndex === 0 && !boundedStructuralFontValidator(bytes, 0).valid)
      throw new Error('Invalid collection face');
    if (!isCollection && face.faceIndex !== 0)
      throw new Error('Font face index requires a collection');
    // FontFace selects the first face from an unchanged collection. Keep its signature and
    // all tables; only other faces require materialization.
    return isCollection && face.faceIndex !== 0
      ? materializeCollectionFace(bytes, face.faceIndex)
      : bytes;
  };
  const descriptorsOf = (face: FontSource): FaceDescriptors => ({
    weight: String(face.request.weight),
    style: face.request.style,
  });
  /** Add one face to another set before it loads; that set renders it once its bytes parse. */
  const addUnloaded = (
    target: EmbeddedFontFaceEnvironment,
    alias: string,
    source: FontSource,
    descriptors: FaceDescriptors
  ): boolean => {
    try {
      const face = target.createFontFace!(alias, browserBytesOf(source), descriptors);
      target.fontSet!.add(face);
      added.push({ set: target.fontSet!, face });
      void face.load().catch(() => {});
      return true;
    } catch {
      // Paint fidelity is best-effort; the face still measures shaped.
      return false;
    }
  };

  await Promise.all(
    [...eager].map(async (family) => {
      attempted.add(family);
      const alias = aliasOf.get(family)!;
      await Promise.all(
        byFamily.get(family)!.map(async (face) => {
          try {
            const descriptors = descriptorsOf(face);
            const fontFace = createFontFace(alias, browserBytesOf(face), descriptors);
            await fontFace.load();
            fontSet.add(fontFace);
            added.push({ set: fontSet, face: fontFace });
            accepted.push({ alias, source: face, descriptors });
            loaded.add(family);
          } catch {
            // Paint fidelity is best-effort; the face still measures shaped.
          }
        })
      );
    })
  );
  if (added.length === 0 && eager.size === byFamily.size) return NO_REGISTRATION;

  let disposed = false;
  /**
   * Register a family that waited for its first use. Synchronous, because the painter is asking
   * for the alias right now: the face is added before it loads, and the browser repaints text
   * under the alias once the bytes parse. Measurement never depends on it, so layout is the same.
   */
  const registerOnFirstUse = (family: string): void => {
    if (attempted.has(family)) return;
    attempted.add(family);
    const alias = aliasOf.get(family)!;
    for (const source of byFamily.get(family)!) {
      const descriptors = descriptorsOf(source);
      if (!addUnloaded(environment, alias, source, descriptors)) continue;
      accepted.push({ alias, source, descriptors });
      loaded.add(family);
      for (const target of shownElsewhere) addUnloaded(target, alias, source, descriptors);
    }
  };
  /**
   * Only advertise an alias the browser accepted: pointing the painter at a family with no
   * registered face would push text to the generic fallback instead of the declared family
   * behind it. A substituted-from name only borrows an alias that its target earned, and
   * never displaces a face registered under that name directly — a real Calibri beats a
   * stand-in for Calibri, the same precedence composition already applies to bytes.
   */
  const aliasFor = (family: string): string | undefined => {
    if (byFamily.has(family)) {
      registerOnFirstUse(family);
      if (loaded.has(family)) return aliasOf.get(family);
    }
    const target = substituted.get(family);
    if (target === undefined) return undefined;
    registerOnFirstUse(target);
    return loaded.has(target) ? aliasOf.get(target) : undefined;
  };

  const shownIn = new WeakSet<object>([fontSet]);
  return {
    get installed() {
      return added.filter((entry) => entry.set === fontSet).length;
    },
    alias: (family) => (disposed ? undefined : aliasFor(familyKey(family))),
    showIn(doc) {
      if (disposed) return;
      const target = fontFaceEnvironmentOf(doc);
      if (!target.fontSet || !target.createFontFace || shownIn.has(target.fontSet)) return;
      shownIn.add(target.fontSet);
      shownElsewhere.push(target);
      for (const { alias, source, descriptors } of accepted) {
        addUnloaded(target, alias, source, descriptors);
      }
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      for (const { set, face } of added) {
        try {
          set.delete(face);
        } catch {
          /* removing from a torn-down set is not an error */
        }
      }
    },
  };
}

/**
 * The sources a shaping snapshot admitted, as the validated, owned copies it holds — never the
 * raw file view. A refused source, or one the snapshot answered with another face, is left out.
 */
export function admittedFontSources(
  sources: readonly FontSource[],
  snapshot: FontResourceSnapshot
): FontSource[] {
  const admitted: FontSource[] = [];
  for (const source of sources) {
    const resolved = snapshot.resolve(source.request);
    if (resolved instanceof FontResolutionError || resolved.id !== source.id) continue;
    const { id, hash, faceIndex } = resolved;
    admitted.push({
      request: source.request,
      id,
      // Read when a face is registered: the owned copy is handed out fresh then, and nothing
      // keeps a second copy of every font for the document's lifetime.
      get bytes() {
        return resolved.bytes;
      },
      hash,
      faceIndex,
    });
  }
  return admitted;
}
