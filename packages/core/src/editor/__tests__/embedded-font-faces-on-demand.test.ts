// Browser faces for families a document does not name register on first use.
//
// A browser keeps its own decoded copy of every registered face, so registering a host's whole
// configuration held every family for a document that names one. The faces a document names still
// register before the first paint; the rest wait until the painter asks for their alias.

import { describe, expect, test } from 'bun:test';
import type { FontSource, FontSourceSubstitution } from '../../contracts/editor.ts';
import {
  registerEmbeddedFontFaces,
  type FontFaceLike,
  type FontFaceSetLike,
} from '../embedded-font-faces.ts';

function source(family: string, weight = 400, style: 'normal' | 'italic' = 'normal'): FontSource {
  return {
    request: { family, weight, style },
    id: `${family}#${weight}#${style}`,
    bytes: new Uint8Array([1, 2, 3]),
    hash: `sha256:${family}-${weight}-${style}`,
    faceIndex: 0,
  };
}

const substitution = (from: string, to: string): FontSourceSubstitution => ({
  from: { family: from, weight: 400, style: 'normal' },
  to: { family: to, weight: 400, style: 'normal' },
});

class FakeFace implements FontFaceLike {
  constructor(
    readonly family: string,
    private readonly fail: boolean
  ) {}
  load(): Promise<unknown> {
    return this.fail ? Promise.reject(new Error('parse failure')) : Promise.resolve(this);
  }
}

function fakeEnvironment(options: { throwOnCreate?: boolean } = {}) {
  const added: FakeFace[] = [];
  const deleted: FakeFace[] = [];
  const fontSet: FontFaceSetLike = {
    add: (face) => added.push(face as FakeFace),
    delete: (face) => {
      deleted.push(face as FakeFace);
      return true;
    },
  };
  return {
    added,
    deleted,
    environment: {
      fontSet,
      createFontFace: (family: string) => {
        if (options.throwOnCreate) throw new Error('unsupported font');
        return new FakeFace(family, false);
      },
    },
  };
}

describe('registerEmbeddedFontFaces on demand', () => {
  const sources = [
    source('Carlito'),
    source('Carlito', 700),
    source('Liberation Serif'),
    source('Liberation Mono'),
  ];
  const substitutions = [
    substitution('Calibri', 'Carlito'),
    substitution('Times New Roman', 'Liberation Serif'),
    substitution('Courier New', 'Liberation Mono'),
  ];

  test('registers only the families a document uses before it resolves', async () => {
    const env = fakeEnvironment();
    const registration = await registerEmbeddedFontFaces(sources, env.environment, substitutions, [
      'Calibri',
    ]);
    // Both Carlito faces, reached through the Calibri substitution.
    expect(env.added).toHaveLength(2);
    expect(registration.installed).toBe(2);
    expect(registration.alias('Calibri')).toBe(env.added[0]!.family);
    expect(registration.alias('Carlito')).toBe(env.added[0]!.family);
  });

  test('registers another family the first time the painter asks for it', async () => {
    const env = fakeEnvironment();
    const registration = await registerEmbeddedFontFaces(sources, env.environment, substitutions, [
      'Calibri',
    ]);
    const alias = registration.alias('Times New Roman');
    expect(alias).toMatch(/^docx-embedded-/);
    expect(env.added).toHaveLength(3);
    expect(env.added.at(-1)!.family).toBe(alias!);
    expect(registration.installed).toBe(3);
    // Asking again registers nothing more.
    expect(registration.alias('Liberation Serif')).toBe(alias);
    expect(env.added).toHaveLength(3);
  });

  test('omitting the used families registers every face up front', async () => {
    const env = fakeEnvironment();
    await registerEmbeddedFontFaces(sources, env.environment, substitutions);
    expect(env.added).toHaveLength(4);
  });

  test('a family the browser refuses on first use has no alias', async () => {
    const env = fakeEnvironment({ throwOnCreate: true });
    const registration = await registerEmbeddedFontFaces(sources, env.environment, substitutions, [
      'Calibri',
    ]);
    expect(registration.alias('Courier New')).toBeUndefined();
    expect(registration.alias('Liberation Mono')).toBeUndefined();
    expect(env.added).toHaveLength(0);
  });

  test('a family registered on first use reaches other windows and is disposed', async () => {
    const env = fakeEnvironment();
    const registration = await registerEmbeddedFontFaces(sources, env.environment, substitutions, [
      'Calibri',
    ]);
    const popout: FakeFace[] = [];
    const popoutDocument = {
      fonts: { add: (face: FakeFace) => popout.push(face), delete: () => true },
      defaultView: {
        FontFace: class extends FakeFace {
          constructor(family: string) {
            super(family, false);
          }
        },
      },
    } as unknown as Document;
    registration.showIn(popoutDocument);
    expect(popout).toHaveLength(2);
    registration.alias('Courier New');
    expect(popout).toHaveLength(3);
    registration.dispose();
    expect(env.deleted).toHaveLength(3);
    expect(registration.alias('Times New Roman')).toBeUndefined();
    expect(env.added).toHaveLength(3);
  });
});
