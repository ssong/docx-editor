// Summarize a V8 heap snapshot: what kinds of objects use the heap, and which named
// variables keep the most memory alive.
//
// Usage: node benchmarks/memory/analyze-snapshot.mjs <file.heapsnapshot> [--top 30]
//
// Three tables:
// - By constructor: count and shallow size of each object kind (Object shapes, arrays,
//   strings, closures, Maps).
// - By module or closure variable: the retained size of each value a `let`/`const` in a
//   module or closure scope holds. This is where module-level caches show up.
// - By property: the retained size held through each named property of an object, summed
//   over all objects. This shows which fields of the tree or layout records cost the most.

import { parseArgs } from 'node:util';
import { config as memlabConfig, getFullHeapFromFile } from 'memlab';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: { top: { type: 'string', default: '30' } },
});
const top = Number(values.top);
memlabConfig.muteConsole = true;
const heap = await getFullHeapFromFile(positionals[0]);
const mb = (bytes) => (bytes / 2 ** 20).toFixed(2).padStart(9);

let total = 0;
const byConstructor = new Map();
const byVariable = new Map();
const byProperty = new Map();
heap.nodes.forEach((node) => {
  total += node.self_size;
  const kind = node.type === 'object' || node.type === 'closure' ? node.name : `(${node.type})`;
  const entry = byConstructor.get(kind) ?? { count: 0, self: 0 };
  entry.count += 1;
  entry.self += node.self_size;
  byConstructor.set(kind, entry);

  for (const edge of node.references) {
    const target = edge.toNode;
    if (edge.type === 'context' && node.name.startsWith('system / Context')) {
      if (target.type === 'code' || target.type === 'hidden') continue;
      const key = `${String(edge.name_or_index)}: ${target.name || target.type}`;
      const held = byVariable.get(key) ?? { count: 0, retained: 0 };
      held.count += 1;
      held.retained += target.retainedSize;
      byVariable.set(key, held);
    } else if (edge.type === 'property' && node.type === 'object') {
      // Retained through this property only when the holder dominates the target.
      if (target.dominatorNode?.id !== node.id) continue;
      const key = `${node.name}.${String(edge.name_or_index)}`;
      const held = byProperty.get(key) ?? { count: 0, retained: 0 };
      held.count += 1;
      held.retained += target.retainedSize;
      byProperty.set(key, held);
    }
  }
});

const table = (title, map, field) => {
  console.log(`\n${title}`);
  for (const [name, entry] of [...map].sort((a, b) => b[1][field] - a[1][field]).slice(0, top)) {
    console.log(
      `${mb(entry[field])} MB  ${String(entry.count).padStart(9)}  ${name.slice(0, 110)}`
    );
  }
};
console.log(`heap: ${mb(total)} MB in ${heap.nodes.length} objects`);
table('By constructor (shallow MB, count):', byConstructor, 'self');
table('By module or closure variable (retained MB, holders):', byVariable, 'retained');
table('By dominating property (retained MB, holders):', byProperty, 'retained');
