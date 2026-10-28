// The field-result addressing mode, and the span granularity, as cache-key components.
//
// In the `editable` mode a saved field result is laid out as ordinary text at its own model
// offsets; in `atomic` it is one unit. Pieces, breaks, and every offset a layout memo stores
// differ between the two, so a cache entry made in one mode must never answer the other. The
// producer string is folded into every paragraph cache key, so the mode rides it. A `run` pass
// stores merged spans in the break it caches (`run-span-merge.ts`), so its mode rides it too.

import { currentFieldResultsMode } from '../store/package/field-result-mode.ts';
import { currentSpanGranularity } from './run-span-merge.ts';

/** `producer`, marked when the layout call in progress addresses saved results as text. */
export function withFieldResultsProducer(producer: string): string {
  const fields = currentFieldResultsMode() === 'editable' ? `${producer}|field-results` : producer;
  return currentSpanGranularity() === 'run' ? `${fields}|spans:run` : fields;
}
