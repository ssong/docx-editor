#!/usr/bin/env node
// Render a memory benchmark report as Markdown tables.
//
// Usage: node benchmarks/memory/report.mjs <report.json> [--compare <earlier.json>]

import { readFileSync } from 'node:fs';
import { parseArgs } from 'node:util';

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: { compare: { type: 'string' } },
});
if (!positionals[0]) throw new Error('Pass a report file.');
const report = JSON.parse(readFileSync(positionals[0], 'utf8'));
const before = values.compare ? JSON.parse(readFileSync(values.compare, 'utf8')) : null;

const stageOf = (row, name) => row.stages.find((stage) => stage.stage === name);
const mb = (value) => (value === undefined ? '–' : value.toFixed(1));
const perPage = (value, pages) =>
  value === undefined || !pages ? '–' : `${Math.round((value * 1024) / pages)}`;
const change = (target, document, stage, field, now) => {
  if (!before || now === undefined) return '';
  const old = before[target]?.find((row) => row.document === document);
  const then = old && stageOf(old, stage)?.[field];
  if (then === undefined || then === 0) return '';
  const pct = Math.round(((now - then) / then) * 100);
  return ` (${pct >= 0 ? '+' : ''}${pct}%)`;
};
const editStage = (row) =>
  row.stages.find((stage) => stage.stage.startsWith('edit ') || stage.stage.startsWith('type '));

const lines = [];
const env = report.environment;
lines.push(
  `Commit \`${report.commit}\`${report.dirty ? ' (with local changes to packages/core or packages/docx-to-pdf)' : ''}, ${report.date.slice(0, 10)}. ` +
    `${env.cpu}, ${env.cpus} CPUs, ${env.memoryGb} GB, Node.js ${env.node}` +
    `${env.chromium ? `, Chromium ${env.chromium}` : ''}. Median of ${report.options.runs} runs. Values in MB unless noted.`
);

if (report.server) {
  lines.push(
    '',
    '### Server (Node.js)',
    '',
    '| Document | Pages | Open | Layout | KB per page | After release | Edit growth | Save | Peak RSS |',
    '| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |'
  );
  for (const row of report.server) {
    const open = stageOf(row, 'open');
    const layout = stageOf(row, 'layout');
    const release = stageOf(row, 'release');
    const edit = editStage(row);
    const save = stageOf(row, 'save');
    const peak = Math.max(...row.stages.map((stage) => stage.peakRssMb));
    lines.push(
      `| ${row.document} | ${row.pages} | ${mb(open.retainedMb)}${change('server', row.document, 'open', 'retainedMb', open.retainedMb)} | ` +
        `${mb(layout.retainedMb)}${change('server', row.document, 'layout', 'retainedMb', layout.retainedMb)} | ` +
        `${perPage(layout.retainedMb, row.pages)} | ${mb(release.retainedMb)} | ${mb(edit.deltaMb)} | ${mb(save.deltaMb)} | ${mb(peak)} |`
    );
  }
  lines.push(
    '',
    '*Open* and *Layout* are the live heap above the module baseline after each stage. *Edit growth* is ' +
      `the heap added by ${report.options.serverEdits} single-character edits. *Peak RSS* is the process high-water mark.`
  );
}

if (report.pdf) {
  lines.push(
    '',
    '### PDF export (Node.js)',
    '',
    '| Document | Pages | Layout ms | Paint ms | Export ms | Peak RSS | KB per page | PDF |',
    '| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |'
  );
  for (const row of report.pdf) {
    const exported = stageOf(row, 'export');
    const base = stageOf(row, 'modules loaded');
    lines.push(
      `| ${row.document} | ${row.pages} | ${row.layoutMs} | ${row.paintMs} | ${exported.ms} | ` +
        `${mb(exported.peakRssMb)}${change('pdf', row.document, 'export', 'peakRssMb', exported.peakRssMb)} | ` +
        `${perPage(exported.peakRssMb - base.rssMb, row.pages)} | \`${row.pdfSha256.slice(0, 12)}\` |`
    );
  }
  lines.push(
    '',
    '*Peak RSS* is the process high-water mark of one `exportPdf` call, packaged fonts only. ' +
      '*KB per page* counts it above the RSS with the modules loaded. *PDF* is the start of the SHA-256 of the output.'
  );
}

if (report.browser) {
  lines.push(
    '',
    '### Browser (Chromium)',
    '',
    '| Document | Pages | Open | KB per page | After scroll | Typing growth | After destroy | DOM nodes | Page total |',
    '| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |'
  );
  for (const row of report.browser) {
    const open = stageOf(row, 'open');
    const scroll = stageOf(row, 'scroll through');
    const edit = editStage(row);
    const destroy = stageOf(row, 'destroy');
    lines.push(
      `| ${row.document} | ${row.pages} | ${mb(open.retainedMb)}${change('browser', row.document, 'open', 'retainedMb', open.retainedMb)} | ` +
        `${perPage(open.retainedMb, row.pages)} | ${mb(scroll.retainedMb)} | ${mb(edit.deltaMb)} | ${mb(destroy.retainedMb)} | ` +
        `${scroll.domNodes} | ${mb(open.pageTotalMb)} |`
    );
  }
  lines.push(
    '',
    '*Open*, *After scroll*, and *After destroy* are the JavaScript heap above the baseline (the page with its fonts loaded). ' +
      `*Typing growth* is the heap added by ${report.options.browserEdits} typed characters. *Page total* is ` +
      '`performance.measureUserAgentSpecificMemory()` after open: JavaScript, DOM, and WebAssembly memory.'
  );
}

console.log(lines.join('\n'));
