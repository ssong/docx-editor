#!/usr/bin/env node
// Layout equivalence check: lays out every fixture with a base commit and with this
// checkout, then compares the fingerprints. CI runs it on every pull request; run it
// locally the same way. See scripts/layout-fingerprint/README.md.

import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { parseArgs } from 'node:util';

const here = import.meta.dirname;
const repo = resolve(here, '../..');

const { values } = parseArgs({
  options: {
    base: { type: 'string', default: 'origin/main' },
    work: { type: 'string' },
    'build-head': { type: 'boolean', default: false },
    measurer: { type: 'string', default: 'fixed,shaped' },
    sizes: { type: 'string', default: '10,100,1000' },
    filter: { type: 'string' },
    jobs: { type: 'string' },
    tolerance: { type: 'string', default: '0' },
    spans: { type: 'string', default: 'exact' },
    markdown: { type: 'string' },
    help: { type: 'boolean', default: false },
  },
});

if (values.help) {
  console.log(`Usage: bun run layout:check [options]

Options:
  --base <ref>        commit to compare against (default: origin/main)
  --work <dir>        where the base checkout and both outputs go
                      (default: a fixed directory under the system temp directory)
  --build-head        build this checkout's packages first (otherwise they must be built)
  --measurer <list>   fixed, shaped, or both (default: fixed,shaped)
  --sizes <a,b,...>   synthetic document sizes ('' for none) (default: 10,100,1000)
  --filter <text>     only documents whose name contains this text
  --jobs <n>          documents laid out at the same time
  --tolerance <pt>    allowed position difference in points (default: 0, exact)
  --spans <mode>      exact or merged (default: exact). merged also lays out this
                      checkout with the layout option spans: 'run'
  --markdown <file>   also write the report as Markdown

Exit code 0 when the layout matches, 1 when it differs.`);
  process.exit(0);
}

const git = (args, cwd = repo) => execFileSync('git', args, { cwd, encoding: 'utf8' }).trim();
const sh = (command, args, cwd) => execFileSync(command, args, { cwd, stdio: 'inherit' });

const baseSha = git(['rev-parse', '--verify', `${values.base}^{commit}`]);
const work = resolve(
  values.work ??
    join(
      tmpdir(),
      `docx-layout-fingerprint-${createHash('sha256').update(repo).digest('hex').slice(0, 8)}`
    )
);
const baseRoot = join(work, 'base');
const builtMarker = join(baseRoot, '.layout-fingerprint-built');

// The packages the fingerprint loads, in dependency order. Same commands as the
// benchmark READMEs, so a base commit builds the way it was meant to.
function buildPackages(root) {
  for (const pkg of ['@docx-editor.dev/i18n', '@docx-editor.dev/core', '@docx-editor.dev/fonts']) {
    sh('bun', ['run', '--filter', pkg, 'build'], root);
  }
}

console.log(`Base: ${values.base} (${baseSha.slice(0, 12)}) in ${baseRoot}`);
if (!existsSync(join(baseRoot, '.git'))) {
  git(['worktree', 'add', '--force', '--detach', baseRoot, baseSha]);
} else if (git(['rev-parse', 'HEAD'], baseRoot) !== baseSha) {
  git(['checkout', '--detach', '--force', baseSha], baseRoot);
}
const built = existsSync(builtMarker) && readFileSync(builtMarker, 'utf8').trim() === baseSha;
if (!built) {
  sh('bun', ['install', '--frozen-lockfile'], baseRoot);
  buildPackages(baseRoot);
  writeFileSync(builtMarker, `${baseSha}\n`);
} else {
  console.log('Base packages already built at this commit.');
}
if (values['build-head']) buildPackages(repo);

const runArgs = ['--measurer', values.measurer, '--sizes', values.sizes];
if (values.filter) runArgs.push('--filter', values.filter);
if (values.jobs) runArgs.push('--jobs', values.jobs);
const run = join(here, 'run.mjs');
const baseOut = join(work, 'out-base');
const headOut = join(work, 'out-head');
sh(process.execPath, [run, '--root', baseRoot, '--out', baseOut, ...runArgs], repo);
// Merged spans compare a head that merges against a base that does not.
const headArgs = values.spans === 'merged' ? ['--layout-spans', 'run'] : [];
sh(process.execPath, [run, '--root', repo, '--out', headOut, ...runArgs, ...headArgs], repo);

const compareArgs = ['--tolerance', values.tolerance, '--spans', values.spans];
if (values.markdown) compareArgs.push('--markdown', resolve(values.markdown));
try {
  sh(process.execPath, [join(here, 'compare.mjs'), baseOut, headOut, ...compareArgs], repo);
} catch (error) {
  console.log(`Fingerprints: ${baseOut} and ${headOut}`);
  process.exit(error.status ?? 1);
}
