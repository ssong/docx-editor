// Runs the screenshot diff gate: renders the fixture pages with a base engine, then renders
// them with the head engine (this checkout) and fails on any changed pixel.
//
// Usage: node e2e/screenshot-diff/run.mjs [--base <git ref> | --base-root <checkout> | --repeat]
//          [--out <dir>] [--chromium <path>] [--skip-build]
//
//   --base <ref>         Check the ref out in a worktree under local/screenshot-diff/, install
//                        and build its packages, and use it as the base (default: origin/main).
//   --base-root <dir>    Use an existing checkout with built packages as the base.
//   --repeat             Use this checkout as the base too: a stability check that must pass.
//   --out <dir>          Where the base PNGs, the head results and the report go
//                        (default: local/screenshot-diff/run).
//   --chromium <path>    Chromium executable (default: Playwright's own).
//   --threshold <0..1>   Per-pixel colour distance that still counts as equal (default: 0).
//   --max-diff-pixels <n>  Differing pixels a page may have (default: 0).
//   --skip-build         Do not build this checkout's packages first.
//
// On failure, `<out>/report/index.html` shows each changed page with the expected, actual
// and diff images.

import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, rmSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { parseArgs } from 'node:util';

const head = resolve(import.meta.dirname, '../..');
const { values } = parseArgs({
  options: {
    base: { type: 'string' },
    'base-root': { type: 'string' },
    repeat: { type: 'boolean', default: false },
    out: { type: 'string', default: join(head, 'local/screenshot-diff/run') },
    chromium: { type: 'string', default: process.env.SCREENSHOT_DIFF_CHROMIUM },
    'skip-build': { type: 'boolean', default: false },
    threshold: { type: 'string' },
    'max-diff-pixels': { type: 'string' },
  },
});

const PACKAGES = ['@docx-editor.dev/i18n', '@docx-editor.dev/core', '@docx-editor.dev/fonts'];
const sh = (cmd, args, cwd) => execFileSync(cmd, args, { cwd, stdio: 'inherit' });
const build = (root) => {
  for (const pkg of PACKAGES) sh('bun', ['run', '--filter', pkg, 'build'], root);
};

function baseFromRef(ref) {
  const sha = execFileSync('git', ['rev-parse', `${ref}^{commit}`], { cwd: head })
    .toString()
    .trim();
  const dir = join(head, 'local/screenshot-diff/base');
  const current = existsSync(join(dir, '.git'))
    ? execFileSync('git', ['rev-parse', 'HEAD'], { cwd: dir }).toString().trim()
    : null;
  if (current !== sha) {
    if (current) sh('git', ['worktree', 'remove', '--force', dir], head);
    sh('git', ['worktree', 'add', '--detach', dir, sha], head);
    sh('bun', ['install', '--frozen-lockfile'], dir);
    build(dir);
  }
  console.log(`base: ${ref} (${sha.slice(0, 10)}) in ${dir}`);
  return dir;
}

const out = resolve(values.out);
const baseRoot = values.repeat
  ? head
  : values['base-root']
    ? resolve(values['base-root'])
    : baseFromRef(values.base ?? 'origin/main');
if (!values['skip-build']) build(head);

rmSync(out, { recursive: true, force: true });
const playwright = (engine, results, extra) =>
  spawnSync('bunx', ['playwright', 'test', '--config', 'e2e/screenshot-diff.config.ts', ...extra], {
    cwd: head,
    stdio: 'inherit',
    env: {
      ...process.env,
      SCREENSHOT_DIFF_ENGINE: engine,
      SCREENSHOT_DIFF_BASELINE: join(out, 'base'),
      SCREENSHOT_DIFF_RESULTS: join(out, results),
      ...(values.chromium ? { SCREENSHOT_DIFF_CHROMIUM: values.chromium } : {}),
      ...(values.threshold ? { SCREENSHOT_DIFF_THRESHOLD: values.threshold } : {}),
      ...(values['max-diff-pixels']
        ? { SCREENSHOT_DIFF_MAX_DIFF_PIXELS: values['max-diff-pixels'] }
        : {}),
    },
  }).status;

console.log(`\n== Rendering the base (${baseRoot})`);
if (playwright(baseRoot, 'base-results', ['--update-snapshots=all']) !== 0) {
  console.error('The base render failed, so there is nothing to compare with.');
  process.exit(2);
}
console.log(`\n== Rendering the head (${head}) and comparing`);
const status = playwright(head, 'results', []);
if (status === 0) console.log('\nScreenshot diff: every captured page matches the base.');
else console.error(`\nScreenshot diff FAILED. Open ${join(out, 'report/index.html')}.`);
process.exit(status ?? 1);
