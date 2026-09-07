import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
test('MVP includes brief workflow', async () => {
  const html = await readFile(resolve(root, 'index.html'), 'utf8');
  const js = await readFile(resolve(root, 'app.js'), 'utf8');
  assert.match(html, /briefDialog/); assert.match(html, /ОТКРЫТЫЕ ЗАДАЧИ/iu); assert.match(html, /аудитор/iu);
  assert.match(js, /localStorage/); assert.match(js, /renderTasks/); assert.match(js, /unshift/);
});
test('report and visual scheme exist', async () => {
  const report = await readFile(resolve(root, 'REPORT.md'), 'utf8');
  const svg = await readFile(resolve(root, 'docs/mindmap.svg'), 'utf8');
  assert.match(report, /Таблица 1/); assert.match(report, /IT CREW DIGITAL/); assert.match(report, /итерацион/iu); assert.match(svg, /МЕТОДОЛОГИЯ/);
});
