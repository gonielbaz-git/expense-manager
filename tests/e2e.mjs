// End-to-end check in headless Chromium: node tests/e2e.mjs   (needs `playwright`)
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
const require = createRequire(import.meta.url);
let pw; try { pw = require('playwright'); } catch { pw = require(path.join(execSync('npm root -g').toString().trim(), 'playwright')); }
const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let failed = 0;
const ok = (cond, msg) => { console.log((cond ? 'ok   ' : 'FAIL ') + msg); if (!cond) failed++; };

const browser = await pw.chromium.launch();
const page = await browser.newPage();
const errors = [];
page.on('pageerror', e => errors.push(String(e)));
page.on('console', m => { if (m.type() === 'error' && !/fonts\.g|net::ERR/.test(m.text())) errors.push(m.text()); });
await page.goto(pathToFileURL(path.join(root, 'index.html')).href);
ok(await page.locator('.dropzone').count() === 1, 'blank app renders');
ok(await page.evaluate(() => Object.keys(state.months).length) === 0, 'blank app carries no data');

const s = f => path.join(root, 'sample', f);
await page.setInputFiles('input[type=file]', [s('card_1234_08_2026.xlsx'), s('card_1234_09_2026.xlsx'), s('bank_report.xlsx')]);
await page.waitForSelector('.kpis');
const T = k => page.evaluate(k => monthTotals(k), k);
const jul = await T('2026-07'), aug = await T('2026-08');
// July: card 212.4+38+310.5+54+400(instalment)+20€+164.9-120, bank: rent 4200 + savings 500
const fx = await page.evaluate(() => fxRate('€'));
ok(Math.abs(jul.exp - (212.4 + 38 + 310.5 + 54 + 400 + 20 * fx + 164.9 - 120 + 4200 + 500)) < 0.02, 'July expenses add up: ' + jul.exp);
ok(jul.inc === 9800, 'July income only counts July credits');
ok(aug.inc === 9800 && aug.pend === 1, 'August income + one pending row');
ok(await page.evaluate(() => state.months['2026-07'].rows.some(r => /רהיטים/.test(r.merchant))), 'instalment lands in the month it is billed for');
const recon = await page.evaluate(() => { ui.month = '2026-07'; render(); return [...document.querySelectorAll('.reconbox')].map(b => b.textContent); });
ok(recon.length === 1 && /תואם/.test(recon[0]), 'July card total reconciles with the bank charge');
const before = await page.evaluate(() => JSON.stringify(state.months));
await page.setInputFiles('input[type=file]', [s('card_1234_08_2026.xlsx'), s('bank_report.xlsx')]);
await page.waitForTimeout(300);
ok(await page.evaluate(() => JSON.stringify(state.months)) === before, 're-importing the same files changes nothing');

// categorise → merchant memory → next import is auto-categorised
await page.evaluate(() => { const r = state.months['2026-07'].rows.find(r => r.merchant === 'סופר השכונה'); r.category = 'סופר וירקות'; rememberMerchant(r.merchant, r.category); });
ok(await page.evaluate(() => lookupCategory('סופר השכונה 12')) === 'סופר וירקות', 'merchant memory matches name variants');

// delete + undo
const n0 = await page.evaluate(() => state.months['2026-07'].rows.length);
await page.locator('#exptable tbody tr .row-actions button').first().click();
ok(await page.evaluate(() => state.months['2026-07'].rows.length) === n0 - 1, 'row deleted');
await page.locator('.toast .toastbtn', { hasText: 'ביטול' }).click();
ok(await page.evaluate(() => state.months['2026-07'].rows.length) === n0, 'undo restores it');

// search hides rows without re-rendering
await page.fill('.search', 'קפה');
ok(await page.locator('#exptable tbody tr:visible').count() === 1, 'search narrows the table');

// a saved copy must reopen with the same data, and a blank copy with none
const doc = await page.evaluate(() => buildDoc());
const p2 = await browser.newPage(); p2.on('pageerror', e => errors.push(String(e)));
await p2.setContent(doc); await p2.waitForSelector('.kpis');
ok((await p2.evaluate(() => monthTotals('2026-07').exp)) === jul.exp, 'downloaded copy reopens with identical totals');
const blank = await page.evaluate(() => buildDoc(blankState()));
ok(!/סופר השכונה|100001/.test(blank), 'blank copy contains no data');
ok(errors.length === 0, 'no console errors ' + errors.join(' | '));
await browser.close();
process.exit(failed ? 1 : 0);
