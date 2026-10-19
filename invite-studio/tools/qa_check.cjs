#!/usr/bin/env node
/* Automated phone checks for a built invite (optional; needs Playwright).
 *
 *   python -m http.server -d dist 8080 &
 *   node tools/qa_check.cjs ramya-karthik-x7q2 [http://localhost:8080] [screenshot-dir]
 *
 * Checks: 320/375/412 px widths without horizontal scroll, nothing loads music before the tap,
 * tap opens + starts music, language toggle, ?to= personalisation + sanitising, watermark/RSVP lock,
 * reduced motion, fonts with Tamil/Devanagari glyphs, real transferred bytes (incl. Google Fonts).
 */
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const slug = process.argv[2] || 'ramya-karthik-x7q2';
const base = (process.argv[3] || 'http://localhost:8080').replace(/\/$/, '');
const shots = process.argv[4] || '';
const url = `${base}/${slug}/`;
const results = [];
const check = (name, ok, info = '') => { results.push({ name, ok, info }); console.log(`${ok ? '✓' : '✗'} ${name}${info ? ' — ' + info : ''}`); };

(async () => {
  const browser = await chromium.launch();
  if (shots) fs.mkdirSync(shots, { recursive: true });

  for (const width of [320, 375, 412]) {
    const ctx = await browser.newContext({ viewport: { width, height: 760 }, deviceScaleFactor: 2, isMobile: true, hasTouch: true });
    const page = await ctx.newPage();
    const errors = [];
    const requests = [];
    let bytes = 0, fontBytes = 0;
    const counted = new Set();
    page.on('pageerror', (e) => errors.push(e.message));
    page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
    page.on('request', (r) => requests.push(r.url()));
    page.on('response', async (r) => {
      if (counted.has(r.url())) return;
      counted.add(r.url());
      try {
        const len = +(r.headers()['content-length'] || (await r.body()).length);
        bytes += len;
        if (/fonts\.gstatic/.test(r.url())) fontBytes += len;
      } catch (e) { /* redirects etc. */ }
    });
    await page.goto(url, { waitUntil: 'networkidle' });
    const before = requests.filter((u) => /\.mp3/.test(u)).length;
    check(`${width}px: no music request before tap`, before === 0);
    if (shots) await page.screenshot({ path: path.join(shots, `opener-${width}.png`) });
    await page.locator('#opener').tap();
    await page.waitForTimeout(1800);
    check(`${width}px: opener removed after tap`, await page.locator('#opener').evaluate((e) => e.classList.contains('is-gone')));
    check(`${width}px: music requested after tap`, requests.some((u) => /\.mp3/.test(u)));
    const overflow = await page.evaluate((w) => Math.max(document.documentElement.scrollWidth, window.innerWidth) - w, width);
    check(`${width}px: no horizontal scroll`, overflow <= 0, `overflow ${overflow}px`);
    // Scroll through so every reveal fires, then screenshot the whole page.
    await page.evaluate(async () => {
      for (let y = 0; y < document.body.scrollHeight; y += 300) { window.scrollTo(0, y); await new Promise((r) => setTimeout(r, 120)); }
      window.scrollTo(0, 0);
    });
    await page.waitForTimeout(600);
    const overflow2 = await page.evaluate((w) => Math.max(document.documentElement.scrollWidth, window.innerWidth) - w, width);
    check(`${width}px: no horizontal scroll after reveals`, overflow2 <= 0, `overflow ${overflow2}px`);
    if (shots) await page.screenshot({ path: path.join(shots, `page-${width}.png`), fullPage: true });
    check(`${width}px: no JS errors`, errors.length === 0, errors.slice(0, 3).join(' | '));
    if (width === 375) check('transfer size incl. fonts and photos', bytes < 2 * 1024 * 1024, `${Math.round(bytes / 1024)} KB total, fonts ${Math.round(fontBytes / 1024)} KB`);
    await ctx.close();
  }

  // Language toggle + fonts
  {
    const ctx = await browser.newContext({ viewport: { width: 375, height: 760 }, isMobile: true, hasTouch: true });
    const page = await ctx.newPage();
    await page.goto(url, { waitUntil: 'networkidle' });
    await page.locator('#opener').tap();
    await page.waitForTimeout(1600);
    const langs = await page.$$eval('#lang-toggle button', (b) => b.map((x) => x.lang));
    for (const l of langs) {
      await page.click(`#lang-toggle button[lang="${l}"]`);
      await page.waitForTimeout(150);
      const h1 = (await page.textContent('h1')).replace(/\s+/g, ' ').trim();
      const eventTitle = await page.textContent('.event-title');
      check(`lang ${l}: names + events switch`, h1.length > 0 && eventTitle.length > 0, `${h1} · ${eventTitle}`);
      if (shots) await page.screenshot({ path: path.join(shots, `lang-${l}.png`) });
    }
    await page.evaluate(() => document.fonts.ready);
    const fonts = await page.evaluate(async () => {
      await Promise.all([
        document.fonts.load('700 40px "Noto Serif Tamil"', 'ரம்யாகார்த்திக்'),
        document.fonts.load('700 40px "Noto Serif Devanagari"', 'रम्याकार्तिक'),
        document.fonts.load('40px "Great Vibes"', 'Ramya'),
      ]);
      return {
        ta: document.fonts.check('700 40px "Noto Serif Tamil"', 'ரம்யாகார்த்திக்'),
        hi: document.fonts.check('700 40px "Noto Serif Devanagari"', 'रम्याकार्तिक'),
        en: document.fonts.check('40px "Great Vibes"', 'Ramya'),
      };
    });
    check('fonts: Tamil, Devanagari and script fonts loaded', fonts.ta && fonts.hi && fonts.en, JSON.stringify(fonts));
    const wm = await page.locator('.watermark').count();
    const disabled = await page.locator('#rsvp-form button[type="submit"]').isDisabled();
    check('watermark shown + RSVP disabled (watermark: true)', wm === 1 && disabled, `watermark ${wm}, submit disabled ${disabled}`);
    await ctx.close();
  }

  // Guest personalisation + sanitising
  for (const [q, expect] of [['Ravi-Uncle', 'Ravi Uncle'], ['<script>alert(1)</script>', null], ['%3Cimg%20src%3Dx%20onerror%3Dalert(1)%3E', null]]) {
    const ctx = await browser.newContext({ viewport: { width: 375, height: 760 } });
    const page = await ctx.newPage();
    let dialog = false;
    page.on('dialog', async (d) => { dialog = true; await d.dismiss(); });
    await page.goto(`${url}?to=${q}`, { waitUntil: 'networkidle' });
    await page.click('#opener');
    await page.waitForTimeout(1500);
    const dear = await page.textContent('.dear-line');
    const injected = await page.locator('.dear-line *').count();
    if (expect) check(`?to=${q} shows greeting`, dear.includes(expect), dear);
    else check(`?to=${q} is sanitised`, !dialog && injected === 0 && !/[<>()=]/.test(dear), JSON.stringify(dear));
    await ctx.close();
  }

  // Reduced motion
  {
    const ctx = await browser.newContext({ viewport: { width: 375, height: 760 }, reducedMotion: 'reduce' });
    const page = await ctx.newPage();
    await page.goto(url, { waitUntil: 'networkidle' });
    await page.click('#opener');
    await page.waitForTimeout(400);
    const canvas = await page.locator('canvas.petals').count();
    const gone = await page.locator('#opener').evaluate((e) => e.classList.contains('is-gone'));
    check('reduced motion: no petals, opener skips animation', canvas === 0 && gone);
    await ctx.close();
  }

  // Private details never appear on the page
  {
    const ctx = await browser.newContext();
    const page = await ctx.newPage();
    await page.goto(url);
    const htmlText = (await page.content()).toLowerCase();
    const leaks = ['ram lakhan', 'inspector', 'cbic'].filter((m) => htmlText.includes(m));
    check('no private details in page', leaks.length === 0, leaks.join(', '));
    const og = await page.$$eval('meta[property^="og:"]', (m) => Object.fromEntries(m.map((x) => [x.getAttribute('property'), x.content])));
    check('OG tags present with absolute image URL', /^https:\/\//.test(og['og:image'] || '') && og['og:image:width'] === '1200' && !!og['og:title'], og['og:image']);
    await ctx.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} checks passed`);
  process.exit(failed ? 1 : 0);
})();
