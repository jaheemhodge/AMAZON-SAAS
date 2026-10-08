// Runs Code.gs in a sandbox with mocked Apps Script services.
// Usage: node --test test/
const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

// Apps Script loads every .gs file into one shared scope, so concatenate them.
const SOURCE = ['Code.gs', 'Fetch.gs', 'Api.gs', 'Sheets.gs']
  .map((f) => fs.readFileSync(path.join(__dirname, '..', 'apps-script', f), 'utf8'))
  .join('\n');

function makeSheet() {
  const rows = [];
  return {
    rows,
    getLastRow: () => rows.length,
    setFrozenRows() {},
    getRange(r, c, nr, nc) {
      return {
        setValues(vals) {
          vals.forEach((v, i) => { rows[r - 1 + i] = v.slice(); });
          return this;
        },
        getValues() {
          const out = [];
          for (let i = 0; i < nr; i++) out.push((rows[r - 1 + i] || []).slice(c - 1, c - 1 + nc));
          return out;
        },
        setFontWeight() { return this; }
      };
    }
  };
}

function load({ pages = {}, key = 'test-key', reviewsStatus = 200, top = [], details = {}, cookie = '' } = {}) {
  const sheets = {};
  const props = { RAPIDAPI_KEY: key, AMAZON_COOKIE: cookie };
  const calls = [];
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet())
  };
  const ctx = {
    console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    PropertiesService: {
      getDocumentProperties: () => ({
        getProperty: (k) => props[k] || null,
        setProperty: (k, v) => { props[k] = v; },
        deleteProperty: (k) => { delete props[k]; }
      })
    },
    Utilities: {
      sleep() {},
      DigestAlgorithm: { MD5: 'md5' },
      computeDigest: (_a, s) => Array.from(crypto.createHash('md5').update(s).digest()).map((b) => (b > 127 ? b - 256 : b))
    },
    UrlFetchApp: {
      fetch(url) {
        const u = new URL(url);
        const q = Object.fromEntries(u.searchParams);
        q.path = u.pathname;
        calls.push(q);
        if (u.pathname === '/product-details') {
          return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ status: 'OK', data: details }) };
        }
        if (u.pathname === '/top-product-reviews') {
          return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ status: 'OK', data: { reviews: top } }) };
        }
        if (reviewsStatus !== 200) {
          return { getResponseCode: () => reviewsStatus, getContentText: () => JSON.stringify({ message: 'Amazon requires login' }) };
        }
        const reviews = (pages[`${q.sort_by}|${q.star_rating || 'ALL'}|${q.page}`]) || [];
        return {
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({ status: 'OK', data: { total_ratings: 1234, reviews } })
        };
      }
    }
  };
  vm.createContext(ctx);
  vm.runInContext(SOURCE, ctx);
  return { ctx, sheets, calls };
}

function reviews(prefix, n) {
  return Array.from({ length: n }, (_, i) => ({
    review_id: `${prefix}${i}`,
    review_title: `Title ${prefix}${i}`,
    review_comment: i === 0 ? '=HYPERLINK("x")' : 'Great for my dog',
    review_star_rating: '5',
    review_author: 'Sam',
    review_date: 'Reviewed in the United States on May 1, 2026',
    is_verified_purchase: true,
    review_images: ['https://img/1.jpg']
  }));
}

test('parseAmazonLink handles common URL shapes', () => {
  const { ctx } = load();
  const p = (l) => JSON.parse(JSON.stringify(ctx.parseAmazonLink(l)));
  assert.deepStrictEqual(p('https://www.amazon.com/Brothly-Bone-Broth/dp/B0ABCDEF12/ref=sr_1_1?keywords=x'),
    { asin: 'B0ABCDEF12', country: 'US', url: 'https://www.amazon.com/dp/B0ABCDEF12' });
  assert.strictEqual(p('https://www.amazon.co.uk/gp/product/B0ABCDEF12').country, 'GB');
  assert.strictEqual(p('https://www.amazon.ca/product-reviews/B0ABCDEF12/ref=cm').asin, 'B0ABCDEF12');
  assert.strictEqual(p('amazon.de/dp/b0abcdef12').country, 'DE');
  assert.strictEqual(p('B0ABCDEF12').asin, 'B0ABCDEF12');
  assert.strictEqual(ctx.parseAmazonLink('https://www.amazon.com/s?k=dog+broth'), null);
  assert.strictEqual(ctx.parseAmazonLink(''), null);
});

test('fetches all pages across filters, writes rows, dedupes', () => {
  const pages = {
    'TOP_REVIEWS|ALL|1': reviews('a', 10),
    'TOP_REVIEWS|ALL|2': reviews('b', 10),
    'TOP_REVIEWS|ALL|3': reviews('c', 4), // short page ends this filter
    'MOST_RECENT|ALL|1': reviews('a', 10), // all duplicates
    'MOST_RECENT|ALL|2': reviews('d', 10)
  };
  const { ctx, sheets } = load({ pages });
  let state = ctx.startProduct('https://www.amazon.com/dp/B0ABCDEF12', false);
  let guard = 0;
  while (!state.done && guard++ < 20) state = ctx.fetchReviewsChunk(state);

  assert.ok(state.done);
  assert.strictEqual(state.added, 34);
  const rows = sheets.Reviews.rows.slice(1);
  assert.strictEqual(rows.length, 34);
  assert.strictEqual(new Set(rows.map((r) => r[2])).size, 34);
  assert.strictEqual(rows[0][5], '\'=HYPERLINK("x")'); // formula-escaped
  assert.strictEqual(rows[0][8], 'Yes');

  const product = sheets.Products.rows[1];
  assert.strictEqual(product[1], 'B0ABCDEF12');
  assert.strictEqual(product[3], 'Done');
  assert.strictEqual(product[4], 34);
  assert.strictEqual(product[6], 1234);

  // Re-running adds nothing new and doesn't duplicate the product row.
  state = ctx.startProduct('https://www.amazon.com/dp/B0ABCDEF12', false);
  while (!state.done && guard++ < 40) state = ctx.fetchReviewsChunk(state);
  assert.strictEqual(state.added, 0);
  assert.strictEqual(sheets.Reviews.rows.length, 35);
  assert.strictEqual(sheets.Products.rows.length, 2);
});

test('stops at max pages per filter', () => {
  const pages = {};
  for (let p = 1; p <= 15; p++) pages[`TOP_REVIEWS|ALL|${p}`] = reviews(`p${p}_`, 10);
  const { ctx, calls } = load({ pages });
  let state = ctx.startProduct('B0ABCDEF12', false);
  while (!state.done) state = ctx.fetchReviewsChunk(state);
  assert.strictEqual(calls.filter((c) => c.sort_by === 'TOP_REVIEWS').length, 10);
  assert.strictEqual(state.added, 100);
});

test('deep mode walks every star rating', () => {
  const { ctx } = load();
  assert.strictEqual(ctx.buildFilters(false).length, 2);
  assert.strictEqual(ctx.buildFilters(true).length, 12);
});

test('missing API key gives a clear error', () => {
  const { ctx } = load({ key: '' });
  assert.throws(() => ctx.startProduct('B0ABCDEF12', false), /No RapidAPI key/);
});

test('falls back to public top reviews when the full list fails', () => {
  const { ctx, sheets, calls } = load({ reviewsStatus: 500, top: reviews('t', 8) });
  let state = ctx.startProduct('B0G1VBDXYF', false);
  let guard = 0;
  while (!state.done && guard++ < 10) state = ctx.fetchReviewsChunk(state);
  assert.strictEqual(state.added, 8);
  assert.match(state.note, /HTTP 500: Amazon requires login/);
  assert.match(state.note, /add an Amazon cookie/);
  assert.strictEqual(sheets.Products.rows[1][3], 'Done (top reviews only)');
  assert.strictEqual(calls.filter((c) => c.path === '/product-reviews').length, 3); // retried, then gave up
});

test('bad key and quota errors stop with the real reason', () => {
  for (const [status, re] of [[403, /HTTP 403\): Amazon requires login/], [429, /quota reached \(HTTP 429\)/]]) {
    const { ctx } = load({ reviewsStatus: status });
    const state = ctx.startProduct('B0G1VBDXYF', false);
    assert.throws(() => ctx.fetchReviewsChunk(state), re);
  }
});

test('sends a trimmed Amazon cookie when set', () => {
  const cookie = 'session-id=1; ubid-main=2; at-main=3; csm-hit=junk; x-main=4; skin=noskin';
  const { ctx, calls } = load({ cookie, pages: { 'TOP_REVIEWS|ALL|1': reviews('a', 3) } });
  let state = ctx.startProduct('B0G1VBDXYF', false);
  state = ctx.fetchReviewsChunk(state);
  assert.strictEqual(calls[0].cookie, 'session-id=1; ubid-main=2; at-main=3; x-main=4');
  assert.strictEqual(calls[0].star_rating, undefined);
});

test('finds reviews under other field names, falling back to product details', () => {
  const details = {
    product_title: 'Foundation',
    product_num_ratings: 57,
    product_information: { top_reviews: [{ rating: '4.0 out of 5 stars', title: 'Nice', body: 'Blends well', author: 'Kim', date: 'May 2, 2026' }] }
  };
  const { ctx, sheets } = load({ reviewsStatus: 500, top: [], details });
  let state = ctx.startProduct('B0G1VBDXYF', false);
  while (!state.done) state = ctx.fetchReviewsChunk(state);
  assert.strictEqual(state.added, 1);
  const row = sheets.Reviews.rows[1];
  assert.deepStrictEqual([row[3], row[4], row[5], row[6]], [4, 'Nice', 'Blends well', 'Kim']);
  assert.strictEqual(sheets.Products.rows[1][6], 57);
});

test('explains what came back when no reviews are found', () => {
  const { ctx } = load({ reviewsStatus: 500, top: [], details: { product_title: 'X' } });
  let state = ctx.startProduct('B0G1VBDXYF', false);
  while (!state.done) state = ctx.fetchReviewsChunk(state);
  assert.match(state.note, /product details keys: product_title/);
});
