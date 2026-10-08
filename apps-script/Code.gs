/**
 * Amazon Reviews -> Google Sheets
 *
 * Paste Amazon product links into the sidebar and every review the API can
 * reach is appended to the "Reviews" tab. Reviews are de-duplicated by review
 * ID, so re-running a product only adds new reviews.
 *
 * Reviews come from the "Real-Time Amazon Data" API on RapidAPI, because
 * Amazon blocks direct scraping and gates most review pages behind a login.
 */

var CONFIG = {
  API_HOST: 'real-time-amazon-data.p.rapidapi.com',
  REVIEWS_SHEET: 'Reviews',
  PRODUCTS_SHEET: 'Products',
  // Amazon serves at most ~10 pages (≈100 reviews) per sort/filter combo.
  MAX_PAGES_PER_FILTER: 10,
  // Pages fetched per server call; keeps each call well under Apps Script's 6-minute limit.
  PAGES_PER_CALL: 4,
  // Stop a filter early after this many pages in a row with no new reviews.
  MAX_STALE_PAGES: 2,
  MAX_CELL_CHARS: 49000,
  API_KEY_PROPERTY: 'RAPIDAPI_KEY',
  // Optional Amazon session cookie. Amazon only shows the first page of reviews
  // (~8 "top reviews") to logged-out visitors; the full review list needs a login.
  COOKIE_PROPERTY: 'AMAZON_COOKIE'
};

var REVIEW_HEADERS = [
  'ASIN', 'Product URL', 'Review ID', 'Rating', 'Title', 'Review', 'Author',
  'Date', 'Verified Purchase', 'Vine Review', 'Helpful Votes', 'Variant',
  'Images', 'Review Link', 'Fetched At'
];

var PRODUCT_HEADERS = [
  'Product Link', 'ASIN', 'Marketplace', 'Status', 'Reviews in Sheet',
  'New This Run', 'Total Ratings on Amazon', 'Last Run'
];

// Marketplace domain suffix -> API country code.
var COUNTRY_BY_DOMAIN = {
  'com': 'US', 'ca': 'CA', 'com.mx': 'MX', 'com.br': 'BR', 'co.uk': 'GB',
  'de': 'DE', 'fr': 'FR', 'it': 'IT', 'es': 'ES', 'nl': 'NL', 'se': 'SE',
  'pl': 'PL', 'com.be': 'BE', 'com.tr': 'TR', 'ae': 'AE', 'sa': 'SA',
  'eg': 'EG', 'in': 'IN', 'co.jp': 'JP', 'sg': 'SG', 'com.au': 'AU'
};

// ---------------------------------------------------------------------------
// Menu & UI
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Amazon Reviews')
    .addItem('Open review fetcher', 'showSidebar')
    .addItem('Set RapidAPI key', 'promptForApiKey')
    .addItem('Set Amazon cookie (optional, unlocks all reviews)', 'promptForCookie')
    .addItem('Test API connection', 'testApiConnection')
    .addSeparator()
    .addItem('Set up sheets', 'setupSheets')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function showSidebar() {
  setupSheets();
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Amazon Reviews');
  SpreadsheetApp.getUi().showSidebar(html);
}

function promptForApiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt(
    'RapidAPI key',
    'Paste your RapidAPI key for "Real-Time Amazon Data" (stored only in this spreadsheet):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getDocumentProperties().setProperty(CONFIG.API_KEY_PROPERTY, key);
  ui.alert('API key saved.');
}

function saveApiKey(key) {
  key = String(key || '').trim();
  if (!key) throw new Error('API key is empty.');
  PropertiesService.getDocumentProperties().setProperty(CONFIG.API_KEY_PROPERTY, key);
  return true;
}

function hasApiKey() {
  return !!getApiKey_();
}

function getApiKey_() {
  return PropertiesService.getDocumentProperties().getProperty(CONFIG.API_KEY_PROPERTY);
}

function getCookie_() {
  return PropertiesService.getDocumentProperties().getProperty(CONFIG.COOKIE_PROPERTY);
}

function promptForCookie() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt(
    'Amazon cookie (optional)',
    'Without this you only get the ~8 top reviews Amazon shows logged-out visitors.\n' +
    'Paste the "cookie" request header from a logged-in Amazon tab (see README). ' +
    'Leave empty and press OK to remove it.',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var cookie = res.getResponseText().trim().replace(/^cookie:\s*/i, '');
  var props = PropertiesService.getDocumentProperties();
  if (cookie) {
    props.setProperty(CONFIG.COOKIE_PROPERTY, cookie);
    ui.alert('Amazon cookie saved.');
  } else {
    props.deleteProperty(CONFIG.COOKIE_PROPERTY);
    ui.alert('Amazon cookie removed.');
  }
}

/** Calls both review endpoints for one product and shows the raw result, for troubleshooting. */
function testApiConnection() {
  var ui = SpreadsheetApp.getUi();
  if (!getApiKey_()) { ui.alert('Set your RapidAPI key first (Amazon Reviews > Set RapidAPI key).'); return; }
  var res = ui.prompt('Test API connection', 'Paste one Amazon product link or ASIN:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var parsed = parseAmazonLink(res.getResponseText());
  if (!parsed) { ui.alert('Could not find an ASIN in that link.'); return; }

  var report = [];
  [
    ['/product-reviews', reviewParams_(parsed.asin, parsed.country, buildFilters(false)[0], 1, null)],
    ['/top-product-reviews', { asin: parsed.asin, country: parsed.country }],
    ['/product-details', { asin: parsed.asin, country: parsed.country }]
  ].forEach(function (c) {
    var r = rawApiCall_(c[0], c[1]);
    var data = r.json && r.json.data;
    var count = extractReviews_(data).length;
    report.push(c[0] + '  →  HTTP ' + r.code + ', ' + count + ' reviews found\n' +
      'keys: ' + describeKeys_(data) + '\n' + r.text.slice(0, 400));
  });
  report.push('Amazon cookie set: ' + (getCookie_() ? 'yes' : 'no'));
  ui.alert('API test for ' + parsed.asin, report.join('\n\n'), ui.ButtonSet.OK);
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, CONFIG.PRODUCTS_SHEET, PRODUCT_HEADERS);
  ensureSheet_(ss, CONFIG.REVIEWS_SHEET, REVIEW_HEADERS);
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Links in the Products tab that have not finished yet, for the sidebar's "load" button. */
function getPendingLinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PRODUCTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows
    .filter(function (r) { return r[0] && String(r[3]).indexOf('Done') !== 0; })
    .map(function (r) { return String(r[0]).trim(); });
}

// ---------------------------------------------------------------------------
// Link parsing
// ---------------------------------------------------------------------------

/**
 * Extracts { asin, country, url } from an Amazon product link.
 * Supports /dp/, /gp/product/, /product-reviews/, /gp/aw/d/ and bare ASINs.
 * Returns null if no ASIN can be found.
 */
function parseAmazonLink(link) {
  link = String(link || '').trim();
  if (!link) return null;

  if (/^[A-Z0-9]{10}$/i.test(link)) {
    return { asin: link.toUpperCase(), country: 'US', url: 'https://www.amazon.com/dp/' + link.toUpperCase() };
  }

  var asinMatch = link.match(/\/(?:dp|gp\/product|gp\/aw\/d|product-reviews|exec\/obidos\/asin)\/([A-Z0-9]{10})(?=[\/?#&]|$)/i);
  if (!asinMatch) return null;
  var asin = asinMatch[1].toUpperCase();

  var country = 'US';
  var host = link.match(/^(?:https?:\/\/)?([^\/?#]+)/i);
  var domainMatch = host && host[1].toLowerCase().match(/amazon\.([a-z.]+)$/);
  if (domainMatch && COUNTRY_BY_DOMAIN[domainMatch[1]]) {
    country = COUNTRY_BY_DOMAIN[domainMatch[1]];
  }
  var domain = domainMatch ? domainMatch[1] : 'com';
  return { asin: asin, country: country, url: 'https://www.amazon.' + domain + '/dp/' + asin };
}

/** Follows redirects on short links (amzn.to, a.co, amzn.eu) to reach the real product URL. */
function resolveShortLink_(link) {
  var url = link;
  for (var i = 0; i < 5; i++) {
    if (!/^(https?:\/\/)?(amzn\.(to|eu|asia)|a\.co)\//i.test(url) && parseAmazonLink(url)) return url;
    var res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    var headers = res.getAllHeaders();
    var location = headers.Location || headers.location;
    if (!location) return url;
    url = Array.isArray(location) ? location[0] : location;
  }
  return url;
}

// ---------------------------------------------------------------------------
// Fetching (driven step by step from the sidebar)
// ---------------------------------------------------------------------------

/**
 * Validates a link and returns the initial job state for fetchReviewsChunk().
 * deepMode also walks every star rating, which reaches far more reviews on
 * popular products at the cost of more API requests.
 */
function startProduct(link, deepMode) {
  if (!getApiKey_()) throw new Error('No RapidAPI key set. Use Amazon Reviews > Set RapidAPI key.');
  setupSheets();

  var parsed = parseAmazonLink(link);
  if (!parsed && /^(https?:\/\/)?(amzn\.(to|eu|asia)|a\.co)\//i.test(String(link).trim())) {
    parsed = parseAmazonLink(resolveShortLink_(String(link).trim()));
  }
  if (!parsed) throw new Error('Could not find a product ASIN in: ' + link);

  var row = upsertProductRow_(link, parsed, { status: 'Running…' });
  return {
    link: link,
    asin: parsed.asin,
    country: parsed.country,
    url: parsed.url,
    productRow: row,
    filters: buildFilters(deepMode),
    filterIndex: 0,
    page: 1,
    stalePages: 0,
    added: 0,
    totalRatings: '',
    done: false
  };
}

function buildFilters(deepMode) {
  var sorts = ['TOP_REVIEWS', 'MOST_RECENT'];
  var stars = deepMode
    ? ['ALL', '5_STARS', '4_STARS', '3_STARS', '2_STARS', '1_STARS']
    : ['ALL'];
  var filters = [];
  stars.forEach(function (star) {
    sorts.forEach(function (sort) { filters.push({ sort: sort, star: star }); });
  });
  return filters;
}

/**
 * Fetches up to CONFIG.PAGES_PER_CALL pages for the job, appends new reviews
 * to the sheet and returns the updated state. The sidebar calls this until
 * state.done is true.
 *
 * If the full review list fails before any page succeeds (usually because
 * Amazon wants a logged-in session), it falls back to the public top reviews.
 */
function fetchReviewsChunk(state) {
  var reviewsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.REVIEWS_SHEET);
  var seen = getExistingReviewIds_(reviewsSheet, state.asin);
  var newRows = [];
  var now = new Date();

  function addReviews(reviews) {
    var added = 0;
    reviews.forEach(function (r) {
      var id = r.review_id || hashReview_(r);
      if (seen[id]) return;
      seen[id] = true;
      added++;
      newRows.push(reviewToRow(r, id, state, now));
    });
    return added;
  }

  try {
    for (var n = 0; n < CONFIG.PAGES_PER_CALL && !state.done; n++) {
      if (state.mode === 'top') {
        var top = callApi_('/top-product-reviews', { asin: state.asin, country: state.country });
        var topReviews = extractReviews_(top);
        if (!topReviews.length) {
          // Product details usually include the same top reviews as part of the listing.
          var details = callApi_('/product-details', { asin: state.asin, country: state.country });
          topReviews = extractReviews_(details);
          if (!topReviews.length) {
            state.note = 'No reviews found in the API response (top reviews keys: ' + describeKeys_(top) +
              '; product details keys: ' + describeKeys_(details) + '). Run Amazon Reviews > Test API connection and share the result.';
          }
          if (!state.totalRatings && details) {
            state.totalRatings = details.product_num_ratings || details.total_ratings || '';
          }
        }
        addReviews(topReviews);
        state.done = true;
        break;
      }

      var filter = state.filters[state.filterIndex];
      var data;
      try {
        data = callApi_('/product-reviews',
          reviewParams_(state.asin, state.country, filter, state.page, state.cursor));
      } catch (e) {
        if (e.fatal) throw e;
        state.lastError = e.message;
        if (!state.pagesOk) {
          state.mode = 'top';
          state.note = 'Full review list unavailable (' + e.message + '). Got Amazon\'s public top reviews instead' +
            (getCookie_() ? '.' : ' — add an Amazon cookie to get all reviews.');
          continue;
        }
        state.note = 'Stopped early: ' + e.message;
        state.done = true;
        break;
      }
      state.pagesOk = (state.pagesOk || 0) + 1;

      var reviews = extractReviews_(data);
      if ((data.total_ratings || data.total_reviews) && !state.totalRatings) {
        state.totalRatings = data.total_ratings || data.total_reviews;
      }
      var newOnPage = addReviews(reviews);
      var nextCursor = data.next_cursor || data.cursor || null;

      state.stalePages = newOnPage === 0 ? state.stalePages + 1 : 0;
      var filterExhausted =
        reviews.length === 0 ||
        reviews.length < 10 ||
        state.page >= CONFIG.MAX_PAGES_PER_FILTER ||
        state.stalePages >= CONFIG.MAX_STALE_PAGES ||
        data.has_next_page === false ||
        (nextCursor && nextCursor === state.cursor);

      if (filterExhausted) {
        state.filterIndex++;
        state.page = 1;
        state.cursor = null;
        state.stalePages = 0;
        if (state.filterIndex >= state.filters.length) state.done = true;
      } else {
        state.page++;
        state.cursor = nextCursor;
      }
    }
  } finally {
    // Save whatever we got even if the API errored mid-chunk.
    if (newRows.length) {
      reviewsSheet.getRange(reviewsSheet.getLastRow() + 1, 1, newRows.length, REVIEW_HEADERS.length)
        .setValues(newRows);
      state.added += newRows.length;
    }
  }

  var inSheet = Object.keys(seen).length;
  state.reviewsInSheet = inSheet;
  state.progress = state.done
    ? 'Done'
    : state.mode === 'top'
      ? 'Getting top reviews'
      : 'Filter ' + (state.filterIndex + 1) + '/' + state.filters.length + ', page ' + state.page;
  upsertProductRow_(state.link, state, {
    status: state.done ? (state.mode === 'top' ? 'Done (top reviews only)' : 'Done') : 'Running…',
    inSheet: inSheet,
    added: state.added,
    totalRatings: state.totalRatings
  });
  return state;
}

/** Marks a product as failed in the Products tab (called by the sidebar on error). */
function markProductError(link, message) {
  var parsed = parseAmazonLink(link) || { asin: '', country: '' };
  upsertProductRow_(link, parsed, { status: 'Error: ' + message });
}

function reviewParams_(asin, country, filter, page, cursor) {
  var params = { asin: asin, country: country, sort_by: filter.sort };
  if (filter.star !== 'ALL') params.star_rating = filter.star;
  if (cursor) params.cursor = cursor;
  else params.page = page;
  var cookie = getCookie_();
  if (cookie) params.cookie = trimCookie_(cookie);
  return params;
}

/**
 * Keeps only the Amazon login cookies the API needs. Apps Script caps request
 * URLs at about 2 KB, and a full browser cookie header is usually longer.
 */
function trimCookie_(cookie) {
  var keep = /^(session-id|session-id-time|session-token|ubid-[a-z]+|at-[a-z]+|sess-at-[a-z]+|x-[a-z]+|i18n-prefs|lc-[a-z]+)$/i;
  return cookie.split(';')
    .map(function (c) { return c.trim(); })
    .filter(function (c) { return keep.test(c.split('=')[0]); })
    .join('; ');
}

/** The review list can come back as data.reviews, data.top_reviews, or a bare array. */
/**
 * Finds the list of reviews anywhere in an API response. Field names differ
 * between endpoints and API versions, so this looks for the first array of
 * objects that look like reviews instead of relying on one key.
 */
function extractReviews_(data, depth) {
  depth = depth || 0;
  if (!data || typeof data !== 'object' || depth > 4) return [];
  if (Array.isArray(data)) {
    if (data.length && looksLikeReview_(data[0])) return data;
    for (var i = 0; i < data.length; i++) {
      var inner = extractReviews_(data[i], depth + 1);
      if (inner.length) return inner;
    }
    return [];
  }
  var preferred = ['reviews', 'top_reviews', 'product_reviews', 'customer_reviews'];
  var keys = preferred.filter(function (k) { return k in data; })
    .concat(Object.keys(data).filter(function (k) { return preferred.indexOf(k) === -1; }));
  for (var j = 0; j < keys.length; j++) {
    var found = extractReviews_(data[keys[j]], depth + 1);
    if (found.length) return found;
  }
  return [];
}

function looksLikeReview_(o) {
  if (!o || typeof o !== 'object' || Array.isArray(o)) return false;
  return Object.keys(o).some(function (k) { return /^review_|^(rating|stars?|body|comment|review)$/i.test(k); });
}

/** Short description of a response's shape, so "no reviews" messages show what actually came back. */
function describeKeys_(data) {
  if (!data || typeof data !== 'object') return String(data);
  return Object.keys(data).slice(0, 15).join(', ') || 'empty';
}

function rawApiCall_(path, params) {
  var query = Object.keys(params)
    .map(function (k) { return k + '=' + encodeURIComponent(params[k]); })
    .join('&');
  var url = 'https://' + CONFIG.API_HOST + path + '?' + query;
  if (url.length > 2000) {
    return { code: 0, text: 'Request too long for Google Apps Script (' + url.length +
      ' chars). Your Amazon cookie is too long; paste only the session-id, session-token, ubid-main, at-main, sess-at-main and x-main values.', json: null };
  }
  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'x-rapidapi-key': getApiKey_(),
      'x-rapidapi-host': CONFIG.API_HOST
    }
  });
  var text = res.getContentText();
  var json = null;
  try { json = JSON.parse(text); } catch (e) { /* not JSON */ }
  return { code: res.getResponseCode(), text: text, json: json };
}

/** Calls the API with retries. Throws errors with .fatal = true when retrying other products is pointless. */
function callApi_(path, params) {
  var r;
  for (var attempt = 0; attempt < 3; attempt++) {
    r = rawApiCall_(path, params);
    if (r.code === 429 || r.code >= 500) {
      if (attempt < 2) Utilities.sleep(2000 * Math.pow(2, attempt));
      continue;
    }
    break;
  }

  var detail = apiErrorDetail_(r);
  if (r.code === 401 || r.code === 403) {
    throw fatalError_('RapidAPI rejected the request (HTTP ' + r.code + '): ' + detail +
      '. Check your key and that you are subscribed to Real-Time Amazon Data.');
  }
  if (r.code === 429) {
    throw fatalError_('RapidAPI rate limit or monthly quota reached (HTTP 429): ' + detail +
      '. Wait a minute, or check your plan usage on RapidAPI.');
  }
  if (r.code !== 200 || !r.json) {
    throw new Error('HTTP ' + r.code + ': ' + detail);
  }
  if (r.json.status && r.json.status !== 'OK') {
    throw new Error('API status ' + r.json.status + ': ' + detail);
  }
  return r.json.data || {};
}

function apiErrorDetail_(r) {
  var j = r.json;
  var msg = j && (j.message || (j.error && (j.error.message || j.error)) || j.status);
  if (msg && typeof msg !== 'string') msg = JSON.stringify(msg);
  return String(msg || r.text || 'no response body').slice(0, 200);
}

function fatalError_(message) {
  var e = new Error(message);
  e.fatal = true;
  return e;
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

function reviewToRow(r, id, state, now) {
  var images = (r.review_images || []).join('\n');
  return [
    state.asin,
    state.url,
    id,
    parseFloat(pick_(r, ['review_star_rating', 'rating', 'stars', 'star_rating'])) ||
      pick_(r, ['review_star_rating', 'rating', 'stars', 'star_rating']),
    clip_(pick_(r, ['review_title', 'title'])),
    clip_(pick_(r, ['review_comment', 'review_text', 'review_body', 'body', 'text', 'comment', 'content'])),
    clip_(pick_(r, ['review_author', 'author', 'reviewer_name', 'author_name'])),
    clip_(pick_(r, ['review_date', 'date'])),
    r.is_verified_purchase === true ? 'Yes' : r.is_verified_purchase === false ? 'No' : '',
    r.is_vine === true ? 'Yes' : r.is_vine === false ? 'No' : '',
    clip_(r.helpful_vote_statement),
    clip_(formatVariant_(r.review_variant || r.product_variant)),
    clip_(images),
    clip_(r.review_link),
    now
  ];
}

function pick_(o, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (o[keys[i]] !== undefined && o[keys[i]] !== null && o[keys[i]] !== '') return o[keys[i]];
  }
  return '';
}

function formatVariant_(v) {
  if (!v) return '';
  if (typeof v === 'string') return v;
  return Object.keys(v).map(function (k) { return k + ': ' + v[k]; }).join(', ');
}

function clip_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v);
  // Stop Sheets from treating review text that starts with = + - @ as a formula.
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s.length > CONFIG.MAX_CELL_CHARS ? s.slice(0, CONFIG.MAX_CELL_CHARS) + '…' : s;
}

function hashReview_(r) {
  var key = [r.review_author, r.review_date, r.review_title, (r.review_comment || '').slice(0, 200)].join('|');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key);
  return 'h_' + bytes.map(function (b) { return ((b + 256) % 256).toString(16); }).join('');
}

function getExistingReviewIds_(sheet, asin) {
  var seen = {};
  var last = sheet.getLastRow();
  if (last < 2) return seen;
  var values = sheet.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === asin && values[i][2]) seen[values[i][2]] = true;
  }
  return seen;
}

/** Finds the product's row by ASIN (or link) and updates it, appending one if missing. Returns the row number. */
function upsertProductRow_(link, info, fields) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PRODUCTS_SHEET);
  var last = sheet.getLastRow();
  var rowNum = -1;
  if (last >= 2) {
    var values = sheet.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if ((info.asin && values[i][1] === info.asin) || String(values[i][0]).trim() === String(link).trim()) {
        rowNum = i + 2;
        break;
      }
    }
  }
  if (rowNum === -1) rowNum = last + 1;

  var current = rowNum <= last
    ? sheet.getRange(rowNum, 1, 1, PRODUCT_HEADERS.length).getValues()[0]
    : ['', '', '', '', '', '', '', ''];
  var row = [
    current[0] || link,
    info.asin || current[1],
    info.country || current[2],
    fields.status !== undefined ? fields.status : current[3],
    fields.inSheet !== undefined ? fields.inSheet : current[4],
    fields.added !== undefined ? fields.added : current[5],
    fields.totalRatings ? fields.totalRatings : current[6],
    new Date()
  ];
  sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  return rowNum;
}
