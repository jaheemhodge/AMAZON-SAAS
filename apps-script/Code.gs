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
  API_KEY_PROPERTY: 'RAPIDAPI_KEY'
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
 */
function fetchReviewsChunk(state) {
  var reviewsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.REVIEWS_SHEET);
  var seen = getExistingReviewIds_(reviewsSheet, state.asin);
  var newRows = [];
  var now = new Date();

  try {
    for (var n = 0; n < CONFIG.PAGES_PER_CALL && !state.done; n++) {
      var filter = state.filters[state.filterIndex];
      var data = callReviewsApi_(state.asin, state.country, filter, state.page);
      var reviews = (data && data.reviews) || [];
      if (data && (data.total_ratings || data.total_reviews) && !state.totalRatings) {
        state.totalRatings = data.total_ratings || data.total_reviews;
      }

      var newOnPage = 0;
      reviews.forEach(function (r) {
        var id = r.review_id || hashReview_(r);
        if (seen[id]) return;
        seen[id] = true;
        newOnPage++;
        newRows.push(reviewToRow(r, id, state, now));
      });

      state.stalePages = newOnPage === 0 ? state.stalePages + 1 : 0;
      var filterExhausted =
        reviews.length === 0 ||
        reviews.length < 10 ||
        state.page >= CONFIG.MAX_PAGES_PER_FILTER ||
        state.stalePages >= CONFIG.MAX_STALE_PAGES ||
        (data && data.has_next_page === false);

      if (filterExhausted) {
        state.filterIndex++;
        state.page = 1;
        state.stalePages = 0;
        if (state.filterIndex >= state.filters.length) state.done = true;
      } else {
        state.page++;
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
    : 'Filter ' + (state.filterIndex + 1) + '/' + state.filters.length + ', page ' + state.page;
  upsertProductRow_(state.link, state, {
    status: state.done ? 'Done' : 'Running…',
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

function callReviewsApi_(asin, country, filter, page) {
  var params = {
    asin: asin,
    country: country,
    page: page,
    sort_by: filter.sort,
    star_rating: filter.star,
    verified_purchases_only: 'false',
    images_or_videos_only: 'false',
    current_format_only: 'false'
  };
  var query = Object.keys(params)
    .map(function (k) { return k + '=' + encodeURIComponent(params[k]); })
    .join('&');
  var url = 'https://' + CONFIG.API_HOST + '/product-reviews?' + query;
  var options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'x-rapidapi-key': getApiKey_(),
      'x-rapidapi-host': CONFIG.API_HOST
    }
  };

  for (var attempt = 0; attempt < 3; attempt++) {
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    if (code === 429 || code >= 500) {
      Utilities.sleep(2000 * Math.pow(2, attempt));
      continue;
    }
    if (code === 401 || code === 403) {
      throw new Error('RapidAPI rejected the key (HTTP ' + code + '). Check your key and that you are subscribed to Real-Time Amazon Data.');
    }
    if (code !== 200) {
      throw new Error('API error HTTP ' + code + ': ' + res.getContentText().slice(0, 200));
    }
    var json = JSON.parse(res.getContentText());
    if (json.status && json.status !== 'OK') {
      throw new Error('API returned status ' + json.status + ': ' + JSON.stringify(json.error || json).slice(0, 200));
    }
    return json.data || {};
  }
  throw new Error('API kept rate-limiting or failing after 3 attempts. Try again in a minute.');
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
    Number(r.review_star_rating) || r.review_star_rating || '',
    clip_(r.review_title),
    clip_(r.review_comment),
    clip_(r.review_author),
    clip_(r.review_date),
    r.is_verified_purchase === true ? 'Yes' : r.is_verified_purchase === false ? 'No' : '',
    r.is_vine === true ? 'Yes' : r.is_vine === false ? 'No' : '',
    clip_(r.helpful_vote_statement),
    clip_(formatVariant_(r.review_variant || r.product_variant)),
    clip_(images),
    clip_(r.review_link),
    now
  ];
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
