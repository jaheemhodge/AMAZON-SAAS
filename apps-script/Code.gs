// Amazon Reviews tool, file 1 of 4 (Code.gs: settings and menu). Also add Fetch.gs, Api.gs and Sheets.gs.

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
