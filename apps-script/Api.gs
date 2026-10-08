// Amazon Reviews tool, file 3 of 4 (Api.gs: RapidAPI calls). All four files must be in the Apps Script project.

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
