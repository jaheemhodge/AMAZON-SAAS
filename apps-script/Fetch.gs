// Amazon Reviews tool, file 2 of 4 (Fetch.gs: link parsing and fetching). All four files must be in the Apps Script project.

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

