// Amazon Reviews tool, file 4 of 4 (Sheets.gs: writing rows). All four files must be in the Apps Script project.


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
