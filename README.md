# Amazon Reviews → Google Sheets

Paste Amazon product links into a sidebar in Google Sheets and every review the API can reach is added to a **Reviews** tab. Each product also gets a status row in a **Products** tab.

- Works with full links (`/dp/`, `/gp/product/`, `/product-reviews/`), short links (`amzn.to`, `a.co`) and bare ASINs
- Supports Amazon US, CA, UK, DE, FR, IT, ES, JP, IN, AU, MX, BR and other marketplaces
- Skips reviews already in the sheet, so re-running a product only adds new ones
- Fetches in small chunks, so large products don't hit Google's 6-minute script limit

## Why it uses an API

Amazon blocks scrapers and now requires a login to view most review pages, so this tool gets reviews from the **Real-Time Amazon Data** API on RapidAPI. The free tier covers testing; paid tiers are inexpensive. Each API request returns up to 10 reviews.

## Setup (about 5 minutes)

1. **Get an API key**
   - Sign up at [rapidapi.com](https://rapidapi.com) and subscribe to **Real-Time Amazon Data** (publisher: letscrape).
   - Copy your `X-RapidAPI-Key` from the endpoint's page.
2. **Create the sheet**
   - Open a new Google Sheet, then go to **Extensions → Apps Script**.
   - Replace the contents of `Code.gs` with [`apps-script/Code.gs`](apps-script/Code.gs).
   - Click **+ → HTML**, name the file `Sidebar`, and paste in [`apps-script/Sidebar.html`](apps-script/Sidebar.html).
   - Optional: under **Project Settings**, enable "Show appsscript.json" and paste in [`apps-script/appsscript.json`](apps-script/appsscript.json). This limits the script's permissions to the current sheet.
   - Save, then reload the Google Sheet.
3. **Run it**
   - Open the new **Amazon Reviews** menu and choose **Open review fetcher**. The first time, Google will ask you to authorize the script.
   - Paste your RapidAPI key when the sidebar asks for it.
   - Paste product links, one per line, and click **Fetch reviews**.

To use the tool again in a new sheet, make a copy of this sheet (**File → Make a copy**). The script is copied with it.

## Getting all reviews (Amazon cookie)

Amazon only shows logged-out visitors the first page of reviews, about 8 "top reviews". Without a login, the tool gets those 8 automatically. To get the full review list, give the API a cookie from a logged-in Amazon session:

1. In Chrome, log in to Amazon. Use a secondary account, not your seller or main buyer account.
2. Open any product page and press **F12** to open DevTools, then click the **Network** tab.
3. Refresh the page and click the first request (the product page itself).
4. Under **Request Headers**, find `cookie:` and copy its value.
5. In the sheet, choose **Amazon Reviews → Set Amazon cookie** and paste it.

The tool keeps only the login cookies the API needs. Cookies expire after a while; if the tool falls back to top reviews again, paste a fresh one.

## Troubleshooting

Use **Amazon Reviews → Test API connection** to see the raw response from the API for one product.

- **HTTP 403**: the key is wrong, or you aren't subscribed to Real-Time Amazon Data.
- **HTTP 429**: you've hit your plan's rate limit or monthly quota.
- **"Done (top reviews only)" in the Products tab**: the full review list needs an Amazon cookie (see above).

## How many reviews will I get?

Amazon only shows about 10 pages (about 100 reviews) for each sort and filter combination. The tool pulls from several combinations and removes duplicates:

| Mode | Combinations | Max reviews per product | Max API requests per product |
|---|---|---|---|
| Normal | Top + Most recent | about 200 | 20 |
| Deep (checkbox) | Above × all ratings plus 5★, 4★, 3★, 2★, 1★ | about 1,200 | 120 |

For products with fewer reviews than these caps, you'll usually get all of them. A filter stops early when it runs out of pages, so small products use far fewer requests.

## Columns in the Reviews tab

ASIN · Product URL · Review ID · Rating · Title · Review · Author · Date · Verified Purchase · Vine Review · Helpful Votes · Variant · Images · Review Link · Fetched At

## Development

```bash
npm test   # runs Code.gs against mocked Apps Script services
```

To change the API provider or the limits, edit `CONFIG` and `callReviewsApi_` at the top of `Code.gs`.
