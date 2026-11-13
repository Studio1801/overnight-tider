# Overnight Tider

Website and ordering system for Overnight Tider (fresh and frozen seafood, retail and wholesale). Static site, no build step, deployed on Vercel from `main`.

## What it does
- **This week's board.** Every item, unit and price in one list. Customers add items without leaving the page.
- **Party sizes.** Pick an occasion and a headcount, get suggested amounts, add them in one tap.
- **Direct orders.** Orders land in a Google Sheet with an order number and are emailed to the owner. The customer sees a confirmation and the next steps. Instagram stays available as a backup.
- **Wholesale quotes.** Restaurants and caterers send a structured request into the same sheet.
- **Text list.** Phone signups for the weekly drop are collected in the sheet.
- **Owner-run stock.** The owner changes prices, marks items sold out or going fast, and edits the special from a Google Sheet. No code involved.
- **Tracking.** Add to order, checkout opened, order sent, quote sent, signups, calls and Instagram taps all go to GA4 (optional) and Vercel Analytics.
- **Search and sharing.** Social share image, FAQ and business structured data, sitemap, robots.txt, canonical URL.

## Launch checklist (everything goes in the `SHOP` block at the top of the script in `index.html`)
1. **Orders backend.** Follow the steps at the top of `apps-script.gs` and paste the web app URL into `endpoint`.
2. **Stock sheet.** In the same Google Sheet, add a tab named `Stock` and paste in `stock-template.csv`. Then File > Share > Publish to web > choose the `Stock` tab > CSV. Paste that link into `stockCSV`.
   - `status`: `in`, `limited` or `out`
   - `price`: any text, e.g. `$34 a dozen`. Prices that contain "special" or "sale" show in the accent color.
   - `special` row: `note` is the top bar text, `status` set to `off` hides the bar.
3. **Owner details.** Fill in `phone`, `address`, `hours`, `pickup`, `cutoff` and `bulkTerms`. Anything left empty stays hidden.
4. **Policy.** Put the owner's refund and live seafood policy in `policy`, in their own words.
5. **Deposits (optional).** Paste a Square payment link into `depositLink`.
6. **Analytics.** Turn on Web Analytics in the Vercel project. Add a GA4 ID to `ga4` if wanted.
7. **Domain.** When a custom domain is connected, replace `overnight-tider.vercel.app` in `index.html` (canonical and og tags), `sitemap.xml`, `robots.txt` and `SHOP.siteUrl`.
8. **Google.** Create the Google Business Profile, link the site, and submit `sitemap.xml` in Search Console.

## Files
- `index.html`: the site
- `404.html`: not-found page
- `apps-script.gs`: Google Apps Script backend for orders, quotes and the text list
- `stock-template.csv`: starting rows for the owner's Stock tab
- `brand/`: logo files. `mark.svg` (icon and favicon), `logo.svg` (light backgrounds), `logo-reverse.svg` (dark backgrounds), `logo-sign.svg` (caps lockup matching the painted building sign). Green #2f6b3d, navy #13202e, band yellow #f2c230
- `img/`: product photos (WebP) and `og.jpg` share image
- `vercel.json`: caching and security headers
