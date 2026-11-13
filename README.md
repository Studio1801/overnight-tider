# Overnight Tider

Single-page site for Overnight Tider (fresh & frozen seafood, retail and bulk). Static HTML, no build step. Deploy as-is on Vercel or any static host.

## Files
- `index.html` : the site. Owner details live in the `SHOP` settings block at the top of the script. Empty values stay hidden on the page.
- `404.html` : not-found page.
- `img/` : product photos from @overnight_tider.
- `alerts-apps-script.gs` : Google Apps Script for the weekly text list. Deploy as a web app and paste the URL into `SHOP.alertsEndpoint`.

## To fill in before launch
`phone`, `address`, `hours`, `pickup`, `cutoff`, `bulkTerms`, `alertsEndpoint`, and item prices in `ITEMS` (currently "Ask").

## Brand
`brand/` holds the logo: `mark.svg` (app icon / favicon), `logo.svg` (horizontal, light backgrounds), `logo-reverse.svg` (dark backgrounds), `logo-sign.svg` (caps lockup matching the painted building sign). Green #2f6b3d, navy #13202e, band yellow #f2c230.
