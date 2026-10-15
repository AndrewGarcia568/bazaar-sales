# Bazaar Sales Mobile — installable offline app

This edition adapts the Bazaar Sales system for phones and tablets. It uses a persistent left sidebar on larger screens and a slide-out side menu on narrow phone screens. It is an installable Progressive Web App (PWA), so users can add it to the home screen and open it in its own app-like window.

## Features retained

- Product setup and editing
- Product selection, quantities, cart, Cash and GCash checkout
- Sales history, deleting individual sales, and clearing sales
- Dashboard totals, product tally, and payment tally
- Five-sheet Excel export
- Local persistence between launches
- Clean first run: no products and no sales are preloaded

## Offline behavior

The app shell, styles, scripts, icons, and Excel library are cached by a service worker. After the app has been loaded once while online, it can be opened offline. All sales and product data stay in this browser's local storage on that device; there is no cloud sync. Export the Excel workbook regularly as a backup.

**Important:** A PWA must be served from HTTPS (or localhost for development) to install and cache offline. You cannot install it just by opening `index.html` from a file manager. The first setup/deployment requires internet access, but day-to-day app use does not.

## Offline Excel export

The Excel workbook is written by a small local JavaScript component in `vendor/xlsx-offline.js`. It has no CDN or npm dependency, so no separate build step is required.

## Publish it so your phone can install it

Upload the **contents** of this folder (including the `vendor/` folder) to a static HTTPS host such as GitHub Pages or Netlify. Keep the folder structure intact. Then open the published HTTPS URL on the phone.

### Android

1. Open the app URL in Chrome.
2. Open the browser menu and choose **Install app** or **Add to Home screen**.
3. Confirm. Launch Bazaar Sales from the home screen.
4. Open it once while online so the offline cache is populated; then test it in airplane mode.

### iPhone / iPad

1. Open the app URL in Safari.
2. Tap **Share** → **Add to Home Screen**.
3. Confirm and launch from the home screen.
4. Keep the app open until its first load completes, then test offline.

## Local preview

Run a static web server in this folder and open its localhost URL. For example, with Python installed:

```sh
python -m http.server 8080
```

Then visit `http://localhost:8080` on that computer. Localhost is treated as a secure context for service workers.

## Storage notes

This mobile PWA uses new storage keys (`bazaarSalesMobileProductsV1` and `bazaarSalesMobileSalesV1`), so it starts with an empty product list and no sales and does not reuse the prior desktop/browser data. Browser site-data clearing can erase the records. Keep regular Excel exports. If you publish an updated version, change `CACHE_NAME` in `service-worker.js` so installed clients fetch the new app files.

## Scope / packaging note

This is an installable PWA package, not a signed Android APK. It is the simplest cross-platform installation path for Android and iPhone without requiring app-store publication. If you specifically need an APK file, that requires an additional Android wrapper/build step and signing/distribution decisions.
