/*
 * Copies a website form submission into the MundyBall Google Sheet (tools/form-to-sheets.gs).
 * Runs only after Web3Forms has accepted the submission, and never blocks or changes what the
 * visitor sees: if the Sheets copy fails, the email from Web3Forms still went out.
 *
 * SHEETS_URL is the Apps Script web app (Deploy > Manage deployments > Web app URL). Empty = off.
 */
(function () {
  var SHEETS_URL = 'https://script.google.com/macros/s/AKfycbyvOsuSemzCcTGBpOMrG7dusSdwy2uTQUM-HnnMOLhQW2K8-fLxt8vBWH7j1XdjfXm4NQ/exec';
  var TOKEN = 'mundyball-forms-v1'; // must match TOKEN in tools/form-to-sheets.gs (not a password)
  var SKIP = { access_key: 1, subject: 1, from_name: 1, redirect: 1 };

  window.sendToSheets = function (formKey, fd) {
    if (!SHEETS_URL) return;
    try {
      var body = new URLSearchParams();
      fd.forEach(function (value, key) {
        if (SKIP[key] || typeof value !== 'string') return;
        body.append(key, value);
      });
      body.set('_form', formKey);
      body.set('_token', TOKEN);
      // sendBeacon survives the page changing or closing; fall back to a fire-and-forget fetch.
      if (navigator.sendBeacon && navigator.sendBeacon(SHEETS_URL, body)) return;
      fetch(SHEETS_URL, { method: 'POST', mode: 'no-cors', keepalive: true, body: body });
    } catch (err) {
      if (window.console) console.warn('Sheets copy skipped:', err);
    }
  };
})();
