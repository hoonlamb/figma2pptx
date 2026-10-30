'use strict';
// The GUI listens on loopback, but that only stops other machines. Any page open in
// the user's browser can still send requests here, so check who is asking. Host
// stops DNS rebinding, Origin stops a cross-site fetch, and Sec-Fetch-Site stops an
// <img> or link, which sends no Origin at all.

function isLocalRequest(req, port) {
  const allowed = new Set([`127.0.0.1:${port}`, `[::1]:${port}`, `localhost:${port}`]);
  if (!allowed.has(req.headers.host)) return false;
  const site = req.headers['sec-fetch-site'];
  if (site && site !== 'same-origin' && site !== 'none') return false;
  const origin = req.headers.origin;
  if (!origin) return true;
  try {
    return allowed.has(new URL(origin).host);
  } catch {
    return false;
  }
}

module.exports = { isLocalRequest };
