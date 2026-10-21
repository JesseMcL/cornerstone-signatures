// Microsoft Entra ID sign-in for the browser picker and administration console.
//
// Implements the OpenID Connect authorization code flow with PKCE as a
// confidential client (the same app registration and client secret used for
// directory sync). The browser never sees Microsoft tokens: after the ID token
// is validated, the server issues its own short-lived session in an encrypted,
// HttpOnly cookie. This replaces Cloudflare Access when AUTH_MODE=entra.

import { createHash, hkdfSync, randomBytes, timingSafeEqual } from 'node:crypto';
import express from 'express';
import { createRemoteJWKSet, EncryptJWT, jwtDecrypt, jwtVerify } from 'jose';
import { HttpError } from './errors.js';

const LOGIN_TTL_SECONDS = 10 * 60;
const DEFAULT_SESSION_TTL_SECONDS = 8 * 60 * 60;
const SESSION_ALG = { alg: 'dir', enc: 'A256GCM' };

const base64url = (buffer) => Buffer.from(buffer).toString('base64url');

function safeEqual(a, b) {
  const left = Buffer.from(String(a ?? ''));
  const right = Buffer.from(String(b ?? ''));
  return left.length === right.length && timingSafeEqual(left, right);
}

export function parseCookies(header) {
  const cookies = {};
  for (const part of String(header || '').split(';')) {
    const index = part.indexOf('=');
    if (index < 1) continue;
    const name = part.slice(0, index).trim();
    if (!name || Object.hasOwn(cookies, name)) continue;
    try { cookies[name] = decodeURIComponent(part.slice(index + 1).trim()); } catch { /* ignore malformed cookie */ }
  }
  return cookies;
}

// Only allow same-site relative paths so the login flow cannot be used as an open redirect.
export function safeReturnTo(value) {
  const path = typeof value === 'string' ? value : '';
  if (!path.startsWith('/') || path.startsWith('//') || path.includes('\\') || /[\u0000-\u001f]/.test(path) || path.length > 512) return '/';
  if (path.startsWith('/auth/')) return '/';
  return path;
}

const escapeHtml = (value) => String(value).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function page(title, message) {
  return `<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title>`
    + '<body style="font-family:system-ui,sans-serif;max-width:32rem;margin:4rem auto;padding:0 1rem;line-height:1.5">'
    + `<h1>${escapeHtml(title)}</h1>${message ? `<p>${escapeHtml(message)}</p>` : ''}<p><a href="/auth/login?returnTo=%2F">Sign in</a></p></body></html>`;
}

// Separate keys per cookie purpose, so a session cookie can never be replayed as a login cookie or vice versa.
function deriveKey(secret, purpose) {
  return new Uint8Array(hkdfSync('sha256', secret, 'cornerstone-signatures', `entra-${purpose}-v1`, 32));
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const TOKEN_TIMEOUT_MS = 10_000;

// B2B guests sign in through the resource tenant, so their tokens carry this tenant's tid.
// They are recognisable by an idp claim naming a different identity provider, or an #EXT# UPN.
export function isGuestIdentity(claims, tenantId) {
  const ownIssuers = new Set([`https://sts.windows.net/${tenantId}/`, `https://login.microsoftonline.com/${tenantId}/v2.0`]);
  if (typeof claims.idp === 'string' && claims.idp && !ownIssuers.has(claims.idp)) return true;
  if (claims.acct === 1 || claims.acct === '1') return true;
  return [claims.upn, claims.preferred_username, claims.unique_name].some((value) => typeof value === 'string' && value.toUpperCase().includes('#EXT#'));
}

const safeLogValue = (value) => String(value).replace(/[^\w.-]/g, '_').slice(0, 64);

export function createEntraLogin({
  tenantId,
  clientId,
  clientSecret,
  publicBaseUrl,
  sessionSecret,
  sessionTtlSeconds = DEFAULT_SESSION_TTL_SECONDS,
  resolveStaffEmail,
  authorityHost = 'https://login.microsoftonline.com',
  jwks,
  fetchImpl = globalThis.fetch,
  logger = console,
}) {
  if (!tenantId || !clientId || !clientSecret) throw new Error('Entra sign-in requires MICROSOFT_TENANT_ID, MICROSOFT_CLIENT_ID and MICROSOFT_GRAPH_CLIENT_SECRET.');
  if (!publicBaseUrl) throw new Error('Entra sign-in requires PUBLIC_BASE_URL.');
  if (!sessionSecret || String(sessionSecret).length < 32) throw new Error('SESSION_SECRET must be at least 32 characters when AUTH_MODE=entra.');
  if (typeof resolveStaffEmail !== 'function') throw new Error('Entra sign-in requires a staff resolver.');
  if (!GUID.test(String(tenantId))) throw new Error('MICROSOFT_TENANT_ID must be the tenant GUID when AUTH_MODE=entra.');
  tenantId = String(tenantId).toLowerCase();

  const baseUrl = String(publicBaseUrl).replace(/\/+$/, '');
  const secure = new URL(baseUrl).protocol === 'https:';
  // __Host- cookies must be Secure, Path=/ and have no Domain, which pins them to this exact origin.
  const sessionCookie = secure ? '__Host-cs_session' : 'cs_session';
  const loginCookie = secure ? '__Host-cs_login' : 'cs_login';
  const redirectUri = `${baseUrl}/auth/callback`;
  const authority = `${authorityHost.replace(/\/+$/, '')}/${encodeURIComponent(tenantId)}`;
  const issuer = `${authority}/v2.0`;
  const keys = jwks || createRemoteJWKSet(new URL(`${authority}/discovery/v2.0/keys`));
  const keys_ = { session: deriveKey(String(sessionSecret), 'session'), login: deriveKey(String(sessionSecret), 'login') };

  const cookieOptions = (maxAgeSeconds) => [
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    ...(secure ? ['Secure'] : []),
    `Max-Age=${maxAgeSeconds}`,
  ].join('; ');
  const setCookie = (res, name, value, maxAgeSeconds) => res.append('Set-Cookie', `${name}=${encodeURIComponent(value)}; ${cookieOptions(maxAgeSeconds)}`);
  const clearCookie = (res, name) => res.append('Set-Cookie', `${name}=; ${cookieOptions(0)}`);

  const seal = (purpose, claims, ttlSeconds) => new EncryptJWT({ ...claims, typ: purpose })
    .setProtectedHeader(SESSION_ALG)
    .setIssuedAt()
    .setExpirationTime(`${ttlSeconds}s`)
    .encrypt(keys_[purpose]);
  const unseal = async (purpose, token) => {
    const { payload } = await jwtDecrypt(token, keys_[purpose], { contentEncryptionAlgorithms: ['A256GCM'], keyManagementAlgorithms: ['dir'] });
    if (payload.typ !== purpose) throw new Error('wrong cookie type');
    return payload;
  };

  const loginUrl = (returnTo = '/') => `/auth/login?returnTo=${encodeURIComponent(safeReturnTo(returnTo))}`;

  async function resolveIdentity(req) {
    const token = parseCookies(req.get('cookie'))[sessionCookie];
    if (!token) throw new HttpError(401, 'authentication_required', 'Sign in with your Microsoft account to continue.', { loginUrl: loginUrl() });
    let session;
    try { session = await unseal('session', token); } catch {
      throw new HttpError(401, 'session_expired', 'Your session has expired. Sign in again to continue.', { loginUrl: loginUrl() });
    }
    // Re-resolve on every request so role and email changes from directory sync apply immediately,
    // and a staff record that has been removed or blocked loses access without waiting for expiry.
    const email = resolveStaffEmail({ oid: session.oid });
    if (!email) throw new HttpError(403, 'not_staff', 'Your Microsoft account is not a managed staff member.');
    return { email, microsoftId: session.oid };
  }

  const router = express.Router();
  router.use((_req, res, next) => { res.set('Cache-Control', 'no-store'); next(); });

  router.get('/login', async (req, res, next) => {
    try {
      const state = base64url(randomBytes(32));
      const nonce = base64url(randomBytes(32));
      const verifier = base64url(randomBytes(48));
      const challenge = base64url(createHash('sha256').update(verifier).digest());
      const returnTo = safeReturnTo(req.query.returnTo);
      setCookie(res, loginCookie, await seal('login', { state, nonce, verifier, returnTo }, LOGIN_TTL_SECONDS), LOGIN_TTL_SECONDS);
      const params = new URLSearchParams({
        client_id: clientId,
        response_type: 'code',
        redirect_uri: redirectUri,
        response_mode: 'query',
        scope: 'openid profile',
        state,
        nonce,
        code_challenge: challenge,
        code_challenge_method: 'S256',
      });
      res.redirect(302, `${authority}/oauth2/v2.0/authorize?${params}`);
    } catch (error) { next(error); }
  });

  router.get('/callback', async (req, res, next) => {
    try {
      const cookie = parseCookies(req.get('cookie'))[loginCookie];
      clearCookie(res, loginCookie);
      if (!cookie) throw new HttpError(400, 'login_expired', 'The sign-in attempt expired or was started in another browser. Try again.', { loginUrl: loginUrl() });
      let pending;
      try { pending = await unseal('login', cookie); } catch {
        throw new HttpError(400, 'login_expired', 'The sign-in attempt expired. Try again.', { loginUrl: loginUrl() });
      }
      if (typeof pending.state !== 'string' || pending.state.length < 32 || typeof pending.verifier !== 'string' || !safeEqual(req.query.state, pending.state)) throw new HttpError(400, 'invalid_state', 'The sign-in response did not match this browser session.');
      if (typeof req.query.error === 'string') {
        logger.warn?.(`[auth] Entra sign-in returned ${safeLogValue(req.query.error)}`);
        throw new HttpError(401, 'sign_in_failed', 'Microsoft sign-in did not complete. Check that you have been assigned access to Cornerstone Signatures.');
      }
      const code = typeof req.query.code === 'string' ? req.query.code : '';
      if (!code || code.length > 4096) throw new HttpError(400, 'invalid_request', 'The sign-in response is missing its authorization code.');

      const tokenResponse = await fetchImpl(`${authority}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          grant_type: 'authorization_code',
          code,
          redirect_uri: redirectUri,
          code_verifier: pending.verifier,
          scope: 'openid profile',
        }),
        signal: AbortSignal.timeout(TOKEN_TIMEOUT_MS),
      });
      const tokens = await tokenResponse.json().catch(() => ({}));
      if (!tokenResponse.ok || typeof tokens.id_token !== 'string') {
        logger.warn?.(`[auth] Entra token exchange failed: ${safeLogValue(tokens.error || tokenResponse.status)}`);
        throw new HttpError(401, 'sign_in_failed', 'Microsoft sign-in could not be completed. Try again.');
      }

      let claims;
      try {
        ({ payload: claims } = await jwtVerify(tokens.id_token, keys, { issuer, audience: clientId, algorithms: ['RS256'], clockTolerance: 60 }));
      } catch {
        throw new HttpError(401, 'invalid_id_token', 'The Microsoft sign-in token could not be verified.');
      }
      if (!safeEqual(claims.nonce, pending.nonce)) throw new HttpError(401, 'invalid_id_token', 'The Microsoft sign-in token could not be verified.');
      if (String(claims.tid).toLowerCase() !== tenantId) throw new HttpError(403, 'wrong_tenant', 'This Microsoft account belongs to a different organization.');
      if (typeof claims.oid !== 'string' || !claims.oid) throw new HttpError(401, 'invalid_id_token', 'The Microsoft sign-in token has no user identifier.');

      if (isGuestIdentity(claims, tenantId)) throw new HttpError(403, 'guest_not_allowed', 'Guest accounts cannot sign in to Cornerstone Signatures.');

      // Object ID first. Until directory sync links a record, fall back to the user principal name, which
      // is unique in the tenant and set by administrators. The mutable email claim is deliberately not used.
      // A fallback match binds the record to this object ID so it cannot be claimed by another account later.
      const principalNames = [claims.upn, claims.preferred_username].filter((value) => typeof value === 'string' && value.includes('@'));
      const email = resolveStaffEmail({ oid: claims.oid, emails: principalNames, bind: true });
      if (!email) throw new HttpError(403, 'not_staff', 'Your Microsoft account is not a managed staff member. Ask IT to run a directory sync or add you as staff.');

      setCookie(res, sessionCookie, await seal('session', { oid: claims.oid, email }, sessionTtlSeconds), sessionTtlSeconds);
      res.redirect(303, safeReturnTo(pending.returnTo));
    } catch (error) { next(error); }
  });

  // Sign out locally, then end the Microsoft session too.
  router.get('/logout', (_req, res) => {
    clearCookie(res, sessionCookie);
    const params = new URLSearchParams({ post_logout_redirect_uri: `${baseUrl}/auth/signed-out` });
    res.redirect(302, `${authority}/oauth2/v2.0/logout?${params}`);
  });

  router.get('/signed-out', (_req, res) => {
    clearCookie(res, sessionCookie);
    res.type('html').send(page('You have signed out', ''));
  });

  // Sign-in happens in a top-level browser navigation, so show errors as a page rather than JSON.
  router.use((error, _req, res, _next) => {
    const status = error instanceof HttpError ? error.status : 500;
    if (status === 500) logger.error?.(error);
    const message = error instanceof HttpError ? error.message : 'An unexpected error occurred during sign-in.';
    res.status(status).type('html').send(page('Sign-in problem', message));
  });

  return { router, resolveIdentity, loginUrl, cookieNames: { session: sessionCookie, login: loginCookie }, redirectUri };
}
