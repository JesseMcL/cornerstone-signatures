import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { after, before, describe, it } from 'node:test';
import { createLocalJWKSet, exportJWK, generateKeyPair, SignJWT } from 'jose';
import { createApp } from '../../server/app.js';
import { createAccessAuthenticator } from '../../server/auth.js';
import { assertProductionConfig } from '../../server/config.js';
import { openDatabase, resolveStaffEmailForIdentity } from '../../server/db.js';
import { createEntraLogin, isGuestIdentity, parseCookies, safeReturnTo } from '../../server/entra-login.js';
import { DEFAULT_DIRECTORY_FILTERS, syncDirectoryUsers } from '../../server/entra.js';

const TENANT = '11111111-1111-1111-1111-111111111111';
const CLIENT = '22222222-2222-2222-2222-222222222222';
const BASE = 'https://signatures.test';
const AUTHORITY = `https://login.microsoftonline.com/${TENANT}`;

describe('Microsoft Entra sign-in', () => {
  let db; let server; let baseUrl; let privateKey; let nextClaims; let tokenRequests;

  const signIdToken = (claims) => new SignJWT({ tid: TENANT, ver: '2.0', ...claims })
    .setProtectedHeader({ alg: 'RS256', kid: 'test-key' })
    .setIssuer(`${AUTHORITY}/v2.0`).setAudience(CLIENT).setIssuedAt().setExpirationTime('5m')
    .sign(privateKey);

  async function startLogin(returnTo = '/admin/') {
    const response = await fetch(`${baseUrl}/auth/login?returnTo=${encodeURIComponent(returnTo)}`, { redirect: 'manual' });
    assert.equal(response.status, 302);
    const location = new URL(response.headers.get('location'));
    const cookie = response.headers.getSetCookie().find((value) => value.startsWith('cs_login='));
    return { location, cookie: cookie.split(';')[0] };
  }

  async function signIn(claims, { returnTo } = {}) {
    const { location, cookie } = await startLogin(returnTo);
    const nonce = location.searchParams.get('nonce');
    nextClaims = { nonce, ...claims };
    const callback = await fetch(`${baseUrl}/auth/callback?code=test-code&state=${location.searchParams.get('state')}`, { headers: { cookie }, redirect: 'manual' });
    const session = callback.headers.getSetCookie().find((value) => value.startsWith('cs_session=') && !value.startsWith('cs_session=;'));
    return { callback, session: session?.split(';')[0], location };
  }

  before(async () => {
    const directory = mkdtempSync(path.join(tmpdir(), 'siggen-entra-test-'));
    mkdirSync(path.join(directory, 'outlook-addin/dist'), { recursive: true });
    writeFileSync(path.join(directory, 'outlook-addin/dist/classic-runtime-entry.js'), '');
    db = openDatabase(':memory:');
    const add = db.prepare('INSERT INTO staff(email,first_name,last_name,entra_object_id,directory_account_enabled) VALUES (?,?,?,?,?)');
    const admin = add.run('ada@example.com', 'Ada', 'Admin', 'oid-ada', 1).lastInsertRowid;
    add.run('pat@example.com', 'Pat', 'Person', 'oid-pat', 1);
    add.run('gone@example.com', 'Gone', 'User', 'oid-gone', 0);
    add.run('bootstrap@example.com', 'Boot', 'Strap', null, null);
    db.prepare(`INSERT INTO role_grants(staff_id,role,granted_by) VALUES (?,'it_admin','test')`).run(admin);

    const keys = await generateKeyPair('RS256');
    privateKey = keys.privateKey;
    const jwk = { ...(await exportJWK(keys.publicKey)), kid: 'test-key', alg: 'RS256', use: 'sig' };
    tokenRequests = [];
    const fetchImpl = async (url, init) => {
      tokenRequests.push({ url, body: new URLSearchParams(init.body) });
      return new Response(JSON.stringify({ id_token: await signIdToken(nextClaims) }), { status: 200, headers: { 'content-type': 'application/json' } });
    };
    // http base keeps cookies un-prefixed so the test client can send them over plain HTTP.
    const login = createEntraLogin({
      tenantId: TENANT, clientId: CLIENT, clientSecret: 'secret', publicBaseUrl: 'http://signatures.test',
      sessionSecret: 'x'.repeat(48), jwks: createLocalJWKSet({ keys: [jwk] }), fetchImpl,
      resolveStaffEmail: (identity) => resolveStaffEmailForIdentity(db, identity), logger: { warn() {}, error() {} },
    });
    const app = createApp({
      db,
      authRouter: login.router,
      auth: { identityResolver: login.resolveIdentity, logoutUrl: '/auth/logout', env: 'production', allowedOrigins: [BASE] },
      microsoftUserResolver: async () => { throw new Error('unused'); },
      publicRoot: directory,
      outlookConfig: { publicBaseUrl: BASE },
      logger: { info() {}, warn() {}, error() {} },
    });
    server = app.listen(0, '127.0.0.1');
    await new Promise((resolve) => server.once('listening', resolve));
    baseUrl = `http://127.0.0.1:${server.address().port}`;
  });

  after(() => { server?.close(); db?.close(); });

  it('redirects to Microsoft with PKCE, state and nonce', async () => {
    const { location, cookie } = await startLogin('/admin/#users');
    assert.equal(`${location.origin}${location.pathname}`, `${AUTHORITY}/oauth2/v2.0/authorize`);
    assert.equal(location.searchParams.get('client_id'), CLIENT);
    assert.equal(location.searchParams.get('response_type'), 'code');
    assert.equal(location.searchParams.get('redirect_uri'), 'http://signatures.test/auth/callback');
    assert.equal(location.searchParams.get('code_challenge_method'), 'S256');
    assert.match(location.searchParams.get('code_challenge'), /^[\w-]{43}$/);
    assert.ok(location.searchParams.get('state').length >= 43);
    assert.ok(location.searchParams.get('nonce').length >= 43);
    assert.match(cookie, /^cs_login=/);
  });

  it('signs in a managed staff member and opens the admin API', async () => {
    const { callback, session } = await signIn({ oid: 'oid-ada', preferred_username: 'ada.upn@example.com' }, { returnTo: '/admin/#users' });
    assert.equal(callback.status, 303);
    assert.equal(callback.headers.get('location'), '/admin/#users');
    assert.ok(session);
    const setCookie = callback.headers.getSetCookie().find((value) => value.startsWith('cs_session='));
    assert.match(setCookie, /HttpOnly/);
    assert.match(setCookie, /SameSite=Lax/);
    const request = tokenRequests.at(-1);
    assert.equal(request.url, `${AUTHORITY}/oauth2/v2.0/token`);
    assert.equal(request.body.get('grant_type'), 'authorization_code');
    assert.ok(request.body.get('code_verifier'));
    const response = await fetch(`${baseUrl}/api/admin/session`, { headers: { cookie: session } });
    assert.equal(response.status, 200);
    const body = await response.json();
    assert.equal(body.email, 'ada@example.com');
    assert.deepEqual(body.roles, ['it_admin']);
    assert.equal(body.logoutUrl, '/auth/logout');
  });

  it('lets staff without a role use the picker but not the admin API', async () => {
    const { session } = await signIn({ oid: 'oid-pat' });
    assert.equal((await fetch(`${baseUrl}/api/picker`, { headers: { cookie: session } })).status, 200);
    assert.equal((await fetch(`${baseUrl}/api/admin/session`, { headers: { cookie: session } })).status, 403);
  });

  it('links a bootstrap admin by principal name once, then refuses other accounts', async () => {
    const { callback } = await signIn({ oid: 'oid-boot', preferred_username: 'bootstrap@example.com' });
    assert.equal(callback.status, 303);
    assert.equal(db.prepare('SELECT entra_object_id FROM staff WHERE email=?').get('bootstrap@example.com').entra_object_id, 'oid-boot');
    assert.equal((await signIn({ oid: 'oid-other', preferred_username: 'bootstrap@example.com' })).callback.status, 403);
  });

  it('ignores the mutable email claim', async () => {
    db.prepare('INSERT INTO staff(email,first_name,last_name) VALUES (?,?,?)').run('unlinked@example.com', 'Un', 'Linked');
    assert.equal((await signIn({ oid: 'oid-x', email: 'unlinked@example.com', preferred_username: 'x@example.com' })).callback.status, 403);
  });

  it('rejects B2B guests even though their token carries this tenant', async () => {
    db.prepare('INSERT INTO staff(email,first_name,last_name) VALUES (?,?,?)').run('guestable@example.com', 'G', 'U');
    assert.equal((await signIn({ oid: 'oid-g1', preferred_username: 'guestable@example.com', idp: 'https://sts.windows.net/99999999-9999-9999-9999-999999999999/' })).callback.status, 403);
    assert.equal((await signIn({ oid: 'oid-g2', preferred_username: 'guestable@example.com', idp: 'live.com' })).callback.status, 403);
    assert.equal((await signIn({ oid: 'oid-g3', upn: 'someone_evil.test#EXT#@corp.onmicrosoft.com', preferred_username: 'guestable@example.com' })).callback.status, 403);
    assert.equal((await signIn({ oid: 'oid-g4', preferred_username: 'guestable@example.com', acct: 1 })).callback.status, 403);
  });

  it('will not accept a session cookie as a login cookie', async () => {
    const { session } = await signIn({ oid: 'oid-ada' });
    const confused = session.replace('cs_session=', 'cs_login=');
    const before = tokenRequests.length;
    const response = await fetch(`${baseUrl}/auth/callback?code=x`, { headers: { cookie: confused }, redirect: 'manual' });
    assert.equal(response.status, 400);
    assert.equal(tokenRequests.length, before);
  });

  it('clears the session cookie on the signed-out page', async () => {
    const response = await fetch(`${baseUrl}/auth/signed-out`);
    assert.ok(response.headers.getSetCookie().some((value) => /^cs_session=;.*Max-Age=0/.test(value)));
  });

  it('returns 401 with a login URL when there is no session', async () => {
    const response = await fetch(`${baseUrl}/api/admin/session`);
    assert.equal(response.status, 401);
    assert.equal((await response.json()).error.details.loginUrl, '/auth/login?returnTo=%2F');
  });

  it('rejects tampered session cookies', async () => {
    const { session } = await signIn({ oid: 'oid-ada' });
    const tampered = `${session.slice(0, -4)}AAAA`;
    assert.equal((await fetch(`${baseUrl}/api/admin/session`, { headers: { cookie: tampered } })).status, 401);
  });

  it('rejects a mismatched state', async () => {
    const { cookie } = await startLogin();
    const response = await fetch(`${baseUrl}/auth/callback?code=x&state=wrong`, { headers: { cookie }, redirect: 'manual' });
    assert.equal(response.status, 400);
  });

  it('rejects a callback without the login cookie', async () => {
    const { location } = await startLogin();
    const response = await fetch(`${baseUrl}/auth/callback?code=x&state=${location.searchParams.get('state')}`, { redirect: 'manual' });
    assert.equal(response.status, 400);
  });

  it('rejects a mismatched nonce', async () => {
    const { location, cookie } = await startLogin();
    nextClaims = { nonce: 'not-the-nonce', oid: 'oid-ada' };
    const response = await fetch(`${baseUrl}/auth/callback?code=x&state=${location.searchParams.get('state')}`, { headers: { cookie }, redirect: 'manual' });
    assert.equal(response.status, 401);
  });

  it('rejects other tenants, unknown users and disabled accounts', async () => {
    assert.equal((await signIn({ oid: 'oid-ada', tid: '99999999-9999-9999-9999-999999999999' })).callback.status, 403);
    assert.equal((await signIn({ oid: 'oid-stranger', preferred_username: 'stranger@example.com' })).callback.status, 403);
    assert.equal((await signIn({ oid: 'oid-gone' })).callback.status, 403);
  });

  it('refuses a principal-name match for a record linked to a different Entra object', async () => {
    assert.equal((await signIn({ oid: 'oid-impostor', preferred_username: 'ada@example.com' })).callback.status, 403);
  });

  it('revokes access when the staff record is removed', async () => {
    db.prepare('INSERT INTO staff(email,first_name,last_name,entra_object_id) VALUES (?,?,?,?)').run('temp@example.com', 'Temp', 'User', 'oid-temp');
    const { session } = await signIn({ oid: 'oid-temp' });
    assert.equal((await fetch(`${baseUrl}/api/picker`, { headers: { cookie: session } })).status, 200);
    db.prepare('DELETE FROM staff WHERE entra_object_id=?').run('oid-temp');
    assert.equal((await fetch(`${baseUrl}/api/picker`, { headers: { cookie: session } })).status, 403);
  });

  it('still requires an allowed Origin for admin changes', async () => {
    const { session } = await signIn({ oid: 'oid-ada' });
    const response = await fetch(`${baseUrl}/api/admin/taglines`, { method: 'POST', headers: { cookie: session, 'content-type': 'application/json' }, body: '{}' });
    assert.equal(response.status, 403);
    assert.equal((await response.json()).error.code, 'invalid_origin');
  });

  it('clears the session and ends the Microsoft session on logout', async () => {
    const response = await fetch(`${baseUrl}/auth/logout`, { redirect: 'manual' });
    assert.equal(response.status, 302);
    assert.ok(response.headers.get('location').startsWith(`${AUTHORITY}/oauth2/v2.0/logout?`));
    assert.ok(response.headers.getSetCookie().some((value) => /^cs_session=;.*Max-Age=0/.test(value)));
  });
});

describe('Entra sign-in helpers', () => {
  it('only allows same-site return paths', () => {
    assert.equal(safeReturnTo('/admin/#users'), '/admin/#users');
    for (const value of ['//evil.test', 'https://evil.test', '/\\evil.test', 'admin', '/auth/callback', undefined, '/x\n']) {
      assert.equal(safeReturnTo(value), '/');
    }
  });

  it('parses cookies without trusting duplicates', () => {
    assert.deepEqual(parseCookies('a=1; b=two%20words; a=2; bad'), { a: '1', b: 'two words' });
  });

  it('refuses weak session secrets and missing settings', () => {
    const options = { tenantId: TENANT, clientId: CLIENT, clientSecret: 's', publicBaseUrl: BASE, resolveStaffEmail: () => null };
    assert.throws(() => createEntraLogin({ ...options, sessionSecret: 'short' }), /SESSION_SECRET/);
    assert.throws(() => createEntraLogin({ ...options, clientSecret: '', sessionSecret: 'x'.repeat(32) }), /CLIENT_SECRET/);
    assert.throws(() => createEntraLogin({ ...options, tenantId: 'contoso.onmicrosoft.com', sessionSecret: 'x'.repeat(32) }), /tenant GUID/);
    const secure = createEntraLogin({ ...options, sessionSecret: 'x'.repeat(32) });
    assert.deepEqual(secure.cookieNames, { session: '__Host-cs_session', login: '__Host-cs_login' });
  });

  it('uses the injected identity resolver instead of Cloudflare Access', async () => {
    const authenticate = createAccessAuthenticator({ identityResolver: async () => ({ email: 'ada@example.com' }), teamDomain: 'team.cloudflareaccess.com', audience: 'aud' });
    const req = { get: () => undefined };
    await new Promise((resolve, reject) => authenticate(req, {}, (error) => (error ? reject(error) : resolve())));
    assert.equal(req.identity.email, 'ada@example.com');
  });

  it('validates production configuration per auth mode', () => {
    const common = {
      env: 'production', adminAllowedOrigins: [BASE], publicBaseUrl: BASE, microsoftClientId: CLIENT, microsoftTenantId: TENANT,
      microsoftApiAudience: `api://${CLIENT}`, outlookAddinId: 'id', sourceCodeUrl: 'https://example.test/src', officeAddinRuntimeUrls: [],
    };
    assert.throws(() => assertProductionConfig({ ...common, authMode: 'cloudflare' }), /CLOUDFLARE_TEAM_DOMAIN/);
    assert.doesNotThrow(() => assertProductionConfig({ ...common, authMode: 'cloudflare', cloudflareTeamDomain: 't', cloudflareAccessAudience: 'a' }));
    assert.throws(() => assertProductionConfig({ ...common, authMode: 'entra', microsoftGraphClientSecret: 's' }), /SESSION_SECRET/);
    assert.throws(() => assertProductionConfig({ ...common, authMode: 'entra', microsoftGraphClientSecret: 's', sessionSecret: 'short' }), /32 characters/);
    assert.doesNotThrow(() => assertProductionConfig({ ...common, authMode: 'entra', microsoftGraphClientSecret: 's', sessionSecret: 'x'.repeat(32) }));
  });

  it('resolves staff by object ID before principal name', () => {
    const db = openDatabase(':memory:');
    db.prepare('INSERT INTO staff(email,first_name,last_name,entra_object_id) VALUES (?,?,?,?)').run('mail@example.com', 'A', 'B', 'oid-1');
    db.prepare('INSERT INTO staff(email,first_name,last_name) VALUES (?,?,?)').run('free@example.com', 'C', 'D');
    assert.equal(resolveStaffEmailForIdentity(db, { oid: 'oid-1', emails: ['upn@example.com'] }), 'mail@example.com');
    assert.equal(resolveStaffEmailForIdentity(db, { oid: 'oid-2', emails: ['mail@example.com'] }), null);
    assert.equal(resolveStaffEmailForIdentity(db, { oid: 'oid-3', emails: [] }), null);
    assert.equal(resolveStaffEmailForIdentity(db, { emails: ['free@example.com'] }), null);
    assert.equal(resolveStaffEmailForIdentity(db, { oid: 'oid-4', emails: ['free@example.com'] }), 'free@example.com');
    assert.equal(db.prepare('SELECT entra_object_id FROM staff WHERE email=?').get('free@example.com').entra_object_id, null);
    db.close();
  });

  it('identifies guest tokens', () => {
    assert.equal(isGuestIdentity({}, TENANT), false);
    assert.equal(isGuestIdentity({ idp: `https://sts.windows.net/${TENANT}/` }, TENANT), false);
    assert.equal(isGuestIdentity({ idp: 'https://sts.windows.net/other/' }, TENANT), true);
    assert.equal(isGuestIdentity({ preferred_username: 'a_b.c#ext#@t.onmicrosoft.com' }, TENANT), true);
  });

  it('merges a duplicate record when a sign-in-linked admin syncs to their primary email', () => {
    const db = openDatabase(':memory:', { initialItAdmins: ['jane.upn@example.com', 'jane@example.com'] });
    assert.equal(resolveStaffEmailForIdentity(db, { oid: 'oid-jane', emails: ['jane.upn@example.com'], bind: true }), 'jane.upn@example.com');
    const result = syncDirectoryUsers(db, [
      { id: 'oid-jane', accountEnabled: true, userType: 'Member', givenName: 'Jane', surname: 'Admin', mail: 'jane@example.com' },
      { id: 'oid-bob', accountEnabled: true, userType: 'Member', givenName: 'Bob', surname: 'New', mail: 'bob@example.com' },
    ], { ...DEFAULT_DIRECTORY_FILTERS, allowedDomains: [] }, 'test');
    assert.equal(result.merged, 1);
    assert.equal(result.created, 1);
    const jane = db.prepare('SELECT id,email FROM staff WHERE entra_object_id=?').get('oid-jane');
    assert.equal(jane.email, 'jane@example.com');
    assert.equal(db.prepare('SELECT COUNT(*) AS n FROM staff WHERE email LIKE ?').get('jane%').n, 1);
    assert.deepEqual(db.prepare('SELECT role FROM role_grants WHERE staff_id=?').all(jane.id).map((row) => row.role), ['it_admin']);
    db.close();
  });
});
