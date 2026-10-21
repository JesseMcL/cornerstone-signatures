# Deployment guide

## 1. Prepare the service

Copy `.env.example` and fill every production setting. Use a stable HTTPS `PUBLIC_BASE_URL`, a stable UUID for `OUTLOOK_ADDIN_ID`, and a public `SOURCE_CODE_URL` containing the exact deployed source. Store `MICROSOFT_GRAPH_CLIENT_SECRET` only in your platform's encrypted or masked secret store.

Build and run the image with a persistent volume at `/app/data`:

```sh
docker build -t cornerstone-signatures:2.1.3 .
docker run --env-file .env -p 3000:3000 -v cornerstone-signatures-data:/app/data cornerstone-signatures:2.1.3
```

The health check is `GET /api/health`. Back up the database from **Manage → Data management** as well as backing up the volume.

## 2. Choose browser sign-in

The browser picker (`/`) and administration console (`/admin/`) need a signed-in user. Set `AUTH_MODE` to one of:

### Option A: Microsoft Entra sign-in (`AUTH_MODE=entra`)

The server signs users in with Microsoft itself using the OpenID Connect authorization code flow with PKCE, reusing the app registration and client secret from step 3. No external access proxy is required.

1. Set `AUTH_MODE=entra` and a random `SESSION_SECRET` of at least 32 characters, for example from `openssl rand -base64 48`. Store it as a secret. Changing it signs everyone out.
2. Optionally set `SESSION_TTL_SECONDS` (default 28800, eight hours). Every request re-checks the session against the staff database, so removing or blocking a staff member revokes access immediately. Disabling an Entra account revokes access after the next directory sync marks it disabled.
3. In the app registration, under **Authentication → Web**, add two redirect URIs: `https://YOUR_PUBLIC_HOST/auth/callback` (sign-in) and `https://YOUR_PUBLIC_HOST/auth/signed-out` (Microsoft only returns to registered URIs after sign-out). Do not configure a front-channel logout URL; signing out elsewhere ends the app session when it expires.
4. **Strongly recommended:** in **Enterprise applications → Cornerstone Signatures → Properties**, set **Assignment required** to **Yes** and assign the staff group. Only assigned users can then sign in, and Conditional Access and MFA policies apply to the picker and admin console like any other Microsoft 365 app.
5. If the container is exposed directly rather than behind a reverse proxy, set `TRUST_PROXY=false` so clients cannot spoof their IP to avoid rate limits.

Signed-in users are matched to staff records by Entra object ID. Until directory sync has linked a record, the user principal name is used instead, and the record is then permanently linked to that account. The mutable `email` claim is never used, and B2B guest accounts are refused. Management roles are still granted inside the application. Set `INITIAL_IT_ADMINS` to each bootstrap admin's user principal name (the address they sign in to Microsoft with) and have them sign in before the first directory sync. That first sign-in links the record to their account, and the next sync updates it to their primary email.

Session cookies are `HttpOnly`, `SameSite=Lax`, AES-256-GCM encrypted, and use the `__Host-` prefix over HTTPS. State-changing admin requests still require an allowed `Origin` from `ADMIN_ALLOWED_ORIGINS`.

### Option B: Cloudflare Access (`AUTH_MODE=cloudflare`, default)

Protect `/`, `/admin/*`, and `/api/admin/*`. Set `CLOUDFLARE_TEAM_DOMAIN`, `CLOUDFLARE_ACCESS_AUD`, and the exact browser origins in `ADMIN_ALLOWED_ORIGINS`.

Outlook cannot reuse a browser's Access session. Bypass Access only for:

```text
/outlook-addin/*
/api/outlook/signature
/.well-known/microsoft-officeaddins-allowed.json
```

The signature endpoint still verifies a Microsoft token for the configured tenant, audience, and delegated scope. Never bypass `/api/admin/*`.

## 3. Register Microsoft Entra

Create a single-tenant application registration and configure:

1. An Application ID URI and delegated `Signature.Read` scope.
2. The SPA redirect URI `brk-multihub://YOUR_PUBLIC_HOST` for nested app authentication.
3. Delegated Microsoft Graph `User.Read` with admin consent.
4. Application Microsoft Graph `User.Read.All` with admin consent for server-side directory synchronization and profile photos.
5. A client secret whose **value** is stored as `MICROSOFT_GRAPH_CLIENT_SECRET`. With `AUTH_MODE=entra` the same secret is used for browser sign-in.
5a. With `AUTH_MODE=entra`, the Web redirect URIs `https://YOUR_PUBLIC_HOST/auth/callback` and `https://YOUR_PUBLIC_HOST/auth/signed-out` (see step 2).
6. The public-host identifier URI `api://YOUR_PUBLIC_HOST/CLIENT_ID` and access-token version 2 for classic Outlook.
7. The Microsoft Office client application as an authorized client for the required delegated scope.

Use the resulting tenant ID, client ID, and API audience in the environment.

## 4. Generate and deploy the Outlook manifest

With the deployment environment loaded, run:

```sh
npm run configure:addin
npm run build:addin
npm run validate:addin
```

Upload `outlook-addin/manifest.xml` through Microsoft 365 **Settings → Integrated apps** and assign it to the intended users or group. The running server also exposes the configured manifest at `/outlook-addin/manifest.xml`. Manifest changes require another Microsoft 365 upload; ordinary application, template, and server changes do not.

Microsoft propagation may take hours. Restart Outlook after the assignment appears. Remove users' old automatic Outlook signatures to avoid duplicates.

## 5. First sign-in

Set one or more emergency addresses in `INITIAL_IT_ADMINS`, sign in (through Microsoft at `/admin/`, or through Access), then configure organization information, Entra filters, location mappings, roles, taglines, and templates in the application. Keep the bootstrap list small.

## Releases and containers

Pull requests and pushes run tests and add-in builds. Tags matching `v*` publish multi-architecture container images to GitHub Container Registry. Creating a GitHub Release also attaches a source archive and a generated sample manifest; see `.github/workflows`.
