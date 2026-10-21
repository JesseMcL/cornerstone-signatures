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
2. Optionally set `SESSION_TTL_SECONDS` (default 28800, eight hours). Sessions are re-checked against the staff database on every request, so removing or blocking a staff member, or disabling their Entra account and syncing, revokes access immediately.
3. In the app registration, add a **Web** platform redirect URI of `https://YOUR_PUBLIC_HOST/auth/callback`, and a front-channel logout URL of `https://YOUR_PUBLIC_HOST/auth/signed-out` if you want sign-out to clear the Microsoft session.
4. Recommended: in **Enterprise applications → Cornerstone Signatures → Properties**, set **Assignment required** to **Yes** and assign the staff group. Conditional Access and MFA policies then apply to the picker and admin console like any other Microsoft 365 app.

Signed-in users are matched to staff records by Entra object ID, falling back to the `email` or `preferred_username` claim until directory sync has linked the record. Only matched, enabled staff can sign in; management roles are still granted inside the application. Put the user principal name or primary email of each bootstrap admin in `INITIAL_IT_ADMINS`.

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
5a. With `AUTH_MODE=entra`, the Web redirect URI `https://YOUR_PUBLIC_HOST/auth/callback` (see step 2).
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
