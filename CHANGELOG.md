# Changelog

## Unreleased

- Added built-in Microsoft Entra ID sign-in for the browser picker and administration console (`AUTH_MODE=entra`), removing the need for Cloudflare Access. Uses the OpenID Connect authorization code flow with PKCE, the existing app registration and client secret, and an encrypted HttpOnly session cookie.
- Signed-in users are matched to staff by Entra object ID, and sessions are re-checked on every request so removed, blocked or disabled staff lose access immediately.
- Added a sign-out link to the administration console when Entra sign-in is enabled.
- Cloudflare Access remains available as `AUTH_MODE=cloudflare` and is still the default.

## 2.1.3

- Updated the transitive `source-map-js` development dependency (Dependabot).

## 2.1.2

- Added drag-and-drop ordering for administrator-managed professional designations. The saved order now controls both the user picker and the rendered `{{designations}}` value.
- Added keyboard reordering with the Up and Down arrow keys on each designation drag handle.

## 2.1.1

- Added Office-provided client diagnostics to distinguish Classic Outlook for Windows, desktop Outlook for Mac, new Outlook for Windows, Outlook on the web by operating system, iOS/iPadOS, and Android.
- Corrected Outlook on the web on macOS being classified as the desktop Mac client when only a browser user agent was available.
- Added accessible pie charts for Outlook client-family and From-address usage analytics while retaining platform and version detail.

## 2.1.0

- Added administrator-managed professional designations and user self-service selection through the `{{designations}}` merge tag.
- Added an optional default location mapping for staff whose Entra office location is blank.
- Added privacy-preserving, aggregate 30-day analytics for successful deliveries by Outlook client/version and primary versus alternate From address.
- Added automatic in-place migration for existing Cornerstone Signatures 2.0 databases and compatible database imports.

## 2.0.3

- Restored Classic Outlook signature retrieval by serving an absolute signature API URL in the classic runtime.

## 2.0.0

- First public Cornerstone Signatures release.
- Clean v2 database baseline with one-way import from final Siggen schema 13 exports.
- Removed organization-specific staff, taglines, locations, links, assets, domains, tenant identifiers, and deployment hosts.
- Added environment-driven Outlook configuration and manifest generation.
- Added container publishing, release automation, deployment documentation, and AGPL-3.0-only licensing.
