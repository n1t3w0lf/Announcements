# Learnings and Discoveries

## Session: 2026-03-23 - Initial Context Review

### Discovery 1: Play/Pause Button Dead Code
- `renderPlayPauseButton()` method exists at `AnnouncementsCarousel.tsx:304` but is never called in the `render()` method
- The `togglePause` handler and `isPaused` state are still wired up
- The styles for `.playPauseButton` exist in the SCSS
- **Impact**: Dead code that could confuse future developers. Was intentionally removed per commit `2c62990`

### Discovery 2: Redundant UI Library Dependencies
- Both `@fluentui/react` (^8.110.10) and `office-ui-fabric-react` (^7.204.0) are in package.json
- Only `@fluentui/react` is imported in the source code
- **Impact**: `office-ui-fabric-react` adds unnecessary bundle size. Should be removed.

### Discovery 3: OData Format Inconsistency
- `AnnouncementDataService.ts` uses `odata=nometadata` format
- `ListProvisioningService.ts` uses `odata=minimalmetadata` format
- Both work but inconsistency could cause confusion
- **Learning**: Different OData formats serve different purposes - provisioning may need metadata for `@odata.type` resolution

### Discovery 4: createAnnouncement Uses Metadata Despite nometadata Header
- In `AnnouncementDataService.ts:130`, the `createAnnouncement` method includes `__metadata` in the request body but sends `odata=nometadata` header
- This is a potential issue - the metadata type reference may be ignored or cause errors with nometadata format
- **Learning**: When using `odata=nometadata`, you don't need `__metadata` in the body; when using `odata=verbose`, you do

### Discovery 5: No Error Boundary
- The component has no React Error Boundary
- If the carousel throws during render, the entire web part will crash
- **Learning**: SPFx best practice is to wrap components in error boundaries

### Discovery 6: Class Component Pattern
- The main component uses class-based React pattern
- Modern SPFx development (1.18.x) supports React hooks via functional components
- **Learning**: Not a bug, but functional components with hooks would be more maintainable

### Discovery 7: Image URL Handling
- `AnnouncementImage` field maps from `item.AnnouncementImage?.Url || item.AnnouncementImage || ''`
- This handles both URL field object format and raw string format
- **Learning**: SP REST API returns URL fields differently based on OData format - the dual fallback handles this correctly

### Discovery 8: Timer Leak Potential
- `startRotation()` at line 104 doesn't check if a timer already exists before creating a new one
- If called twice without `stopRotation()`, the old timer leaks
- `componentDidMount` calls it after `loadAnnouncements`, but `componentDidUpdate` properly stops before restarting

---

## Session: 2026-03-23 - Management Panel Feature

### Discovery 9: OData __metadata Causes 400 with nometadata Content-Type
- Sending `__metadata` blocks in the request body while using `Content-Type: application/json;odata=nometadata` can cause SharePoint to return 400 Bad Request
- **Fix**: Either remove `__metadata` (for nometadata) or switch headers to `odata=verbose` (if metadata is needed)
- **Learning**: OData format in headers and body must be consistent. For URL fields with nometadata, simple `{Url: '...'}` objects work fine

### Discovery 10: EMPTY_FORM_DATA as Module Constant Freezes Dates
- `export const EMPTY_FORM_DATA = { ValidFrom: new Date().toISOString(), ... }` evaluates `new Date()` once at module import time
- In long-running SPA sessions (common in SharePoint), the dates become stale hours later
- **Fix**: Use a factory function `createEmptyFormData()` that creates fresh dates on each call
- **Learning**: Never use `new Date()` or `Date.now()` in module-level constants for form defaults

### Discovery 11: OData $filter and Single Quotes in Login Names
- SharePoint Online login names like `i:0#.f|membership|o'brien@contoso.com` contain single quotes
- `encodeURIComponent` does NOT escape OData string delimiters — only URL-unsafe chars
- Unescaped single quotes in `$filter=LoginName eq '...'` break the OData parser (400 error)
- **Fix**: Double the single quotes (`replace(/'/g, "''")`) BEFORE `encodeURIComponent`
- **Learning**: OData string escaping and URL encoding are separate concerns; both must be applied

### Discovery 12: SPHttpClient.post Body Accepts ArrayBuffer Despite String Typing
- `SPHttpClient.post()` TypeScript signature declares `body: string`
- At runtime, it passes through to native `fetch()` which accepts `ArrayBuffer`, `Blob`, etc.
- For binary file uploads, pass `arrayBuffer as any` — works correctly in all SP Online environments
- **Learning**: SPFx TypeScript types are occasionally narrower than the runtime API supports

### Discovery 13: Site Assets Library May Not Exist on New Sites
- Brand new SharePoint sites may not have a "Site Assets" library provisioned
- `GET /_api/web/lists/getByTitle('Site Assets')` returns 404
- **Fix**: Auto-create with `BaseTemplate: 101` (Document Library) — mirrors ListProvisioningService pattern
- **Learning**: Never assume standard libraries exist; always provision-or-verify

### Discovery 14: ETag GET Not Validated Before MERGE/DELETE
- The existing `updateAnnouncement` and `deleteAnnouncement` pattern does GET (for ETag) then write
- If the GET fails (404 — item deleted by another user, or 403), `response.ok` was never checked
- Code silently falls back to `ETag: *` which bypasses concurrency control
- **Fix**: Check `response.ok` after the GET and throw with a descriptive error before attempting write

### Discovery 15: Panel Render Must Not Be Gated by Early Returns
- If the "Manage Announcements" button is only rendered in the final render branch (after all early returns), it becomes invisible when announcements list is empty or in error state
- This means site owners cannot create the FIRST announcement through the management UI
- **Fix**: Include `renderManageButton()` in all non-loading, non-provisioning render branches

---

## Session: 2026-03-23 - Image Size Controls & Permission Fix

### Discovery 16: Claims-Encoded LoginName Breaks OData $filter
- SharePoint Online login names like `i:0#.f|membership|user@tenant.com` contain `#`, `|`, `:` characters
- Even with `encodeURIComponent`, OData `$filter=LoginName eq '...'` fails to match these values
- **Fix**: Use numeric user ID lookup instead: `AssociatedOwnerGroup/users/getById(userId)`
- The user ID is available from `context.pageContext.legacyPageContext.userId` (fast, no API call)
- **Learning**: Never use string-based OData filtering on claims-encoded SharePoint login names. Always prefer numeric ID lookups.

### Discovery 17: Hardcoded CSS Dimensions Block Runtime Customization
- The image container had `height: 400px` hardcoded in SCSS and `object-fit: cover` on the image
- These cannot be overridden by inline styles because CSS Modules generate scoped class names with equal specificity
- **Fix**: Remove hardcoded values from SCSS, use inline `style` attribute for runtime-configurable dimensions
- **Learning**: For properties that need to be dynamic (per-item or per-config), prefer inline styles over SCSS rules. Use SCSS for structural layout, inline for configurable values.

### Discovery 18: Number Columns Need Explicit FieldTypeKind
- SharePoint REST API requires `'FieldTypeKind': 9` for Number fields and `'@odata.type': 'SP.FieldNumber'`
- Optional constraints: `MinimumValue`, `MaximumValue`
- **Learning**: Always include both `@odata.type` and `FieldTypeKind` in column creation for reliability across environments

---

## Session: 2026-03-23 - Configurable Gradients & Image Width

### Discovery 19: CSS Background Shorthand Overrides Inline Background
- When SCSS has `background: linear-gradient(...)` and inline style also sets `background`, CSS specificity rules apply
- CSS Modules scoped classes and inline styles have equivalent specificity, but the SCSS `background` shorthand resets all background sub-properties
- **Fix**: Remove hardcoded `background` from SCSS entirely when the value needs to be runtime-configurable
- **Learning**: For any CSS property that will vary at runtime, the SCSS should not set it at all — not even as a "default". Use inline styles exclusively.

---
*Last Updated: 2026-03-23*
