# Project Working File - Announcements Carousel

## Session Log

### Session: 2026-03-23 - Initial Context Review

#### Status: Complete

#### What Was Done
1. Full codebase review completed
2. All source files read and analyzed
3. Project documentation files created:
   - `CODEBASE_CONTEXT.md` - Technical architecture and structure
   - `UI_UX_STRATEGY.md` - Design philosophy and user journeys
   - `PROJECT_WORKING_FILE.md` - This file (task tracking)
   - `LEARNINGS_AND_DISCOVERIES.md` - Insights and patterns found
4. Memory system initialized

---

### Session: 2026-03-23 - Announcement Management Panel Feature

#### Status: COMPLETE - Build passing, peer reviewed

#### Task
Add a management panel for site owners to create, edit, delete announcements and upload images directly.

#### Architecture Decisions
| Decision | Choice | Rationale |
|---|---|---|
| Management UI | Fluent UI `Panel` (PanelType.large) | DetailsList needs ~940px; matches SP management patterns |
| Permission check | `AssociatedOwnerGroup/users` + `IsSiteAdmin` fallback | Lightweight, single API call, graceful fallback |
| Image upload | POST binary to `Site Assets/Files/add()` | Native SP REST, no extra dependencies |
| Form placement | Inline collapsible section in Panel | Avoids z-index issues, keeps context visible |
| Date defaults | Factory function `createEmptyFormData()` | Ensures dates are fresh, not stale from module load |
| OData format | `nometadata` without `__metadata` in body | Consistent with read endpoints, avoids 400 errors |

#### Files Created (7)
| File | Lines | Purpose |
|---|---|---|
| `services/PermissionService.ts` | ~85 | Owners group membership + SiteAdmin fallback |
| `services/ImageUploadService.ts` | ~155 | Upload images to Site Assets with validation |
| `components/IAnnouncementFormFieldsProps.ts` | ~28 | Form component props interface |
| `components/AnnouncementFormFields.tsx` | ~220 | Reusable form with all fields + image upload |
| `components/IAnnouncementManagePanelProps.ts` | ~14 | Panel component props interface |
| `components/AnnouncementManagePanel.module.scss` | ~100 | Panel styles |
| `components/AnnouncementManagePanel.tsx` | ~450 | Panel with DetailsList + full CRUD |

#### Files Modified (4)
| File | Changes |
|---|---|
| `models/IAnnouncement.ts` | Added `IAnnouncementFormData` interface + `createEmptyFormData()` factory |
| `services/AnnouncementDataService.ts` | Fixed OData metadata mismatch, added `CustomIconUrl` handling, added ETag response check |
| `components/AnnouncementsCarousel.tsx` | Added `isOwner`/`isManagePanelOpen` state, permission check, manage button + panel integration |
| `components/AnnouncementsCarousel.module.scss` | Added `.manageButtonContainer` and `.manageButton` styles |

#### Peer Review Findings & Fixes
| # | Issue | Severity | Fix Applied |
|---|---|---|---|
| 1 | `__metadata` in body with `nometadata` header | Critical | Removed all `__metadata` blocks, use simple `{Url:'...'}` for URL fields |
| 3 | `ServerRelativeUrl` null not guarded in ImageUploadService | Important | Added null check with descriptive error |
| 4 | `EMPTY_FORM_DATA` static dates freeze at module load | Important | Replaced with `createEmptyFormData()` factory function |
| 5 | ETag GET response not checked in `updateAnnouncement` | Important | Added `response.ok` check before using ETag |
| 6 | Manage button not shown when announcements list empty | Important | Added `renderManageButton()` to error and empty state renders |
| 7 | Single quotes in login names break OData filter | Important | Added `replace(/'/g, "''")` before `encodeURIComponent` |
| 8 | Error MessageBar used `successBar` CSS class | Minor | Renamed to `feedbackBar` for both message bars |

#### Build Results
- First build: PASS (0 errors, 0 warnings)
- Post-review build: PASS (0 errors, 0 warnings)
- Node version: v18.20.8
- Build time: ~5.5s

#### User Journey
```
Owner sees "Manage Announcements" button (top-right of carousel)
  Click --> Panel slides in from right
    Sees all announcements in table (Image, Title, From, To, Status, Icon, Actions)
    "Add New Announcement" --> Form expands at top
      Fill Title, Description, Dates, Icon settings
      "Upload Image" --> File picker --> Uploads to Site Assets --> Thumbnail preview
      "Add Announcement" --> Saves to list, refreshes table
    Edit icon on row --> Form pre-populated --> "Save Changes"
    Delete icon --> Confirmation dialog --> "Delete"
    Close panel --> Carousel auto-refreshes if changes were made
Non-owners --> Button hidden, standard carousel view only
```

#### Next Steps
- User to test the feature manually
- Consider adding "Discard changes?" guard when closing panel with unsaved form data
- Consider sanitizing Description before `dangerouslySetInnerHTML` rendering (DOMPurify)
- Consider removing redundant `office-ui-fabric-react` dependency

---

### Session: 2026-03-23 - Image Size & Fit Controls

#### Status: COMPLETE - Build passing

#### Task
Add per-announcement and web part-level controls for image container height and image fit mode.

#### What Was Done
1. Added `ImageFitMode` enum (cover, contain, fill, none) to models
2. Added `ImageHeight` (Number) and `ImageFit` (Choice) columns to list provisioning
3. Added `imageHeight` slider (100-800px) and `imageFit` dropdown to property pane "Image Settings" group
4. Updated `AnnouncementDataService` to handle new fields in select, mapping, create, update
5. Added image sizing controls to the management form (per-announcement overrides)
6. Updated carousel rendering: per-announcement values override web part defaults
7. Removed hardcoded `height: 400px` and `object-fit: cover` from SCSS (now inline styles)
8. Responsive SCSS uses `max-height: 250px` instead of fixed height on mobile

#### Architecture
- **Two levels of control**: Web part property pane sets global defaults; per-announcement list columns override
- **Priority**: `announcement.ImageHeight || props.imageHeight` (per-item wins)
- **New list columns**: `ImageHeight` (Number, 100-1200, optional) + `ImageFit` (Choice: cover/contain/fill/none, default: cover)

#### Bug Fix (Bonus)
- Fixed `PermissionService` to use numeric user ID lookup (`getById`) instead of unreliable `$filter` on claims-encoded `LoginName` strings — this was causing the manage button to not appear on the workbench

#### Build Result
- PASS (0 errors, 0 warnings)

---

### Session: 2026-03-23 - Configurable Gradients & Image Width

#### Status: COMPLETE - Build passing

#### Task
Make background gradients configurable from the property pane. Add image width as a configurable property (both web part level and per-announcement).

#### What Was Done
1. Added 4 gradient properties to web part: `gradientStartColor`, `gradientEndColor`, `gradientDirection`, `overlayGradientColor`
2. Added `imageWidth` property (slider 0-2000px, 0 = full width)
3. Added "Background Gradient" group to property pane (page 2) with 4 controls
4. Added `ImageWidth` number column (100-2000) to list provisioning + data service
5. Added `ImageWidth` TextField to management form (per-announcement override)
6. Updated carousel rendering: gradient + width from inline styles, per-announcement overrides
7. Removed hardcoded `background: linear-gradient(...)` from `.imageContainer` and `.imageOverlay` in SCSS
8. Updated `AnnouncementManagePanel` edit prepopulation and submit payload for `ImageWidth`

#### Architecture
- **Gradient config**: 4 property pane fields → inline `style` on `.imageContainer` (container gradient) and `.imageOverlay` (overlay gradient)
- **Image width**: `imageWidth` slider at web part level, `ImageWidth` number column per-announcement; 0 = full width, >0 = fixed px
- **Two-level override pattern** (same as height/fit): `announcement.ImageWidth || props.imageWidth`

#### Files Modified (6)
| File | Changes |
|---|---|
| `IAnnouncementsCarouselProps.ts` | Added `imageWidth`, `gradientStartColor`, `gradientEndColor`, `gradientDirection`, `overlayGradientColor` |
| `AnnouncementsCarouselWebPart.ts` | Added properties + "Image Settings" width slider + "Background Gradient" group |
| `AnnouncementsCarousel.tsx` | Updated `renderAnnouncement()` with gradient + width inline styles |
| `AnnouncementsCarousel.module.scss` | Removed hardcoded gradients from `.imageContainer` and `.imageOverlay` |
| `AnnouncementFormFields.tsx` | Added `ImageWidth` TextField alongside `ImageHeight` |
| `AnnouncementManagePanel.tsx` | Added `ImageWidth` to edit prepopulation + submit payload |

#### Build Result
- PASS (0 errors, 0 warnings)
- Node: v18.20.8 | Build time: ~6s

---
*Last Updated: 2026-03-23*
