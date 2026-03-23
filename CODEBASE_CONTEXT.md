# Codebase Context - Announcements Carousel Web Part

## Project Overview
SPFx 1.18.2 web part that displays announcements in a carousel format with celebration icons, auto-provisioning, and rich customization options.

## Tech Stack
| Technology | Version |
|---|---|
| SPFx | 1.18.2 |
| TypeScript | 4.7.4 |
| React | 17.0.1 |
| Fluent UI React | ^8.110.10 |
| Node.js | >=16.13.0 <17.0.0 |
| Build Tool | Gulp 4.0.2 |

## Architecture

### Directory Structure
```
src/webparts/announcementsCarousel/
├── AnnouncementsCarouselWebPart.ts          # Web part entry point, property pane config
├── AnnouncementsCarouselWebPart.manifest.json
├── components/
│   ├── AnnouncementsCarousel.tsx             # Main React component (class-based)
│   ├── AnnouncementsCarousel.module.scss     # Component styles (403 lines)
│   ├── AnnouncementsCarousel.module.scss.ts  # Generated CSS module types
│   └── IAnnouncementsCarouselProps.ts        # Props interface (21 properties)
├── models/
│   └── IAnnouncement.ts                      # Data models + enums
├── services/
│   ├── AnnouncementDataService.ts            # CRUD operations via SP REST API
│   ├── CelebrationIconService.ts             # Icon config (emoji, colors, animations)
│   └── ListProvisioningService.ts            # Auto-creates SP list with columns
└── loc/
    ├── en-us.js                              # Localization
    └── mystrings.d.ts                        # String types
```

### Key Components

#### AnnouncementsCarouselWebPart.ts
- Entry point extending `BaseClientSideWebPart`
- 18 configurable properties across 3 property pane pages
- Auto-provisions list on `onInit()`

#### AnnouncementsCarousel.tsx (Main Component)
- Class-based React component
- State: announcements[], currentIndex, loading, error, isProvisioning, isPaused
- Handles carousel rotation with `setInterval`
- Renders: slides, navigation dots, arrows, celebration icons
- Transition effects: fade, slide, zoom

#### Services
- **AnnouncementDataService**: Full CRUD via SP REST API with OData nometadata
- **CelebrationIconService**: 8 predefined icon configs with animations (bounce, pulse, wave, shine, sparkle)
- **ListProvisioningService**: Creates "Announcements Carousel" list with 7 custom columns

### Data Model
```typescript
IAnnouncement {
  Id, Title, Description, AnnouncementImage,
  ValidFrom, ValidTo, CelebrationIcon, CelebrationIconPosition, CustomIconUrl, IsActive
}

CelebrationType: None | Birthday | Anniversary | Achievement | Celebration | NewHire | Promotion | Holiday | Custom
IconPosition: topLeft | topRight | bottomLeft | bottomRight | center
```

### SharePoint List Columns
| Column | Type | Required |
|---|---|---|
| Title | Single Line Text | Yes (default) |
| Description | Multi-line Rich Text | No |
| AnnouncementImage | URL | No |
| ValidFrom | DateTime | Yes |
| ValidTo | DateTime | Yes |
| CelebrationIcon | Choice (9 options) | No |
| CelebrationIconPosition | Choice (5 options) | No |
| CustomIconUrl | URL | No |

### Property Pane Configuration (3 Pages)
1. **Data Source & Behavior**: List name, auto-play, rotation interval, transitions
2. **Appearance**: Title/description display, font sizes, colors, height, border radius, shadow, overlay
3. **Celebration Icons**: Show/hide icons, icon size

### ESLint Configuration
- Extends: `@microsoft/eslint-config-spfx/lib/profiles/react`
- `@typescript-eslint/no-explicit-any`: off
- `@typescript-eslint/no-unused-vars`: warn
- `react/jsx-no-bind`: off
- `@typescript-eslint/explicit-function-return-type`: off

### Build Commands
- `npm run build` - Bundle
- `npm run serve` - Dev server (port 4321)
- `npm run package` - Production build + .sppkg

### Recent Git History
- `da1d02b` - fix: Remove unused imports and escape quotes in JSX
- `88df804` - feat: Add separate color controls for title and description
- `2c62990` - feat: Add custom icon support and remove play/pause button and date display
- `0ea8d1b` - fix: Use OData minimalmetadata format instead of verbose
- `9b3ef5d` - fix: Add delay and better error handling for list provisioning
