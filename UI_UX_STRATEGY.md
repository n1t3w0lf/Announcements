# UI/UX Strategy - Announcements Carousel Web Part

## Design Philosophy
The announcements carousel follows a content-first approach, prioritizing readability and visual impact while maintaining consistency with SharePoint's modern design language and Fluent UI patterns.

## Current User Journey

### Content Creator Flow
1. Navigate to "Announcements Carousel" SharePoint list (auto-provisioned)
2. Create new list item with Title, Description, Image URL, date range
3. Optionally select celebration icon type and position
4. Announcement automatically appears in carousel when within valid date range

### End User Flow
1. View carousel on SharePoint page
2. Announcements auto-rotate on configured interval
3. Navigate manually via dots or arrow buttons
4. View celebration icons with animated effects
5. Click links within rich-text descriptions

## Current UI Components

### Carousel Display
- Full-width slide with image, title, and description
- Configurable height (300-800px)
- Background color customization
- Border radius and shadow options
- Image overlay with adjustable opacity

### Navigation
- **Dots**: Centered below carousel, active state with blue highlight
- **Arrows**: Left/right chevrons, positioned mid-slide, semi-transparent background
- Both independently toggleable via property pane

### Celebration Icons
- Positioned absolute within image container (5 positions)
- Circular badge with emoji, colored background
- 5 animation types: bounce, pulse, wave, shine, sparkle
- Custom icon support via URL

### State Handling
- **Loading**: Centered Fluent UI Spinner
- **Error**: Fluent UI MessageBar (error type)
- **Empty**: Fluent UI MessageBar (info type) with list name guidance
- **Provisioning**: Spinner with "Setting up" message

## Responsive Design
- Breakpoint at 768px for mobile
- Reduced image container height (400px -> 250px)
- Compressed padding and navigation sizing
- Icon positioning adjustments for smaller screens

## Accessibility
- ARIA labels on navigation buttons and dots
- Semantic HTML structure (h2 for titles)
- Keyboard-navigable controls
- Alt text on images

## Areas for Future Enhancement
- Touch/swipe gesture support for mobile
- Keyboard arrow key navigation
- Reduced motion media query support
- High contrast theme support
- Focus management for screen readers
- Announcement count indicator
- Progress bar for auto-rotation timing
