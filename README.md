# Announcements Carousel Web Part

A beautiful, feature-rich SharePoint Framework (SPFx) web part for displaying announcements and celebrations with automatic list provisioning and extensive customization options.

## Features

### 🎉 Core Functionality
- **Automatic List Provisioning**: Creates SharePoint list with all required columns on first use
- **Carousel Rotation**: Displays one announcement at a time with automatic rotation
- **Date-Based Filtering**: Only shows announcements within their valid date range
- **Celebration Icons**: Add themed celebration icons (birthdays, anniversaries, achievements, etc.)
- **Beautiful Animations**: Multiple transition effects (fade, slide, zoom)

### 🎨 Customization Options

#### Carousel Behavior
- Auto-play with configurable rotation interval (3-60 seconds)
- Manual navigation with arrows and dots
- Pause/play controls
- Three transition effects: Fade, Slide, Zoom

#### Content Display
- Toggle title and description visibility
- Adjustable font sizes for title (16-72px) and description (12-32px)
- Rich text support in descriptions
- Date range display

#### Visual Styling
- Custom background and text colors
- Adjustable carousel height (300-800px)
- Border radius control (0-50px)
- Optional shadow effect
- Image overlay with adjustable opacity (0-100%)

#### Celebration Icons
- 8 predefined celebration types with custom emojis:
  - 🎂 Birthday
  - 🎊 Anniversary
  - 🏆 Achievement
  - 🎉 Celebration
  - 👋 New Hire
  - ⭐ Promotion
  - 🎄 Holiday
  - ✨ Custom
- Adjustable icon size (30-120px)
- 5 positioning options: Top-Left, Top-Right, Bottom-Left, Bottom-Right, Center
- Animated effects: Bounce, Pulse, Wave, Shine, Sparkle

#### Navigation
- Navigation dots for quick slide access
- Arrow buttons for previous/next
- Responsive design for mobile devices

## SharePoint List Structure

The web part automatically creates a list named "Announcements Carousel" with these columns:

| Column Name | Type | Required | Description |
|------------|------|----------|-------------|
| Title | Single line text | Yes | Announcement title |
| Description | Multiple lines text (Rich) | No | Detailed description with formatting |
| AnnouncementImage | URL | No | Image URL for the announcement |
| ValidFrom | Date/Time | Yes | Start date for displaying the announcement |
| ValidTo | Date/Time | Yes | End date for displaying the announcement |
| CelebrationIcon | Choice | No | Type of celebration icon to display |
| CelebrationIconPosition | Choice | No | Position of the celebration icon |

## Installation

### Prerequisites
- Node.js v16.x
- SharePoint Online tenant
- SPFx development environment

### Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd Announcements
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Trust the development certificate** (first time only)
   ```bash
   gulp trust-dev-cert
   ```

4. **Start the local development server**
   ```bash
   gulp serve
   ```

5. **Test in SharePoint workbench**
   - Navigate to: `https://<your-tenant>.sharepoint.com/_layouts/workbench.aspx`
   - Add the "Announcements Carousel" web part to the page

## Building for Production

1. **Bundle the solution**
   ```bash
   gulp bundle --ship
   ```

2. **Package the solution**
   ```bash
   gulp package-solution --ship
   ```

3. **Deploy to SharePoint**
   - Navigate to your tenant's App Catalog
   - Upload the `.sppkg` file from `./sharepoint/solution/`
   - Deploy the solution
   - Trust the solution when prompted

## Usage

### Adding the Web Part

1. Edit a SharePoint page
2. Click the "+" icon to add a web part
3. Search for "Announcements Carousel"
4. Add the web part to your page

**Note**: On first use, the web part will automatically create the "Announcements Carousel" list if it doesn't exist.

### Creating Announcements

1. Navigate to the "Announcements Carousel" list in your site
2. Click "New" to create an announcement
3. Fill in the required fields:
   - **Title**: Your announcement headline
   - **Description**: Detailed information (supports rich text)
   - **AnnouncementImage**: URL to an image (can use SharePoint image library)
   - **ValidFrom**: Start date
   - **ValidTo**: End date
   - **CelebrationIcon**: (Optional) Choose a celebration type
   - **CelebrationIconPosition**: (Optional) Choose icon position

### Configuring the Web Part

1. Click the edit (pencil) icon on the web part
2. Click the property pane icon (gear) to open settings
3. Configure options across three tabs:
   - **Page 1**: Data Source, Carousel Behavior, Navigation
   - **Page 2**: Content Display, Colors & Styling, Dimensions & Effects
   - **Page 3**: Celebration Icons

## Architecture

### Services
- **ListProvisioningService**: Handles automatic SharePoint list creation and column provisioning
- **AnnouncementDataService**: Manages CRUD operations for announcements
- **CelebrationIconService**: Provides celebration icon configurations and animations

### Components
- **AnnouncementsCarousel**: Main React component with carousel logic
- **AnnouncementsCarousel.module.scss**: Comprehensive styling with animations

### Models
- **IAnnouncement**: TypeScript interface for announcement data
- **CelebrationType**: Enum for celebration icon types
- **IconPosition**: Enum for icon positioning

## Best Practices

### For Content Creators
1. Use high-quality images (recommended: 1200x600px)
2. Keep titles concise (under 60 characters)
3. Use rich text formatting in descriptions for better readability
4. Set appropriate date ranges to avoid expired announcements
5. Choose celebration icons that match the announcement context

### For Administrators
1. Regular cleanup of expired announcements
2. Monitor list size (consider archiving old items)
3. Use consistent naming conventions for images
4. Test the web part appearance on different devices
5. Configure rotation interval based on announcement count

### For Developers
1. The solution follows SPFx best practices
2. All services are stateless and reusable
3. React components use modern hooks and lifecycle methods
4. SCSS modules prevent style conflicts
5. TypeScript ensures type safety

## Customization

### Modifying Celebration Icons

Edit `src/webparts/announcementsCarousel/services/CelebrationIconService.ts`:

```typescript
[CelebrationType.Custom]: {
  emoji: '🎯', // Change emoji
  backgroundColor: '#FF5722', // Change background color
  color: '#FFFFFF', // Change text color
  label: 'Custom Label', // Change label
  animation: 'bounce' // Change animation
}
```

### Adding New Celebration Types

1. Add to `CelebrationType` enum in `models/IAnnouncement.ts`
2. Add configuration in `CelebrationIconService.ts`
3. Update SharePoint column choices in `ListProvisioningService.ts`

### Styling Customization

Edit `src/webparts/announcementsCarousel/components/AnnouncementsCarousel.module.scss` to customize:
- Colors and gradients
- Animation timings
- Responsive breakpoints
- Hover effects

## Troubleshooting

### List Not Created
- **Issue**: The list doesn't appear after adding the web part
- **Solution**: Check your permissions. You need at least "Manage Lists" permission.

### Announcements Not Showing
- **Issue**: No announcements appear even though items exist
- **Solution**: Verify that:
  - ValidFrom date is in the past
  - ValidTo date is in the future
  - All required fields are filled

### Images Not Loading
- **Issue**: Images don't display in the carousel
- **Solution**: Ensure:
  - Image URLs are accessible
  - URLs are absolute (not relative)
  - Images are in SharePoint or publicly accessible

### Performance Issues
- **Issue**: Slow loading or rotation
- **Solution**:
  - Reduce image sizes
  - Limit number of active announcements
  - Increase rotation interval
  - Disable transitions if needed

## Browser Support

- Microsoft Edge (Latest)
- Google Chrome (Latest)
- Mozilla Firefox (Latest)
- Safari (Latest)

## License

This project is licensed under the MIT License.

## Support

For issues, questions, or contributions, please contact your SharePoint administrator or development team.

## Version History

### 1.0.0
- Initial release
- Automatic list provisioning
- Carousel with rotation
- Celebration icons with animations
- Comprehensive configuration options
- Responsive design
- Date-based filtering

## Credits

Built with ❤️ using SharePoint Framework (SPFx), React, and TypeScript.
