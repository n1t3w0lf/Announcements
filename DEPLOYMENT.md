# Deployment Guide

This guide provides detailed instructions for deploying the Announcements Carousel Web Part to SharePoint Online.

## Prerequisites

Before deploying, ensure you have:

1. **SharePoint Online** tenant with administrator access
2. **App Catalog** set up in your tenant
3. **Node.js v16.x** installed on your development machine
4. **Permissions**:
   - Tenant administrator or access to App Catalog
   - Site collection administrator rights on target sites

## Step-by-Step Deployment

### 1. Prepare the Solution

```bash
# Navigate to the project directory
cd Announcements

# Install dependencies (if not already done)
npm install

# Clean any previous builds
npm run clean
```

### 2. Build for Production

```bash
# Create production bundle
gulp bundle --ship

# Package the solution
gulp package-solution --ship
```

This will create a `.sppkg` file in the `sharepoint/solution/` directory.

### 3. Upload to App Catalog

1. Navigate to your **App Catalog** site:
   - Tenant App Catalog: `https://<tenant>-admin.sharepoint.com/_layouts/15/online/TenantAdminApps.aspx`
   - Site Collection App Catalog: `https://<site>/_layouts/15/appstore.aspx`

2. **Upload the package**:
   - Go to "Apps for SharePoint"
   - Click "Upload"
   - Select the `.sppkg` file from `sharepoint/solution/`

3. **Deploy the app**:
   - A dialog will appear asking to "Make this solution available to all sites in the organization"
   - Check this box if you want tenant-wide deployment (recommended)
   - Click "Deploy"

### 4. Trust the Solution

When prompted:
- Review the API permissions (none required for this solution)
- Click "Trust It" or "Enable app"

### 5. Add to SharePoint Site

#### Option A: If Tenant-Wide Deployment Enabled

1. Go to any SharePoint site
2. Edit a page
3. Click "+" to add a web part
4. Search for "Announcements Carousel"
5. Add to the page
6. The list will be created automatically

#### Option B: If Site Collection Deployment

1. Go to **Site Contents**
2. Click "New" → "App"
3. Find "Announcements Carousel" and click "Add"
4. Wait for installation to complete
5. Go to a page and add the web part

## Post-Deployment Configuration

### Initial Setup

After adding the web part for the first time:

1. **Verify List Creation**:
   - Go to Site Contents
   - Look for "Announcements Carousel" list
   - If not present, check browser console for errors

2. **Configure Permissions**:
   - Ensure users have at least "Read" access to the list
   - Content creators need "Contribute" access

3. **Add Sample Content**:
   ```
   Title: Welcome to Our New Portal
   Description: We're excited to announce...
   AnnouncementImage: <URL to image>
   ValidFrom: <Today's date>
   ValidTo: <30 days from now>
   CelebrationIcon: Celebration
   CelebrationIconPosition: top-right
   ```

### Web Part Configuration

Configure the web part properties:

1. **Click Edit** on the web part
2. **Open Property Pane** (gear icon)
3. **Recommended Settings**:
   - Rotation Interval: 10 seconds
   - Auto Play: On
   - Show Navigation Arrows: On
   - Show Navigation Dots: On
   - Transition Effect: Fade
   - Height: 500px
   - Border Radius: 8px
   - Show Shadow: On

## Updating the Solution

When you need to update the web part:

### 1. Update Version Number

Edit `config/package-solution.json`:

```json
{
  "solution": {
    "version": "1.0.1.0"  // Increment version
  }
}
```

### 2. Rebuild and Repackage

```bash
gulp clean
gulp bundle --ship
gulp package-solution --ship
```

### 3. Upload New Version

1. Go to App Catalog
2. Upload the new `.sppkg` file
3. When prompted, choose "Replace It"
4. The update will propagate to all sites using the web part

### 4. Refresh Pages

Users may need to refresh their browser to see updates.

## Troubleshooting Deployment

### Issue: Package Upload Fails

**Possible Causes**:
- File is corrupted
- App Catalog permissions issue
- Version conflict

**Solutions**:
1. Rebuild the package from scratch
2. Verify App Catalog administrator access
3. Delete old version before uploading new one

### Issue: Web Part Not Appearing

**Possible Causes**:
- App not deployed
- App not added to site
- Browser cache

**Solutions**:
1. Verify app is deployed in App Catalog
2. Add app to site from Site Contents
3. Clear browser cache (Ctrl+F5)
4. Try in incognito/private browsing mode

### Issue: List Not Created

**Possible Causes**:
- Insufficient permissions
- JavaScript error
- SharePoint API throttling

**Solutions**:
1. Check browser console for errors
2. Verify "Manage Lists" permission
3. Try refreshing the page
4. Check SharePoint health status

### Issue: Images Not Loading

**Possible Causes**:
- CORS restrictions
- Invalid image URLs
- Permission issues

**Solutions**:
1. Use images from SharePoint document library
2. Verify image URLs are absolute
3. Check image library permissions
4. Test image URL in browser directly

## Environment-Specific Configurations

### Development Environment

```bash
# Local workbench
gulp serve

# SharePoint workbench
gulp serve --nobrowser
# Then navigate to: https://<tenant>.sharepoint.com/_layouts/workbench.aspx
```

### Staging Environment

1. Create a separate App Catalog for staging
2. Deploy to staging first for testing
3. Use different version numbers (e.g., 1.0.0-rc1)

### Production Environment

1. Always test in staging first
2. Schedule deployments during low-traffic periods
3. Communicate changes to users
4. Have rollback plan ready

## Security Considerations

### Permissions Required

The web part requires:
- **Read** access to the announcements list (all users)
- **Contribute** access for content creators
- **Manage Lists** permission for initial list creation

### Data Security

- All data stays within SharePoint
- No external API calls
- No sensitive data in web part properties
- Respects SharePoint permissions

### Content Security Policy

The solution:
- Uses only inline styles (no external stylesheets)
- No external scripts loaded
- Complies with SharePoint CSP

## Performance Optimization

### Image Optimization

1. **Recommended Image Specs**:
   - Format: JPG or PNG
   - Size: 1200x600px
   - File size: < 500KB

2. **Use SharePoint Image Renditions**:
   - Store images in SharePoint
   - Use image renditions for automatic sizing
   - Reduces bandwidth usage

### List Performance

1. **Index Columns**:
   - Add index on ValidFrom and ValidTo
   - Improves query performance

2. **Limit Active Items**:
   - Archive old announcements
   - Keep active items under 100

3. **Caching**:
   - The web part caches data client-side
   - Reduce server requests

## Monitoring and Maintenance

### Regular Tasks

**Weekly**:
- Review and remove expired announcements
- Check for broken image links
- Verify web part functioning on key pages

**Monthly**:
- Review list size and archive old items
- Check error logs in SharePoint
- Update any outdated content

**Quarterly**:
- Review and update web part configuration
- Test on different devices/browsers
- Check for SPFx updates

### Health Checks

1. **Verify List Exists**:
   ```
   Site Contents → Announcements Carousel
   ```

2. **Check Active Announcements**:
   ```
   Filter list by: ValidFrom <= Today AND ValidTo >= Today
   ```

3. **Test Web Part**:
   - Add to test page
   - Verify carousel rotation
   - Test navigation controls
   - Check celebration icons

## Support and Resources

### Getting Help

1. **Check Browser Console**: F12 → Console tab for errors
2. **SharePoint Logs**: Site Settings → Site Collection Administration → Health Reports
3. **Network Tab**: F12 → Network tab to see API calls

### Useful Commands

```bash
# Check SPFx version
npm list @microsoft/sp-core-library

# Update dependencies
npm update

# Rebuild from scratch
npm run clean && npm install && gulp bundle --ship && gulp package-solution --ship

# Check for outdated packages
npm outdated
```

### Documentation Links

- [SPFx Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [SharePoint REST API](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui#/controls/web)

## Rollback Procedure

If you need to rollback to a previous version:

1. **Go to App Catalog**
2. **Select the app**
3. **Click "Files" tab**
4. **Find previous version** in version history
5. **Restore previous version**
6. **Redeploy**

Alternatively, keep backup copies of `.sppkg` files for each version.

## Conclusion

This deployment guide should help you successfully deploy and maintain the Announcements Carousel Web Part. For specific issues not covered here, consult your SharePoint administrator or refer to the main README.md file.
