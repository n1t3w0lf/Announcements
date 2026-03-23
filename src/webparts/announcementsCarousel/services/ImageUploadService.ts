import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/** Result returned after a successful image upload to Site Assets. */
export interface IImageUploadResult {
  absoluteUrl: string;
  serverRelativeUrl: string;
  fileName: string;
}

/** Maximum allowed file size for image uploads (10 MB). */
const MAX_IMAGE_SIZE_BYTES = 10 * 1024 * 1024;

/** Accepted image MIME types for upload validation. */
const ALLOWED_IMAGE_TYPES: string[] = [
  'image/jpeg',
  'image/png',
  'image/gif',
  'image/webp'
];

/**
 * Service to upload images to the SharePoint Site Assets library.
 * Validates file type and size before uploading.
 */
export class ImageUploadService {

  /** Cached server-relative URL of the Site Assets folder — constant within a session. */
  private static siteAssetsFolderUrl: string | null = null;

  /**
   * Uploads an image file to the Site Assets library.
   * Validates type and size before upload, sanitizes the file name, and returns
   * the absolute URL suitable for use as an AnnouncementImage field value.
   */
  public static async uploadImageToSiteAssets(
    context: WebPartContext,
    imageFile: File
  ): Promise<IImageUploadResult> {
    // Validate file size
    if (imageFile.size > MAX_IMAGE_SIZE_BYTES) {
      throw new Error(`Image must be under 10 MB. Selected file is ${(imageFile.size / (1024 * 1024)).toFixed(1)} MB.`);
    }

    // Validate file type
    if (!ALLOWED_IMAGE_TYPES.includes(imageFile.type)) {
      throw new Error(`Only JPEG, PNG, GIF, and WebP images are allowed. Selected file type: ${imageFile.type}`);
    }

    // Get Site Assets folder URL (cached after first call)
    const siteAssetsFolderRelativeUrl = await this.getSiteAssetsFolderUrl(context);

    // Sanitize filename to prevent special character issues
    const sanitizedFileName = this.sanitizeFileName(imageFile.name);

    // Read file into ArrayBuffer for binary upload
    const fileArrayBuffer = await imageFile.arrayBuffer();

    // Build the upload endpoint
    const uploadEndpoint =
      `${context.pageContext.web.absoluteUrl}/_api/web/` +
      `GetFolderByServerRelativeUrl('${encodeURIComponent(siteAssetsFolderRelativeUrl)}')` +
      `/Files/add(url='${encodeURIComponent(sanitizedFileName)}',overwrite=true)`;

    const uploadResponse: SPHttpClientResponse = await context.spHttpClient.post(
      uploadEndpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/octet-stream',
          'odata-version': ''
        },
        body: fileArrayBuffer as any // SPHttpClient types expect string but ArrayBuffer works at runtime
      }
    );

    if (!uploadResponse.ok) {
      const errorText = await uploadResponse.text();
      throw new Error(`Image upload failed: ${errorText}`);
    }

    const uploadData = await uploadResponse.json();
    const serverRelativeUrl: string = uploadData.ServerRelativeUrl;
    const siteOrigin = new URL(context.pageContext.web.absoluteUrl).origin;

    return {
      absoluteUrl: `${siteOrigin}${serverRelativeUrl}`,
      serverRelativeUrl: serverRelativeUrl,
      fileName: sanitizedFileName
    };
  }

  /**
   * Retrieves the server-relative URL of the Site Assets library root folder.
   * Caches the result for the duration of the session.
   * If the Site Assets library does not exist, attempts to create it.
   */
  private static async getSiteAssetsFolderUrl(context: WebPartContext): Promise<string> {
    if (this.siteAssetsFolderUrl) {
      return this.siteAssetsFolderUrl;
    }

    const siteAssetsEndpoint =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Site Assets')/RootFolder?$select=ServerRelativeUrl`;

    const siteAssetsResponse: SPHttpClientResponse = await context.spHttpClient.get(
      siteAssetsEndpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      }
    );

    if (siteAssetsResponse.status === 404) {
      // Site Assets library doesn't exist — create it
      await this.createSiteAssetsLibrary(context);
      return this.getSiteAssetsFolderUrl(context);
    }

    if (!siteAssetsResponse.ok) {
      throw new Error('Cannot access the Site Assets library. Please check your permissions.');
    }

    const siteAssetsData = await siteAssetsResponse.json();

    if (!siteAssetsData.ServerRelativeUrl) {
      throw new Error('Site Assets library returned no ServerRelativeUrl. The library may not be properly configured.');
    }

    this.siteAssetsFolderUrl = siteAssetsData.ServerRelativeUrl;
    return this.siteAssetsFolderUrl;
  }

  /**
   * Creates the Site Assets library if it does not exist.
   * Uses BaseTemplate 101 (Document Library) which is the standard for Site Assets.
   */
  private static async createSiteAssetsLibrary(context: WebPartContext): Promise<void> {
    const createLibraryEndpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists`;

    const libraryDefinition = {
      '@odata.type': 'SP.List',
      'BaseTemplate': 101,
      'Title': 'Site Assets',
      'Description': 'Use this library to store files used across the site.',
      'ContentTypesEnabled': false
    };

    const createResponse: SPHttpClientResponse = await context.spHttpClient.post(
      createLibraryEndpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=minimalmetadata',
          'Content-Type': 'application/json;odata=minimalmetadata'
        },
        body: JSON.stringify(libraryDefinition)
      }
    );

    if (!createResponse.ok) {
      const errorText = await createResponse.text();
      throw new Error(`Failed to create Site Assets library: ${errorText}`);
    }
  }

  /**
   * Sanitizes a filename for safe SharePoint storage.
   * Removes special characters, adds a timestamp prefix to prevent collisions.
   */
  private static sanitizeFileName(originalFileName: string): string {
    const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
    const baseNameWithoutExtension = originalFileName
      .substring(0, originalFileName.lastIndexOf('.'))
      .replace(/[^a-zA-Z0-9\-_]/g, '_')
      .substring(0, 50);
    const uploadTimestamp = Date.now();

    return `announcement_${uploadTimestamp}_${baseNameWithoutExtension}${fileExtension}`;
  }
}
