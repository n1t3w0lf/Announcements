import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Service to check whether the current user belongs to the site Owners group.
 * Used to gate the management UI — this is a UI-only check; SharePoint list
 * permissions enforce actual write access server-side.
 */
export class PermissionService {

  /**
   * Checks if the current user is a member of the site's Associated Owner Group.
   * Uses the numeric user ID (reliable) rather than LoginName string filtering
   * (which breaks with claims-encoded names containing special characters).
   * Falls back to checking IsSiteAdmin if the group endpoint returns 403.
   * Never throws — returns false on any error to avoid blocking the carousel.
   */
  public static async checkIsOwner(context: WebPartContext): Promise<boolean> {
    try {
      // Step 1: Get the current user's numeric ID
      const currentUserId = await this.getCurrentUserId(context);
      if (currentUserId === null) {
        console.warn('PermissionService: Could not determine current user ID');
        return false;
      }

      // Step 2: Try to get this user from the Owners group by their numeric ID.
      // If the user is in the group, this returns 200. If not, it returns 404.
      const ownerCheckEndpoint =
        `${context.pageContext.web.absoluteUrl}/_api/web/AssociatedOwnerGroup/users/getById(${currentUserId})`;

      const ownerCheckResponse: SPHttpClientResponse = await context.spHttpClient.get(
        ownerCheckEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      // 200 = user is in owners group
      if (ownerCheckResponse.ok) {
        return true;
      }

      // 403/401 = endpoint restricted, fall back to site admin check
      if (ownerCheckResponse.status === 403 || ownerCheckResponse.status === 401) {
        return await this.checkIsSiteAdmin(context);
      }

      // 404 = user is NOT in owners group, but they might still be a site admin
      if (ownerCheckResponse.status === 404) {
        return await this.checkIsSiteAdmin(context);
      }

      return false;
    } catch (permissionCheckError) {
      console.error('PermissionService: Error checking owner membership:', permissionCheckError);
      return false;
    }
  }

  /**
   * Gets the current user's numeric SharePoint user ID.
   * Tries pageContext first (fastest), falls back to REST API.
   */
  private static async getCurrentUserId(context: WebPartContext): Promise<number | null> {
    try {
      // Try to get from legacy page context (available on real SP pages and workbench)
      const legacyContext = context.pageContext.legacyPageContext;
      if (legacyContext && legacyContext.userId) {
        return legacyContext.userId;
      }

      // Fallback: REST API call
      const currentUserEndpoint =
        `${context.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=Id`;

      const currentUserResponse: SPHttpClientResponse = await context.spHttpClient.get(
        currentUserEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!currentUserResponse.ok) {
        return null;
      }

      const currentUserData = await currentUserResponse.json();
      return currentUserData.Id || null;
    } catch (userIdError) {
      console.error('PermissionService: Error getting current user ID:', userIdError);
      return null;
    }
  }

  /**
   * Fallback check: site administrators always have owner-level permissions.
   * Used when the AssociatedOwnerGroup endpoint is restricted or user is not in the group.
   */
  private static async checkIsSiteAdmin(context: WebPartContext): Promise<boolean> {
    try {
      const siteAdminEndpoint =
        `${context.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=IsSiteAdmin`;

      const siteAdminResponse: SPHttpClientResponse = await context.spHttpClient.get(
        siteAdminEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!siteAdminResponse.ok) {
        return false;
      }

      const siteAdminData = await siteAdminResponse.json();
      return siteAdminData.IsSiteAdmin === true;
    } catch (siteAdminCheckError) {
      console.error('PermissionService: Error checking site admin status:', siteAdminCheckError);
      return false;
    }
  }
}
