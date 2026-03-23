import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAnnouncement, CelebrationType, IconPosition, ImageFitMode, ImageSizeMode, ImageBackgroundType } from '../models/IAnnouncement';

/** All fields requested from SharePoint list items. */
const SELECT_FIELDS = [
  'Id', 'Title', 'Description', 'AnnouncementImage', 'ValidFrom', 'ValidTo',
  'CelebrationIcon', 'CelebrationIconPosition', 'CustomIconUrl',
  'ImageHeight', 'ImageWidth', 'ImageFit', 'ImageSizeMode',
  'ImageBackgroundType', 'ImageBackgroundColor',
  'BackgroundColor', 'TitleColor', 'DescriptionColor',
  'CardHeight', 'BorderRadius', 'ShowShadow', 'ShowCelebrationIcon', 'CelebrationIconSize',
  'GradientStartColor', 'GradientEndColor', 'GradientDirection',
  'OverlayGradientColor', 'OverlayOpacity',
  'RedirectUrl', 'RedirectTarget'
].join(',');

export class AnnouncementDataService {
  private context: WebPartContext;
  private listName: string;

  constructor(context: WebPartContext, listName: string) {
    this.context = context;
    this.listName = listName;
  }

  public async getActiveAnnouncements(): Promise<IAnnouncement[]> {
    try {
      const today = new Date().toISOString();

      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.listName}')/items?` +
        `$select=${SELECT_FIELDS}&` +
        `$filter=ValidFrom le datetime'${today}' and ValidTo ge datetime'${today}'&` +
        `$orderby=ValidFrom desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch announcements: ${response.statusText}`);
      }

      const data = await response.json();
      return this.mapListItemsToAnnouncements(data.value);
    } catch (error) {
      console.error('Error fetching announcements:', error);
      return [];
    }
  }

  public async getAllAnnouncements(): Promise<IAnnouncement[]> {
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.listName}')/items?` +
        `$select=${SELECT_FIELDS}&` +
        `$orderby=ValidFrom desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch announcements: ${response.statusText}`);
      }

      const data = await response.json();
      return this.mapListItemsToAnnouncements(data.value);
    } catch (error) {
      console.error('Error fetching announcements:', error);
      return [];
    }
  }

  public async getAnnouncementById(id: number): Promise<IAnnouncement | null> {
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.listName}')/items(${id})?` +
        `$select=${SELECT_FIELDS}`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        return null;
      }

      const data = await response.json();
      const mapped = this.mapListItemsToAnnouncements([data]);
      return mapped.length > 0 ? mapped[0] : null;
    } catch (error) {
      console.error('Error fetching announcement:', error);
      return null;
    }
  }

  private mapListItemsToAnnouncements(items: any[]): IAnnouncement[] {
    return items.map(item => ({
      Id: item.Id,
      Title: item.Title || '',
      Description: item.Description || '',
      AnnouncementImage: item.AnnouncementImage?.Url || item.AnnouncementImage || '',
      ValidFrom: new Date(item.ValidFrom),
      ValidTo: new Date(item.ValidTo),
      CelebrationIcon: (item.CelebrationIcon as CelebrationType) || CelebrationType.None,
      CelebrationIconPosition: (item.CelebrationIconPosition as IconPosition) || IconPosition.TopRight,
      CustomIconUrl: item.CustomIconUrl?.Url || item.CustomIconUrl || '',
      IsActive: this.isAnnouncementActive(new Date(item.ValidFrom), new Date(item.ValidTo)),
      ImageHeight: item.ImageHeight || 400,
      ImageWidth: item.ImageWidth != null ? item.ImageWidth : 0,
      ImageFit: (item.ImageFit as ImageFitMode) || ImageFitMode.Cover,
      ImageSizeMode: (item.ImageSizeMode as ImageSizeMode) || ImageSizeMode.Fit,
      ImageBackgroundType: (item.ImageBackgroundType as ImageBackgroundType) || ImageBackgroundType.Gradient,
      ImageBackgroundColor: item.ImageBackgroundColor || '#667eea',
      BackgroundColor: item.BackgroundColor || '#ffffff',
      TitleColor: item.TitleColor || '#333333',
      DescriptionColor: item.DescriptionColor || '#333333',
      CardHeight: item.CardHeight || 500,
      BorderRadius: item.BorderRadius != null ? item.BorderRadius : 8,
      ShowShadow: item.ShowShadow != null ? item.ShowShadow : true,
      ShowCelebrationIcon: item.ShowCelebrationIcon != null ? item.ShowCelebrationIcon : true,
      CelebrationIconSize: item.CelebrationIconSize || 60,
      GradientStartColor: item.GradientStartColor || '#667eea',
      GradientEndColor: item.GradientEndColor || '#764ba2',
      GradientDirection: item.GradientDirection != null ? item.GradientDirection : 135,
      OverlayGradientColor: item.OverlayGradientColor || '#000000',
      OverlayOpacity: item.OverlayOpacity != null ? item.OverlayOpacity : 30,
      RedirectUrl: item.RedirectUrl?.Url || item.RedirectUrl || '',
      RedirectTarget: item.RedirectTarget || '_self'
    }));
  }

  private isAnnouncementActive(validFrom: Date, validTo: Date): boolean {
    const now = new Date();
    return now >= validFrom && now <= validTo;
  }

  public async createAnnouncement(announcement: Partial<IAnnouncement>): Promise<number> {
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.listName}')/items`;

      const itemData: any = {
        'Title': announcement.Title,
        'Description': announcement.Description,
        'ValidFrom': announcement.ValidFrom?.toISOString(),
        'ValidTo': announcement.ValidTo?.toISOString()
      };

      if (announcement.AnnouncementImage) {
        itemData.AnnouncementImage = { 'Url': announcement.AnnouncementImage };
      }

      if (announcement.CelebrationIcon) itemData.CelebrationIcon = announcement.CelebrationIcon;
      if (announcement.CelebrationIconPosition) itemData.CelebrationIconPosition = announcement.CelebrationIconPosition;
      if (announcement.CustomIconUrl) itemData.CustomIconUrl = { 'Url': announcement.CustomIconUrl };

      if (announcement.ImageHeight) itemData.ImageHeight = announcement.ImageHeight;
      if (announcement.ImageWidth != null) itemData.ImageWidth = announcement.ImageWidth;
      if (announcement.ImageFit) itemData.ImageFit = announcement.ImageFit;
      if (announcement.ImageSizeMode) itemData.ImageSizeMode = announcement.ImageSizeMode;
      if (announcement.ImageBackgroundType) itemData.ImageBackgroundType = announcement.ImageBackgroundType;
      if (announcement.ImageBackgroundColor) itemData.ImageBackgroundColor = announcement.ImageBackgroundColor;

      if (announcement.BackgroundColor) itemData.BackgroundColor = announcement.BackgroundColor;
      if (announcement.TitleColor) itemData.TitleColor = announcement.TitleColor;
      if (announcement.DescriptionColor) itemData.DescriptionColor = announcement.DescriptionColor;
      if (announcement.CardHeight) itemData.CardHeight = announcement.CardHeight;
      if (announcement.BorderRadius != null) itemData.BorderRadius = announcement.BorderRadius;
      if (announcement.ShowShadow != null) itemData.ShowShadow = announcement.ShowShadow;
      if (announcement.ShowCelebrationIcon != null) itemData.ShowCelebrationIcon = announcement.ShowCelebrationIcon;
      if (announcement.CelebrationIconSize) itemData.CelebrationIconSize = announcement.CelebrationIconSize;
      if (announcement.GradientStartColor) itemData.GradientStartColor = announcement.GradientStartColor;
      if (announcement.GradientEndColor) itemData.GradientEndColor = announcement.GradientEndColor;
      if (announcement.GradientDirection != null) itemData.GradientDirection = announcement.GradientDirection;
      if (announcement.OverlayGradientColor) itemData.OverlayGradientColor = announcement.OverlayGradientColor;
      if (announcement.OverlayOpacity != null) itemData.OverlayOpacity = announcement.OverlayOpacity;
      if (announcement.RedirectUrl) itemData.RedirectUrl = { 'Url': announcement.RedirectUrl };
      if (announcement.RedirectTarget) itemData.RedirectTarget = announcement.RedirectTarget;

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify(itemData)
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to create announcement: ${response.statusText}`);
      }

      const data = await response.json();
      return data.Id;
    } catch (error) {
      console.error('Error creating announcement:', error);
      throw error;
    }
  }

  public async updateAnnouncement(id: number, announcement: Partial<IAnnouncement>): Promise<void> {
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.listName}')/items(${id})`;

      const getResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!getResponse.ok) {
        throw new Error(`Failed to retrieve announcement for update: ${getResponse.statusText}`);
      }

      const etag = getResponse.headers.get('ETag') || '*';

      const itemData: any = {};

      if (announcement.Title !== undefined) itemData.Title = announcement.Title;
      if (announcement.Description !== undefined) itemData.Description = announcement.Description;
      if (announcement.ValidFrom !== undefined) itemData.ValidFrom = announcement.ValidFrom.toISOString();
      if (announcement.ValidTo !== undefined) itemData.ValidTo = announcement.ValidTo.toISOString();

      if (announcement.AnnouncementImage !== undefined) {
        itemData.AnnouncementImage = { 'Url': announcement.AnnouncementImage };
      }

      if (announcement.CelebrationIcon !== undefined) itemData.CelebrationIcon = announcement.CelebrationIcon;
      if (announcement.CelebrationIconPosition !== undefined) itemData.CelebrationIconPosition = announcement.CelebrationIconPosition;
      if (announcement.CustomIconUrl !== undefined) itemData.CustomIconUrl = { 'Url': announcement.CustomIconUrl };

      if (announcement.ImageHeight !== undefined) itemData.ImageHeight = announcement.ImageHeight;
      if (announcement.ImageWidth !== undefined) itemData.ImageWidth = announcement.ImageWidth;
      if (announcement.ImageFit !== undefined) itemData.ImageFit = announcement.ImageFit;
      if (announcement.ImageSizeMode !== undefined) itemData.ImageSizeMode = announcement.ImageSizeMode;
      if (announcement.ImageBackgroundType !== undefined) itemData.ImageBackgroundType = announcement.ImageBackgroundType;
      if (announcement.ImageBackgroundColor !== undefined) itemData.ImageBackgroundColor = announcement.ImageBackgroundColor;

      if (announcement.BackgroundColor !== undefined) itemData.BackgroundColor = announcement.BackgroundColor;
      if (announcement.TitleColor !== undefined) itemData.TitleColor = announcement.TitleColor;
      if (announcement.DescriptionColor !== undefined) itemData.DescriptionColor = announcement.DescriptionColor;
      if (announcement.CardHeight !== undefined) itemData.CardHeight = announcement.CardHeight;
      if (announcement.BorderRadius !== undefined) itemData.BorderRadius = announcement.BorderRadius;
      if (announcement.ShowShadow !== undefined) itemData.ShowShadow = announcement.ShowShadow;
      if (announcement.ShowCelebrationIcon !== undefined) itemData.ShowCelebrationIcon = announcement.ShowCelebrationIcon;
      if (announcement.CelebrationIconSize !== undefined) itemData.CelebrationIconSize = announcement.CelebrationIconSize;
      if (announcement.GradientStartColor !== undefined) itemData.GradientStartColor = announcement.GradientStartColor;
      if (announcement.GradientEndColor !== undefined) itemData.GradientEndColor = announcement.GradientEndColor;
      if (announcement.GradientDirection !== undefined) itemData.GradientDirection = announcement.GradientDirection;
      if (announcement.OverlayGradientColor !== undefined) itemData.OverlayGradientColor = announcement.OverlayGradientColor;
      if (announcement.OverlayOpacity !== undefined) itemData.OverlayOpacity = announcement.OverlayOpacity;
      if (announcement.RedirectUrl !== undefined) itemData.RedirectUrl = { 'Url': announcement.RedirectUrl };
      if (announcement.RedirectTarget !== undefined) itemData.RedirectTarget = announcement.RedirectTarget;

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': etag,
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify(itemData)
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to update announcement: ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error updating announcement:', error);
      throw error;
    }
  }

  public async deleteAnnouncement(id: number): Promise<void> {
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.listName}')/items(${id})`;

      const getResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      const etag = getResponse.headers.get('ETag') || '*';

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'IF-MATCH': etag,
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to delete announcement: ${response.statusText}`);
      }
    } catch (error) {
      console.error('Error deleting announcement:', error);
      throw error;
    }
  }
}
