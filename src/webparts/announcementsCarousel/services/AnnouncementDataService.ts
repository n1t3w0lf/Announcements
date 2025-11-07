import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAnnouncement, CelebrationType, IconPosition } from '../models/IAnnouncement';

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

      // Get announcements where ValidFrom <= Today <= ValidTo
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.listName}')/items?` +
        `$select=Id,Title,Description,AnnouncementImage,ValidFrom,ValidTo,CelebrationIcon,CelebrationIconPosition,CustomIconUrl&` +
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
        `$select=Id,Title,Description,AnnouncementImage,ValidFrom,ValidTo,CelebrationIcon,CelebrationIconPosition,CustomIconUrl&` +
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
        `$select=Id,Title,Description,AnnouncementImage,ValidFrom,ValidTo,CelebrationIcon,CelebrationIconPosition,CustomIconUrl`;

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
      IsActive: this.isAnnouncementActive(new Date(item.ValidFrom), new Date(item.ValidTo))
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
        '__metadata': { 'type': `SP.Data.${this.listName.replace(/\s/g, '_x0020_')}ListItem` },
        'Title': announcement.Title,
        'Description': announcement.Description,
        'ValidFrom': announcement.ValidFrom?.toISOString(),
        'ValidTo': announcement.ValidTo?.toISOString()
      };

      if (announcement.AnnouncementImage) {
        itemData.AnnouncementImage = {
          '__metadata': { 'type': 'SP.FieldUrlValue' },
          'Url': announcement.AnnouncementImage
        };
      }

      if (announcement.CelebrationIcon) {
        itemData.CelebrationIcon = announcement.CelebrationIcon;
      }

      if (announcement.CelebrationIconPosition) {
        itemData.CelebrationIconPosition = announcement.CelebrationIconPosition;
      }

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

      // Get item etag first
      const getResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      const etag = getResponse.headers.get('ETag') || '*';

      const itemData: any = {
        '__metadata': { 'type': `SP.Data.${this.listName.replace(/\s/g, '_x0020_')}ListItem` }
      };

      if (announcement.Title !== undefined) itemData.Title = announcement.Title;
      if (announcement.Description !== undefined) itemData.Description = announcement.Description;
      if (announcement.ValidFrom !== undefined) itemData.ValidFrom = announcement.ValidFrom.toISOString();
      if (announcement.ValidTo !== undefined) itemData.ValidTo = announcement.ValidTo.toISOString();

      if (announcement.AnnouncementImage !== undefined) {
        itemData.AnnouncementImage = {
          '__metadata': { 'type': 'SP.FieldUrlValue' },
          'Url': announcement.AnnouncementImage
        };
      }

      if (announcement.CelebrationIcon !== undefined) itemData.CelebrationIcon = announcement.CelebrationIcon;
      if (announcement.CelebrationIconPosition !== undefined) itemData.CelebrationIconPosition = announcement.CelebrationIconPosition;

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

      // Get item etag first
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
