import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class ListProvisioningService {
  private static readonly LIST_NAME = 'Announcements Carousel';
  private static readonly LIST_DESCRIPTION = 'Stores announcements and celebrations for the carousel web part';

  public static async ensureList(context: WebPartContext): Promise<boolean> {
    try {
      // Check if list exists
      const listExists = await this.checkListExists(context);

      if (!listExists) {
        console.log('List does not exist. Creating...');
        await this.createList(context);
        console.log('List created. Creating columns...');
        await this.createColumns(context);
        console.log('Columns created. Configuring list settings...');
        await this.configureListSettings(context);
        console.log('List created successfully. Waiting for SharePoint to fully provision...');

        // Wait for SharePoint to fully provision the list
        await this.delay(3000);

        // Verify the list is accessible
        const listNowExists = await this.checkListExists(context);
        if (!listNowExists) {
          console.warn('List was created but is not yet accessible. Waiting longer...');
          await this.delay(2000);
        }

        console.log('List is ready');
        return true;
      } else {
        console.log('List already exists');
        // Verify all columns exist
        await this.verifyColumns(context);
        return true;
      }
    } catch (error) {
      console.error('Error provisioning list:', error);
      throw error;
    }
  }

  private static delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  private static async checkListExists(context: WebPartContext): Promise<boolean> {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')`;

    try {
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  private static async createList(context: WebPartContext): Promise<void> {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists`;

    const listData = {
      '@odata.type': 'SP.List',
      'BaseTemplate': 100,
      'Title': this.LIST_NAME,
      'Description': this.LIST_DESCRIPTION,
      'ContentTypesEnabled': false,
      'EnableVersioning': true
    };

    const response: SPHttpClientResponse = await context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=minimalmetadata',
          'Content-Type': 'application/json;odata=minimalmetadata'
        },
        body: JSON.stringify(listData)
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to create list: ${errorText}`);
    }
  }

  private static async createColumns(context: WebPartContext): Promise<void> {
    const baseUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/fields`;

    // Define columns with proper metadata
    const columns = [
      {
        '@odata.type': 'SP.FieldMultiLineText',
        'FieldTypeKind': 3,
        'Title': 'Description',
        'Required': false,
        'RichText': true,
        'NumberOfLines': 6
      },
      {
        '@odata.type': 'SP.FieldUrl',
        'FieldTypeKind': 11,
        'Title': 'AnnouncementImage',
        'Required': false
      },
      {
        '@odata.type': 'SP.FieldDateTime',
        'FieldTypeKind': 4,
        'Title': 'ValidFrom',
        'Required': true,
        'DisplayFormat': 1
      },
      {
        '@odata.type': 'SP.FieldDateTime',
        'FieldTypeKind': 4,
        'Title': 'ValidTo',
        'Required': true,
        'DisplayFormat': 1
      },
      {
        '@odata.type': 'SP.FieldChoice',
        'FieldTypeKind': 6,
        'Title': 'CelebrationIcon',
        'Required': false,
        'Choices': ['None', 'Birthday', 'Anniversary', 'Achievement', 'Celebration', 'NewHire', 'Promotion', 'Holiday', 'Custom'],
        'DefaultValue': 'None'
      },
      {
        '@odata.type': 'SP.FieldChoice',
        'FieldTypeKind': 6,
        'Title': 'CelebrationIconPosition',
        'Required': false,
        'Choices': ['topLeft', 'topRight', 'bottomLeft', 'bottomRight', 'center'],
        'DefaultValue': 'topRight'
      },
      {
        '@odata.type': 'SP.FieldUrl',
        'FieldTypeKind': 11,
        'Title': 'CustomIconUrl',
        'Required': false
      },
      {
        '@odata.type': 'SP.FieldNumber',
        'FieldTypeKind': 9,
        'Title': 'ImageHeight',
        'Required': false,
        'MinimumValue': 100,
        'MaximumValue': 1200,
        'Description': 'Image container height in pixels. Leave empty to use the web part default.'
      },
      {
        '@odata.type': 'SP.FieldNumber',
        'FieldTypeKind': 9,
        'Title': 'ImageWidth',
        'Required': false,
        'MinimumValue': 100,
        'MaximumValue': 2000,
        'Description': 'Image width in pixels. Leave empty to use the web part default.'
      },
      {
        '@odata.type': 'SP.FieldChoice',
        'FieldTypeKind': 6,
        'Title': 'ImageFit',
        'Required': false,
        'Choices': ['cover', 'contain', 'fill', 'none'],
        'DefaultValue': 'cover',
        'Description': 'How the image fills its container.'
      },
      {
        '@odata.type': 'SP.FieldChoice',
        'FieldTypeKind': 6,
        'Title': 'ImageSizeMode',
        'Required': false,
        'Choices': ['fit', 'manual'],
        'DefaultValue': 'fit',
        'Description': 'Whether image uses CSS fit mode or manual pixel dimensions.'
      },
      {
        '@odata.type': 'SP.FieldChoice',
        'FieldTypeKind': 6,
        'Title': 'ImageBackgroundType',
        'Required': false,
        'Choices': ['solid', 'gradient'],
        'DefaultValue': 'gradient',
        'Description': 'Whether image container background is a solid color or gradient.'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 2,
        'Title': 'ImageBackgroundColor',
        'Required': false,
        'DefaultValue': '#667eea',
        'Description': 'Solid background color for image container (hex).'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 2,
        'Title': 'BackgroundColor',
        'Required': false,
        'DefaultValue': '#ffffff',
        'Description': 'Card background color (hex).'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 2,
        'Title': 'TitleColor',
        'Required': false,
        'DefaultValue': '#333333',
        'Description': 'Title text color (hex).'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 2,
        'Title': 'DescriptionColor',
        'Required': false,
        'DefaultValue': '#333333',
        'Description': 'Description text color (hex).'
      },
      {
        '@odata.type': 'SP.FieldNumber',
        'FieldTypeKind': 9,
        'Title': 'CardHeight',
        'Required': false,
        'MinimumValue': 200,
        'MaximumValue': 1000,
        'Description': 'Card height in pixels.'
      },
      {
        '@odata.type': 'SP.FieldNumber',
        'FieldTypeKind': 9,
        'Title': 'BorderRadius',
        'Required': false,
        'MinimumValue': 0,
        'MaximumValue': 50,
        'Description': 'Border radius in pixels.'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 8,
        'Title': 'ShowShadow',
        'Required': false,
        'DefaultValue': '1',
        'Description': 'Show drop shadow on card.'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 8,
        'Title': 'ShowCelebrationIcon',
        'Required': false,
        'DefaultValue': '1',
        'Description': 'Show celebration icon.'
      },
      {
        '@odata.type': 'SP.FieldNumber',
        'FieldTypeKind': 9,
        'Title': 'CelebrationIconSize',
        'Required': false,
        'MinimumValue': 20,
        'MaximumValue': 200,
        'Description': 'Celebration icon size in pixels.'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 2,
        'Title': 'GradientStartColor',
        'Required': false,
        'DefaultValue': '#667eea',
        'Description': 'Image background gradient start color (hex).'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 2,
        'Title': 'GradientEndColor',
        'Required': false,
        'DefaultValue': '#764ba2',
        'Description': 'Image background gradient end color (hex).'
      },
      {
        '@odata.type': 'SP.FieldNumber',
        'FieldTypeKind': 9,
        'Title': 'GradientDirection',
        'Required': false,
        'MinimumValue': 0,
        'MaximumValue': 360,
        'Description': 'Gradient direction in degrees.'
      },
      {
        '@odata.type': 'SP.Field',
        'FieldTypeKind': 2,
        'Title': 'OverlayGradientColor',
        'Required': false,
        'DefaultValue': '#000000',
        'Description': 'Overlay gradient color (hex).'
      },
      {
        '@odata.type': 'SP.FieldNumber',
        'FieldTypeKind': 9,
        'Title': 'OverlayOpacity',
        'Required': false,
        'MinimumValue': 0,
        'MaximumValue': 100,
        'Description': 'Overlay opacity percentage.'
      },
      {
        '@odata.type': 'SP.FieldUrl',
        'FieldTypeKind': 11,
        'Title': 'RedirectUrl',
        'Required': false,
        'Description': 'Optional URL to navigate to when the announcement is clicked.'
      },
      {
        '@odata.type': 'SP.FieldChoice',
        'FieldTypeKind': 6,
        'Title': 'RedirectTarget',
        'Required': false,
        'Choices': ['_self', '_blank'],
        'DefaultValue': '_self',
        'Description': 'Open redirect URL in same tab or new tab.'
      }
    ];

    for (const column of columns) {
      try {
        const response: SPHttpClientResponse = await context.spHttpClient.post(
          baseUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata',
              'Content-Type': 'application/json;odata=minimalmetadata'
            },
            body: JSON.stringify(column)
          }
        );

        if (!response.ok) {
          const errorText = await response.text();
          console.warn(`Warning creating column ${column.Title}: ${errorText}`);
        } else {
          console.log(`Column ${column.Title} created successfully`);
        }
      } catch (error) {
        console.warn(`Error creating column ${column.Title}:`, error);
      }
    }
  }

  private static async verifyColumns(context: WebPartContext): Promise<void> {
    // Check if all required columns exist, if not create them
    const requiredColumns = ['Description', 'AnnouncementImage', 'ValidFrom', 'ValidTo', 'CelebrationIcon', 'CelebrationIconPosition', 'CustomIconUrl', 'ImageHeight', 'ImageWidth', 'ImageFit', 'ImageSizeMode', 'ImageBackgroundType', 'ImageBackgroundColor', 'BackgroundColor', 'TitleColor', 'DescriptionColor', 'CardHeight', 'BorderRadius', 'ShowShadow', 'ShowCelebrationIcon', 'CelebrationIconSize', 'GradientStartColor', 'GradientEndColor', 'GradientDirection', 'OverlayGradientColor', 'OverlayOpacity', 'RedirectUrl', 'RedirectTarget'];
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/fields?$select=Title`;

    try {
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const existingColumns = data.value.map((field: any) => field.Title);
        const missingColumns = requiredColumns.filter(col => !existingColumns.includes(col));

        if (missingColumns.length > 0) {
          console.log('Missing columns detected, creating them...');
          await this.createColumns(context);
        }
      }
    } catch (error) {
      console.warn('Error verifying columns:', error);
    }
  }

  private static async configureListSettings(context: WebPartContext): Promise<void> {
    // Configure list view to show relevant columns
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/DefaultView/ViewFields`;

    const viewFields = ['Title', 'Description', 'AnnouncementImage', 'ValidFrom', 'ValidTo', 'CelebrationIcon'];

    for (const field of viewFields) {
      try {
        await context.spHttpClient.post(
          `${endpoint}/addViewField('${field}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=minimalmetadata',
              'Content-Type': 'application/json;odata=minimalmetadata'
            }
          }
        );
      } catch (error) {
        console.warn(`Error adding field ${field} to view:`, error);
      }
    }
  }

  public static getListName(): string {
    return this.LIST_NAME;
  }
}
