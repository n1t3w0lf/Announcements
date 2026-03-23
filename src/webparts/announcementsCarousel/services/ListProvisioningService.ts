import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class ListProvisioningService {
  private static readonly LIST_NAME = 'Announcements Carousel';
  private static readonly LIST_DESCRIPTION = 'Stores announcements and celebrations for the carousel web part';

  /** Skip column verification after first successful check within a browser session. */
  private static columnsVerified = false;

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

        this.columnsVerified = true;
        console.log('List is ready');
        return true;
      } else {
        console.log('List already exists');
        // Only verify columns once per browser session
        if (!this.columnsVerified) {
          await this.verifyColumns(context);
          this.columnsVerified = true;
        }
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

  /**
   * Creates the list. Uses odata=nometadata — no type annotations needed for list creation.
   */
  private static async createList(context: WebPartContext): Promise<void> {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists`;

    const listData = {
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
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(listData)
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to create list: ${errorText}`);
    }
  }

  /**
   * Returns all column definitions in odata=verbose format.
   * - Uses __metadata.type instead of @odata.type
   * - Wraps Choices arrays as { results: [...] } (required by verbose format)
   */
  private static getColumnDefinitions(): any[] {
    return [
      {
        '__metadata': { 'type': 'SP.FieldMultiLineText' },
        'FieldTypeKind': 3,
        'Title': 'Description',
        'Required': false,
        'RichText': true,
        'NumberOfLines': 6
      },
      {
        '__metadata': { 'type': 'SP.FieldUrl' },
        'FieldTypeKind': 11,
        'Title': 'AnnouncementImage',
        'Required': false
      },
      {
        '__metadata': { 'type': 'SP.FieldDateTime' },
        'FieldTypeKind': 4,
        'Title': 'ValidFrom',
        'Required': true,
        'DisplayFormat': 1
      },
      {
        '__metadata': { 'type': 'SP.FieldDateTime' },
        'FieldTypeKind': 4,
        'Title': 'ValidTo',
        'Required': true,
        'DisplayFormat': 1
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'FieldTypeKind': 6,
        'Title': 'CelebrationIcon',
        'Required': false,
        'Choices': { 'results': ['None', 'Birthday', 'Anniversary', 'Achievement', 'Celebration', 'NewHire', 'Promotion', 'Holiday', 'Custom'] },
        'DefaultValue': 'None'
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'FieldTypeKind': 6,
        'Title': 'CelebrationIconPosition',
        'Required': false,
        'Choices': { 'results': ['topLeft', 'topRight', 'bottomLeft', 'bottomRight', 'center'] },
        'DefaultValue': 'topRight'
      },
      {
        '__metadata': { 'type': 'SP.FieldUrl' },
        'FieldTypeKind': 11,
        'Title': 'CustomIconUrl',
        'Required': false
      },
      {
        '__metadata': { 'type': 'SP.FieldNumber' },
        'FieldTypeKind': 9,
        'Title': 'ImageHeight',
        'Required': false,
        'MinimumValue': 100,
        'MaximumValue': 1200,
        'Description': 'Image container height in pixels. Leave empty to use the web part default.'
      },
      {
        '__metadata': { 'type': 'SP.FieldNumber' },
        'FieldTypeKind': 9,
        'Title': 'ImageWidth',
        'Required': false,
        'MinimumValue': 100,
        'MaximumValue': 2000,
        'Description': 'Image width in pixels. Leave empty to use the web part default.'
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'FieldTypeKind': 6,
        'Title': 'ImageFit',
        'Required': false,
        'Choices': { 'results': ['cover', 'contain', 'fill', 'none'] },
        'DefaultValue': 'cover',
        'Description': 'How the image fills its container.'
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'FieldTypeKind': 6,
        'Title': 'ImageSizeMode',
        'Required': false,
        'Choices': { 'results': ['fit', 'manual'] },
        'DefaultValue': 'fit',
        'Description': 'Whether image uses CSS fit mode or manual pixel dimensions.'
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'FieldTypeKind': 6,
        'Title': 'ImageBackgroundType',
        'Required': false,
        'Choices': { 'results': ['solid', 'gradient'] },
        'DefaultValue': 'gradient',
        'Description': 'Whether image container background is a solid color or gradient.'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 2,
        'Title': 'ImageBackgroundColor',
        'Required': false,
        'DefaultValue': '#667eea',
        'Description': 'Solid background color for image container (hex).'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 2,
        'Title': 'BackgroundColor',
        'Required': false,
        'DefaultValue': '#ffffff',
        'Description': 'Card background color (hex).'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 2,
        'Title': 'TitleColor',
        'Required': false,
        'DefaultValue': '#333333',
        'Description': 'Title text color (hex).'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 2,
        'Title': 'DescriptionColor',
        'Required': false,
        'DefaultValue': '#333333',
        'Description': 'Description text color (hex).'
      },
      {
        '__metadata': { 'type': 'SP.FieldNumber' },
        'FieldTypeKind': 9,
        'Title': 'CardHeight',
        'Required': false,
        'MinimumValue': 200,
        'MaximumValue': 1000,
        'Description': 'Card height in pixels.'
      },
      {
        '__metadata': { 'type': 'SP.FieldNumber' },
        'FieldTypeKind': 9,
        'Title': 'BorderRadius',
        'Required': false,
        'MinimumValue': 0,
        'MaximumValue': 50,
        'Description': 'Border radius in pixels.'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 8,
        'Title': 'ShowShadow',
        'Required': false,
        'DefaultValue': '1',
        'Description': 'Show drop shadow on card.'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 8,
        'Title': 'ShowCelebrationIcon',
        'Required': false,
        'DefaultValue': '1',
        'Description': 'Show celebration icon.'
      },
      {
        '__metadata': { 'type': 'SP.FieldNumber' },
        'FieldTypeKind': 9,
        'Title': 'CelebrationIconSize',
        'Required': false,
        'MinimumValue': 20,
        'MaximumValue': 200,
        'Description': 'Celebration icon size in pixels.'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 2,
        'Title': 'GradientStartColor',
        'Required': false,
        'DefaultValue': '#667eea',
        'Description': 'Image background gradient start color (hex).'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 2,
        'Title': 'GradientEndColor',
        'Required': false,
        'DefaultValue': '#764ba2',
        'Description': 'Image background gradient end color (hex).'
      },
      {
        '__metadata': { 'type': 'SP.FieldNumber' },
        'FieldTypeKind': 9,
        'Title': 'GradientDirection',
        'Required': false,
        'MinimumValue': 0,
        'MaximumValue': 360,
        'Description': 'Gradient direction in degrees.'
      },
      {
        '__metadata': { 'type': 'SP.Field' },
        'FieldTypeKind': 2,
        'Title': 'OverlayGradientColor',
        'Required': false,
        'DefaultValue': '#000000',
        'Description': 'Overlay gradient color (hex).'
      },
      {
        '__metadata': { 'type': 'SP.FieldNumber' },
        'FieldTypeKind': 9,
        'Title': 'OverlayOpacity',
        'Required': false,
        'MinimumValue': 0,
        'MaximumValue': 100,
        'Description': 'Overlay opacity percentage.'
      },
      {
        '__metadata': { 'type': 'SP.FieldUrl' },
        'FieldTypeKind': 11,
        'Title': 'RedirectUrl',
        'Required': false,
        'Description': 'Optional URL to navigate to when the announcement is clicked.'
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'FieldTypeKind': 6,
        'Title': 'RedirectTarget',
        'Required': false,
        'Choices': { 'results': ['_self', '_blank'] },
        'DefaultValue': '_self',
        'Description': 'Open redirect URL in same tab or new tab.'
      }
    ];
  }

  /**
   * Creates columns sequentially using odata=verbose Content-Type.
   * Verbose format is required for typed fields (Choice, MultiLineText, etc.)
   * that need __metadata.type and { results: [...] } for Choices arrays.
   * Accept header remains nometadata for cleaner responses.
   */
  private static async createColumns(context: WebPartContext, onlyColumns?: string[]): Promise<void> {
    const baseUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/fields`;

    let columns = this.getColumnDefinitions();
    if (onlyColumns && onlyColumns.length > 0) {
      columns = columns.filter(col => onlyColumns.includes(col.Title));
    }

    for (const column of columns) {
      try {
        const response: SPHttpClientResponse = await context.spHttpClient.post(
          baseUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
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
    const allColumns = this.getColumnDefinitions();
    const requiredColumns = allColumns.map(col => col.Title);
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
          console.log(`Missing columns detected: ${missingColumns.join(', ')}. Creating only those...`);
          await this.createColumns(context, missingColumns);
        }
      }
    } catch (error) {
      console.warn('Error verifying columns:', error);
    }
  }

  private static async configureListSettings(context: WebPartContext): Promise<void> {
    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${this.LIST_NAME}')/DefaultView/ViewFields`;

    const viewFields = ['Title', 'Description', 'AnnouncementImage', 'ValidFrom', 'ValidTo', 'CelebrationIcon'];

    for (const field of viewFields) {
      try {
        await context.spHttpClient.post(
          `${endpoint}/addViewField('${field}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'odata-version': ''
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
