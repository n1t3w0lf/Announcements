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
      '__metadata': {
        'type': 'SP.List'
      },
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
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
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
        'Choices': {
          '__metadata': { 'type': 'Collection(Edm.String)' },
          'results': ['None', 'Birthday', 'Anniversary', 'Achievement', 'Celebration', 'NewHire', 'Promotion', 'Holiday', 'Custom']
        },
        'DefaultValue': 'None'
      },
      {
        '__metadata': { 'type': 'SP.FieldChoice' },
        'FieldTypeKind': 6,
        'Title': 'CelebrationIconPosition',
        'Required': false,
        'Choices': {
          '__metadata': { 'type': 'Collection(Edm.String)' },
          'results': ['topLeft', 'topRight', 'bottomLeft', 'bottomRight', 'center']
        },
        'DefaultValue': 'topRight'
      }
    ];

    for (const column of columns) {
      try {
        const response: SPHttpClientResponse = await context.spHttpClient.post(
          baseUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose'
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
    const requiredColumns = ['Description', 'AnnouncementImage', 'ValidFrom', 'ValidTo', 'CelebrationIcon', 'CelebrationIconPosition'];
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
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose'
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
