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
        await this.createColumns(context);
        await this.configureListSettings(context);
        console.log('List created successfully');
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

    // Define columns with their XML schema
    const columns = [
      {
        parameters: {
          'FieldTypeKind': 3,
          'Title': 'Description',
          'Required': false,
          'RichText': true,
          'NumberOfLines': 6
        }
      },
      {
        parameters: {
          'FieldTypeKind': 11,
          'Title': 'AnnouncementImage',
          'Required': false
        }
      },
      {
        parameters: {
          'FieldTypeKind': 4,
          'Title': 'ValidFrom',
          'Required': true,
          'DisplayFormat': 1
        }
      },
      {
        parameters: {
          'FieldTypeKind': 4,
          'Title': 'ValidTo',
          'Required': true,
          'DisplayFormat': 1
        }
      },
      {
        parameters: {
          'FieldTypeKind': 6,
          'Title': 'CelebrationIcon',
          'Required': false,
          'Choices': ['None', 'Birthday', 'Anniversary', 'Achievement', 'Celebration', 'NewHire', 'Promotion', 'Holiday', 'Custom'],
          'DefaultValue': 'None'
        }
      },
      {
        parameters: {
          'FieldTypeKind': 6,
          'Title': 'CelebrationIconPosition',
          'Required': false,
          'Choices': ['topLeft', 'topRight', 'bottomLeft', 'bottomRight', 'center'],
          'DefaultValue': 'topRight'
        }
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
            body: JSON.stringify(column.parameters)
          }
        );

        if (!response.ok) {
          const errorText = await response.text();
          console.warn(`Warning creating column ${column.parameters.Title}: ${errorText}`);
        } else {
          console.log(`Column ${column.parameters.Title} created successfully`);
        }
      } catch (error) {
        console.warn(`Error creating column ${column.parameters.Title}:`, error);
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
