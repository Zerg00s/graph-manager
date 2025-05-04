import { Client } from '@microsoft/microsoft-graph-client';
import {
  AuthCodeMSALBrowserAuthenticationProvider,
} from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import {
  IPublicClientApplication,
  PublicClientApplication,
  InteractionType,
} from '@azure/msal-browser';
import { Site, Drive } from '@microsoft/microsoft-graph-types';
import { SharePointContainer } from '../types/microsoft-graph-extended';

export interface PagedResult<T> {
  value: T[];
  nextLink?: string;
}

export class GraphService {
  private graphClient: Client;

  constructor(msal: IPublicClientApplication) {
    // The runtime object returned by useMsal() *is* a PublicClientApplication;
    // we just have to convince the compiler.
    const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(
      msal as PublicClientApplication,
      {
        account: msal.getAllAccounts()[0] ?? undefined,
        scopes: ['User.Read', 'Sites.Read.All', 'Files.Read.All', 'FileStorageContainer.Selected'],
        interactionType: InteractionType.Popup, // or InteractionType.Redirect
      },
    );

    this.graphClient = Client.initWithMiddleware({ authProvider });
  }

  async getSharePointSites(
    pageSize: number = 100,
    nextLink?: string,
  ): Promise<PagedResult<Site>> {
    const response = nextLink
      ? await this.graphClient.api(nextLink).get()
      : await this.graphClient
          .api('/sites')
          .select('id,displayName,webUrl,createdDateTime')
          .top(pageSize)
          .get();

    return {
      value: response.value,
      nextLink: response['@odata.nextLink'],
    };
  }

  async searchSharePointSites(
    searchTerm: string,
    pageSize: number = 100,
  ): Promise<PagedResult<Site>> {
    const response = await this.graphClient
      .api('/sites')
      .search(searchTerm)
      .select('id,displayName,webUrl,createdDateTime')
      .top(pageSize)
      .get();

    return {
      value: response.value,
      nextLink: response['@odata.nextLink'],
    };
  }

  async getDrives(
    pageSize: number = 100,
    nextLink?: string,
  ): Promise<PagedResult<Drive>> {
    const response = nextLink
      ? await this.graphClient.api(nextLink).get()
      : await this.graphClient
          .api('/me/drives')
          .select('id,name,driveType,webUrl,createdDateTime,owner')
          .top(pageSize)
          .get();

    return {
      value: response.value,
      nextLink: response['@odata.nextLink'],
    };
  }

  async getSiteDrives(
    siteId: string,
    pageSize: number = 100,
    nextLink?: string,
  ): Promise<PagedResult<Drive>> {
    const response = nextLink
      ? await this.graphClient.api(nextLink).get()
      : await this.graphClient
          .api(`/sites/${siteId}/drives`)
          .select('id,name,driveType,webUrl,createdDateTime,owner')
          .top(pageSize)
          .get();

    return {
      value: response.value,
      nextLink: response['@odata.nextLink'],
    };
  }

  async getAllDrives(
    pageSize: number = 100,
    nextLink?: string,
  ): Promise<PagedResult<Drive>> {
    // Fetch drives from both user's drives and root site drives
    const response = nextLink
      ? await this.graphClient.api(nextLink).get()
      : await this.graphClient
          .api('/drives')
          .select('id,name,driveType,webUrl,createdDateTime,owner')
          .top(pageSize)
          .get();

    return {
      value: response.value,
      nextLink: response['@odata.nextLink'],
    };
  }

  async getContainerTypes(): Promise<PagedResult<any>> {
    try {
      const response = await this.graphClient
        .api('/storage/fileStorage/containerTypes')
        .version('beta')  // Use beta API
        .get();  // Get all properties to see what's available

      // Log to see what properties exist
      console.log('Container types response:', response);

      return {
        value: response.value,
        nextLink: response['@odata.nextLink'],
      };
    } catch (error: any) {
      if (error.statusCode === 403 || error.message?.includes('403')) {
        throw new Error('Access denied. Additional permissions may be required to list container types.');
      }
      throw error;
    }
  }

  async getSharePointContainers(
    pageSize: number = 100,
    nextLink?: string,
    containerTypeId?: string,
  ): Promise<PagedResult<SharePointContainer>> {
    try {
      // If no containerTypeId provided, show a helpful error
      if (!containerTypeId) {
        throw new Error('A container type ID is required. You can find available container types using the Container Types feature.');
      }

      const response = nextLink
        ? await this.graphClient.api(nextLink).version('beta').get()
        : await this.graphClient
            .api('/storage/fileStorage/containers')
            .version('beta')  // Use beta API
            .top(pageSize)
            .filter(`containerTypeId eq '${containerTypeId}'`)  // Always use quotes for the filter
            .get();  // Get all properties first to see what's available

      // Log to see what properties exist
      console.log('SharePoint containers response:', response);

      return {
        value: response.value,
        nextLink: response['@odata.nextLink'],
      };
    } catch (error: any) {
      console.error('Error fetching containers:', error);
      
      if (error.statusCode === 403 || error.message?.includes('403')) {
        throw new Error('Access denied. The FileStorageContainer.Selected permission may be required.');
      }
      
      if (error.message?.includes('containerTypeId filter parameter is required')) {
        throw new Error('A container type ID is required to list SharePoint containers.');
      }
      
      if (error.message?.includes('Invalid filter clause') || error.message?.includes('incompatible types')) {
        throw new Error('Invalid container type ID format. The container type ID should be a GUID (e.g., "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"). You may need to get the correct container type ID from your SharePoint administrator.');
      }
      
      // If it's a 404, the API might not be available
      if (error.statusCode === 404) {
        throw new Error('SharePoint Embedded API not found. Please ensure SharePoint Embedded is enabled for your tenant and you have the required permissions.');
      }
      
      throw error;
    }
  }
  
  

  async getContainerPermissions(
    containerId: string,
  ): Promise<PagedResult<any>> {
    const response = await this.graphClient
      .api(`/storage/fileStorage/containers/${containerId}/permissions`)
      .version('beta')  // Use beta API
      .get();

    return {
      value: response.value,
      nextLink: response['@odata.nextLink'],
    };
  }
}
