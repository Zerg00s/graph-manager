import { Identity } from '@microsoft/microsoft-graph-types';

declare module '@microsoft/microsoft-graph-types' {
  interface IdentitySet {
    /**
     * Optional. The application associated with this action.
     */
    application?: Identity;
    /**
     * Optional. The device associated with this action.
     */
    device?: Identity;
    /**
     * Optional. The user associated with this action.
     */
    user?: Identity;
    /**
     * Optional. The group associated with this action.
     */
    group?: Identity;
  }
  
  interface Drive {
    /**
     * Identity of the user, device, or application which created the item. Read-only.
     */
    createdBy?: IdentitySet;
    /**
     * Identity of the user, device, application which last modified the item. Read-only.
     */
    lastModifiedBy?: IdentitySet;
    /**
     * Optional. The user account that owns the drive. Read-only.
     */
    owner?: IdentitySet;
  }
}

// SharePoint Container type definition
export interface SharePointContainer {
  id: string;
  displayName?: string;
  description?: string;
  containerTypeId?: string;
  status?: string;
  createdDateTime?: string;
  storageUsedInBytes?: number;
  customProperties?: Record<string, string>;
}
