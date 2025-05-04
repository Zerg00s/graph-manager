# SharePoint Containers Implementation for Microsoft Graph API Manager

## Overview

This document describes the implementation of the SharePoint Containers feature for the Microsoft Graph API Manager application. SharePoint Embedded containers enable developers to store files directly within their applications while leveraging SharePoint's powerful content management capabilities.

## Features Implemented

1. **List SharePoint Containers**: View all accessible SharePoint Embedded containers in a paginated data grid
2. **Search Containers**: Filter containers by name, description, or container type ID
3. **Container Details**: Display container information including:
   - Container name
   - Description
   - Container Type ID
   - Status
   - Creation date
4. **Pagination**: Support for loading more containers with infinite scroll
5. **Refresh**: Ability to refresh the container list

## Technical Implementation

### Components

1. **SharePointContainers.tsx**: Main component that displays the containers list
   - Uses Fluent UI React components for consistent UI
   - Implements search functionality with debouncing
   - Provides pagination support
   - Handles error states and loading indicators
   - Displays appropriate error messages for permission issues

### Hooks

2. **useSharePointContainers.ts**: Custom hook for managing SharePoint containers data
   - Handles API calls through GraphService
   - Manages loading states and error handling
   - Implements client-side search filtering
   - Provides pagination functionality
   - Handles permission errors with helpful messages

### Services

3. **GraphService.ts**: Updated to include SharePoint container-related methods
   - `getSharePointContainers()`: Fetches all SharePoint containers
   - `getContainerPermissions()`: Fetches permissions for a specific container

### Type Definitions

4. **microsoft-graph-extended.d.ts**: Extended type definitions
   - Added `SharePointContainer` interface with proper TypeScript types

## API Endpoints Used

- `/storage/fileStorage/containers` (beta endpoint): Get all SharePoint Embedded containers
- `/storage/fileStorage/containers/{containerId}/permissions` (beta endpoint): Get permissions for a specific container

Note: This feature uses the Microsoft Graph beta API for better support of SharePoint Embedded features.

## Permissions Required

The following Microsoft Graph API permissions may be required:
- `FileStorageContainer.Selected`: Access to SharePoint Embedded containers

**IMPORTANT NOTE**: The `FileStorageContainer.Selected` permission is a specialized permission for SharePoint Embedded. This permission:
- May require your tenant to have SharePoint Embedded enabled
- May require your application to be registered as a SharePoint Embedded application
- May need to be configured by your tenant administrator

## Known Issues and Limitations

1. **Container Type ID Required**: The SharePoint containers API requires a containerTypeId filter parameter to list containers. You must specify a container type ID before you can view any containers.
2. **Permission Configuration**: The `FileStorageContainer.Selected` permission is not a standard Graph permission and may require special configuration in Azure AD
3. **Tenant Requirements**: SharePoint Embedded must be enabled for your tenant
4. **API Availability**: The SharePoint containers API is relatively new and may not be available in all tenants
5. **Search is Client-Side**: Search is performed client-side as the API doesn't support server-side search

## Error Handling

The implementation includes specific error handling for common scenarios:
- 403 Forbidden: Indicates missing permissions
- 400 Bad Request: Indicates the API might not be available or properly configured
- Generic errors: Caught and displayed with user-friendly messages

## Usage

1. Sign in to the application with your Microsoft 365 account
2. Click on "List SharePoint Containers" button on the home page
3. **Enter a Container Type ID**: The SharePoint containers API requires a container type ID filter
   - Enter a container type ID in GUID format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - Example: `f47ac10b-58cc-4372-a567-0e02b2c3d479`
   - **Important**: This is different from individual container IDs (which start with `b!`)
   - Get the correct container type ID from your SharePoint Embedded administrator
   - You can also try to list container types using the Graph API: `/storage/fileStorage/containerTypes`
4. If you see an error about permissions or API availability, contact your administrator
5. Use the search box to filter containers by name, description, or container type ID
6. Click "Load More Containers" to see additional results
7. Use the "Back" button to return to the home page API requires a container type ID filter
   - Get the container type ID from your SharePoint Embedded administrator
   - Enter the ID in the input field and click "Load Containers"
4. If you see an error about permissions or API availability, contact your administrator
5. Use the search box to filter containers by name, description, or container type ID
6. Click "Load More Containers" to see additional results
7. Use the "Back" button to return to the home page

## SharePoint Embedded Concepts

SharePoint Embedded is a cloud-based solution that allows developers to store and manage content within their applications using SharePoint's powerful capabilities. Key concepts include:

1. **Container Types**: Define the schema and capabilities for containers
2. **Containers**: Instances of container types that store actual content
3. **Permissions**: Granular access control for containers and their content
4. **API-first approach**: All functionality accessible through Microsoft Graph APIs

## Troubleshooting

If you encounter errors when trying to access SharePoint containers:

1. **"containerTypeId filter parameter is required" Error**: Enter a valid container type ID before trying to load containers. This is a requirement of the SharePoint containers API.
2. **"Invalid filter clause" Error**: Make sure the container type ID is in GUID format (e.g., `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`)
3. **"Incompatible types" Error**: The API expects a GUID for containerTypeId, not a container ID (b!xxx format)
4. **403 Forbidden Error**: Check that your application has the required permissions in Azure AD
5. **400 Bad Request Error**: Verify that SharePoint Embedded is enabled for your tenant
6. **No Data Returned**: Ensure that SharePoint containers exist in your tenant and that you have the correct container type ID

## Finding Your Container Type ID

To find a valid container type ID:

1. Contact your SharePoint Embedded administrator
2. Use the Microsoft Graph Explorer to list container types:
   ```
   GET https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypes
   ```
3. The container type ID should be in GUID format (e.g., `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`)
4. This is different from individual container IDs which start with `b!`

## Understanding Container IDs vs Container Type IDs

- **Container Type ID**: A GUID that identifies the type/template of containers (e.g., `f47ac10b-58cc-4372-a567-0e02b2c3d479`)
- **Container ID**: A unique identifier for individual containers, typically starting with `b!` (e.g., `b!nxwKP_hVfUODdMXxUFSumTA8v7Gs311DrxNBlyiAP23e8tPTs1_MRaK1qrY8Ecyv`)

The API requires filtering by container type ID (GUID format) to list containers.

For more information about SharePoint Embedded, refer to the official Microsoft documentation.
