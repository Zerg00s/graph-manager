# Container Types Feature for Microsoft Graph API Manager

## Overview

The Container Types feature allows you to discover and list available SharePoint Embedded container types in your tenant. This is essential for finding the correct container type IDs (GUIDs) needed to query SharePoint containers.

## How to Find Container Type IDs

### Method 1: Using This Application

1. Sign in to the application
2. Click "View Container Types" button on the home page
3. The application will query the `/storage/fileStorage/containerTypes` endpoint
4. You'll see a list of available container types with their:
   - ID (GUID format)
   - Display Name
   - Description
5. Click the copy button next to any ID to copy it
6. Use the "Use This Type" button to automatically populate it in the SharePoint Containers view

### Method 2: Using Microsoft Graph Explorer

1. Go to [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Sign in with your Microsoft 365 account
3. Run this query:
   ```
   GET https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypes
   ```
4. The response will contain container types with their GUIDs

### Method 3: Using PowerShell

```powershell
# Install the Microsoft Graph PowerShell module
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "FileStorageContainer.Selected"

# Get container types
Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/storage/fileStorage/containerTypes"
```

## Understanding Container Type IDs

- Container Type IDs are GUIDs (e.g., `f47ac10b-58cc-4372-a567-0e02b2c3d479`)
- They define the template/schema for SharePoint Embedded containers
- One container type can have multiple containers
- You need the container type ID to query containers via the Graph API

## Required Permissions

To view container types, you need:
- `FileStorageContainer.Selected` permission
- Appropriate admin consent

## Integration with SharePoint Containers

1. View Container Types to find available types
2. Copy the GUID of the desired container type
3. Go to SharePoint Containers view
4. Either:
   - Paste the GUID and click "Load Containers"
   - Click "Browse Types" to select from the list

## Troubleshooting

1. **No container types returned**: Ensure SharePoint Embedded is enabled in your tenant
2. **Access denied**: Check that your application has the required permissions
3. **Empty list**: Verify that container types have been created in your tenant

## API Reference

Endpoint: `GET /storage/fileStorage/containerTypes` (beta endpoint)

Note: This feature uses the Microsoft Graph beta API for better support of SharePoint Embedded features.

Response format:
```json
{
  "value": [
    {
      "id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "displayName": "Project Files",
      "description": "Container type for project files"
    }
  ]
}
```
