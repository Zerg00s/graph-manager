# Drives Implementation for Microsoft Graph API Manager

## Overview

This document describes the implementation of the Drives feature for the Microsoft Graph API Manager application. The feature allows users to list and search through their Microsoft 365 drives, including OneDrive personal drives, SharePoint document libraries, and shared drives.

## Features Implemented

1. **List Drives**: View all accessible drives in a paginated data grid
2. **Search Drives**: Filter drives by name, type, or owner
3. **Drive Details**: Display drive information including:
   - Drive name
   - Drive type (personal, business, documentLibrary, etc.)
   - Web URL with direct link
   - Owner information (user, group, application, or device)
   - Creation date
4. **Pagination**: Support for loading more drives with infinite scroll
5. **Refresh**: Ability to refresh the drive list

## Technical Implementation

### Components

1. **Drives.tsx**: Main component that displays the drives list
   - Uses Fluent UI React components for consistent UI
   - Implements search functionality with debouncing
   - Provides pagination support
   - Handles error states and loading indicators
   - Properly handles Microsoft Graph API's IdentitySet structure for owner information

### Hooks

2. **useDrives.ts**: Custom hook for managing drives data
   - Handles API calls through GraphService
   - Manages loading states and error handling
   - Implements search with client-side filtering
   - Provides pagination functionality
   - Supports searching across different identity types (user, group, application, device)

### Services

3. **GraphService.ts**: Updated to include drive-related methods
   - `getDrives()`: Fetches user's personal drives
   - `getSiteDrives()`: Fetches drives for a specific SharePoint site
   - `getAllDrives()`: Fetches all accessible drives

### Type Definitions

4. **microsoft-graph-extended.d.ts**: Extended type definitions
   - Provides proper TypeScript types for IdentitySet structure
   - Extends Drive interface with proper owner type definitions

## API Endpoints Used

- `/me/drives`: Get current user's drives
- `/sites/{site-id}/drives`: Get drives for a specific site
- `/drives`: Get all drives (requires appropriate permissions)

## Permissions Required

The following Microsoft Graph API permissions are required:
- `User.Read`: Read user profile
- `Sites.Read.All`: Read all SharePoint sites
- `Files.Read.All`: Read all files user can access

## Data Structure

The Microsoft Graph API returns drives with an IdentitySet structure for owner information:

```typescript
interface IdentitySet {
  user?: Identity;
  group?: Identity;
  application?: Identity;
  device?: Identity;
}

interface Identity {
  displayName?: string;
  id?: string;
}
```

The implementation properly handles all types of identities (user, group, application, device) when displaying owner information.

## Usage

1. Sign in to the application with your Microsoft 365 account
2. Click on "View Drives" button on the home page
3. The drives list will load automatically
4. Use the search box to filter drives by name, type, or owner
5. Click "Load More Drives" to see additional results
6. Click on any web URL to open the drive in your browser
7. Use the "Back" button to return to the home page

## Known Limitations

1. Search is performed client-side as the Microsoft Graph API doesn't provide native search functionality for drives
2. Some SharePoint document libraries may not be accessible based on permissions
3. Owner information may vary depending on the drive type and permissions

## Future Enhancements

1. Add ability to view files within each drive
2. Implement drive-specific operations (create folder, upload file, etc.)
3. Add support for filtering by drive type
4. Implement server-side search when API support becomes available
5. Add support for displaying quota information
6. Add ability to browse drive contents inline
