import { useState, useCallback, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import { SharePointContainer } from '../types/microsoft-graph-extended';
import { GraphService } from '../services/graphService';

interface UseSharePointContainersResult {
  containers: SharePointContainer[];
  loading: boolean;
  error: string | null;
  hasMore: boolean;
  containerTypeId: string | null;
  setContainerTypeId: (id: string | null) => void;
  loadContainers: () => Promise<void>;
  loadMore: () => Promise<void>;
  searchContainers: (searchTerm: string) => Promise<void>;
  clearContainers: () => void;
}

export const useSharePointContainers = (
  pageSize: number = 100,
): UseSharePointContainersResult => {
  const { instance } = useMsal();

  // Create the service exactly once per component instance.
  const graphService = useMemo(() => new GraphService(instance), [instance]);

  const [containers, setContainers] = useState<SharePointContainer[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [hasMore, setHasMore] = useState(false);
  const [containerTypeId, setContainerTypeId] = useState<string | null>(null);

  const loadContainers = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      // If no containerTypeId is specified, show an informative error
      if (!containerTypeId) {
        setError('Please specify a container type ID to list SharePoint containers. The API requires this parameter.');
        setContainers([]);
        return;
      }

      const result = await graphService.getSharePointContainers(pageSize, undefined, containerTypeId);
      setContainers(result.value);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err: any) {
      // Check if it's a consent error
      if (err && (err.message?.includes('consent_required') || err.message?.includes('65001'))) {
        setError('Additional permissions are required. Please sign out and sign back in to grant the required permissions for accessing SharePoint containers.');
      } else if (err && (err.message?.includes('403') || err.statusCode === 403 || err.message?.includes('Access denied'))) {
        setError('Access denied. The FileStorageContainer.Selected permission may be required for your application to access SharePoint containers. Please check with your administrator.');
      } else if (err && (err.message?.includes('400') || err.statusCode === 400)) {
        setError('Bad request. The SharePoint containers API requires a containerTypeId filter parameter.');
      } else if (err && err.message?.includes('containerTypeId')) {
        setError(err.message);
      } else {
        setError(
          err?.message || 'Failed to load SharePoint containers',
        );
      }
      setContainers([]);
    } finally {
      setLoading(false);
    }
  }, [graphService, pageSize, containerTypeId]);

  const loadMore = useCallback(async () => {
    if (!nextLink || loading || !containerTypeId) return;

    setLoading(true);
    setError(null);

    try {
      const result = await graphService.getSharePointContainers(pageSize, nextLink, containerTypeId);
      setContainers(prev => [...prev, ...result.value]);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err: any) {
      setError(
        err?.message || 'Failed to load more SharePoint containers',
      );
    } finally {
      setLoading(false);
    }
  }, [graphService, nextLink, loading, pageSize, containerTypeId]);

  const searchContainers = useCallback(
    async (searchTerm: string) => {
      if (!searchTerm.trim()) {
        return loadContainers();
      }

      if (!containerTypeId) {
        setError('Please specify a container type ID to search SharePoint containers.');
        return;
      }

      setLoading(true);
      setError(null);

      try {
        // Since the Microsoft Graph API doesn't support direct search on containers,
        // we'll filter the results client-side
        const result = await graphService.getSharePointContainers(pageSize, undefined, containerTypeId);
        const filteredContainers = result.value.filter(container => {
          const displayNameMatch = container.displayName?.toLowerCase().includes(searchTerm.toLowerCase());
          const descriptionMatch = container.description?.toLowerCase().includes(searchTerm.toLowerCase());
          const typeMatch = container.containerTypeId?.toLowerCase().includes(searchTerm.toLowerCase());
          
          return displayNameMatch || descriptionMatch || typeMatch;
        });
        
        setContainers(filteredContainers);
        setNextLink(result.nextLink);
        setHasMore(Boolean(result.nextLink));
      } catch (err: any) {
        setError(
          err?.message || 'Failed to search SharePoint containers',
        );
        setContainers([]);
      } finally {
        setLoading(false);
      }
    },
    [graphService, loadContainers, pageSize, containerTypeId],
  );

  const clearContainers = useCallback(() => {
    setContainers([]);
    setNextLink(undefined);
    setHasMore(false);
    setError(null);
  }, []);

  return {
    containers,
    loading,
    error,
    hasMore,
    containerTypeId,
    setContainerTypeId,
    loadContainers,
    loadMore,
    searchContainers,
    clearContainers,
  };
};
