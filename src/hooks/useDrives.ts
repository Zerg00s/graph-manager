import { useState, useCallback, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import { Drive } from '@microsoft/microsoft-graph-types';
import { GraphService } from '../services/graphService';

interface UseDrivesResult {
  drives: Drive[];
  loading: boolean;
  error: string | null;
  hasMore: boolean;
  loadDrives: () => Promise<void>;
  loadMore: () => Promise<void>;
  searchDrives: (searchTerm: string) => Promise<void>;
  clearDrives: () => void;
}

export const useDrives = (
  pageSize: number = 100,
): UseDrivesResult => {
  const { instance } = useMsal();

  // Create the service exactly once per component instance.
  const graphService = useMemo(() => new GraphService(instance), [instance]);

  const [drives, setDrives] = useState<Drive[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [hasMore, setHasMore] = useState(false);

  const loadDrives = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      // First try to get all drives (including shared ones)
      const result = await graphService.getAllDrives(pageSize);
      setDrives(result.value);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err) {
      try {
        // If that fails, fallback to just user's drives
        const result = await graphService.getDrives(pageSize);
        setDrives(result.value);
        setNextLink(result.nextLink);
        setHasMore(Boolean(result.nextLink));
      } catch (fallbackErr) {
        setError(
          fallbackErr instanceof Error ? fallbackErr.message : 'Failed to load drives',
        );
        setDrives([]);
      }
    } finally {
      setLoading(false);
    }
  }, [graphService, pageSize]);

  const loadMore = useCallback(async () => {
    if (!nextLink || loading) return;

    setLoading(true);
    setError(null);

    try {
      const result = await graphService.getAllDrives(pageSize, nextLink);
      setDrives(prev => [...prev, ...result.value]);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to load more drives',
      );
    } finally {
      setLoading(false);
    }
  }, [graphService, nextLink, loading, pageSize]);

  const searchDrives = useCallback(
    async (searchTerm: string) => {
      if (!searchTerm.trim()) {
        return loadDrives();
      }

      setLoading(true);
      setError(null);

      try {
        // Since the Microsoft Graph API doesn't support direct search on drives,
        // we'll filter the results client-side
        const result = await graphService.getAllDrives(pageSize);
        const filteredDrives = result.value.filter(drive => {
          const nameMatch = drive.name?.toLowerCase().includes(searchTerm.toLowerCase());
          const typeMatch = drive.driveType?.toLowerCase().includes(searchTerm.toLowerCase());
          
          // Check owner identity - handle the IdentitySet structure
          let ownerMatch = false;
          if (drive.owner) {
            const ownerUser = drive.owner.user?.displayName?.toLowerCase();
            const ownerGroup = drive.owner.group?.displayName?.toLowerCase();
            const ownerApp = drive.owner.application?.displayName?.toLowerCase();
            const ownerDevice = drive.owner.device?.displayName?.toLowerCase();
            
            ownerMatch = !!(
              ownerUser?.toLowerCase().includes(searchTerm.toLowerCase()) ||
              ownerGroup?.toLowerCase().includes(searchTerm.toLowerCase()) ||
              ownerApp?.toLowerCase().includes(searchTerm.toLowerCase()) ||
              ownerDevice?.toLowerCase().includes(searchTerm.toLowerCase())
            );
          }
          
          return nameMatch || typeMatch || ownerMatch;
        });
        
        setDrives(filteredDrives);
        setNextLink(result.nextLink);
        setHasMore(Boolean(result.nextLink));
      } catch (err) {
        setError(
          err instanceof Error ? err.message : 'Failed to search drives',
        );
        setDrives([]);
      } finally {
        setLoading(false);
      }
    },
    [graphService, loadDrives, pageSize],
  );

  const clearDrives = useCallback(() => {
    setDrives([]);
    setNextLink(undefined);
    setHasMore(false);
    setError(null);
  }, []);

  return {
    drives,
    loading,
    error,
    hasMore,
    loadDrives,
    loadMore,
    searchDrives,
    clearDrives,
  };
};
