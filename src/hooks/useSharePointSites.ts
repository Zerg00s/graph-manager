import { useState, useCallback, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import { Site } from '@microsoft/microsoft-graph-types';
import { GraphService } from '../services/graphService';

interface UseSharePointSitesResult {
  sites: Site[];
  loading: boolean;
  error: string | null;
  hasMore: boolean;
  loadSites: () => Promise<void>;
  loadMore: () => Promise<void>;
  searchSites: (searchTerm: string) => Promise<void>;
  clearSites: () => void;
}

export const useSharePointSites = (
  pageSize: number = 100,
): UseSharePointSitesResult => {
  const { instance } = useMsal();

  // Create the service exactly once per component instance.
  const graphService = useMemo(() => new GraphService(instance), [instance]);

  const [sites, setSites] = useState<Site[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [hasMore, setHasMore] = useState(false);

  const loadSites = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      const result = await graphService.getSharePointSites(pageSize);
      setSites(result.value);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to load SharePoint sites',
      );
      setSites([]);
    } finally {
      setLoading(false);
    }
  }, [graphService, pageSize]);

  const loadMore = useCallback(async () => {
    if (!nextLink || loading) return;

    setLoading(true);
    setError(null);

    try {
      const result = await graphService.getSharePointSites(pageSize, nextLink);
      setSites(prev => [...prev, ...result.value]);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to load more SharePoint sites',
      );
    } finally {
      setLoading(false);
    }
  }, [graphService, nextLink, loading, pageSize]);

  const searchSites = useCallback(
    async (searchTerm: string) => {
      if (!searchTerm.trim()) {
        return loadSites();
      }

      setLoading(true);
      setError(null);

      try {
        const result = await graphService.searchSharePointSites(
          searchTerm,
          pageSize,
        );
        setSites(result.value);
        setNextLink(result.nextLink);
        setHasMore(Boolean(result.nextLink));
      } catch (err) {
        setError(
          err instanceof Error ? err.message : 'Failed to search SharePoint sites',
        );
        setSites([]);
      } finally {
        setLoading(false);
      }
    },
    [graphService, loadSites, pageSize],
  );

  const clearSites = useCallback(() => {
    setSites([]);
    setNextLink(undefined);
    setHasMore(false);
    setError(null);
  }, []);

  return {
    sites,
    loading,
    error,
    hasMore,
    loadSites,
    loadMore,
    searchSites,
    clearSites,
  };
};
