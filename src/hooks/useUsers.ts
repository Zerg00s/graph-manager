import { useState, useCallback, useMemo } from 'react';
import { useMsal } from '@azure/msal-react';
import { User } from '@microsoft/microsoft-graph-types';
import { GraphService } from '../services/graphService';

interface UseUsersResult {
  users: User[];
  loading: boolean;
  error: string | null;
  hasMore: boolean;
  loadUsers: () => Promise<void>;
  loadMore: () => Promise<void>;
  searchUsers: (searchTerm: string) => Promise<void>;
  clearUsers: () => void;
}

export const useUsers = (
  pageSize: number = 100,
): UseUsersResult => {
  const { instance } = useMsal();

  // Create the service exactly once per component instance.
  const graphService = useMemo(() => new GraphService(instance), [instance]);

  const [users, setUsers] = useState<User[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [hasMore, setHasMore] = useState(false);

  const loadUsers = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      const result = await graphService.getUsers(pageSize);
      setUsers(result.value);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to load users',
      );
      setUsers([]);
    } finally {
      setLoading(false);
    }
  }, [graphService, pageSize]);

  const loadMore = useCallback(async () => {
    if (!nextLink || loading) return;

    setLoading(true);
    setError(null);

    try {
      const result = await graphService.getUsers(pageSize, nextLink);
      setUsers(prev => [...prev, ...result.value]);
      setNextLink(result.nextLink);
      setHasMore(Boolean(result.nextLink));
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to load more users',
      );
    } finally {
      setLoading(false);
    }
  }, [graphService, nextLink, loading, pageSize]);

  const searchUsers = useCallback(
    async (searchTerm: string) => {
      if (!searchTerm.trim()) {
        return loadUsers();
      }

      setLoading(true);
      setError(null);

      try {
        const result = await graphService.searchUsers(searchTerm, pageSize);
        setUsers(result.value);
        setNextLink(result.nextLink);
        setHasMore(Boolean(result.nextLink));
      } catch (err) {
        setError(
          err instanceof Error ? err.message : 'Failed to search users',
        );
        setUsers([]);
      } finally {
        setLoading(false);
      }
    },
    [graphService, loadUsers, pageSize],
  );

  const clearUsers = useCallback(() => {
    setUsers([]);
    setNextLink(undefined);
    setHasMore(false);
    setError(null);
  }, []);

  return {
    users,
    loading,
    error,
    hasMore,
    loadUsers,
    loadMore,
    searchUsers,
    clearUsers,
  };
};
