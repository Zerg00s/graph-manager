import React, { useEffect, useState, useCallback } from 'react';
import {
  makeStyles,
  tokens,
  SearchBox,
  DataGrid,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridBody,
  DataGridRow,
  DataGridCell,
  TableColumnDefinition,
  createTableColumn,
  Button,
  Spinner,
  Card,
  CardHeader,
  Title3,
  Link,
  Text,
} from '@fluentui/react-components';
import { 
  ArrowLeftRegular,
  OpenRegular,
  PeopleRegular,
} from '@fluentui/react-icons';
import { User } from '@microsoft/microsoft-graph-types';
import { useUsers } from '../hooks/useUsers';

const useStyles = makeStyles({
  container: {
    display: 'flex' as unknown as never,
    flexDirection: 'column' as unknown as never,
    height: 'calc(100vh - 200px)' as unknown as never,
    gap: tokens.spacingVerticalM as unknown as never,
  },
  header: {
    display: 'flex' as unknown as never,
    alignItems: 'center' as unknown as never,
    gap: tokens.spacingHorizontalM as unknown as never,
    marginBottom: tokens.spacingVerticalM as unknown as never,
  },
  searchBox: {
    width: '300px' as unknown as never,
  },
  dataGrid: {
    flex: 1 as unknown as never,
    overflow: 'auto' as unknown as never,
  },
  loadMoreContainer: {
    display: 'flex' as unknown as never,
    justifyContent: 'center' as unknown as never,
    padding: tokens.spacingVerticalL as unknown as never,
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1 as unknown as never,
    padding: tokens.spacingVerticalM as unknown as never,
  },
  linkCell: {
    display: 'flex' as unknown as never,
    alignItems: 'center' as unknown as never,
    gap: tokens.spacingHorizontalS as unknown as never,
  },
});

interface UsersProps {
  onBack: () => void;
}

export const Users: React.FC<UsersProps> = ({ onBack }) => {
  const styles = useStyles();
  const [searchQuery, setSearchQuery] = useState('');
  const { users, loading, error, hasMore, loadUsers, loadMore, searchUsers, clearUsers } = useUsers(100);

  useEffect(() => {
    loadUsers();
  }, [loadUsers]);

  const handleSearch = useCallback((value: string) => {
    setSearchQuery(value);
    // Debounce search
    const timeoutId = setTimeout(() => {
      searchUsers(value);
    }, 300);
    return () => clearTimeout(timeoutId);
  }, [searchUsers]);

  const handleRefresh = useCallback(() => {
    clearUsers();
    loadUsers();
  }, [clearUsers, loadUsers]);

  const columns: TableColumnDefinition<User>[] = [
    createTableColumn<User>({
      columnId: 'displayName',
      renderHeaderCell: () => 'Display Name',
      renderCell: (item) => item.displayName || 'Unnamed User',
    }),
    createTableColumn<User>({
      columnId: 'userPrincipalName',
      renderHeaderCell: () => 'Email',
      renderCell: (item) => item.userPrincipalName || '-',
    }),
    createTableColumn<User>({
      columnId: 'jobTitle',
      renderHeaderCell: () => 'Job Title',
      renderCell: (item) => item.jobTitle || '-',
    }),
    createTableColumn<User>({
      columnId: 'department',
      renderHeaderCell: () => 'Department',
      renderCell: (item) => item.department || '-',
    }),
    createTableColumn<User>({
      columnId: 'officeLocation',
      renderHeaderCell: () => 'Office',
      renderCell: (item) => item.officeLocation || '-',
    }),
    createTableColumn<User>({
      columnId: 'mail',
      renderHeaderCell: () => 'Mail',
      renderCell: (item) => item.mail || '-',
    }),
  ];

  return (
    <Card className={styles.container}>
      <CardHeader
        header={
          <div className={styles.header}>
            <Button
              appearance="subtle"
              icon={<ArrowLeftRegular />}
              onClick={onBack}
            >
              Back
            </Button>
            <Title3>Users</Title3>
            <Button
              appearance="subtle"
              icon={<PeopleRegular />}
              onClick={handleRefresh}
              disabled={loading}
            >
              Refresh
            </Button>
          </div>
        }
      />
      
      <div style={{ padding: '0 16px' }}>
        <SearchBox
          className={styles.searchBox}
          placeholder="Search users..."
          value={searchQuery}
          onChange={(_, data) => handleSearch(data.value)}
        />
      </div>

      {error && (
        <Text className={styles.errorText}>{error}</Text>
      )}

      <div className={styles.dataGrid}>
        <DataGrid
          items={users}
          columns={columns}
          sortable
          resizableColumns
          columnSizingOptions={{
            displayName: {
              minWidth: 150,
              defaultWidth: 200,
            },
            userPrincipalName: {
              minWidth: 200,
              defaultWidth: 250,
            },
            jobTitle: {
              minWidth: 150,
              defaultWidth: 200,
            },
            department: {
              minWidth: 150,
              defaultWidth: 200,
            },
            officeLocation: {
              minWidth: 100,
              defaultWidth: 150,
            },
            mail: {
              minWidth: 200,
              defaultWidth: 250,
            },
          }}
        >
          <DataGridHeader>
            <DataGridRow>
              {({ renderHeaderCell }) => (
                <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
              )}
            </DataGridRow>
          </DataGridHeader>
          <DataGridBody<User>>
            {({ item, rowId }) => (
              <DataGridRow<User> key={rowId}>
                {({ renderCell }) => (
                  <DataGridCell>{renderCell(item)}</DataGridCell>
                )}
              </DataGridRow>
            )}
          </DataGridBody>
        </DataGrid>
      </div>

      {hasMore && (
        <div className={styles.loadMoreContainer}>
          {loading ? (
            <Spinner size="small" label="Loading more users..." />
          ) : (
            <Button onClick={loadMore}>
              Load More Users
            </Button>
          )}
        </div>
      )}

      {loading && users.length === 0 && (
        <div className={styles.loadMoreContainer}>
          <Spinner label="Loading users..." />
        </div>
      )}
    </Card>
  );
};
