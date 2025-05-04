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
  BoxRegular,
} from '@fluentui/react-icons';
import { Drive } from '@microsoft/microsoft-graph-types';
import { useDrives } from '../hooks/useDrives';

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

interface DrivesProps {
  onBack: () => void;
}

export const Drives: React.FC<DrivesProps> = ({ onBack }) => {
  const styles = useStyles();
  const [searchQuery, setSearchQuery] = useState('');
  const { drives, loading, error, hasMore, loadDrives, loadMore, searchDrives, clearDrives } = useDrives(100);

  useEffect(() => {
    loadDrives();
  }, [loadDrives]);

  const handleSearch = useCallback((value: string) => {
    setSearchQuery(value);
    // Debounce search
    const timeoutId = setTimeout(() => {
      searchDrives(value);
    }, 300);
    return () => clearTimeout(timeoutId);
  }, [searchDrives]);

  const handleRefresh = useCallback(() => {
    clearDrives();
    loadDrives();
  }, [clearDrives, loadDrives]);

  const columns: TableColumnDefinition<Drive>[] = [
    createTableColumn<Drive>({
      columnId: 'name',
      renderHeaderCell: () => 'Drive Name',
      renderCell: (item) => item.name || 'Unnamed Drive',
    }),
    createTableColumn<Drive>({
      columnId: 'driveType',
      renderHeaderCell: () => 'Type',
      renderCell: (item) => item.driveType || '-',
    }),
    createTableColumn<Drive>({
      columnId: 'webUrl',
      renderHeaderCell: () => 'URL',
      renderCell: (item) => (
        <div className={styles.linkCell}>
          {item.webUrl ? (
            <Link href={item.webUrl} target="_blank" rel="noopener noreferrer">
              {item.webUrl}
            </Link>
          ) : (
            '-'
          )}
          <OpenRegular />
        </div>
      ),
    }),
    createTableColumn<Drive>({
      columnId: 'owner',
      renderHeaderCell: () => 'Owner',
      renderCell: (item) => {
        // Handle the IdentitySet structure in the Graph API response
        if (item.owner) {
          // Check for user identity
          if (item.owner.user?.displayName) {
            return item.owner.user.displayName;
          }
          // Check for group identity
          if (item.owner.group?.displayName) {
            return `Group: ${item.owner.group.displayName}`;
          }
          // Check for application identity
          if (item.owner.application?.displayName) {
            return `App: ${item.owner.application.displayName}`;
          }
          // Check for device identity
          if (item.owner.device?.displayName) {
            return `Device: ${item.owner.device.displayName}`;
          }
        }
        return '-';
      },
    }),
    createTableColumn<Drive>({
      columnId: 'createdDateTime',
      renderHeaderCell: () => 'Created Date',
      renderCell: (item) => 
        item.createdDateTime ? new Date(item.createdDateTime).toLocaleDateString() : '-',
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
            <Title3>Drives</Title3>
            <Button
              appearance="subtle"
              icon={<BoxRegular />}
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
          placeholder="Search drives..."
          value={searchQuery}
          onChange={(_, data) => handleSearch(data.value)}
        />
      </div>

      {error && (
        <Text className={styles.errorText}>{error}</Text>
      )}

      <div className={styles.dataGrid}>
        <DataGrid
          items={drives}
          columns={columns}
          sortable
          resizableColumns
          columnSizingOptions={{
            name: {
              minWidth: 150,
              defaultWidth: 250,
            },
            driveType: {
              minWidth: 100,
              defaultWidth: 120,
            },
            webUrl: {
              minWidth: 200,
              defaultWidth: 350,
            },
            owner: {
              minWidth: 150,
              defaultWidth: 200,
            },
            createdDateTime: {
              minWidth: 120,
              defaultWidth: 150,
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
          <DataGridBody<Drive>>
            {({ item, rowId }) => (
              <DataGridRow<Drive> key={rowId}>
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
            <Spinner size="small" label="Loading more drives..." />
          ) : (
            <Button onClick={loadMore}>
              Load More Drives
            </Button>
          )}
        </div>
      )}

      {loading && drives.length === 0 && (
        <div className={styles.loadMoreContainer}>
          <Spinner label="Loading drives..." />
        </div>
      )}
    </Card>
  );
};
