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
import { Site } from '@microsoft/microsoft-graph-types';
import { useSharePointSites } from '../hooks/useSharePointSites';

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

interface SharePointSitesProps {
  onBack: () => void;
}

export const SharePointSites: React.FC<SharePointSitesProps> = ({ onBack }) => {
  const styles = useStyles();
  const [searchQuery, setSearchQuery] = useState('');
  const { sites, loading, error, hasMore, loadSites, loadMore, searchSites, clearSites } = useSharePointSites(100);

  useEffect(() => {
    loadSites();
  }, [loadSites]);

  const handleSearch = useCallback((value: string) => {
    setSearchQuery(value);
    // Debounce search
    const timeoutId = setTimeout(() => {
      searchSites(value);
    }, 300);
    return () => clearTimeout(timeoutId);
  }, [searchSites]);

  const handleRefresh = useCallback(() => {
    clearSites();
    loadSites();
  }, [clearSites, loadSites]);

  const columns: TableColumnDefinition<Site>[] = [
    createTableColumn<Site>({
      columnId: 'displayName',
      renderHeaderCell: () => 'Site Name',
      renderCell: (item) => item.displayName || 'Unnamed Site',
    }),
    createTableColumn<Site>({
      columnId: 'webUrl',
      renderHeaderCell: () => 'URL',
      renderCell: (item) => (
        <div className={styles.linkCell}>
        {item.webUrl ? (
          <Link href={item.webUrl} target="_blank" rel="noopener noreferrer">
            {item.webUrl}
          </Link>
        ) : (
          '-' /* or whatever placeholder you like */
        )}
        <OpenRegular />
      </div>
      ),
    }),
    createTableColumn<Site>({
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
            <Title3>SharePoint Sites</Title3>
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
          placeholder="Search sites..."
          value={searchQuery}
          onChange={(_, data) => handleSearch(data.value)}
        />
      </div>

      {error && (
        <Text className={styles.errorText}>{error}</Text>
      )}

      <div className={styles.dataGrid}>
        <DataGrid
          items={sites}
          columns={columns}
          sortable
          resizableColumns
          columnSizingOptions={{
            displayName: {
              minWidth: 150,
              defaultWidth: 250,
            },
            webUrl: {
              minWidth: 200,
              defaultWidth: 400,
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
          <DataGridBody<Site>>
            {({ item, rowId }) => (
              <DataGridRow<Site> key={rowId}>
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
            <Spinner size="small" label="Loading more sites..." />
          ) : (
            <Button onClick={loadMore}>
              Load More Sites
            </Button>
          )}
        </div>
      )}

      {loading && sites.length === 0 && (
        <div className={styles.loadMoreContainer}>
          <Spinner label="Loading SharePoint sites..." />
        </div>
      )}
    </Card>
  );
};
