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
  Text,
  Input,
  Label,
} from '@fluentui/react-components';
import { 
  ArrowLeftRegular,
  BoxMultipleRegular,
  CutRegular,
  BoxSearchRegular,
} from '@fluentui/react-icons';
import { SharePointContainer } from '../types/microsoft-graph-extended';
import { useSharePointContainers } from '../hooks/useSharePointContainers';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from '../authConfig';
import { ContainerTypes } from './ContainerTypes';

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

interface SharePointContainersProps {
  onBack: () => void;
}

export const SharePointContainers: React.FC<SharePointContainersProps> = ({ onBack }) => {
  const styles = useStyles();
  const { instance } = useMsal();
  const [searchQuery, setSearchQuery] = useState('');
  const [inputContainerTypeId, setInputContainerTypeId] = useState('');
  const [showContainerTypes, setShowContainerTypes] = useState(false);
  const { 
    containers, 
    loading, 
    error, 
    hasMore,
    containerTypeId,
    setContainerTypeId, 
    loadContainers, 
    loadMore, 
    searchContainers, 
    clearContainers 
  } = useSharePointContainers(100);

  useEffect(() => {
    if (containerTypeId) {
      loadContainers();
    }
  }, [loadContainers, containerTypeId]);

  const handleSearch = useCallback((value: string) => {
    setSearchQuery(value);
    // Debounce search
    const timeoutId = setTimeout(() => {
      searchContainers(value);
    }, 300);
    return () => clearTimeout(timeoutId);
  }, [searchContainers]);

  const handleRefresh = useCallback(() => {
    clearContainers();
    if (containerTypeId) {
      loadContainers();
    }
  }, [clearContainers, loadContainers, containerTypeId]);

  const handleReConsent = useCallback(() => {
    instance.loginPopup({
      ...loginRequest,
      prompt: "consent"
    }).then(() => {
      // After re-consent, refresh the containers
      handleRefresh();
    }).catch((error) => {
      console.error("Re-consent failed:", error);
    });
  }, [instance, handleRefresh]);

  const handleSetContainerTypeId = useCallback(() => {
    if (inputContainerTypeId.trim()) {
      setContainerTypeId(inputContainerTypeId.trim());
    }
  }, [inputContainerTypeId, setContainerTypeId]);

  const handleSelectContainerType = useCallback((containerTypeId: string) => {
    setInputContainerTypeId(containerTypeId);
    setContainerTypeId(containerTypeId);
    setShowContainerTypes(false);
  }, [setContainerTypeId]);

  if (showContainerTypes) {
    return (
      <ContainerTypes
        onBack={() => setShowContainerTypes(false)}
        onSelectContainerType={handleSelectContainerType}
      />
    );
  }

  const columns: TableColumnDefinition<SharePointContainer>[] = [
    createTableColumn<SharePointContainer>({
      columnId: 'id',
      renderHeaderCell: () => 'Container ID',
      renderCell: (item) => item.id || 'Unknown ID',
    }),
    createTableColumn<SharePointContainer>({
      columnId: 'containerTypeId',
      renderHeaderCell: () => 'Container Type ID',
      renderCell: (item) => item.containerTypeId || '-',
    }),
    createTableColumn<SharePointContainer>({
      columnId: 'allProperties',
      renderHeaderCell: () => 'All Properties',
      renderCell: (item) => (
        <pre style={{ 
          fontSize: '12px', 
          maxHeight: '100px', 
          overflow: 'auto',
          margin: 0,
          padding: '4px',
          backgroundColor: 'var(--colorNeutralBackground2)',
          borderRadius: '4px',
          width: '400px',
        }}>
          {JSON.stringify(item, null, 2)}
        </pre>
      ),
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
            <Title3>SharePoint Containers</Title3>
            <Button
              appearance="subtle"
              icon={<BoxMultipleRegular />}
              onClick={handleRefresh}
              disabled={loading}
            >
              Refresh
            </Button>
          </div>
        }
      />
      
      <div style={{ padding: '0 16px', marginBottom: '16px' }}>
        <Label htmlFor="containerTypeId">Container Type ID (GUID)</Label>
        <div style={{ display: 'flex', gap: '8px', marginTop: '4px' }}>
          <Input
            id="containerTypeId"
            value={inputContainerTypeId}
            onChange={(_, data) => setInputContainerTypeId(data.value)}
            placeholder="Enter container type GUID (e.g., xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)"
            style={{ flex: 1 }}
          />
          <Button
            appearance="primary"
            onClick={handleSetContainerTypeId}
            disabled={!inputContainerTypeId.trim() || loading}
          >
            Load Containers
          </Button>
          <Button
            appearance="secondary"
            icon={<BoxSearchRegular />}
            onClick={() => setShowContainerTypes(true)}
          >
            Browse Types
          </Button>
        </div>
        {!containerTypeId && (
          <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginTop: '4px' }}>
            Enter a container type ID in GUID format or click "Browse Types" to see available container types.
          </Text>
        )}
      </div>
      
      <div style={{ padding: '0 16px' }}>
        <SearchBox
          className={styles.searchBox}
          placeholder="Search containers..."
          value={searchQuery}
          onChange={(_, data) => handleSearch(data.value)}
          disabled={!containerTypeId}
        />
      </div>

      {containerTypeId && error && (
        <div style={{ padding: '0 16px' }}>
          <Text className={styles.errorText}>{error}</Text>
          {error.includes('permissions are required') && (
            <Button
              appearance="primary"
              icon={<CutRegular />}
              onClick={handleReConsent}
              style={{ marginTop: '8px' }}
            >
              Grant Permission
            </Button>
          )}
        </div>
      )}

      <div className={styles.dataGrid}>
        <DataGrid
          items={containers}
          columns={columns}
          sortable
          resizableColumns
          columnSizingOptions={{
            id: {
              minWidth: 250,
              defaultWidth: 300,
            },
            containerTypeId: {
              minWidth: 300,
              defaultWidth: 350,
            },
            allProperties: {
              minWidth: 400,
              defaultWidth: 500,
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
          <DataGridBody<SharePointContainer>>
            {({ item, rowId }) => (
              <DataGridRow<SharePointContainer> key={rowId}>
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
            <Spinner size="small" label="Loading more containers..." />
          ) : (
            <Button onClick={loadMore}>
              Load More Containers
            </Button>
          )}
        </div>
      )}

      {loading && containers.length === 0 && (
        <div className={styles.loadMoreContainer}>
          <Spinner label="Loading SharePoint containers..." />
        </div>
      )}
    </Card>
  );
};
