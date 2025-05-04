import React, { useEffect, useState, useCallback } from 'react';
import {
  makeStyles,
  tokens,
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
} from '@fluentui/react-components';
import { 
  ArrowLeftRegular,
  BoxMultipleRegular,
  CopyRegular,
} from '@fluentui/react-icons';
import { useMsal } from '@azure/msal-react';
import { GraphService } from '../services/graphService';

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
  dataGrid: {
    flex: 1 as unknown as never,
    overflow: 'auto' as unknown as never,
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1 as unknown as never,
    padding: tokens.spacingVerticalM as unknown as never,
  },
  copyButton: {
    minWidth: 'auto' as unknown as never,
  },
});

interface ContainerTypesProps {
  onBack: () => void;
  onSelectContainerType?: (containerTypeId: string) => void;
}

interface ContainerType {
  id: string;
  displayName?: string;
  description?: string;
}

export const ContainerTypes: React.FC<ContainerTypesProps> = ({ onBack, onSelectContainerType }) => {
  const styles = useStyles();
  const { instance } = useMsal();
  const [containerTypes, setContainerTypes] = useState<ContainerType[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);

  const loadContainerTypes = useCallback(async () => {
    setLoading(true);
    setError(null);

    try {
      const graphService = new GraphService(instance);
      const result = await graphService.getContainerTypes();
      setContainerTypes(result.value);
    } catch (err: any) {
      setError(err?.message || 'Failed to load container types');
      setContainerTypes([]);
    } finally {
      setLoading(false);
    }
  }, [instance]);

  useEffect(() => {
    loadContainerTypes();
  }, [loadContainerTypes]);

  const handleCopyId = useCallback((id: string) => {
    navigator.clipboard.writeText(id).then(() => {
      setCopiedId(id);
      setTimeout(() => setCopiedId(null), 2000);
    });
  }, []);

  const handleSelectContainerType = useCallback((id: string) => {
    if (onSelectContainerType) {
      onSelectContainerType(id);
      onBack();
    }
  }, [onSelectContainerType, onBack]);

  const columns: TableColumnDefinition<ContainerType>[] = [
    createTableColumn<ContainerType>({
      columnId: 'id',
      renderHeaderCell: () => 'Container Type ID (GUID)',
      renderCell: (item) => (
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span>{item.id}</span>
          <Button
            appearance="subtle"
            icon={<CopyRegular />}
            onClick={() => handleCopyId(item.id)}
            className={styles.copyButton}
            title="Copy ID"
          />
          {copiedId === item.id && <Text size={200} style={{ color: 'green' }}>Copied!</Text>}
        </div>
      ),
    }),
    createTableColumn<ContainerType>({
      columnId: 'properties',
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
        }}>
          {JSON.stringify(item, null, 2)}
        </pre>
      ),
    }),
    createTableColumn<ContainerType>({
      columnId: 'actions',
      renderHeaderCell: () => 'Actions',
      renderCell: (item) => (
        <Button
          appearance="primary"
          size="small"
          onClick={() => handleSelectContainerType(item.id)}
          disabled={!onSelectContainerType}
        >
          Use This Type
        </Button>
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
            <Title3>Container Types</Title3>
            <Button
              appearance="subtle"
              icon={<BoxMultipleRegular />}
              onClick={loadContainerTypes}
              disabled={loading}
            >
              Refresh
            </Button>
          </div>
        }
      />

      {error && (
        <Text className={styles.errorText}>{error}</Text>
      )}

      <div className={styles.dataGrid}>
        <DataGrid
          items={containerTypes}
          columns={columns}
          sortable
          resizableColumns
          columnSizingOptions={{
            id: {
              minWidth: 300,
              defaultWidth: 350,
            },
            properties: {
              minWidth: 300,
              defaultWidth: 400,
            },
            actions: {
              minWidth: 150,
              defaultWidth: 200,
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
          <DataGridBody<ContainerType>>
            {({ item, rowId }) => (
              <DataGridRow<ContainerType> key={rowId}>
                {({ renderCell }) => (
                  <DataGridCell>{renderCell(item)}</DataGridCell>
                )}
              </DataGridRow>
            )}
          </DataGridBody>
        </DataGrid>
      </div>

      {loading && containerTypes.length === 0 && (
        <div style={{ display: 'flex', justifyContent: 'center', padding: tokens.spacingVerticalL }}>
          <Spinner label="Loading container types..." />
        </div>
      )}
    </Card>
  );
};
