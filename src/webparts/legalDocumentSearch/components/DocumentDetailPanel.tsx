import * as React from 'react';
import { Link, Panel, PanelType, Spinner } from '@fluentui/react';
import styles from './LegalDocumentSearch.module.scss';
import { DocumentService } from '../services/DocumentService';
import { IDocumentDetails, IDocumentItem } from '../models/IDocumentItem';

export interface IDocumentDetailPanelProps {
  driveId: string;
  item: IDocumentItem | undefined;
  additionalColumns: string[];
  service: DocumentService;
  onDismiss: () => void;
}

interface IDetailValueProps {
  label: string;
  value?: string;
}

const formatDate = (value: string, includeTime: boolean = true): string => {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  const options: Intl.DateTimeFormatOptions = {
    day: 'numeric',
    month: 'numeric',
    year: 'numeric'
  };

  if (includeTime) {
    options.hour = 'numeric';
    options.minute = '2-digit';
  }

  return date.toLocaleString(undefined, options);
};

const formatSize = (value?: number): string => {
  if (!value) {
    return '';
  }

  if (value < 1024) {
    return `${value} B`;
  }

  const units: string[] = ['KB', 'MB', 'GB', 'TB'];
  let size: number = value / 1024;
  let unitIndex: number = 0;

  while (size >= 1024 && unitIndex < units.length - 1) {
    size = size / 1024;
    unitIndex++;
  }

  return `${size.toFixed(size >= 10 ? 0 : 1)} ${units[unitIndex]}`;
};

const DetailValue: React.FC<IDetailValueProps> = (props: IDetailValueProps) => {
  if (!props.value) {
    return null;
  }

  return (
    <div className={styles.detailField}>
      <span className={styles.detailLabel}>{props.label}</span>
      <span className={styles.detailValue}>{props.value}</span>
    </div>
  );
};

export const DocumentDetailPanel: React.FC<IDocumentDetailPanelProps> = (props: IDocumentDetailPanelProps) => {
  const [details, setDetails] = React.useState<IDocumentDetails | undefined>();
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [hasError, setHasError] = React.useState<boolean>(false);
  const item = props.item;

  React.useEffect(() => {
    if (!item || item.isFolder || !item.sourceId) {
      setDetails(undefined);
      setHasError(false);
      return;
    }

    setIsLoading(true);
    setHasError(false);

    props.service.getGraphDriveItemDetails(props.driveId, item.sourceId)
      .then((documentDetails: IDocumentDetails) => {
        setDetails(documentDetails);
        setIsLoading(false);
      })
      .catch(() => {
        setHasError(true);
        setIsLoading(false);
      });
  }, [item, props.driveId, props.service]);

  const detailFields = React.useMemo(() => {
    if (!details) {
      return [];
    }

    const priorityFields: string[] = ['GLPI_DocumentId', 'GLPI Document ID', ...props.additionalColumns];
    const fieldNames: string[] = Object.keys(details.fields);
    const selectedFields: string[] = priorityFields
      .filter((fieldName: string, index: number, array: string[]) => !!fieldName && array.indexOf(fieldName) === index)
      .filter((fieldName: string) => fieldNames.indexOf(fieldName) > -1);

    if (selectedFields.length > 0) {
      return selectedFields;
    }

    return fieldNames.slice(0, 8);
  }, [details, props.additionalColumns]);

  return (
    <Panel
      isOpen={!!item}
      onDismiss={props.onDismiss}
      type={PanelType.medium}
      headerText={item?.name || 'Detalles'}
      isLightDismiss={true}
      closeButtonAriaLabel="Cerrar"
    >
      {item && (
        <div className={styles.detailPanel}>
          <DetailValue label="Nombre" value={item.name} />
          <DetailValue label="Ruta" value={item.path} />
          <DetailValue label="Tipo" value={item.fileType} />
          <DetailValue label="Creado" value={formatDate(item.created)} />
          <DetailValue label="Modificado" value={formatDate(item.modified)} />
          <DetailValue label="Modificado por" value={item.modifiedBy} />
          <DetailValue label="Tamaño" value={formatSize(item.size)} />
          {item.url && (
            <div className={styles.detailField}>
              <span className={styles.detailLabel}>Documento</span>
              <Link href={item.url} target="_blank" rel="noreferrer">Abrir</Link>
            </div>
          )}
          {isLoading && <Spinner label="Cargando..." />}
          {hasError && <div className={styles.status}>No se pudo cargar la información.</div>}
          {details && detailFields.map((fieldName: string) => (
            <DetailValue key={fieldName} label={fieldName} value={details.fields[fieldName]} />
          ))}
        </div>
      )}
    </Panel>
  );
};
