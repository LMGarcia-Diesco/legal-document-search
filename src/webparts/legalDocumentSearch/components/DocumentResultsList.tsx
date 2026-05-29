import * as React from 'react';
import { Icon, Link } from '@fluentui/react';
import styles from './LegalDocumentSearch.module.scss';
import { IDocumentItem } from '../models/IDocumentItem';

export interface IDocumentResultsListProps {
  items: IDocumentItem[];
  additionalColumns: string[];
  onOpenFolder: (item: IDocumentItem) => void;
}

interface IMetaProps {
  label: string;
  value?: string;
}

const formatDate = (value: string): string => {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
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

const getFileIconName = (item: IDocumentItem): string => {
  if (item.isFolder) {
    return 'FabricFolder';
  }

  switch ((item.fileType || '').toLowerCase()) {
    case 'doc':
    case 'docx':
      return 'WordDocument';
    case 'xls':
    case 'xlsx':
    case 'csv':
      return 'ExcelDocument';
    case 'ppt':
    case 'pptx':
      return 'PowerPointDocument';
    case 'pdf':
      return 'PDF';
    case 'png':
    case 'jpg':
    case 'jpeg':
    case 'gif':
    case 'bmp':
    case 'svg':
    case 'webp':
      return 'Photo2';
    case 'zip':
    case 'rar':
    case '7z':
      return 'ZipFolder';
    case 'txt':
    case 'md':
      return 'TextDocument';
    default:
      return 'Page';
  }
};

const getIconClassName = (item: IDocumentItem): string => {
  if (item.isFolder) {
    return styles.folderIcon;
  }

  switch ((item.fileType || '').toLowerCase()) {
    case 'doc':
    case 'docx':
      return styles.wordIcon;
    case 'xls':
    case 'xlsx':
    case 'csv':
      return styles.excelIcon;
    case 'ppt':
    case 'pptx':
      return styles.powerPointIcon;
    case 'pdf':
      return styles.pdfIcon;
    case 'png':
    case 'jpg':
    case 'jpeg':
    case 'gif':
    case 'bmp':
    case 'svg':
    case 'webp':
      return styles.imageIcon;
    case 'zip':
    case 'rar':
    case '7z':
      return styles.archiveIcon;
    default:
      return styles.fileIcon;
  }
};

const Meta: React.FC<IMetaProps> = (props: IMetaProps) => {
  return (
    <span className={styles.metaItem}>
      <span className={styles.metaLabel}>{props.label}</span>
      <span className={styles.metaValue}>{props.value || ''}</span>
    </span>
  );
};

export const DocumentResultsList: React.FC<IDocumentResultsListProps> = (props: IDocumentResultsListProps) => {
  return (
    <div className={styles.resultList}>
      {props.items.map((item: IDocumentItem) => (
        <article key={`${item.path}|${item.name}|${item.id}`} className={styles.resultRow}>
          <button
            type="button"
            className={styles.itemMain}
            onClick={() => item.isFolder ? props.onOpenFolder(item) : undefined}
            disabled={!item.isFolder}
          >
            <span className={getIconClassName(item)}>
              <Icon iconName={getFileIconName(item)} />
            </span>
            <span className={styles.itemText}>
              <span className={styles.itemName}>{item.name}</span>
              <span className={styles.itemPath}>{item.path}</span>
            </span>
          </button>
          <div className={styles.itemMeta}>
            <Meta label="Modificado" value={formatDate(item.modified)} />
            <Meta label="Por" value={item.modifiedBy} />
            <Meta label="Tipo" value={item.fileType} />
            <Meta label="Creado" value={formatDate(item.created)} />
            <Meta label="Tamaño" value={formatSize(item.size)} />
            {props.additionalColumns.map((column: string) => (
              <Meta key={column} label={column} value={item.additionalValues[column]} />
            ))}
          </div>
          <div className={styles.itemAction}>
            {!item.isFolder && item.url && (
              <Link href={item.url} target="_blank" rel="noreferrer">
                Abrir
              </Link>
            )}
          </div>
        </article>
      ))}
    </div>
  );
};
