import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './LegalDocumentSearch.module.scss';
import type { ILegalDocumentSearchProps } from './ILegalDocumentSearchProps';
import { DocumentService } from '../services/DocumentService';
import { useInfiniteDocuments } from '../hooks/useInfiniteDocuments';
import { SearchBar } from './SearchBar';
import { SearchFilters } from './SearchFilters';
import { DocumentResultsList } from './DocumentResultsList';
import { DocumentDetailPanel } from './DocumentDetailPanel';
import { IDocumentItem } from '../models/IDocumentItem';

const parseAdditionalColumns = (value: string): string[] => {
  return value
    .split(',')
    .map((column: string) => column.trim())
    .filter((column: string) => !!column);
};

interface IFolderBreadcrumbProps {
  folderPath: string;
  onNavigate: (folderPath: string) => void;
}

interface IFolderPanelProps {
  currentFolderPath: string;
  folders: IDocumentItem[];
  searchText: string;
  onNavigate: (folderPath: string) => void;
  onOpenFolder: (item: IDocumentItem) => void;
}

const getNextFolderPath = (currentFolderPath: string, folderName: string): string => {
  const currentPath = currentFolderPath.trim().replace(/^\/+|\/+$/g, '');
  return currentPath ? `${currentPath}/${folderName}` : folderName;
};

const FolderBreadcrumb: React.FC<IFolderBreadcrumbProps> = (props: IFolderBreadcrumbProps) => {
  const segments = props.folderPath
    .split('/')
    .map((segment: string) => segment.trim())
    .filter((segment: string) => !!segment);

  return (
    <nav className={styles.breadcrumb}>
      <button type="button" className={styles.breadcrumbButton} onClick={() => props.onNavigate('')}>
        Documentos
      </button>
      {segments.map((segment: string, index: number) => {
        const path = segments.slice(0, index + 1).join('/');

        return (
          <React.Fragment key={path}>
            <span className={styles.breadcrumbSeparator}>/</span>
            <button type="button" className={styles.breadcrumbButton} onClick={() => props.onNavigate(path)}>
              {segment}
            </button>
          </React.Fragment>
        );
      })}
    </nav>
  );
};

const FolderPanel: React.FC<IFolderPanelProps> = (props: IFolderPanelProps) => {
  const segments = props.currentFolderPath
    .split('/')
    .map((segment: string) => segment.trim())
    .filter((segment: string) => !!segment);
  const showCurrentFolders = !props.searchText;

  return (
    <aside className={styles.folderPanel}>
      <div className={styles.folderPanelTitle}>Carpetas</div>
      <button
        type="button"
        className={`${styles.folderNode} ${!props.currentFolderPath ? styles.folderNodeActive : ''}`}
        onClick={() => props.onNavigate('')}
      >
        <Icon iconName="FabricFolder" />
        <span>Documentos</span>
      </button>
      {segments.map((segment: string, index: number) => {
        const path = segments.slice(0, index + 1).join('/');
        const level = Math.min(index + 1, 5);

        return (
          <button
            type="button"
            key={path}
            className={`${styles.folderNode} ${index === segments.length - 1 ? styles.folderNodeActive : ''}`}
            style={{ paddingLeft: `${12 + level * 14}px` }}
            onClick={() => props.onNavigate(path)}
          >
            <Icon iconName="ChevronRight" />
            <Icon iconName="FabricFolder" />
            <span>{segment}</span>
          </button>
        );
      })}
      {showCurrentFolders && props.folders.map((folder: IDocumentItem) => (
        <button
          type="button"
          key={`${folder.path}|${folder.name}|${folder.id}`}
          className={styles.folderNode}
          style={{ paddingLeft: `${12 + Math.min(segments.length + 1, 5) * 14}px` }}
          onClick={() => props.onOpenFolder(folder)}
        >
          <Icon iconName="ChevronRight" />
          <Icon iconName="FabricFolder" />
          <span>{folder.name}</span>
        </button>
      ))}
    </aside>
  );
};

const LegalDocumentSearch: React.FC<ILegalDocumentSearchProps> = (props: ILegalDocumentSearchProps) => {
  const [searchText, setSearchText] = React.useState<string>('');
  const [debouncedSearchText, setDebouncedSearchText] = React.useState<string>('');
  const [currentFolderPath, setCurrentFolderPath] = React.useState<string>(props.folderPath || '');
  const [selectedItem, setSelectedItem] = React.useState<IDocumentItem | undefined>();
  const scrollContainerRef = React.useRef<HTMLDivElement | null>(null);
  const sentinelRef = React.useRef<HTMLDivElement | null>(null);
  const service = React.useMemo(
    () => new DocumentService(props.spHttpClient, props.msGraphClientFactory),
    [props.msGraphClientFactory, props.spHttpClient]
  );
  const additionalColumns = React.useMemo(() => parseAdditionalColumns(props.additionalColumns), [props.additionalColumns]);
  const siteUrl = props.siteUrl.trim() || props.currentWebUrl;

  React.useEffect(() => {
    const timeoutId: number = window.setTimeout(() => setDebouncedSearchText(searchText), 450);
    return () => window.clearTimeout(timeoutId);
  }, [searchText]);

  const {
    items,
    isLoading,
    error,
    hasMore,
    loadMore
  } = useInfiniteDocuments({
    service,
    siteUrl,
    driveId: props.driveId,
    libraryUrl: props.libraryUrl,
    folderPath: currentFolderPath,
    additionalColumns,
    pageSize: props.pageSize,
    searchProvider: props.searchProvider,
    filters: {
      searchText: debouncedSearchText
    }
  });

  React.useEffect(() => {
    setCurrentFolderPath(props.folderPath || '');
  }, [props.folderPath]);

  React.useEffect(() => {
    const sentinel = sentinelRef.current;
    const container = scrollContainerRef.current;

    if (!sentinel || !container) {
      return;
    }

    const observer = new IntersectionObserver((entries: IntersectionObserverEntry[]) => {
      if (entries[0].isIntersecting && hasMore && !isLoading) {
        loadMore();
      }
    }, {
      root: container,
      rootMargin: '160px'
    });

    observer.observe(sentinel);

    return () => observer.disconnect();
  }, [hasMore, isLoading, loadMore]);

  const handleOpenFolder = React.useCallback((item: IDocumentItem): void => {
    if (!item.isFolder || debouncedSearchText) {
      return;
    }

    setCurrentFolderPath(getNextFolderPath(currentFolderPath, item.name));
  }, [currentFolderPath, debouncedSearchText]);

  const handleNavigateFolder = React.useCallback((folderPath: string): void => {
    setCurrentFolderPath(folderPath);
  }, []);

  const handleOpenDetails = React.useCallback((item: IDocumentItem): void => {
    setSelectedItem(item);
  }, []);

  const handleDismissDetails = React.useCallback((): void => {
    setSelectedItem(undefined);
  }, []);

  const currentFolders = React.useMemo(() => {
    return items.filter((item: IDocumentItem) => item.isFolder);
  }, [items]);

  return (
    <section className={styles.legalDocumentSearch}>
      <div className={styles.header}>
        <h2 className={styles.title}>Búsqueda documental</h2>
        <SearchBar value={searchText} onChange={setSearchText} />
      </div>
      <SearchFilters />
      <FolderBreadcrumb folderPath={currentFolderPath} onNavigate={handleNavigateFolder} />
      {error && <div className={styles.status}>No se pudo cargar la información.</div>}
      <div className={styles.workspace}>
        <FolderPanel
          currentFolderPath={currentFolderPath}
          folders={currentFolders}
          searchText={debouncedSearchText}
          onNavigate={handleNavigateFolder}
          onOpenFolder={handleOpenFolder}
        />
        <div className={styles.results} ref={scrollContainerRef}>
          <DocumentResultsList
            items={items}
            additionalColumns={additionalColumns}
            onOpenFolder={handleOpenFolder}
            onOpenDetails={handleOpenDetails}
          />
          {!isLoading && !error && items.length === 0 && <div className={styles.status}>No se encontraron documentos.</div>}
          {isLoading && <div className={styles.status}>Cargando...</div>}
          <div ref={sentinelRef} className={styles.sentinel} />
        </div>
      </div>
      <DocumentDetailPanel
        driveId={props.driveId}
        item={selectedItem}
        additionalColumns={additionalColumns}
        service={service}
        onDismiss={handleDismissDetails}
      />
    </section>
  );
};

export default LegalDocumentSearch;
