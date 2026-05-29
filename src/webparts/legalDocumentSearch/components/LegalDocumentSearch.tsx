import * as React from 'react';
import styles from './LegalDocumentSearch.module.scss';
import type { ILegalDocumentSearchProps } from './ILegalDocumentSearchProps';
import { DocumentService } from '../services/DocumentService';
import { useInfiniteDocuments } from '../hooks/useInfiniteDocuments';
import { SearchBar } from './SearchBar';
import { SearchFilters } from './SearchFilters';
import { DocumentResultsList } from './DocumentResultsList';

const parseAdditionalColumns = (value: string): string[] => {
  return value
    .split(',')
    .map((column: string) => column.trim())
    .filter((column: string) => !!column);
};

const LegalDocumentSearch: React.FC<ILegalDocumentSearchProps> = (props: ILegalDocumentSearchProps) => {
  const [searchText, setSearchText] = React.useState<string>('');
  const [debouncedSearchText, setDebouncedSearchText] = React.useState<string>('');
  const scrollContainerRef = React.useRef<HTMLDivElement | null>(null);
  const sentinelRef = React.useRef<HTMLDivElement | null>(null);
  const service = React.useMemo(
    () => new DocumentService(props.spHttpClient, props.msGraphClientFactory),
    [props.msGraphClientFactory, props.spHttpClient]
  );
  const additionalColumns = React.useMemo(() => parseAdditionalColumns(props.additionalColumns), [props.additionalColumns]);
  const siteUrl = props.siteUrl.trim() || props.currentWebUrl;

  React.useEffect(() => {
    const timeoutId: number = window.setTimeout(() => setDebouncedSearchText(searchText), 350);
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
    folderPath: props.folderPath,
    additionalColumns,
    pageSize: props.pageSize,
    searchProvider: props.searchProvider,
    filters: {
      searchText: debouncedSearchText
    }
  });

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

  return (
    <section className={styles.legalDocumentSearch}>
      <h2 className={styles.title}>Búsqueda documental</h2>
      <div className={styles.controls}>
        <SearchBar value={searchText} onChange={setSearchText} />
        <SearchFilters />
      </div>
      {error && <div className={styles.status}>No se pudo cargar la información.</div>}
      <div className={styles.results} ref={scrollContainerRef}>
        <DocumentResultsList items={items} additionalColumns={additionalColumns} />
        {!isLoading && !error && items.length === 0 && <div className={styles.status}>No se encontraron documentos.</div>}
        {isLoading && <div className={styles.status}>Cargando...</div>}
        <div ref={sentinelRef} className={styles.sentinel} />
      </div>
    </section>
  );
};

export default LegalDocumentSearch;
