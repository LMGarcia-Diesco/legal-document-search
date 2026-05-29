import * as React from 'react';
import { DocumentService, SearchProvider } from '../services/DocumentService';
import { IDocumentItem } from '../models/IDocumentItem';
import { ISearchFilters } from '../models/ISearchFilters';

export interface IUseInfiniteDocumentsOptions {
  service: DocumentService;
  siteUrl: string;
  driveId: string;
  libraryUrl: string;
  folderPath: string;
  additionalColumns: string[];
  pageSize: number;
  searchProvider: string;
  filters: ISearchFilters;
}

export interface IUseInfiniteDocumentsResult {
  items: IDocumentItem[];
  isLoading: boolean;
  error: boolean;
  hasMore: boolean;
  loadMore: () => void;
}

export const useInfiniteDocuments = (options: IUseInfiniteDocumentsOptions): IUseInfiniteDocumentsResult => {
  const [items, setItems] = React.useState<IDocumentItem[]>([]);
  const [nextStartRow, setNextStartRow] = React.useState<number | undefined>();
  const [nextLink, setNextLink] = React.useState<string | undefined>();
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<boolean>(false);
  const [hasMore, setHasMore] = React.useState<boolean>(true);
  const requestIdRef = React.useRef<number>(0);
  const loadingRef = React.useRef<boolean>(false);

  const loadPage = React.useCallback(async (reset: boolean): Promise<void> => {
    if (loadingRef.current) {
      return;
    }

    loadingRef.current = true;
    setIsLoading(true);
    setError(false);

    const requestId: number = reset ? requestIdRef.current + 1 : requestIdRef.current;
    if (reset) {
      requestIdRef.current = requestId;
    }

    try {
      const page = await options.service.getDocuments({
        siteUrl: options.siteUrl,
        driveId: options.driveId,
        libraryUrl: options.libraryUrl,
        folderPath: options.folderPath,
        additionalColumns: options.additionalColumns,
        pageSize: options.pageSize,
        searchProvider: options.searchProvider as SearchProvider,
        filters: options.filters,
        startRow: reset ? 0 : nextStartRow,
        nextLink: reset ? undefined : nextLink
      });

      if (requestId !== requestIdRef.current) {
        return;
      }

      setItems((currentItems: IDocumentItem[]) => reset ? page.items : [...currentItems, ...page.items]);
      setNextStartRow(page.nextStartRow);
      setNextLink(page.nextLink);
      setHasMore(page.nextStartRow !== undefined || page.nextLink !== undefined);
    } catch (requestError) {
      if (requestId === requestIdRef.current) {
        // Keep the failing query from being retried continuously by the infinite-scroll sentinel.
        console.error(requestError);
        setError(true);
        setHasMore(false);
      }
    } finally {
      if (requestId === requestIdRef.current) {
        // The ref is intentionally updated after the request completes to prevent duplicate page loads.
        // eslint-disable-next-line require-atomic-updates
        loadingRef.current = false;
        setIsLoading(false);
      }
    }
  }, [
    nextStartRow,
    nextLink,
    options.additionalColumns,
    options.driveId,
    options.filters,
    options.folderPath,
    options.libraryUrl,
    options.pageSize,
    options.searchProvider,
    options.service,
    options.siteUrl
  ]);

  React.useEffect(() => {
    setItems([]);
    setNextStartRow(undefined);
    setNextLink(undefined);
    setHasMore(true);
    loadingRef.current = false;
    loadPage(true).catch(() => undefined);
  }, [
    options.additionalColumns.join('|'),
    options.driveId,
    options.filters.searchText,
    options.folderPath,
    options.libraryUrl,
    options.pageSize,
    options.searchProvider,
    options.siteUrl
  ]);

  const loadMore = React.useCallback((): void => {
    if (!isLoading && hasMore) {
      loadPage(false).catch(() => undefined);
    }
  }, [hasMore, isLoading, loadPage]);

  return {
    items,
    isLoading,
    error,
    hasMore,
    loadMore
  };
};
