import { MSGraphClientFactory, MSGraphClientV3, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IDocumentItem, IDocumentPage } from '../models/IDocumentItem';
import { ISearchFilters } from '../models/ISearchFilters';

export type SearchProvider = 'graphDrive' | 'sharePointSearch';

export interface IDocumentQuery {
  /**
   * Motor principal recomendado.
   * graphDrive: usa directamente el Drive ID de Microsoft Graph.
   * sharePointSearch: usa SharePoint Search KQL con Path:{libraryUrl}.
   */
  searchProvider: SearchProvider;

  /**
   * Microsoft Graph Drive ID.
   * Ejemplo:
   * b!JfpcK5VR6UeaMO4R50UtnD27UAJqcP5LpFat3NJzH2_40IFm6suUQaTYJcC9uuRy
   *
   * Este ID representa directamente el document library.
   */
  driveId?: string;

  /**
   * URL del sitio SharePoint.
   * Se usa principalmente para llamar /_api/search/query en el motor KQL.
   * Si no viene, se intenta derivar desde libraryUrl.
   */
  siteUrl?: string;

  /**
   * URL real de la biblioteca documental para KQL.
   * Ejemplo:
   * https://diesco.sharepoint.com/sites/prueba/Documentos%20compartidos
   *
   * KQL NO usa driveId. KQL usa Path:"{libraryUrl}".
   */
  libraryUrl?: string;

  /**
   * Compatibilidad con la configuración vieja.
   * Ya no se usa para Graph como campo obligatorio.
   */
  libraryName?: string;

  /**
   * Carpeta opcional dentro del document library.
   * Para Graph: relativa al root del drive.
   * Para KQL: relativa a libraryUrl.
   *
   * Si quieres buscar/listar toda la biblioteca, déjalo vacío.
   */
  folderPath?: string;

  /**
   * Compatibilidad con la propiedad vieja.
   * Preferir folderPath en adelante.
   */
  baseFolderPath?: string;

  additionalColumns: string[];
  pageSize: number;
  filters: ISearchFilters;

  /**
   * SharePoint Search KQL usa startRow.
   */
  startRow?: number;

  /**
   * Microsoft Graph Drive usa @odata.nextLink.
   */
  nextLink?: string;
}

interface ISearchCell {
  Key?: string;
  Value?: string;
}

interface ISearchRow {
  Cells?: ISearchCell[] | {
    results?: ISearchCell[];
  };
}

interface ISearchResponse {
  PrimaryQueryResult?: {
    RelevantResults?: {
      RowCount?: number;
      TotalRows?: number;
      Table?: {
        Rows?: ISearchRow[] | {
          results?: ISearchRow[];
        };
      };
    };
  };
  d?: {
    query?: ISearchResponse;
  };
}

interface IGraphDriveItem {
  id?: string;
  name?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  size?: number;
  file?: {
    mimeType?: string;
  };
  folder?: {
    childCount?: number;
  };
  lastModifiedBy?: {
    user?: {
      id?: string;
      email?: string;
      displayName?: string;
    };
    application?: {
      id?: string;
      displayName?: string;
    };
  };
  parentReference?: {
    driveType?: string;
    driveId?: string;
    id?: string;
    name?: string;
    path?: string;
    siteId?: string;
  };
  [key: string]: unknown;
}

interface IGraphCollectionResponse<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

export class DocumentService {
  public constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly msGraphClientFactory: MSGraphClientFactory
  ) {}

  public async getDocuments(query: IDocumentQuery): Promise<IDocumentPage> {
    if (query.searchProvider === 'graphDrive') {
      return this.searchWithGraphDrive(query);
    }

    return this.searchWithSharePoint(query);
  }

  /**
   * MOTOR 1: Microsoft Graph Drive
   *
   * Usa directamente:
   * /drives/{driveId}/root/children
   * /drives/{driveId}/root/search(q='texto')
   *
   * No resuelve siteId.
   * No lista drives.
   * No usa KQL.
   * No usa /search/query global.
   */
  private async searchWithGraphDrive(query: IDocumentQuery): Promise<IDocumentPage> {
    const graphClient: MSGraphClientV3 = await this.msGraphClientFactory.getClient('3');

    const response: IGraphCollectionResponse<IGraphDriveItem> = query.nextLink
      ? await graphClient.api(this.getGraphApiPath(query.nextLink)).get()
      : await this.getGraphDrivePage(graphClient, query);

    const hasSearchText: boolean = this.hasSearchText(query);

    const items: IDocumentItem[] = (response.value || [])
      .filter((item: IGraphDriveItem) => {
        /**
         * Con búsqueda, mostramos solo archivos.
         * Sin búsqueda, permitimos archivos y carpetas para navegación.
         */
        return hasSearchText ? !!item.file : (!!item.file || !!item.folder);
      })
      .map((item: IGraphDriveItem, index: number) => this.mapGraphDriveItem(item, query.additionalColumns, index));

    return {
      items,
      nextLink: response['@odata.nextLink']
    };
  }

  private async getGraphDrivePage(
    graphClient: MSGraphClientV3,
    query: IDocumentQuery
  ): Promise<IGraphCollectionResponse<IGraphDriveItem>> {
    const driveId: string = (query.driveId || '').trim();

    if (!driveId) {
      throw new Error('Debe configurar el Drive ID.');
    }

    const searchText: string = this.getSearchText(query);
    const folderPath: string = this.getEffectiveFolderPath(query);

    const endpoint: string = searchText
      ? `/drives/${driveId}/root/search(q='${this.escapeGraphSearchText(searchText)}')`
      : this.buildGraphChildrenEndpoint(driveId, folderPath);

    return graphClient
      .api(endpoint)
      .select(this.getGraphDriveSelectFields().join(','))
      .top(query.pageSize)
      .get();
  }

  private buildGraphChildrenEndpoint(driveId: string, folderPath: string): string {
    if (!folderPath) {
      return `/drives/${driveId}/root/children`;
    }

    return `/drives/${driveId}/root:/${this.encodeRelativePath(folderPath)}:/children`;
  }

  private getGraphDriveSelectFields(): string[] {
    /**
     * No agrego additionalColumns aquí porque driveItem no expone
     * columnas personalizadas de SharePoint directamente como propiedades simples.
     *
     * Para columnas personalizadas vía Graph habría que enriquecer con listItem/fields.
     * Para KQL, esas columnas se manejan mediante managed properties en selectproperties.
     */
    return [
      'id',
      'name',
      'webUrl',
      'createdDateTime',
      'lastModifiedDateTime',
      'lastModifiedBy',
      'parentReference',
      'file',
      'folder',
      'size'
    ];
  }

  /**
   * MOTOR 2: SharePoint Search KQL
   *
   * Usa:
   * /_api/search/query
   *
   * KQL NO usa driveId.
   * KQL usa Path:"{libraryUrl}".
   */
  private async searchWithSharePoint(query: IDocumentQuery): Promise<IDocumentPage> {
    const startRow: number = query.startRow || 0;
    const url: string = this.buildSearchUrl(query, startRow);

    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(`SharePoint search request failed: ${response.status}`);
    }

    const data: ISearchResponse = await response.json();
    const relevantResults = (data.d?.query || data).PrimaryQueryResult?.RelevantResults;
    const rows = this.toArray(relevantResults?.Table?.Rows);

    const rowCount: number = relevantResults?.RowCount || rows.length;
    const totalRows: number = relevantResults?.TotalRows || 0;

    const items: IDocumentItem[] = rows.map((row: ISearchRow, index: number) =>
      this.mapSearchRow(row, query.additionalColumns, startRow + index)
    );

    const nextStartRow: number = startRow + rowCount;

    return {
      items,
      nextStartRow: nextStartRow < totalRows ? nextStartRow : undefined
    };
  }

  private buildSearchUrl(query: IDocumentQuery, startRow: number): string {
    const searchApiSiteUrl: string = this.getSearchApiSiteUrl(query);
    const queryText: string = this.buildSharePointQueryText(query);

    const selectProperties: string[] = this.cleanStringArray([
      'Title',
      'Path',
      'Filename',
      'ParentLink',
      'Created',
      'LastModifiedTime',
      'FileExtension',
      'Size',
      'Author',
      'EditorOWSUSER',
      ...query.additionalColumns
    ]);

    const params: string[] = [
      this.buildQuotedQueryParam('querytext', queryText),
      `rowlimit=${query.pageSize}`,
      `startrow=${startRow}`,
      this.buildQuotedQueryParam('selectproperties', selectProperties.join(',')),
      this.buildQuotedQueryParam('sortlist', 'LastModifiedTime:descending'),
      'trimduplicates=false'
    ];

    return `${searchApiSiteUrl}/_api/search/query?${params.join('&')}`;
  }

  private buildSharePointQueryText(query: IDocumentQuery): string {
    const searchText: string = this.getSearchText(query);

    const queryParts: string[] = [
      searchText ? this.escapeKqlText(searchText) : '*',
      'IsDocument:1'
    ];

    const scopePath: string = this.getKqlScopePath(query.libraryUrl, this.getEffectiveFolderPath(query));

    if (scopePath) {
      queryParts.push(`Path:"${this.escapeKqlText(scopePath)}"`);
    }

    return queryParts.join(' ');
  }

  private getKqlScopePath(libraryUrl?: string, folderPath?: string): string {
    const cleanLibraryUrl: string = this.normalizeAbsoluteUrl(libraryUrl || '');
    const cleanFolderPath: string = this.cleanRelativePath(folderPath || '');

    if (!cleanLibraryUrl) {
      return '';
    }

    if (!cleanFolderPath) {
      return cleanLibraryUrl;
    }

    return `${cleanLibraryUrl}/${this.encodeRelativePath(cleanFolderPath)}`;
  }

  private getSearchApiSiteUrl(query: IDocumentQuery): string {
    const explicitSiteUrl: string = this.normalizeAbsoluteUrl(query.siteUrl || '');

    if (explicitSiteUrl) {
      return explicitSiteUrl;
    }

    const libraryUrl: string = this.normalizeAbsoluteUrl(query.libraryUrl || '');

    if (!libraryUrl) {
      throw new Error('Debe configurar URL del sitio o URL de biblioteca para SharePoint Search KQL.');
    }

    return this.getParentAbsoluteUrl(libraryUrl);
  }

  private mapSearchRow(row: ISearchRow, additionalColumns: string[], fallbackId: number): IDocumentItem {
    const values: Record<string, string> = {};
    const cells = this.toArray(row.Cells);

    cells.forEach((cell: ISearchCell) => {
      if (cell.Key) {
        values[cell.Key] = cell.Value || '';
      }
    });

    const url: string = values.Path || '';
    const name: string = values.Filename || values.Title || this.getNameFromUrl(url);
    const modifiedBy: string = this.getDisplayName(values.EditorOWSUSER || values.Author || '');

    const additionalValues: Record<string, string> = {};
    additionalColumns.forEach((column: string) => {
      additionalValues[column] = values[column] || '';
    });

    return {
      id: fallbackId,
      name,
      path: values.ParentLink || this.getParentPath(url),
      created: values.Created || '',
      modified: values.LastModifiedTime || '',
      modifiedBy,
      fileType: values.FileExtension || '',
      size: Number(values.Size) || undefined,
      url,
      isFolder: false,
      additionalValues
    };
  }

  private mapGraphDriveItem(item: IGraphDriveItem, additionalColumns: string[], fallbackId: number): IDocumentItem {
    const url: string = item.webUrl || '';

    const additionalValues: Record<string, string> = {};

    /**
     * driveItem no trae columnas personalizadas de SharePoint por defecto.
     * Si algún valor estándar llega con ese nombre, lo mostramos; si no, queda vacío.
     */
    additionalColumns.forEach((column: string) => {
      const value: unknown = item[column];
      additionalValues[column] = value === undefined || value === null ? '' : String(value);
    });

    return {
      id: fallbackId,
      name: item.name || this.getNameFromUrl(url),
      path: this.getGraphParentPath(item.parentReference?.path || url),
      created: item.createdDateTime || '',
      modified: item.lastModifiedDateTime || '',
      modifiedBy: item.lastModifiedBy?.user?.displayName || item.lastModifiedBy?.application?.displayName || '',
      fileType: item.file ? this.getFileType(item.name || '', item.file.mimeType || '') : 'folder',
      size: item.size,
      url,
      isFolder: !!item.folder,
      additionalValues
    };
  }

  private getGraphApiPath(value: string): string {
    /**
     * @odata.nextLink normalmente viene como:
     * https://graph.microsoft.com/v1.0/drives/...?$skiptoken=...
     *
     * MSGraphClient funciona mejor recibiendo la ruta relativa:
     * /drives/...?$skiptoken=...
     */
    if (!/^https?:\/\//i.test(value)) {
      return value;
    }

    const url = new URL(value);
    const path = url.pathname.replace(/^\/v1\.0/i, '');

    return `${path}${url.search}`;
  }

  private getSearchText(query: IDocumentQuery): string {
    return (query.filters?.searchText || '').trim();
  }

  private hasSearchText(query: IDocumentQuery): boolean {
    return this.getSearchText(query).length > 0;
  }

  private getEffectiveFolderPath(query: IDocumentQuery): string {
    return this.cleanRelativePath(query.folderPath || query.baseFolderPath || '');
  }

  private buildQuotedQueryParam(name: string, value: string): string {
    return `${name}='${encodeURIComponent(value.replace(/'/g, "''"))}'`;
  }

  private toArray<T>(value: T[] | { results?: T[] } | undefined): T[] {
    if (!value) {
      return [];
    }

    return Array.isArray(value) ? value : value.results || [];
  }

  private getNameFromUrl(url: string): string {
    const cleanUrl: string = url.split('?')[0];
    const parts: string[] = cleanUrl.split('/');
    return this.safeDecodeURIComponent(parts[parts.length - 1] || '');
  }

  private getParentPath(url: string): string {
    const cleanUrl: string = url.split('?')[0];
    const lastSlashIndex: number = cleanUrl.lastIndexOf('/');
    return lastSlashIndex > -1 ? cleanUrl.substring(0, lastSlashIndex) : '';
  }

  private getParentAbsoluteUrl(url: string): string {
    const cleanUrl: string = this.trimTrailingSlash(url);
    const lastSlashIndex: number = cleanUrl.lastIndexOf('/');

    if (lastSlashIndex <= cleanUrl.indexOf('//') + 1) {
      return cleanUrl;
    }

    return cleanUrl.substring(0, lastSlashIndex);
  }

  private getGraphParentPath(value: string): string {
    const driveRootMarker: string = 'root:';
    const markerIndex: number = value.indexOf(driveRootMarker);

    if (markerIndex > -1) {
      const path = value.substring(markerIndex + driveRootMarker.length);
      return this.safeDecodeURIComponent(path || '/');
    }

    return this.getParentPath(value);
  }

  private getFileType(name: string, mimeType: string): string {
    const lastDotIndex: number = name.lastIndexOf('.');

    if (lastDotIndex > -1 && lastDotIndex < name.length - 1) {
      return name.substring(lastDotIndex + 1).toLowerCase();
    }

    return mimeType;
  }

  private getDisplayName(value: string): string {
    const parts: string[] = value.split('|');
    return parts[parts.length - 1] || value;
  }

  private escapeKqlText(value: string): string {
    return value.replace(/"/g, '\\"');
  }

  private escapeGraphSearchText(value: string): string {
    /**
     * Graph usa:
     * /root/search(q='texto')
     *
     * Duplicamos comillas simples y luego encodeamos.
     */
    return encodeURIComponent(value.replace(/'/g, "''"));
  }

  private normalizeAbsoluteUrl(value: string): string {
    const cleanValue: string = this.trimTrailingSlash(value.trim());

    if (!cleanValue) {
      return '';
    }

    const url = new URL(cleanValue);
    const normalizedPath: string = this.encodeAbsolutePath(url.pathname);

    return `${url.origin}${normalizedPath}`;
  }

  private encodeAbsolutePath(path: string): string {
    const startsWithSlash: boolean = path.startsWith('/');
    const encodedPath: string = path
      .split('/')
      .filter(Boolean)
      .map((segment: string) => encodeURIComponent(this.safeDecodeURIComponent(segment)))
      .join('/');

    return startsWithSlash ? `/${encodedPath}` : encodedPath;
  }

  private encodeRelativePath(path: string): string {
    return this.cleanRelativePath(path)
      .split('/')
      .filter(Boolean)
      .map((segment: string) => encodeURIComponent(this.safeDecodeURIComponent(segment)))
      .join('/');
  }

  private cleanRelativePath(value: string): string {
    return value.trim().replace(/^\/+|\/+$/g, '');
  }

  private cleanStringArray(values: string[]): string[] {
    return values
      .map((value: string) => value.trim())
      .filter((value: string, index: number, array: string[]) => {
        return !!value && array.indexOf(value) === index;
      });
  }

  private safeDecodeURIComponent(value: string): string {
    try {
      return decodeURIComponent(value);
    } catch {
      return value;
    }
  }

  private trimTrailingSlash(value: string): string {
    return value.replace(/\/$/, '');
  }
}
