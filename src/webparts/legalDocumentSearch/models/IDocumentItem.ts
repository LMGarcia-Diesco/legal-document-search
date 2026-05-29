export interface IDocumentItem {
  id: number;
  sourceId?: string;
  name: string;
  path: string;
  created: string;
  modified: string;
  modifiedBy: string;
  fileType: string;
  size?: number;
  url: string;
  isFolder: boolean;
  additionalValues: Record<string, string>;
}

export interface IDocumentPage {
  items: IDocumentItem[];
  nextStartRow?: number;
  nextLink?: string;
}

export interface IDocumentDetails {
  fields: Record<string, string>;
}
