export interface IDocumentItem {
  id: number;
  name: string;
  path: string;
  modified: string;
  modifiedBy: string;
  fileType: string;
  url: string;
  additionalValues: Record<string, string>;
}

export interface IDocumentPage {
  items: IDocumentItem[];
  nextStartRow?: number;
  nextLink?: string;
}
