import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';

export interface ILegalDocumentSearchProps {
  spHttpClient: SPHttpClient;
  msGraphClientFactory: MSGraphClientFactory;
  currentWebUrl: string;
  siteUrl: string;
  driveId: string;
  libraryUrl: string;
  folderPath: string;
  additionalColumns: string;
  pageSize: number;
  searchProvider: string;
}
