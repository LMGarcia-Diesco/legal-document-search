import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneField,
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'LegalDocumentSearchWebPartStrings';
import LegalDocumentSearch from './components/LegalDocumentSearch';
import { ILegalDocumentSearchProps } from './components/ILegalDocumentSearchProps';

export interface ILegalDocumentSearchWebPartProps {
  siteUrl: string;
  driveId: string;
  libraryUrl: string;
  folderPath: string;
  additionalColumns: string;
  pageSize: number;
  searchProvider: string;
}

export default class LegalDocumentSearchWebPart extends BaseClientSideWebPart<ILegalDocumentSearchWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ILegalDocumentSearchProps> = React.createElement(
      LegalDocumentSearch,
      {
        spHttpClient: this.context.spHttpClient,
        msGraphClientFactory: this.context.msGraphClientFactory,
        currentWebUrl: this.context.pageContext.web.absoluteUrl,
        siteUrl: this.properties.siteUrl,
        driveId: this.properties.driveId,
        libraryUrl: this.properties.libraryUrl,
        folderPath: this.properties.folderPath,
        additionalColumns: this.properties.additionalColumns,
        pageSize: this.properties.pageSize || 50,
        searchProvider: this.properties.searchProvider || 'graphDrive'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const searchProvider: string = this.properties.searchProvider || 'graphDrive';
    const groupFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneDropdown('searchProvider', {
        label: strings.SearchProviderFieldLabel,
        options: [
          { key: 'graphDrive', text: 'Microsoft Graph Drive' },
          { key: 'sharePointSearch', text: 'SharePoint Search KQL' }
        ]
      })
    ];

    if (searchProvider === 'graphDrive') {
      groupFields.push(
        PropertyPaneTextField('driveId', {
          label: strings.DriveIdFieldLabel
        })
      );
    }

    if (searchProvider === 'sharePointSearch') {
      groupFields.push(
        PropertyPaneTextField('libraryUrl', {
          label: strings.LibraryUrlFieldLabel
        })
      );
    }

    groupFields.push(
      PropertyPaneTextField('folderPath', {
        label: strings.FolderPathFieldLabel
      }),
      PropertyPaneDropdown('pageSize', {
        label: strings.PageSizeFieldLabel,
        options: [
          { key: 25, text: '25' },
          { key: 50, text: '50' },
          { key: 100, text: '100' }
        ]
      }),
      PropertyPaneTextField('additionalColumns', {
        label: strings.AdditionalColumnsFieldLabel
      })
    );

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields
            }
          ]
        }
      ]
    };
  }
}
