import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react';
import { IDocumentItem } from '../models/IDocumentItem';
import { renderCellValue, renderDocumentLink } from './DocumentRow';

export interface IDocumentResultsListProps {
  items: IDocumentItem[];
  additionalColumns: string[];
}

const formatDate = (value: string): string => {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
};

export const DocumentResultsList: React.FC<IDocumentResultsListProps> = (props: IDocumentResultsListProps) => {
  const columns = React.useMemo<IColumn[]>(() => {
    const baseColumns: IColumn[] = [
      {
        key: 'name',
        name: 'Nombre',
        fieldName: 'name',
        minWidth: 180,
        isResizable: true,
        onRender: (item: IDocumentItem) => renderCellValue(item.name)
      },
      {
        key: 'path',
        name: 'Ruta/carpeta',
        fieldName: 'path',
        minWidth: 220,
        isResizable: true,
        onRender: (item: IDocumentItem) => renderCellValue(item.path)
      },
      {
        key: 'modified',
        name: 'Fecha de modificación',
        fieldName: 'modified',
        minWidth: 150,
        isResizable: true,
        onRender: (item: IDocumentItem) => renderCellValue(formatDate(item.modified))
      },
      {
        key: 'modifiedBy',
        name: 'Modificado por',
        fieldName: 'modifiedBy',
        minWidth: 150,
        isResizable: true,
        onRender: (item: IDocumentItem) => renderCellValue(item.modifiedBy)
      },
      {
        key: 'fileType',
        name: 'Tipo de archivo',
        fieldName: 'fileType',
        minWidth: 90,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: IDocumentItem) => renderCellValue(item.fileType)
      }
    ];

    const customColumns: IColumn[] = props.additionalColumns.map((column: string) => ({
      key: column,
      name: column,
      minWidth: 120,
      isResizable: true,
      onRender: (item: IDocumentItem) => renderCellValue(item.additionalValues[column])
    }));

    return [
      ...baseColumns,
      ...customColumns,
      {
        key: 'open',
        name: '',
        minWidth: 60,
        maxWidth: 80,
        onRender: (item: IDocumentItem) => renderDocumentLink(item.url)
      }
    ];
  }, [props.additionalColumns]);

  return (
    <DetailsList
      items={props.items}
      columns={columns}
      compact={true}
      selectionMode={SelectionMode.none}
      layoutMode={DetailsListLayoutMode.justified}
      setKey="documentResults"
    />
  );
};
