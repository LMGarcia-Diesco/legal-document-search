import * as React from 'react';
import { Link } from '@fluentui/react';

export const renderCellValue = (value?: string): JSX.Element => {
  return <span title={value || ''}>{value || ''}</span>;
};

export const renderDocumentLink = (url: string): JSX.Element => {
  if (!url) {
    return <span />;
  }

  return (
    <Link href={url} target="_blank" rel="noreferrer">
      Abrir
    </Link>
  );
};
