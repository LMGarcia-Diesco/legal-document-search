import * as React from 'react';
import { SearchBox } from '@fluentui/react';

export interface ISearchBarProps {
  value: string;
  onChange: (value: string) => void;
}

export const SearchBar: React.FC<ISearchBarProps> = (props: ISearchBarProps) => {
  return (
    <SearchBox
      placeholder="Buscar documentos"
      value={props.value}
      onChange={(_, newValue?: string) => props.onChange(newValue || '')}
      onClear={() => props.onChange('')}
    />
  );
};
