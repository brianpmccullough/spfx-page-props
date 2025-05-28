import * as React from 'react';
import { ITheme, useTheme } from '@fluentui/react';

export interface IChipProps {
  label: string;
  className?: string;
}

export const Chip: React.FC<IChipProps> = ({ label, className }) => {
  const theme: ITheme = useTheme();
  
  const chipStyle: React.CSSProperties = {
    display: 'inline-block',
    padding: '4px 8px',
    margin: '2px',
    borderRadius: '16px',
    backgroundColor: theme.palette.neutralLighter,
    color: theme.palette.neutralPrimary,
  };

  return (
    <span style={chipStyle} className={className}>
      {label}
    </span>
  );
};