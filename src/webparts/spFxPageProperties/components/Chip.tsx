import * as React from 'react';
import styles from './PageProperties.module.scss';

export interface IChipProps {
  label?: string;
}

export const Chip: React.FC<IChipProps> = ({ label }) => {

  return (
    <span className={styles.chip}>
      {label}
    </span>
  );
};