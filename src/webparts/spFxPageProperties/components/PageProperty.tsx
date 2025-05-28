import * as React from 'react';
import { Persona, PersonaSize, PersonaPresence } from '@fluentui/react/lib/Persona';
import { IListColumnWithValue, LookupFieldValue, TaxonomyFieldValue, UserFieldValue } from '../models/IListSiteColumn';
import styles from './SpFxPageProperties.module.scss';
import { Chip } from './Chip';

export default class PageProperty extends React.Component<IListColumnWithValue> {

  public render(): React.ReactElement<IListColumnWithValue> {
    const { internalName, title } = this.props;

    return (
      <div key={internalName} className="property">
        <h3 className="ms-fontSize-20 ms-fontWeight-regular property-name">{title}</h3>
        <div className="property-value">
          {this._renderValue(this.props)}
        </div>
      </div>
    );
  }

  private _renderValue(property: IListColumnWithValue): JSX.Element {
    const { value, fieldType } = property;

    if (!value || value === null || value === undefined) {
      return <span className="empty"/>;
    }

    const ensureMultiValue = <T,>(value: T | T[]): T[] => {
      return Array.isArray(value) ? value : [value];
    };

    if (fieldType === 'User' || fieldType === 'UserMulti') {
      const values = ensureMultiValue(value as UserFieldValue | UserFieldValue[]);
      return (<div className="personas">
          {values.map((user, index) => (
            <Persona
              key={index}
              text={user.name}
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageUrl={user.profilePictureUrl}
              imageInitials={user.initials}
            />
          ))}
        </div>);
    } else if (fieldType === 'Lookup' || fieldType === 'LookupMulti') {
      const values = ensureMultiValue(value as LookupFieldValue | LookupFieldValue[]);
      return (
        <div className={styles.chipContainer}>
          {values.map((item, index) => (
            <Chip key={index} label={item.title} />
          ))}
        </div>
      );
    } else if (fieldType === 'TaxonomyFieldType' || fieldType === 'TaxonomyFieldTypeMulti') {
      const values = ensureMultiValue(value as TaxonomyFieldValue | TaxonomyFieldValue[]);
      return (
        <div className={styles.chipContainer}>
          {values.map((term, index) => (
            <Chip key={index} label={term.term} />
          ))}
        </div>
      );
    } else if (fieldType === 'Choice' || fieldType === 'ChoiceMulti') {
      const values = ensureMultiValue(value as string | string[]);
      return (
        <div className={styles.chipContainer}>
          {values.map((choice, index) => (
            <Chip key={index} label={choice} />
          ))}
        </div>
      );
    } else if (fieldType === 'DateTime') {
      return <span>{new Date(value as string).toLocaleDateString()}</span>;
    } else if (fieldType === 'Boolean') {
      return <span>{value ? 'Yes' : 'No'}</span>;
    }

    if (typeof value === 'string') {
      const newLineRegex = /\r?\n/g;
      const urlRegex = /\b((?:https?:\/\/|www\.)[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?:\/[^\s<]*)?)/gi;
      const emailRegex = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;

      const formattedValue = value
        .replace(newLineRegex, '<br/>')
        .replace(urlRegex, (url) => `<a href="${url}" target="_blank" rel="noopener noreferrer">${url}</a>`)
        .replace(emailRegex, (email) => `<a href="mailto:${email}">${email}</a>`);
      return <span dangerouslySetInnerHTML={{ __html: formattedValue }} />;
    }

    return <span>{value}</span>;
  }
}
