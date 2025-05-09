export interface IListColumn {
  /** The unique identifier (GUID) of the field */
  id: string;
  /** The display name of the field */
  title: string;
  /** The internal name of the field */
  internalName: string;
  /** The field type as a string (e.g. User, Hyperlink, Text, Date, Lookup, TaxonomyFieldType) */
  fieldType: string;
  /** Whether the field is hidden */
  hidden: boolean;
  /** The group the field belongs to */
  group: string;
}
