import { SharePointFieldType } from "../services/SharePointFieldType";

export interface IListColumn {
  /** The unique identifier (GUID) of the field */
  id: string;
  /** The display name of the field */
  title: string;
  /** The internal name of the field */
  internalName: string;
  /** The field type as a string (e.g. User, Hyperlink, Text, Date, Lookup, TaxonomyFieldType) */
  fieldType: SharePointFieldType;
  /** Whether the field is hidden */
  hidden: boolean;
  /** The group the field belongs to */
  group: string;
}

export interface UserFieldValue {
  /** The ID of the user */
  id: number;
  /** The display name of the user */
  name: string;
  /** The login name of the user */
  userName?: string;
  /** The profile picture of the user */
  profilePictureUrl?: string;
  /** The initials of the user */
  initials: string;
}

export interface LookupFieldValue {
  /** The ID of the lookup item */
  id: number;
  /** The title of the lookup item */
  title: string;
}

export interface TaxonomyFieldValue {
  /** The ID of the taxonomy term */
  id: string;
  /** The taxonomy term */
  term: string;
}

export interface UrlFieldValue {
  /** The URL of the link */
  url: string;
  /** The description of the link */
  displayText: string;
}

export interface IListColumnWithValue extends IListColumn {
  /** The value of the field */
  // eslint-disable-next-line @rushstack/no-new-null
  value: string | number | boolean | Date | UserFieldValue | UserFieldValue[] | LookupFieldValue | LookupFieldValue[] | TaxonomyFieldValue | TaxonomyFieldValue[] | UrlFieldValue | null;
}