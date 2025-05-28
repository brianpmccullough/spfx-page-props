import { IListColumn, IListColumnWithValue } from '../models/IListSiteColumn';

/**
 * Interface for a service that retrieves the current page's properties.
 */
export interface IPagePropertiesService {
  /**
   * Gets the properties of the current page.
   * @returns Promise resolving to a dictionary of page properties.
   */
  getPageProperties(listColumns: IListColumn[]): Promise<IListColumnWithValue[]>;

  getListColumns(): Promise<IListColumn[]>;
}
