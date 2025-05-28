import { DisplayMode } from "@microsoft/sp-core-library";
import { IListColumnWithValue } from "../models/IListSiteColumn";

export interface ISpFxPagePropertiesProps {
  title: string;
  displayMode: DisplayMode
  updateTitle: (value: string) => void;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  pageProperties: IListColumnWithValue[];
  selectedPageProperties: string[];
}
