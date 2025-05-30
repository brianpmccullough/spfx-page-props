import { DisplayMode } from "@microsoft/sp-core-library";
import { IListColumnWithValue } from "../models/IListSiteColumn";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IPagePropertiesProps {
  theme: IReadonlyTheme;
  title: string;
  displayMode: DisplayMode
  updateTitle: (value: string) => void;
  pageProperties: IListColumnWithValue[];
  selectedPageProperties: string[];
}
