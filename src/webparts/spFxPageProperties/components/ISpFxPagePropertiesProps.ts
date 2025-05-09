export interface ISpFxPagePropertiesProps {
  description: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
  pageProperties: Record<string, unknown>;
  selectedPageProperties: string[];
}
