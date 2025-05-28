declare interface ISpFxPagePropertiesWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  PropertyPaneTitleLabel: string;
}

declare module 'SpFxPagePropertiesWebPartStrings' {
  const strings: ISpFxPagePropertiesWebPartStrings;
  export = strings;
}
