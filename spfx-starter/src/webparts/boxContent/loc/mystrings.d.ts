declare interface IBoxContentWebPartStrings {
  PropertyPaneDescription: string;
  ContentGroupName: string;
  DesignGroupName: string;
  TitleFieldLabel: string;
  DescriptionFieldLabel: string;
  DurationFieldLabel: string;
  ButtonTextFieldLabel: string;
  ButtonUrlFieldLabel: string;
  ButtonIconFieldLabel: string;
  ShowDurationFieldLabel: string;
  BackgroundColorFieldLabel: string;
  TitleColorFieldLabel: string;
  DescriptionColorFieldLabel: string;
  ButtonColorFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'BoxContentWebPartStrings' {
  const strings: IBoxContentWebPartStrings;
  export = strings;
}
