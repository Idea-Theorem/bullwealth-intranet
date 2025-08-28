declare interface IHeadingAndSubheadingWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  HeadingFieldLabel: string;
  SubheadingFieldLabel: string;
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

declare module 'HeadingAndSubheadingWebPartStrings' {
  const strings: IHeadingAndSubheadingWebPartStrings;
  export = strings;
}
