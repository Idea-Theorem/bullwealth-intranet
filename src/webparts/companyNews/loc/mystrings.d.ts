declare interface ICompanyNewsWebPartStrings {
  AppSharePointEnvironment: any;
  AppLocalEnvironmentSharePoint: any;
  AppOfficeEnvironment: string;
  AppOfficeEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOutlookEnvironment: string;
  AppLocalEnvironmentOutlook: string;
  AppLocalEnvironmentOffice: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  ItemsToShowFieldLabel: string;
}

declare module 'CompanyNewsWebPartStrings' {
  const strings: ICompanyNewsWebPartStrings;
  export = strings;
}