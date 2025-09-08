declare interface ITwoColBoxContentWebPartStrings {
  PropertyPaneDescription: string;
  LayoutGroupName: string;
  LeftCardGroupName: string;
  RightCardGroupName: string;
  DesignGroupName: string;
  
  ColumnLayoutFieldLabel: string;
  ContainerBackgroundColorFieldLabel: string;
  CardSpacingFieldLabel: string;
  
  CardTitleFieldLabel: string;
  CardSubtitleFieldLabel: string;
  NameFieldLabel: string;
  EmailFieldLabel: string;
  PhoneFieldLabel: string;
  EmailButtonTextFieldLabel: string;
  PhoneButtonTextFieldLabel: string;
  ShowEmailButtonFieldLabel: string;
  ShowPhoneButtonFieldLabel: string;
  CardBackgroundColorFieldLabel: string;
  TitleColorFieldLabel: string;
  SubtitleColorFieldLabel: string;
  NameColorFieldLabel: string;
  ContactColorFieldLabel: string;
  EmailButtonColorFieldLabel: string;
  PhoneButtonColorFieldLabel: string;
  
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

declare module 'TwoColBoxContentWebPartStrings' {
  const strings: ITwoColBoxContentWebPartStrings;
  export = strings;
}
