export interface IContactCard {
  title: string;
  subtitle: string;
  name: string;
  email: string;
  phone: string;
  emailButtonText: string;
  phoneButtonText: string;
  showEmailButton: boolean;
  showPhoneButton: boolean;
  cardBackgroundColor: string;
  titleColor: string;
  subtitleColor: string;
  nameColor: string;
  contactColor: string;
  emailButtonColor: string;
  phoneButtonColor: string;
}

export interface I2ColBoxContentProps {
  leftCard: IContactCard;
  rightCard: IContactCard;
  columnLayout: string; // 'left-right' or 'right-left'
  containerBackgroundColor: string;
  cardSpacing: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
