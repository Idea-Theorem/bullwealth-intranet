export interface IWelcomeBannerPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  // Configurable properties from Property Pane
  messageTitle?: string;
  ceoName?: string;
  ceoTitle?: string;
  ceoMessage?: string;
  ceoExpandedMessage?: string;
  ceoImageUrl?: string;
  backgroundImageUrl?: string;
  videoUrl?: string;
  showVideo?: boolean;
  readMoreText?: string;
  readLessText?: string;
}