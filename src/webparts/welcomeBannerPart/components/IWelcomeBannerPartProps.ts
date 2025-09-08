export interface IWelcomeBannerPartProps {
 // Basic properties
  description: string;

  // Theme and context properties
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  // CEO Message properties
  messageTitle: string;
  ceoName: string;
  ceoTitle: string;
  ceoMessage: string;
  ceoExpandedMessage: string;
  
  // Image URLs
  backgroundImageUrl: string;
  ceoImageUrl: string;
  
  // Video properties
  videoUrl: string;
  showVideo: boolean;
  
  // Button text
  readMoreText: string;
  readLessText: string;
}