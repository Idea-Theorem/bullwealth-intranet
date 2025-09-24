import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IVideoBannerProps {
  title?: string;
  message?: string;
  buttonText?: string;
  videoUrl?: string;
  thumbnailUrl?: string;
  backgroundImageUrl?: string;
  publishedDate?: string;
  readMoreUrl?: string;
  showInModal?: boolean;
  autoPlay?: boolean;
  context: WebPartContext; // Required for SharePoint API calls
}
