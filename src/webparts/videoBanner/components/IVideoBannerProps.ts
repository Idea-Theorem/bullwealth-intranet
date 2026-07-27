import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IVideoBannerProps {
  context: WebPartContext;
  backgroundImages: string[];
  autoSlide: boolean;
  slideInterval: number;
}
