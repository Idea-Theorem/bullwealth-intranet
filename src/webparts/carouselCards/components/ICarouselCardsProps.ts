export interface ICarouselCardsProps {
  title: string;
  subtitle: string;
  cards: ICarouselCard[];
}

export interface ICarouselCard {
  id: string;
  icon: string; // This will store the uploaded image URL or Fluent UI icon name
  iconType: 'upload' | 'fluent'; // Type of icon
  iconColor: string;
  title: string;
  description: string;
  bulletPoints: string[];
  isVisible: boolean; // Show/Hide card
}
