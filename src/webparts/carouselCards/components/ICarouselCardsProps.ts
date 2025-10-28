export interface ICarouselCardsProps {
  title: string;
  subtitle: string;
  cards: ICarouselCard[];
}

export interface ICarouselCard {
  id: string;
  icon: string;
  iconType: 'upload' | 'fluent';
  iconColor: string;
  title: string;
  description: string;
  bulletPoints: string[]; // ✅ Keep this - it's still used
  isVisible: boolean;
}
