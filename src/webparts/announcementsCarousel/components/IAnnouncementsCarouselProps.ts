import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAnnouncementsCarouselProps {
  context: WebPartContext;
  listName: string;
  rotationInterval: number; // in seconds
  showTitle: boolean;
  showDescription: boolean;
  titleFontSize: number;
  descriptionFontSize: number;
  backgroundColor: string;
  titleColor: string;
  descriptionColor: string;
  height: number;
  enableTransitions: boolean;
  transitionEffect: 'fade' | 'slide' | 'zoom';
  showNavigationDots: boolean;
  showNavigationArrows: boolean;
  autoPlay: boolean;
  celebrationIconSize: number;
  showCelebrationIcon: boolean;
  borderRadius: number;
  showShadow: boolean;
  overlayOpacity: number;
}
