import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAnnouncementsCarouselProps {
  context: WebPartContext;
  listName: string;
  rotationInterval: number; // in seconds
  showTitle: boolean;
  showDescription: boolean;
  titleFontSize: number;
  descriptionFontSize: number;
  enableTransitions: boolean;
  transitionEffect: 'fade' | 'slide' | 'zoom';
  showNavigationDots: boolean;
  showNavigationArrows: boolean;
  autoPlay: boolean;
}
