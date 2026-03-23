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
  /** Message shown when there are no active announcements. */
  emptyStateMessage: string;
  /** Fluent UI icon name for the empty state. */
  emptyStateIcon: string;
  /** Color of the empty state icon (hex). */
  emptyStateIconColor: string;
  /** Background type for the empty state: 'solid', 'gradient', or 'image'. */
  emptyStateBackgroundType: 'solid' | 'gradient' | 'image';
  /** Background color when type is solid. */
  emptyStateBackgroundColor: string;
  /** Gradient start color when type is gradient. */
  emptyStateGradientStart: string;
  /** Gradient end color when type is gradient. */
  emptyStateGradientEnd: string;
  /** Gradient direction in degrees when type is gradient. */
  emptyStateGradientDirection: number;
  /** Background image URL when type is image. */
  emptyStateBackgroundImage: string;
  /** Text color for the empty state message. */
  emptyStateTextColor: string;
}
