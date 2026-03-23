/** How the image fills its container — maps to CSS object-fit values. */
export enum ImageFitMode {
  Cover = 'cover',
  Contain = 'contain',
  Fill = 'fill',
  None = 'none'
}

/** Whether image dimensions use CSS fit mode or manual pixel values. */
export enum ImageSizeMode {
  Fit = 'fit',
  Manual = 'manual'
}

/** Whether the image container background is a solid color or gradient. */
export enum ImageBackgroundType {
  Solid = 'solid',
  Gradient = 'gradient'
}

export interface IAnnouncement {
  Id: number;
  Title: string;
  Description: string;
  AnnouncementImage: string;
  ValidFrom: Date;
  ValidTo: Date;
  CelebrationIcon?: CelebrationType;
  CelebrationIconPosition?: IconPosition;
  CustomIconUrl?: string;
  IsActive?: boolean;
  /** Image container height in px. */
  ImageHeight?: number;
  /** Image width in px. 0 = full width. */
  ImageWidth?: number;
  /** Image fit mode. */
  ImageFit?: ImageFitMode;
  /** Whether image uses CSS fit or manual dimensions. */
  ImageSizeMode?: ImageSizeMode;
  /** Card background color (hex). */
  BackgroundColor?: string;
  /** Title text color (hex). */
  TitleColor?: string;
  /** Description text color (hex). */
  DescriptionColor?: string;
  /** Overall card height in px. */
  CardHeight?: number;
  /** Border radius in px. */
  BorderRadius?: number;
  /** Whether to show a drop shadow on the card. */
  ShowShadow?: boolean;
  /** Whether to show the celebration icon. */
  ShowCelebrationIcon?: boolean;
  /** Celebration icon size in px. */
  CelebrationIconSize?: number;
  /** Whether image background is solid or gradient. */
  ImageBackgroundType?: ImageBackgroundType;
  /** Solid background color for image container (hex). Used when ImageBackgroundType is solid. */
  ImageBackgroundColor?: string;
  /** Gradient start color (hex). */
  GradientStartColor?: string;
  /** Gradient end color (hex). */
  GradientEndColor?: string;
  /** Gradient direction in degrees. */
  GradientDirection?: number;
  /** Overlay gradient color (hex). */
  OverlayGradientColor?: string;
  /** Overlay opacity percentage (0-100). */
  OverlayOpacity?: number;
  /** Optional URL to navigate to when the announcement is clicked. */
  RedirectUrl?: string;
  /** Whether to open the redirect URL in the same tab or a new tab. */
  RedirectTarget?: '_self' | '_blank';
}

export enum CelebrationType {
  None = 'None',
  Birthday = 'Birthday',
  Anniversary = 'Anniversary',
  Achievement = 'Achievement',
  Celebration = 'Celebration',
  NewHire = 'NewHire',
  Promotion = 'Promotion',
  Holiday = 'Holiday',
  Custom = 'Custom'
}

export enum IconPosition {
  TopLeft = 'topLeft',
  TopRight = 'topRight',
  BottomLeft = 'bottomLeft',
  BottomRight = 'bottomRight',
  Center = 'center'
}

export interface IAnnouncementListItem {
  Id: number;
  Title: string;
  Description: string;
  AnnouncementImage: {
    Url: string;
  };
  ValidFrom: string;
  ValidTo: string;
  CelebrationIcon: string;
  CelebrationIconPosition: string;
  CustomIconUrl: string;
  ImageHeight: number;
  ImageWidth: number;
  ImageFit: string;
  ImageSizeMode: string;
  ImageBackgroundType: string;
  ImageBackgroundColor: string;
  BackgroundColor: string;
  TitleColor: string;
  DescriptionColor: string;
  CardHeight: number;
  BorderRadius: number;
  ShowShadow: boolean;
  ShowCelebrationIcon: boolean;
  CelebrationIconSize: number;
  GradientStartColor: string;
  GradientEndColor: string;
  GradientDirection: number;
  OverlayGradientColor: string;
  OverlayOpacity: number;
  RedirectUrl: string;
  RedirectTarget: string;
}

/**
 * Form data shape used by the announcement management form.
 * All values stored as strings for easy form binding.
 */
export interface IAnnouncementFormData {
  Title: string;
  Description: string;
  AnnouncementImage: string;
  ValidFrom: string;
  ValidTo: string;
  CelebrationIcon: CelebrationType;
  CelebrationIconPosition: IconPosition;
  CustomIconUrl: string;
  ImageHeight: string;
  ImageWidth: string;
  ImageFit: string;
  ImageSizeMode: string;
  ImageBackgroundType: string;
  ImageBackgroundColor: string;
  BackgroundColor: string;
  TitleColor: string;
  DescriptionColor: string;
  CardHeight: string;
  BorderRadius: string;
  ShowShadow: string;
  ShowCelebrationIcon: string;
  CelebrationIconSize: string;
  GradientStartColor: string;
  GradientEndColor: string;
  GradientDirection: string;
  OverlayGradientColor: string;
  OverlayOpacity: string;
  RedirectUrl: string;
  RedirectTarget: string;
}

/** Creates fresh empty form data for new announcements with sensible defaults. */
export function createEmptyFormData(): IAnnouncementFormData {
  return {
    Title: '',
    Description: '',
    AnnouncementImage: '',
    ValidFrom: new Date().toISOString(),
    ValidTo: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString(),
    CelebrationIcon: CelebrationType.None,
    CelebrationIconPosition: IconPosition.TopRight,
    CustomIconUrl: '',
    ImageHeight: '400',
    ImageWidth: '0',
    ImageFit: 'cover',
    ImageSizeMode: 'fit',
    ImageBackgroundType: 'gradient',
    ImageBackgroundColor: '#667eea',
    BackgroundColor: '#ffffff',
    TitleColor: '#333333',
    DescriptionColor: '#333333',
    CardHeight: '500',
    BorderRadius: '8',
    ShowShadow: 'true',
    ShowCelebrationIcon: 'true',
    CelebrationIconSize: '60',
    GradientStartColor: '#667eea',
    GradientEndColor: '#764ba2',
    GradientDirection: '135',
    OverlayGradientColor: '#000000',
    OverlayOpacity: '30',
    RedirectUrl: '',
    RedirectTarget: '_self'
  };
}
