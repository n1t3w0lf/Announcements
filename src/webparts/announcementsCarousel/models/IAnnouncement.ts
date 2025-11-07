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
}
