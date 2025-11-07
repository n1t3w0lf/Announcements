export interface IAnnouncement {
  Id: number;
  Title: string;
  Description: string;
  AnnouncementImage: string;
  ValidFrom: Date;
  ValidTo: Date;
  CelebrationIcon?: CelebrationType;
  CelebrationIconPosition?: IconPosition;
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
  TopLeft = 'top-left',
  TopRight = 'top-right',
  BottomLeft = 'bottom-left',
  BottomRight = 'bottom-right',
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
}
