import { CelebrationType } from '../models/IAnnouncement';

export interface ICelebrationIconConfig {
  emoji: string;
  backgroundColor: string;
  color: string;
  label: string;
  animation?: string;
}

export class CelebrationIconService {
  public static getIconConfig(type: CelebrationType): ICelebrationIconConfig | null {
    const configs: Record<CelebrationType, ICelebrationIconConfig | null> = {
      [CelebrationType.None]: null,
      [CelebrationType.Birthday]: {
        emoji: '🎂',
        backgroundColor: '#FF6B9D',
        color: '#FFFFFF',
        label: 'Birthday',
        animation: 'bounce'
      },
      [CelebrationType.Anniversary]: {
        emoji: '🎊',
        backgroundColor: '#9C27B0',
        color: '#FFFFFF',
        label: 'Anniversary',
        animation: 'pulse'
      },
      [CelebrationType.Achievement]: {
        emoji: '🏆',
        backgroundColor: '#FFB300',
        color: '#FFFFFF',
        label: 'Achievement',
        animation: 'bounce'
      },
      [CelebrationType.Celebration]: {
        emoji: '🎉',
        backgroundColor: '#FF5722',
        color: '#FFFFFF',
        label: 'Celebration',
        animation: 'pulse'
      },
      [CelebrationType.NewHire]: {
        emoji: '👋',
        backgroundColor: '#4CAF50',
        color: '#FFFFFF',
        label: 'New Hire',
        animation: 'wave'
      },
      [CelebrationType.Promotion]: {
        emoji: '⭐',
        backgroundColor: '#2196F3',
        color: '#FFFFFF',
        label: 'Promotion',
        animation: 'shine'
      },
      [CelebrationType.Holiday]: {
        emoji: '🎄',
        backgroundColor: '#D32F2F',
        color: '#FFFFFF',
        label: 'Holiday',
        animation: 'pulse'
      },
      [CelebrationType.Custom]: {
        emoji: '✨',
        backgroundColor: '#607D8B',
        color: '#FFFFFF',
        label: 'Special',
        animation: 'sparkle'
      }
    };

    return configs[type] || null;
  }

  public static getAllIconTypes(): Array<{ key: CelebrationType; text: string; icon: string }> {
    return [
      { key: CelebrationType.None, text: 'None', icon: '' },
      { key: CelebrationType.Birthday, text: 'Birthday 🎂', icon: '🎂' },
      { key: CelebrationType.Anniversary, text: 'Anniversary 🎊', icon: '🎊' },
      { key: CelebrationType.Achievement, text: 'Achievement 🏆', icon: '🏆' },
      { key: CelebrationType.Celebration, text: 'Celebration 🎉', icon: '🎉' },
      { key: CelebrationType.NewHire, text: 'New Hire 👋', icon: '👋' },
      { key: CelebrationType.Promotion, text: 'Promotion ⭐', icon: '⭐' },
      { key: CelebrationType.Holiday, text: 'Holiday 🎄', icon: '🎄' },
      { key: CelebrationType.Custom, text: 'Custom ✨', icon: '✨' }
    ];
  }
}
