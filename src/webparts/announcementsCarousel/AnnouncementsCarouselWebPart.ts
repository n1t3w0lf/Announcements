import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AnnouncementsCarousel from './components/AnnouncementsCarousel';
import { IAnnouncementsCarouselProps } from './components/IAnnouncementsCarouselProps';
import { ListProvisioningService } from './services/ListProvisioningService';

export interface IAnnouncementsCarouselWebPartProps {
  listName: string;
  rotationInterval: number;
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

export default class AnnouncementsCarouselWebPart extends BaseClientSideWebPart<IAnnouncementsCarouselWebPartProps> {

  protected async onInit(): Promise<void> {
    await super.onInit();

    // Set default values if not set
    if (!this.properties.listName) {
      this.properties.listName = ListProvisioningService.getListName();
    }

    // Ensure list exists on first load
    try {
      await ListProvisioningService.ensureList(this.context);
    } catch (error) {
      console.error('Error ensuring list exists:', error);
    }
  }

  public render(): void {
    const element: React.ReactElement<IAnnouncementsCarouselProps> = React.createElement(
      AnnouncementsCarousel,
      {
        context: this.context,
        listName: this.properties.listName || ListProvisioningService.getListName(),
        rotationInterval: this.properties.rotationInterval || 10,
        showTitle: this.properties.showTitle !== false,
        showDescription: this.properties.showDescription !== false,
        titleFontSize: this.properties.titleFontSize || 32,
        descriptionFontSize: this.properties.descriptionFontSize || 16,
        backgroundColor: this.properties.backgroundColor || '#ffffff',
        titleColor: this.properties.titleColor || '#333333',
        descriptionColor: this.properties.descriptionColor || '#333333',
        height: this.properties.height || 500,
        enableTransitions: this.properties.enableTransitions !== false,
        transitionEffect: this.properties.transitionEffect || 'fade',
        showNavigationDots: this.properties.showNavigationDots !== false,
        showNavigationArrows: this.properties.showNavigationArrows !== false,
        autoPlay: this.properties.autoPlay !== false,
        celebrationIconSize: this.properties.celebrationIconSize || 60,
        showCelebrationIcon: this.properties.showCelebrationIcon !== false,
        borderRadius: this.properties.borderRadius || 8,
        showShadow: this.properties.showShadow !== false,
        overlayOpacity: this.properties.overlayOpacity || 30
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure your announcements carousel settings'
          },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'List Name',
                  description: 'SharePoint list name for announcements',
                  value: this.properties.listName || ListProvisioningService.getListName()
                })
              ]
            },
            {
              groupName: 'Carousel Behavior',
              groupFields: [
                PropertyPaneToggle('autoPlay', {
                  label: 'Auto Play',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.autoPlay !== false
                }),
                PropertyPaneSlider('rotationInterval', {
                  label: 'Rotation Interval (seconds)',
                  min: 3,
                  max: 60,
                  value: this.properties.rotationInterval || 10,
                  showValue: true,
                  step: 1
                }),
                PropertyPaneToggle('enableTransitions', {
                  label: 'Enable Transitions',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.enableTransitions !== false
                }),
                PropertyPaneDropdown('transitionEffect', {
                  label: 'Transition Effect',
                  options: [
                    { key: 'fade', text: 'Fade' },
                    { key: 'slide', text: 'Slide' },
                    { key: 'zoom', text: 'Zoom' }
                  ],
                  selectedKey: this.properties.transitionEffect || 'fade'
                })
              ]
            },
            {
              groupName: 'Navigation',
              groupFields: [
                PropertyPaneToggle('showNavigationDots', {
                  label: 'Show Navigation Dots',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.showNavigationDots !== false
                }),
                PropertyPaneToggle('showNavigationArrows', {
                  label: 'Show Navigation Arrows',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.showNavigationArrows !== false
                })
              ]
            }
          ]
        },
        {
          header: {
            description: 'Customize the appearance of your announcements'
          },
          groups: [
            {
              groupName: 'Content Display',
              groupFields: [
                PropertyPaneToggle('showTitle', {
                  label: 'Show Title',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.showTitle !== false
                }),
                PropertyPaneSlider('titleFontSize', {
                  label: 'Title Font Size (px)',
                  min: 16,
                  max: 72,
                  value: this.properties.titleFontSize || 32,
                  showValue: true,
                  step: 2
                }),
                PropertyPaneToggle('showDescription', {
                  label: 'Show Description',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.showDescription !== false
                }),
                PropertyPaneSlider('descriptionFontSize', {
                  label: 'Description Font Size (px)',
                  min: 12,
                  max: 32,
                  value: this.properties.descriptionFontSize || 16,
                  showValue: true,
                  step: 1
                })
              ]
            },
            {
              groupName: 'Colors & Styling',
              groupFields: [
                PropertyPaneTextField('backgroundColor', {
                  label: 'Background Color',
                  description: 'Hex color code (e.g., #ffffff)',
                  value: this.properties.backgroundColor || '#ffffff'
                }),
                PropertyPaneTextField('titleColor', {
                  label: 'Title Color',
                  description: 'Hex color code (e.g., #333333)',
                  value: this.properties.titleColor || '#333333'
                }),
                PropertyPaneTextField('descriptionColor', {
                  label: 'Description Color',
                  description: 'Hex color code (e.g., #333333)',
                  value: this.properties.descriptionColor || '#333333'
                }),
                PropertyPaneSlider('overlayOpacity', {
                  label: 'Image Overlay Opacity (%)',
                  min: 0,
                  max: 100,
                  value: this.properties.overlayOpacity || 30,
                  showValue: true,
                  step: 5
                })
              ]
            },
            {
              groupName: 'Dimensions & Effects',
              groupFields: [
                PropertyPaneSlider('height', {
                  label: 'Carousel Height (px)',
                  min: 300,
                  max: 800,
                  value: this.properties.height || 500,
                  showValue: true,
                  step: 50
                }),
                PropertyPaneSlider('borderRadius', {
                  label: 'Border Radius (px)',
                  min: 0,
                  max: 50,
                  value: this.properties.borderRadius || 8,
                  showValue: true,
                  step: 2
                }),
                PropertyPaneToggle('showShadow', {
                  label: 'Show Shadow',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.showShadow !== false
                })
              ]
            }
          ]
        },
        {
          header: {
            description: 'Configure celebration icon settings'
          },
          groups: [
            {
              groupName: 'Celebration Icons',
              groupFields: [
                PropertyPaneToggle('showCelebrationIcon', {
                  label: 'Show Celebration Icons',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.showCelebrationIcon !== false
                }),
                PropertyPaneSlider('celebrationIconSize', {
                  label: 'Icon Size (px)',
                  min: 30,
                  max: 120,
                  value: this.properties.celebrationIconSize || 60,
                  showValue: true,
                  step: 10
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
