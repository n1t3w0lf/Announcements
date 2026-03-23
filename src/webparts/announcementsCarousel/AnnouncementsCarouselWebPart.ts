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
  enableTransitions: boolean;
  transitionEffect: 'fade' | 'slide' | 'zoom';
  showNavigationDots: boolean;
  showNavigationArrows: boolean;
  autoPlay: boolean;
}

export default class AnnouncementsCarouselWebPart extends BaseClientSideWebPart<IAnnouncementsCarouselWebPartProps> {

  protected async onInit(): Promise<void> {
    await super.onInit();

    if (!this.properties.listName) {
      this.properties.listName = ListProvisioningService.getListName();
    }

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
        enableTransitions: this.properties.enableTransitions !== false,
        transitionEffect: this.properties.transitionEffect || 'fade',
        showNavigationDots: this.properties.showNavigationDots !== false,
        showNavigationArrows: this.properties.showNavigationArrows !== false,
        autoPlay: this.properties.autoPlay !== false
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
            },
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
            }
          ]
        }
      ]
    };
  }
}
