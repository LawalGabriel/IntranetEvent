/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

//import * as strings from 'IntranetEventWebPartStrings';
import IntranetEvent from './components/IntranetEvent';
import { IIntranetEventProps } from './components/IIntranetEventProps';

export interface IIntranetEventWebPartProps {
  description: string;
  listTitle: string;
  headerBgColor: string;
  headerTextColor: string;
  headerFontSize: string;
  headerFontWeight: string;
  dateBgColor: string;
  dateTextColor: string;
  eventBgColor: string;
  eventBgColorAlt: string;
  eventTextColor: string;
  categoryBgColor: string;
  categoryTextColor: string;
  maxRows: number;
  rowHeight: string;
  webPartTitle: string;
  timeLocationColor: string;
}

export default class IntranetEventWebPart extends BaseClientSideWebPart<IIntranetEventWebPartProps> {
  
  protected async onInit(): Promise<void> {
    await super.onInit();
    
    // Initialize default values if not set
    if (!this.properties.headerFontSize) {
      this.properties.headerFontSize = 'clamp(18px, 2vw, 24px)';
    }
    if (!this.properties.headerFontWeight) {
      this.properties.headerFontWeight = '600';
    }
    if (!this.properties.eventBgColorAlt) {
      this.properties.eventBgColorAlt = '#f8f9fa';
    }
  }

  public render(): void {
    const element: React.ReactElement<IIntranetEventProps> = React.createElement(
      IntranetEvent,
      {
        context: this.context,
        listTitle: this.properties.listTitle || "Events",
        headerBgColor: this.properties.headerBgColor,
        headerTextColor: this.properties.headerTextColor,
        headerFontSize: this.properties.headerFontSize,
        headerFontWeight: this.properties.headerFontWeight,
        dateBgColor: this.properties.dateBgColor,
        dateTextColor: this.properties.dateTextColor,
        eventBgColor: this.properties.eventBgColor,
        eventBgColorAlt: this.properties.eventBgColorAlt,
        eventTextColor: this.properties.eventTextColor,
        categoryBgColor: this.properties.categoryBgColor,
        categoryTextColor: this.properties.categoryTextColor,
        description: this.properties.description,
        isDarkTheme: !!this.context.sdks?.microsoftTeams?.context && this.context.sdks.microsoftTeams.context.theme === 'dark',
        environmentMessage: '',
        hasTeamsContext: !!this.context.sdks?.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        maxRows: this.properties.maxRows,
        rowHeight: this.properties.rowHeight,
        webPartTitle: this.properties.webPartTitle || 'EVENTS',
        timeLocationColor: this.properties.timeLocationColor
      }
    );
    
    ReactDom.render(element, this.domElement);
  }

  // REMOVE THIS ENTIRE METHOD - IT'S CAUSING THE ERROR:
  // domElement(element: ReactElement<IIntranetEventProps, string | JSXElementConstructor<any>>, domElement: any) {
  //   throw new Error('Method not implemented.')
  // }

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
            description: 'Configure Events Web Part'
          },
          groups: [
            {
              groupName: 'Basic Settings',
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: 'List Title',
                  value: this.properties.listTitle || 'Events'
                }),
                PropertyPaneSlider('maxRows', {
                  label: 'Maximum Rows (before scrolling)',
                  min: 1,
                  max: 20,
                  value: this.properties.maxRows || 4,
                  showValue: true
                }),
                PropertyPaneTextField('rowHeight', {
                  label: 'Row Height',
                  description: 'Enter height with unit (e.g., 70px, 5rem)',
                  value: this.properties.rowHeight || '70px'
                }),
                PropertyPaneTextField('webPartTitle', {
                  label: 'Web Part Title',
                  description: 'Enter the title to display in the web part header',
                  value: this.properties.webPartTitle || 'EVENTS'
                })
              ]
            },
            {
              groupName: 'Header Styling',
              groupFields: [
                PropertyPaneTextField('headerBgColor', {
                  label: 'Background Color',
                  description: 'Enter hex color code (e.g., #2c3e50)',
                  value: this.properties.headerBgColor || '#2c3e50'
                }),
                PropertyPaneTextField('headerTextColor', {
                  label: 'Text Color',
                  description: 'Enter hex color code (e.g., #ffffff)',
                  value: this.properties.headerTextColor || '#ffffff'
                }),
                PropertyPaneTextField('headerFontSize', {
                  label: 'Font Size',
                  description: 'Enter font size with unit (e.g., 24px, 1.5rem, clamp(18px, 2vw, 24px))',
                  value: this.properties.headerFontSize || 'clamp(18px, 2vw, 24px)'
                }),
                PropertyPaneTextField('headerFontWeight', {
                  label: 'Font Weight',
                  description: 'Enter font weight (e.g., 600, bold, normal)',
                  value: this.properties.headerFontWeight || '600'
                })
              ]
            },
            {
              groupName: 'Event Row Colors',
              groupFields: [
                PropertyPaneTextField('eventBgColor', {
                  label: 'Even Row Background',
                  description: 'Enter hex color code (e.g., #ffffff)',
                  value: this.properties.eventBgColor || '#ffffff'
                }),
                PropertyPaneTextField('eventBgColorAlt', {
                  label: 'Odd Row Background',
                  description: 'Enter hex color code (e.g., #f8f9fa)',
                  value: this.properties.eventBgColorAlt || '#f8f9fa'
                }),
                PropertyPaneTextField('eventTextColor', {
                  label: 'Text Color',
                  description: 'Enter hex color code (e.g., #333333)',
                  value: this.properties.eventTextColor || '#333333'
                })
              ]
            },
            {
              groupName: 'Date Box Colors',
              groupFields: [
                PropertyPaneTextField('dateBgColor', {
                  label: 'Background Color',
                  description: 'Enter hex color code (e.g., #e74c3c)',
                  value: this.properties.dateBgColor || '#e74c3c'
                }),
                PropertyPaneTextField('dateTextColor', {
                  label: 'Text Color',
                  description: 'Enter hex color code (e.g., #ffffff)',
                  value: this.properties.dateTextColor || '#ffffff'
                })
              ]
            },
            {
              groupName: 'Category & Time Colors',
              groupFields: [
                PropertyPaneTextField('categoryBgColor', {
                  label: 'Category Background',
                  description: 'Enter hex color code (e.g., #3498db)',
                  value: this.properties.categoryBgColor || '#3498db'
                }),
                PropertyPaneTextField('categoryTextColor', {
                  label: 'Category Text Color',
                  description: 'Enter hex color code (e.g., #ffffff)',
                  value: this.properties.categoryTextColor || '#ffffff'
                }),
                PropertyPaneTextField('timeLocationColor', {
                  label: 'Time & Location Color',
                  description: 'Enter hex color code (e.g., #666666)',
                  value: this.properties.timeLocationColor || '#666666'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    
    // Re-render the web part when properties change
    this.render();
  }
}