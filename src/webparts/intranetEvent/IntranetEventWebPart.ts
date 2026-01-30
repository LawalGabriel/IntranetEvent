/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,PropertyPaneSlider
  
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

//import * as strings from 'IntranetEventWebPartStrings';
import IntranetEvent from './components/IntranetEvent';
import { IIntranetEventProps } from './components/IIntranetEventProps';

export interface IIntranetEventWebPartProps {
  webPartTitle: string;
  rowHeight: string;
  maxRows: number | undefined;
  description: string;
  listTitle: string;
  headerBgColor: string;
  headerTextColor: string;
  dateBgColor: string;
  dateTextColor: string;
  eventBgColor: string;
  eventTextColor: string;
  categoryBgColor: string;
  categoryTextColor: string;
}

export default class IntranetEventWebPart extends BaseClientSideWebPart<IIntranetEventWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IIntranetEventProps> = React.createElement(
  IntranetEvent,
  {
    context: this.context,
    listTitle: this.properties.listTitle || "Events",  // Add this line
    headerBgColor: this.properties.headerBgColor,
    headerTextColor: this.properties.headerTextColor,
    dateBgColor: this.properties.dateBgColor,
    dateTextColor: this.properties.dateTextColor,
    eventBgColor: this.properties.eventBgColor,
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
    webPartTitle: this.properties.webPartTitle || 'EVENTS'
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
          description: 'Configure Events Web Part'
        },
        groups: [
          {
            groupName: 'Basic Settings',
            groupFields: [
              PropertyPaneTextField('listTitle', {
                label: 'List Title',
                value: this.properties.listTitle
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
  value: 'EVENTS'
})
              
            ]
          },
          {
            groupName: 'Header Colors',
            groupFields: [
              PropertyPaneTextField('headerBgColor', {
                label: 'Background Color',
                description: 'Enter hex color code (e.g., #2c3e50)'
              }),
              PropertyPaneTextField('headerTextColor', {
                label: 'Text Color',
                description: 'Enter hex color code (e.g., #ffffff)'
              })
            ]
          },
          {
            groupName: 'Event Row Colors',
            groupFields: [
              PropertyPaneTextField('eventBgColor', {
                label: 'Even Row Background',
                description: 'Enter hex color code (e.g., #ffffff)'
              }),
              PropertyPaneTextField('eventBgColorAlt', {
                label: 'Odd Row Background',
                description: 'Enter hex color code (e.g., #f8f9fa)'
              }),
              PropertyPaneTextField('eventTextColor', {
                label: 'Text Color',
                description: 'Enter hex color code (e.g., #333333)'
              })
            ]
          },
          {
            groupName: 'Date Box Colors',
            groupFields: [
              PropertyPaneTextField('dateBgColor', {
                label: 'Background Color',
                description: 'Enter hex color code (e.g., #e74c3c)'
              }),
              PropertyPaneTextField('dateTextColor', {
                label: 'Text Color',
                description: 'Enter hex color code (e.g., #ffffff)'
              })
            ]
          },
          {
            groupName: 'Category & Time Colors',
            groupFields: [
              PropertyPaneTextField('categoryBgColor', {
                label: 'Category Background',
                description: 'Enter hex color code (e.g., #3498db)'
              }),
              PropertyPaneTextField('categoryTextColor', {
                label: 'Category Text Color',
                description: 'Enter hex color code (e.g., #ffffff)'
              }),
              PropertyPaneTextField('timeLocationColor', {
                label: 'Time & Location Color',
                description: 'Enter hex color code (e.g., #666666)'
              })
            ]
          }
        ]
      }
    ]
  };
}

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (newValue !== oldValue) {
      this.render();
    }
  }
}