import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IIntranetEvent {
  Id: number;
  Title: string;
  EventDate: string;
  EndDate: string;
  Location?: string;
  Description?: string;
  Category?: string;
}

export interface IIntranetEventProps {
  description?: string;
  listTitle: string;
  context: WebPartContext;
   isDarkTheme: boolean;
    webPartTitle: string;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  // Color properties
  headerBgColor?: string;
  headerTextColor?: string;
  dateBgColor?: string;
  dateTextColor?: string;
  eventBgColor?: string;
  eventBgColorAlt?: string;  // Alternate row color
  eventTextColor?: string;
  categoryBgColor?: string;
  categoryTextColor?: string;
  // Row limit configuration
  maxRows?: number;
  timeLocationColor?: string; 
  rowHeight?: string;

}