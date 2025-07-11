import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITeamCalendarProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
