import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeDirectoryProps {
  title: string;
  maxEmployeesToShow: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}

export interface IEmployee {
  id: number;
  name: string;
  title: string;
  email: string;
  phone: string;
  profileImage: string;
  department?: string;
}