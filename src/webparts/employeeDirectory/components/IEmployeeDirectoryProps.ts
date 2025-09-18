import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeDirectoryProps {
  title: string;
  maxEmployeesToShow: number;
  orgChartLink: string;
  sections: ISection[];
  context: WebPartContext;
}

export interface ISection {
  title: string;
  listName: string;
}

export interface IEmployee {
  id: number;
  name: string;
  title: string;
  email: string;
  phone: string;
  profileImage: string;
}
