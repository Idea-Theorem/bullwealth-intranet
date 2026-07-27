import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeDirectoryProps {
  title: string;
  maxEmployeesToShow: number;
  orgChartLink: string;
  listName: string;
  selectedCompany: string; // ✅ NEW: Pre-selected company filter
  context: WebPartContext;
}

export interface IEmployee {
  id: number;
  name: string;
  title: string; // CompanyName field
  email: string;
  phone: string;
  profileImage: string;
  groupName: string;
}
