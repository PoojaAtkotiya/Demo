import { IWebPartContext } from "@microsoft/sp-webpart-base";

export interface IViewRequestProps {
  description: string;
  context:IWebPartContext;
}


export interface ISPList {
  ID: string;
  Title:string;
  Category:string;
  SubCategory :string;
  Description:string;
  Status: string;
  ApproverComment : string;
  ApprovedBy :string;
  ApprovedRejectedDate :string;
  } 