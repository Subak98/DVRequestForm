import { IList } from "../common/IObject";

export interface IServiceProvider {
  saveDraftRequest(listid: string, request: any): Promise<IList[]>;
  EditRequest(listid: string, request: any): Promise<IList[]>;
  ChangeStatus(
    listid: string,
    Id: any,
    status?: string,
    comments?: string,
    useremail?: string,
  ): Promise<IList[]>;
  fetchDepartments(sitetype: string): Promise<string[]>;
  fetchFormDescription(formDesclistid: string): Promise<any[]>;
  fetchUniqueDepartment(filter: string): Promise<any>;
  getReviewData(listName: string, Id: number, filter?: string): Promise<any>;
  sendEmail(
    context: any,
    siteName: string,
    siteDescription: string,
    primaryOwners: any[],
    secondaryOwners: any[],
    department: string,
    approver: any,
    reason: string,
    approverOptions: any[],
    Id: any,
    PAURL: string,
  ): Promise<any>;
}
