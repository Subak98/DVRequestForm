import { IServiceProvider } from "./IServiceProvider";

import { BaseWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";

import { SPFI } from "@pnp/sp";
import { getSP } from "./pnpjsConfig";
import { IList } from "../common/IObject";
import moment from "moment";
import { Field } from "@pnp/sp/fields/types";
import {
  IHttpClientOptions,
  HttpClientResponse,
  HttpClient,
} from "@microsoft/sp-http";
export default class SharepointServiceProvider implements IServiceProvider {
  _webPartContext: BaseWebPartContext;
  _webAbsoluteUrl: string;
  public sp: SPFI;
  //   private _webPartContext: BaseWebPartContext;
  //   private _webAbsoluteUrl: string;

  constructor(_context: BaseWebPartContext) {
    this._webPartContext = _context;
    this._webAbsoluteUrl = _context.pageContext.web.absoluteUrl;
    this.sp = getSP(_context as WebPartContext);
  }

  public async getReviewData(
    listid: string,
    itemId?: number,
    filter?: string,
  ): Promise<IList[]> {
    let _items: IList[];
    let getitems: any[];
    if (itemId) {
      getitems = await this.sp.web.lists
        .getById(listid)
        .items.getById(itemId)
        .select(
          "Title,SiteDescription,SiteJustification,SiteType,Status,Department,Created,Author/Title,PrimaryOwner/Title,Approvers/Title,SecondaryOwner/Title,PrimaryOwner/EMail,Approvers/EMail,SecondaryOwner/EMail,Id,ApprovedBy/Title",
        )
        .expand("Author,PrimaryOwner,Approvers,SecondaryOwner,ApprovedBy")();

      // Convert single item to array
      getitems = [getitems];
    } else {
      let query: any = this.sp.web.lists.getById(listid).items;
      if (filter) {
        query = query.filter(filter);
      }

      getitems = await query
        .select(
          "Title,SiteDescription,SiteJustification,SiteType,Status,Department,Created,Author/Title,PrimaryOwner/Title,Approvers/Title,SecondaryOwner/Title,PrimaryOwner/EMail,Approvers/EMail,SecondaryOwner/EMail,Id,ApprovedBy/Title",
        )
        .expand("Author,PrimaryOwner,Approvers,SecondaryOwner,ApprovedBy")
        .orderBy("Modified", false)();
    }
    _items = [];
    console.log(getitems);

    for (let i = 0; i < getitems.length; i++) {
      let item = getitems[i];
      let lst: any = {
        Title: item.Title,
        Description: item.SiteDescription,
        SiteJustification: item.SiteJustification,
        Status: item.Status,
        SiteType: item.SiteType,
        Department: item.Department,
        Date: moment(item.Created).format("MM/DD/YYYY"),
        Requestor: item.Author.Title,
        ApprovedBy: item.ApprovedBy,
        PrimaryOwnerTitle: item.PrimaryOwner
          ? item.PrimaryOwner.map(
              (owner: { Title: string }) => owner.Title,
            ).join(", ")
          : "",
        PrimaryOwnerEmail: item.PrimaryOwner
          ? item.PrimaryOwner.map((owner: { EMail: string }) => owner.EMail)
          : [],
        PrimaryOwner: item.PrimaryOwner
          ? item.PrimaryOwner.map(
              (owner: { Title: string; EMail: string }) => ({
                Title: owner.Title,
                Email: owner.EMail,
              }),
            )
          : [],
        Approver: item.Approvers
          ? item.Approvers.map(
              (approver: { Title: string; EMail: string }) => ({
                Title: approver.Title,
                Email: approver.EMail,
              }),
            ).join(", ")
          : "",
        ApproverTitle: item.Approvers
          ? item.Approvers.map(
              (approver: { Title: string }) => approver.Title,
            ).join(", ")
          : "",
        ApproverEmail: item.Approvers
          ? item.Approvers.map((approver: { EMail: string }) => approver.EMail)
          : [],
        SecondaryOwnerTitle: item.SecondaryOwner
          ? item.SecondaryOwner.map(
              (owner: { Title: string }) => owner.Title,
            ).join(", ")
          : "",
        SecondaryOwnerEmail: item.SecondaryOwner
          ? item.SecondaryOwner.map((owner: { EMail: string }) => owner.EMail)
          : [],
        SecondaryOwner: item.SecondaryOwner
          ? item.SecondaryOwner.map(
              (owner: { Title: string; EMail: string }) => ({
                Title: owner.Title,
                Email: owner.EMail,
              }),
            )
          : [],
        // Order0: item.Order0,

        Id: item.Id,
      };

      _items.push(lst);
    }
    // console.log(_items);

    return _items;
  }
  public async saveDraftRequest(listid: string, request: any): Promise<any[]> {
    // console.log(request);

    try {
      let newresponse = await this.sp.web.lists
        .getById(listid)
        .items.add({
          Title: request.siteName,
          SiteDescription: request.siteDescription,
          SiteJustification: request.reason,
          Department: request.department,
          SiteType: request.sitetype,
          Status: request.status,
        })
        .then(async (response: any) => {
          // console.log(response);

          await this.sp.web.lists
            .getById(listid)
            .items.getById(response.Id)
            .validateUpdateListItem(
              [
                ...(request.primaryOwners && request.primaryOwners.length > 0
                  ? [
                      {
                        FieldName: "PrimaryOwner",
                        FieldValue: JSON.stringify(
                          request.primaryOwners.map((person: { id: any }) => ({
                            Key: person.id,
                          })),
                        ),
                      },
                    ]
                  : []),
                ...(request.approverOptions &&
                request.approverOptions.length > 0
                  ? [
                      {
                        FieldName: "Approvers",
                        FieldValue: JSON.stringify(
                          request.approverOptions.map((email: string) => ({
                            Key: `i:0#.f|membership|${email}`,
                          })),
                        ),
                      },
                    ]
                  : []),
                ...(request.secondaryOwners &&
                request.secondaryOwners.length > 0
                  ? [
                      {
                        FieldName: "SecondaryOwner",
                        FieldValue: JSON.stringify(
                          request.secondaryOwners.map(
                            (person: { id: any }) => ({
                              Key: person.id,
                            }),
                          ),
                        ),
                      },
                    ]
                  : []),
              ],
              true,
            );
          return response;
        });
      return newresponse;
      // Item creation result is not returned since the method's return type is void
    } catch (error) {
      console.error("Error adding item with validation:", error);
      throw error;
    }
  }
  public async EditRequest(listid: string, request: any): Promise<any[]> {
    console.log(request, "request");
    try {
      await this.sp.web.lists
        .getById(listid)
        .items.getById(request.id)
        .validateUpdateListItem(
          [
            {
              FieldName: "Title",
              FieldValue: request.siteName,
            },
            {
              FieldName: "SiteType",
              FieldValue: request.SiteType,
            },
            {
              FieldName: "SiteDescription",
              FieldValue: request.siteDescription,
            },
            {
              FieldName: "SiteJustification",
              FieldValue: request.reason,
            },
            {
              FieldName: "Department",
              FieldValue: request.department,
            },
            {
              FieldName: "PrimaryOwner",
              FieldValue:
                request.primaryOwners.length > 0
                  ? JSON.stringify(
                      request.primaryOwners.map((person: { id: any }) => ({
                        Key: person.id
                          ? person.id
                          : `i:0#.f|membership|${person}`,
                      })),
                    )
                  : "",
            },
            // ...(request.primaryOwners && request.primaryOwners.length > 0
            //   ? [
            //       {
            //         FieldName: "PrimaryOwner",
            //         FieldValue: JSON.stringify(
            //           request.primaryOwners.map((person: { id: any }) => ({
            //             Key: person.id
            //               ? person.id
            //               : `i:0#.f|membership|${person}`,
            //           })),
            //         ),
            //       },
            //     ]
            //   : []),
            ...(request.approverOptions && request.approverOptions.length > 0
              ? [
                  {
                    FieldName: "Approvers",
                    FieldValue: JSON.stringify(
                      request.approverOptions.map((email: string) => ({
                        Key: `i:0#.f|membership|${email}`,
                      })),
                    ),
                  },
                ]
              : []),
            {
              FieldName: "SecondaryOwner",
              FieldValue:
                request.secondaryOwners.length > 0
                  ? JSON.stringify(
                      request.secondaryOwners.map((person: { id: any }) => ({
                        Key: person.id
                          ? person.id
                          : `i:0#.f|membership|${person}`,
                      })),
                    )
                  : "",
            },
            // ...(request.secondaryOwners && request.secondaryOwners.length > 0
            //   ? [
            //       {
            //         FieldName: "SecondaryOwner",
            //         FieldValue: JSON.stringify(
            //           request.secondaryOwners.map((person: { id: any }) => ({
            //             Key: person.id
            //               ? person.id
            //               : `i:0#.f|membership|${person}`,
            //           })),
            //         ),
            //       },
            //     ]
            //   : []),
          ],
          true,
        );

      // Item creation result is not returned since the method's return type is void
    } catch (error) {
      console.error("Error adding item with validation:", error);
      throw error;
    }
    return ["Draft saved successfully!"];
  }

  public async ChangeStatus(
    listid: string,
    Id: any,
    status?: string,
    comments?: string,
    useremail?: string,
  ): Promise<any[]> {
    try {
      await this.sp.web.lists
        .getById(listid)
        .items.getById(Id)
        .validateUpdateListItem(
          [
            {
              FieldName: "Status",
              FieldValue: status || "Submitted",
            },
            {
              FieldName: "Comments",
              FieldValue: comments || "",
            },
            ...(useremail && status == "Approved"
              ? [
                  {
                    FieldName: "ApprovedBy",
                    FieldValue: JSON.stringify([
                      {
                        Key: `i:0#.f|membership|${useremail}`,
                      },
                    ]),
                  },
                ]
              : []),
          ],
          true,
        );

      // Item creation result is not returned since the method's return type is void
    } catch (error) {
      console.error("Error adding item with validation:", error);
      throw error;
    }
    return ["Draft saved successfully!"];
  }
  public async fetchDepartments(): Promise<string[]> {
    // Simulate fetching departments from SharePoint
    //   let _items: IList[];
    const getitems = await this.sp.web.lists.getByTitle("Departments").items();
    // console.log("getitems", getitems);
    return getitems.map((item: any) => item.Title);
  }
  public async fetchFormDescription(formDesclistid: string): Promise<string[]> {
    // Simulate fetching departments from SharePoint
    //   let _items: IList[];
    console.log(formDesclistid);

    const getitems = await this.sp.web.lists
      .getById(formDesclistid)
      .items.select("Description", "ID")
      .top(1)();
    // console.log("getitems", getitems);
    return getitems;
  }
  public async fetchUniqueDepartment(filter: string): Promise<string[]> {
    const getitems = await this.sp.web.lists
      .getByTitle("Departments")
      .items.filter(filter)();
    return getitems;
  }
  public sendEmail(
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
  ): Promise<HttpClientResponse> {
    const postURL = PAURL;
    // "https://defaultb14a476cad7448f8986f8942642cad.2e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f645d31e5aa448999fc7a1d80294e8a1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=iit1NX7jswZ2xWpxRJkqrmebmxlr9ObSheoq5JndZrY";
    const body: string = JSON.stringify({
      title: siteName,
      department: department,
      reason: reason,
      userEmail: context.pageContext.user.email,
      primaryOwners: primaryOwners,
      secondaryOwners: secondaryOwners,
      Id: Id,
    });
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const httpClientOptions: IHttpClientOptions = {
      body: body,
      headers: requestHeaders,
    };
    console.log("Sending Email");
    return context.httpClient
      .post(postURL, HttpClient.configurations.v1, httpClientOptions)
      .then((response: HttpClientResponse): Promise<HttpClientResponse> => {
        console.log("Email sent.");
        return response.json();
      });
  }
}
