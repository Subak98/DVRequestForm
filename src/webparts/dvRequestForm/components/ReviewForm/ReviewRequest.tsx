import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  DetailsList,
  IColumn,
  IconButton,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  Dropdown,
  PrimaryButton,
  DefaultButton,
  SelectionMode,
} from "@fluentui/react";
import { SPHttpClient } from "@microsoft/sp-http";

import styles from "./ReviewRequest.module.scss";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
interface IRequestRow {
  key: string;
  Title: string;
  owner: string;
  Date: string;
  Department: string;
  Description: string;
  SiteJustification: string;
  SiteType?: string;
  PrimaryOwnerTitle: any;
  SecondaryOwnerEmail: any;
  PrimaryOwnerEmail: any;
  SecondaryOwnerTitle: string;
  primaryOwners: any;
  ApproverTitle: string;
  Id: number;
  Status: "Draft" | "Submitted" | "Approved" | "Rejected" | "InProgress";
}
interface IApproverOption {
  key: string;
  text: string;
  value: string;
}

interface IReviewRequestProps {
  props: any;
}
const ReviewRequest: React.FC<IReviewRequestProps> = (props) => {
  const [requests, setRequests] = useState<IRequestRow[]>([]);
  const [message, setMessage] = useState<string | undefined>();
  const [messageType, setMessageType] = useState<MessageBarType>(
    MessageBarType.success,
  );
  const [isEditDialogOpen, setIsEditDialogOpen] = useState(false);

  React.useEffect(() => {
    if (!message) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      setMessage(undefined);
    }, 2000);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [message]);
  const [editingItem, setEditingItem] = useState<any>(null);
  const [editFormData, setEditFormData] = useState<any>(null);
  const [isSuccessDialogOpen, setIsSuccessDialogOpen] = useState(false);
  const [submittedTitle, setSubmittedTitle] = useState<string | null>(null);
  const [isValidationDialogOpen, setIsValidationDialogOpen] = useState(false);
  const [validationMessage, setValidationMessage] = useState<string | null>(
    null,
  );
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 10;
  const [primaryOwners, setPrimaryOwners] = React.useState<any[]>([]);
  const [primaryEmails, setPrimaryEmails] = React.useState<any[]>([]);

  const [secondaryOwners, setSecondaryOwners] = React.useState<any[]>([]);
  const [secondaryEmails, setSecondaryEmails] = React.useState<any[]>([]);

  const [departmentOptions, setDepartmentOptions] = useState<IApproverOption[]>(
    [],
  );
  const [department, setDepartment] = useState(editFormData?.Department || "");
  const [approverOptions, setApproverOptions] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.props.context.msGraphClientFactory,
    spHttpClient: props.props.context.spHttpClient,
  };
  const handleEdit = (item: IRequestRow) => {
    console.log(item);
    setDepartment(item.Department);
    setPrimaryOwners(item.PrimaryOwnerTitle);
    setPrimaryEmails(item.PrimaryOwnerEmail);
    setSecondaryEmails(item.SecondaryOwnerEmail);
    setEditingItem(item);
    setEditFormData({ ...item });
    setIsEditDialogOpen(true);
  };
  const getReviewData = async () => {
    await props.props.provider
      .getReviewData(
        props.props.RequestFormId,
        "",
        `Author/EMail eq '${props.props.userEmail}' or PrimaryOwner/EMail eq '${props.props.userEmail}'`,
      )
      .then((response: any) => {
        console.log("response", response);
        setRequests(response);
      });
  };
  const handleSendRequest = async (item: IRequestRow) => {
    console.log(item);

    // validate required fields before submitting
    if (item.Status === "Submitted") {
      setMessageType(MessageBarType.error);
      setMessage(`Request "${item.Title}" has already been submitted.`);
      return;
    }

    const requiredChecks: { name: string; value: any }[] = [
      { name: "Site Title", value: item.Title },
      { name: "Primary Owner", value: item.PrimaryOwnerTitle },
      { name: "Department", value: item.Department },
      { name: "Approver", value: item.ApproverTitle },
      { name: "Justification", value: item.SiteJustification },
    ];

    const missing = requiredChecks
      .filter(
        (c) => !c.value || (Array.isArray(c.value) && c.value.length === 0),
      )
      .map((c) => c.name);
    if (missing.length > 0) {
      setValidationMessage(
        `Please complete required fields: ${missing.join(", ")}`,
      );
      setIsValidationDialogOpen(true);
      return;
    }
    await props.props.provider
      .ChangeStatus(props.props.RequestFormId, item.Id)
      .then(async (response: any) => {
        setMessageType(MessageBarType.success);
        getReviewData().catch((error) => {
          setMessageType(MessageBarType.error);
          setMessage("Error refreshing data after save.");
          console.error(error);
        });
        setIsEditDialogOpen(false);
        setEditingItem(null);
        setEditFormData(null);
        await props.props.provider.sendEmail(
          props.props.context,
          item.Title,
          "",
          item.PrimaryOwnerTitle,
          item.SecondaryOwnerTitle,
          department,
          [...approverOptions],
          item.SiteJustification,
          item.ApproverTitle,
          item.Id,
          props.props.PAURL,
        );
      });
    // mark as submitted locally and show success popup
    // setRequests((current) =>
    //   current.map((row) =>
    //     row.key === item.key ? { ...row, status: "Submitted" } : row,
    //   ),
    // );
    setMessageType(MessageBarType.success);
    setMessage(undefined);
    setSubmittedTitle(item.Title);
    setIsSuccessDialogOpen(true);
  };

  const renderTooltipCell =
    (fieldName: keyof IRequestRow) => (item: IRequestRow) => (
      <span title={item[fieldName] != null ? String(item[fieldName]) : ""}>
        {item[fieldName]}
      </span>
    );

  const columns: IColumn[] = useMemo(
    () => [
      {
        key: "column1",
        name: "Site Title",
        fieldName: "Title",
        minWidth: 150,
        maxWidth: 220,
        isResizable: true,
        onRender: renderTooltipCell("Title"),
      },
      {
        key: "column2",
        name: "Primary Owner",
        fieldName: "PrimaryOwnerTitle",
        minWidth: 140,
        maxWidth: 180,
        isResizable: true,
        onRender: renderTooltipCell("PrimaryOwnerTitle"),
      },
      {
        key: "column3",
        name: "Date",
        fieldName: "Date",
        minWidth: 100,
        maxWidth: 120,
        // onRender: renderTooltipCell("Date"),
      },
      {
        key: "column4",
        name: "Approvers",
        fieldName: "ApproverTitle",
        minWidth: 120,
        maxWidth: 140,
        onRender: renderTooltipCell("ApproverTitle"),
      },
      // {
      //   key: "column5",
      //   name: "Approver",
      //   fieldName: "approver",
      //   minWidth: 160,
      //   maxWidth: 180,
      // },

      {
        key: "column5",
        name: "Site Justification",
        fieldName: "SiteJustification",
        minWidth: 120,
        maxWidth: 140,
        onRender: renderTooltipCell("SiteJustification"),
      },
      {
        key: "column6",
        name: "Status",
        fieldName: "Status",
        minWidth: 120,
        maxWidth: 140,
        // onRender: renderTooltipCell("Status"),
      },
      {
        key: "column7",
        name: "Edit",
        minWidth: 80,
        maxWidth: 100,
        isResizable: false,
        onRender: (item: IRequestRow) => (
          <IconButton
            iconProps={{ iconName: "Edit" }}
            title="Edit"
            ariaLabel="Edit"
            disabled={
              item.Status === "Submitted" ||
              item.Status === "Approved" ||
              item.Status === "InProgress"
            }
            onClick={() => handleEdit(item)}
          />
        ),
      },
      {
        key: "column8",
        name: "Submit request",
        minWidth: 110,
        maxWidth: 130,
        isResizable: false,
        onRender: (item: IRequestRow) => (
          <IconButton
            iconProps={{ iconName: "Send" }}
            title="Send request"
            ariaLabel="Send request"
            disabled={
              item.Status === "Submitted" ||
              item.Status === "Approved" ||
              item.Status === "InProgress"
            }
            onClick={() => handleSendRequest(item)}
          />
        ),
      },
    ],
    [],
  );
  const fetchDepartments = async () => {
    await props.props.provider
      .fetchDepartments()
      .then((response: any) => {
        setDepartmentOptions(
          response.map((dept: string) => ({
            key: dept,
            text: dept,
            value: dept,
          })),
        );
      })
      .catch((error: any) => {
        console.error("Error fetching departments:", error);
      });
  };
  React.useEffect(() => {
    fetchDepartments().catch((error) => {
      setMessageType(MessageBarType.error);
      setMessage("Error loading departments.");
      console.error(error);
    });
  }, [department]);

  React.useEffect(() => {
    getReviewData().catch((error) => {
      setMessageType(MessageBarType.error);
      setMessage("Error loading review data.");
      console.error(error);
    });
  }, []);

  const pageCount = Math.max(1, Math.ceil(requests.length / pageSize));
  const pagedRequests = React.useMemo(() => {
    const startIndex = (currentPage - 1) * pageSize;
    return requests.slice(startIndex, startIndex + pageSize);
  }, [requests, currentPage]);

  React.useEffect(() => {
    if (currentPage > pageCount) {
      setCurrentPage(pageCount);
    }
  }, [currentPage, pageCount]);
  console.log(secondaryEmails);
  console.log(secondaryOwners);

  // console.log(editFormData);

  const handleSaveEdit = async () => {
    if (!editFormData) return;
    return await props.props.provider
      .EditRequest(props.props.RequestFormId, {
        siteName: editFormData.Title,
        SiteType: editFormData.SiteType,
        siteDescription: editFormData.Description,
        primaryOwners: Array.isArray(primaryOwners)
          ? primaryOwners
          : primaryEmails, // ...(primaryOwners.length > 0 && {
        //   primaryOwners: primaryOwners,
        // }),
        // primaryOwners: primaryOwners,
        secondaryOwners: Array.isArray(secondaryOwners)
          ? secondaryOwners
          : secondaryEmails,
        department: department || editFormData.Department,
        approverOptions: approverOptions,
        reason: editFormData.SiteJustification,
        id: editingItem.Id,
      })
      .then((response: any) => {
        // console.log("Draft saved:", response);
        // setRequests((current) =>
        //   current.map((row) =>
        //     row.key === editFormData.key ? editFormData : row,
        //   ),
        // );
        setMessageType(MessageBarType.success);
        getReviewData().catch((error) => {
          setMessageType(MessageBarType.error);
          setMessage("Error refreshing data after save.");
          console.error(error);
        });
        setMessage(`Request "${editFormData.Title}" was updated successfully.`);
        // keep edit dialog open here; caller will close if desired
      })
      .catch((error: any) => {
        console.error("Error saving draft:", error);
      });
  };
  React.useEffect(() => {
    setApproverOptions([]);
    if (!department) {
      return;
    }

    const loadApprovers = async () => {
      const filter = `Title eq '${department.replace(/'/g, "''")}'`;
      const requestUrl = `${props.props.currentWebUrl}/_api/web/lists/GetByTitle('Departments')/items?$filter=${encodeURIComponent(filter)}&$select=*,Approvers/Title,Approvers/EMail,Approvers/Id&$expand=Approvers&$top=1`;
      try {
        const response = await props.props.spHttpClient.get(
          requestUrl,
          SPHttpClient.configurations.v1,
          {
            headers: { Accept: "application/json;odata=nometadata" },
          },
        );
        // console.log(response);

        if (!response.ok) {
          throw new Error("Failed to load approvers.");
        }

        const json = await response.json();
        const items = json.value || [];
        // console.log(items);
        const approversarr = items.flatMap(
          (item: any) =>
            item.Approvers?.map((approver: any) => approver.EMail) || [],
        );
        console.log(approversarr, "approversarr");

        setApproverOptions(approversarr);

        if (approversarr.length === 0) {
          setMessageType(MessageBarType.error);
          setMessage(`No approvers were found for department ${department}.`);
        } else {
          setMessage(undefined);
        }
      } catch (error) {
        setMessageType(MessageBarType.error);
        setMessage("Unable to load approvers for the selected department.");
      }
    };

    void loadApprovers();
  }, [props.props.currentWebUrl, department, props.props.spHttpClient]);

  const handleEditFormChange = (field: keyof IRequestRow, value: any) => {
    if (editFormData) {
      setEditFormData({ ...editFormData, [field]: value });
    }
  };

  return (
    <div className={styles.reviewRequest}>
      <Stack tokens={{ childrenGap: 16 }}>
        {/* <div className={styles.reviewHeader}>
          <Text variant="xLarge">Review Requests</Text>
          <Text variant="medium">
            View all draft and submitted requests in a table.
          </Text>
        </div> */}

        {message && (
          <MessageBar
            messageBarType={messageType}
            onDismiss={() => setMessage(undefined)}
          >
            {message}
          </MessageBar>
        )}

        <div className={styles.tableWrapper}>
          <DetailsList
            items={pagedRequests}
            selectionMode={SelectionMode.none}
            columns={columns}
            setKey="none"
          />
        </div>
      </Stack>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <DefaultButton
          text="Previous"
          disabled={currentPage <= 1}
          onClick={() => setCurrentPage((page) => Math.max(page - 1, 1))}
        />
        <Text>{`Page ${currentPage} of ${pageCount}`}</Text>
        <DefaultButton
          text="Next"
          disabled={currentPage >= pageCount}
          onClick={() =>
            setCurrentPage((page) => Math.min(page + 1, pageCount))
          }
        />
      </Stack>
      <Dialog
        hidden={!isEditDialogOpen}
        onDismiss={() => {
          setIsEditDialogOpen(false);
          setEditingItem(null);
          setEditFormData(null);
        }}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: "Edit Request",
          subText: `Edit the details for "${editFormData?.Title || ""}"`,
        }}
        modalProps={{ isBlocking: false }}
        className="editRequestDialog"
      >
        {editFormData && (
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
              <Stack.Item grow={1}>
                <TextField
                  label="Site name"
                  value={editFormData.Title}
                  onChange={(_, value) =>
                    handleEditFormChange("Title", value || "")
                  }
                  placeholder="Enter the requested site name"
                  required
                />
              </Stack.Item>
              <Stack.Item grow={1}>
                {" "}
                <Dropdown
                  label="Site Type"
                  selectedKey={editFormData.SiteType}
                  onChange={(_, option) =>
                    handleEditFormChange("SiteType", option?.key as string)
                  }
                  placeholder="Select Site Type"
                  options={[
                    { key: "Communication Site", text: "Communication Site" },
                    { key: "Team Site", text: "Team Site" },
                  ]}
                  required
                  defaultValue={editFormData.SiteType}
                />
              </Stack.Item>
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
              <Stack.Item grow={1}>
                <PeoplePicker
                  context={peoplePickerContext}
                  titleText="Select Primary owners"
                  personSelectionLimit={10}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required={true}
                  //  disabled={isSubmitting}
                  searchTextLimit={2}
                  onChange={(items) => setPrimaryOwners(items)}
                  showHiddenInUI={false}
                  defaultSelectedUsers={editFormData.PrimaryOwnerEmail}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </Stack.Item>
              <Stack.Item grow={1}>
                <PeoplePicker
                  context={peoplePickerContext}
                  titleText="Select Secondary owners"
                  personSelectionLimit={10}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required={false}
                  //  disabled={isSubmitting}
                  searchTextLimit={2}
                  onChange={(items) => setSecondaryOwners(items)}
                  defaultSelectedUsers={editFormData.SecondaryOwnerEmail}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </Stack.Item>
            </Stack>

            <TextField
              label="Site description"
              value={editFormData.Description}
              onChange={(_, value) =>
                handleEditFormChange("Description", value || "")
              }
              placeholder="Add a short description of the site"
              multiline
              rows={3}
            />

            <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
              <Stack.Item grow={1}>
                <Dropdown
                  label="Department"
                  selectedKey={department}
                  onChange={(_, option) => setDepartment(option?.key as string)}
                  placeholder="Select department"
                  options={departmentOptions}
                  required
                  defaultValue={editFormData.Department}
                />
              </Stack.Item>
              <Stack.Item grow={1}>
                <PeoplePicker
                  context={peoplePickerContext}
                  titleText="Approvers"
                  personSelectionLimit={10}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required={true}
                  disabled={true}
                  searchTextLimit={2}
                  onChange={(items) => setApproverOptions(items)}
                  showHiddenInUI={false}
                  defaultSelectedUsers={approverOptions}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </Stack.Item>
            </Stack>

            <TextField
              label="Reason for site creation"
              value={editFormData.SiteJustification}
              onChange={(_, value) =>
                handleEditFormChange("SiteJustification", value || "")
              }
              placeholder="Describe why the site is needed"
              multiline
              rows={4}
              required
            />

            {/* {message && (
              <MessageBar messageBarType={messageType}>{message}</MessageBar>
            )} */}

            <Stack
              horizontal
              horizontalAlign="space-between"
              verticalAlign="end"
              tokens={{ childrenGap: 12 }}
              wrap
            >
              <div style={{ display: "flex", gap: "15px" }}>
                <DefaultButton
                  text="Cancel"
                  onClick={() => {
                    setIsEditDialogOpen(false);
                    setEditingItem(null);
                    setEditFormData(null);
                  }}
                />
                <PrimaryButton
                  text="Save Changes"
                  onClick={async () => {
                    try {
                      await handleSaveEdit();
                      setIsEditDialogOpen(false);
                      setEditingItem(null);
                      setEditFormData(null);
                    } catch (err) {
                      console.error("Error saving changes:", err);
                    }
                  }}
                />
                <PrimaryButton
                  text="Submit Request"
                  onClick={async () => {
                    const toSend = editFormData;
                    if (!toSend) return;
                    const idToFetch = editingItem?.Id || toSend.Id;
                    await handleSaveEdit();

                    let sendItem = toSend;
                    try {
                      const updatedArr: any =
                        await props.props.provider.getReviewData(
                          props.props.RequestFormId,
                          Number(idToFetch),
                        );
                      if (Array.isArray(updatedArr) && updatedArr.length > 0) {
                        sendItem = updatedArr[0];
                      }
                    } catch (err) {
                      console.log(
                        "Could not reload saved item, using local item",
                        err,
                      );
                    }

                    await handleSendRequest(sendItem);
                  }}
                />
              </div>
            </Stack>
          </Stack>
        )}
      </Dialog>

      <Dialog
        hidden={!isSuccessDialogOpen}
        onDismiss={() => setIsSuccessDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Request Submitted",
          subText: submittedTitle
            ? `Request "${submittedTitle}" has been submitted successfully.`
            : "The request has been submitted.",
        }}
        modalProps={{ isBlocking: false }}
      >
        <DialogFooter>
          <PrimaryButton
            text="OK"
            onClick={() => {
              setIsSuccessDialogOpen(false);
              setSubmittedTitle(null);
            }}
          />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!isValidationDialogOpen}
        onDismiss={() => setIsValidationDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Missing required fields",
          subText: validationMessage || "Please complete the required fields.",
        }}
        modalProps={{ isBlocking: false }}
      >
        <DialogFooter>
          <PrimaryButton
            text="OK"
            onClick={() => {
              setIsValidationDialogOpen(false);
              setValidationMessage(null);
            }}
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default ReviewRequest;
