import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  Dialog,
  DialogType,
  DialogFooter,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import type { IApproveFormProps } from "./IApproveFormProps";
import styles from "./ApproveForm.module.scss";

interface IRequestData {
  Id: string;
  Title: string;
  Description: string;
  SiteType: string;
  SiteJustification: string;
  Department: string;
  PrimaryOwnerTitle: string;
  SecondaryOwnerTitle?: string;
  PrimaryOwner?: any;
  SecondaryOwner?: any;
  ApproverTitle: string;
  Status: string;
  ApprovedBy: any;
  Requestor: string;
  RequestApprover: any;
  RequestorEmail: string;
}
interface IDeptData {
  Title: string;
  TemplateId: string;
}
const ApproveForm: React.FC<IApproveFormProps> = ({
  currentWebUrl,
  spHttpClient,
  context,
  provider,
  requestId,
  listid,
  useremail,
  username,
  ApprovalResponseURL,
}) => {
  const [requestData, setRequestData] = useState<IRequestData | null>(null);
  const [comments, setComments] = useState("");
  const [isLoading, setIsLoading] = useState(true);
  const [message, setMessage] = useState<string | undefined>();
  const [messageType, setMessageType] = useState<MessageBarType>(
    MessageBarType.success,
  );
  const [approvalStatus, setApprovalStatus] = useState<
    "Hold" | "Approved" | "Rejected" | null
  >(null);
  const [departmentdata, SetDepartmentData] = useState<IDeptData | null>(null);
  const fetchRequestData = async () => {
    if (!requestId) {
      setIsLoading(false);
      return;
    }

    try {
      await provider
        .getReviewData(listid, Number(requestId))
        .then(async (data: any) => {
          console.log(data);

          setRequestData(data[0]);
          await provider
            .fetchUniqueDepartment(
              `Title eq '${data[0].Department}' and SiteType eq '${data[0].SiteType}'`,
            )
            .then((dep: any) => {
              SetDepartmentData(dep[0]);
            });
          setMessageType(MessageBarType.success);
        });
    } catch (error) {
      console.error("Error fetching request:", error);
      setMessageType(MessageBarType.error);
      setMessage("Unable to load request details.");
    } finally {
      setIsLoading(false);
    }
  };
  useEffect(() => {
    fetchRequestData().catch((error) => {
      setMessageType(MessageBarType.error);
      setMessage("Error loading request details.");
      console.error(error);
    });
  }, [requestId]);
  console.log(requestData);
  // console.log(context);
  // console.log(departmentdata);

  const normalizePeople = (
    people?: {
      Title?: string;
      Email?: string;
      title?: string;
      email?: string;
    }[],
    fallbackText?: string,
  ) => {
    if (Array.isArray(people) && people.length > 0) {
      return people.map((person) => ({
        title: person.Title || person.title || "",
        email: person.Email || person.email || "",
      }));
    }

    if (fallbackText) {
      return fallbackText
        .split(",")
        .map((title) => ({ title: title.trim(), email: "" }))
        .filter((person) => person.title);
    }

    return [];
  };

  const handleApprovalSubmit = async (
    status: "Hold" | "Approved" | "Rejected",
  ) => {
    if (status === "Hold" && !comments) {
      setMessageType(MessageBarType.error);
      setMessage("Please provide comments when putting a request on hold.");
      return;
    }

    if (!requestData) return;

    try {
      await provider
        .ChangeStatus(listid, requestData.Id, status, comments, useremail)
        .then(async (response: any) => {
          setApprovalStatus(status);
          setMessageType(MessageBarType.success);
          setMessage(
            `Request has been ${status.toLowerCase()}ed successfully.`,
          );
          console.log(requestData);

          // console.log(tenantName);
          // Close dialog and navigate after successful submission
        });
    } catch (error) {
      setMessageType(MessageBarType.error);
      setMessage("Error processing approval.");
      console.error(error);
    }
    if (status === "Hold" || status === "Rejected") {
      await provider.ApprovalResponse(
        context,
        requestData.Title,
        requestData.Description,
        requestData.PrimaryOwner,
        requestData.SecondaryOwner,
        requestData.Department,
        requestData.ApproverTitle,
        comments,
        requestData.RequestApprover,
        requestData.Id,
        requestData.Requestor,
        requestData.RequestorEmail,
        status,
        ApprovalResponseURL,
        username,
      );
    }
    const tenantName = new URL(
      context.pageContext.web.absoluteUrl,
    ).hostname.split(".")[0];
    const urlJson = {
      sitetitle: requestData.Title,
      tenantname: tenantName,
      sitetype: requestData.SiteType,
      templateid: departmentdata?.TemplateId,
      primaryowner: normalizePeople(
        requestData.PrimaryOwner,
        requestData.PrimaryOwnerTitle,
      ),
      secondaryowner: normalizePeople(
        requestData.SecondaryOwner,
        requestData.SecondaryOwnerTitle,
      ),
      RequesterEmail: [
        {
          title: requestData.Requestor,
          email: requestData.RequestorEmail,
        },
      ],
    };
    const encoded = encodeURIComponent(JSON.stringify(urlJson));
    console.log(encoded);
    if (status === "Hold" || status === "Rejected") {
      window.location.href = `${currentWebUrl}/SitePages/DV-Request-Form.aspx?tab=review`;
    } else {
      window.open(
        `${currentWebUrl}/SitePages/SPOProvisioning.aspx?config=${encoded}`,
        "_blank",
        "noopener,noreferrer",
      );
      window.setTimeout(() => {
        window.location.href = `${currentWebUrl}/SitePages/DV-Request-Form.aspx?tab=review`;
      }, 300);
    }
  };

  if (isLoading) {
    return (
      <Dialog
        hidden={false}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: "Loading...",
        }}
        modalProps={{ isBlocking: true }}
      >
        <Text>Loading request details...</Text>
      </Dialog>
    );
  }
  console.log(requestData);

  if (!requestData) {
    return null;
  }
  const baseSiteUrl = currentWebUrl.split("/").slice(0, -1).join("/");
  return (
    <div className={styles.approveForm}>
      {requestData.Status == "Approved" && requestData.ApprovedBy ? (
        <MessageBar messageBarType={MessageBarType.warning}>
          This request has already been approved by{" "}
          <Text styles={{ root: { fontWeight: 600 } }}>
            {requestData.ApprovedBy.Title}
          </Text>
          .
        </MessageBar>
      ) : (
        <Dialog
          hidden={false}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: "Approve Request",
            subText: `Review and approve/reject the request: "${requestData.Title}"`,
          }}
          className="approveDialog"
          modalProps={{ isBlocking: false }}
        >
          <Stack
            tokens={{ childrenGap: 16 }}
            styles={{ root: { padding: "0 24px" } }}
          >
            <div>
              <Text className={styles.boldtext}>Site Title: </Text>
              <Text>{requestData.Title}</Text>
            </div>

            <div>
              <Text className={styles.boldtext}>Description: </Text>
              <Text>{requestData.Description || "-"}</Text>
            </div>

            <div>
              <Text className={styles.boldtext}>Department: </Text>
              <Text>{requestData.Department}</Text>
            </div>

            <div>
              <Text className={styles.boldtext}>Primary Owner: </Text>
              <Text>{requestData.PrimaryOwnerTitle}</Text>
            </div>
            <div>
              <Text className={styles.boldtext}>Secondary Owner: </Text>
              <Text>{requestData.SecondaryOwnerTitle}</Text>
            </div>
            <div>
              <Text className={styles.boldtext}>Approver: </Text>
              <Text>{requestData.ApproverTitle}</Text>
            </div>

            <div>
              <Text className={styles.boldtext}>Justification: </Text>
              <Text>{requestData.SiteJustification}</Text>
            </div>

            <div>
              <Text className={styles.boldtext}>Status: </Text>
              <Text>{requestData.Status}</Text>
            </div>
            <div>
              <Text className={styles.boldtext}>Requested URL: </Text>
              <Text>{`${baseSiteUrl}/${requestData.Title.replace(/\s+/g, "")}`}</Text>
            </div>
            <TextField
              label="Comments"
              placeholder="Add comments for this approval decision"
              multiline
              rows={4}
              value={comments}
              onChange={(_, value) => setComments(value || "")}
            />
            {message && (
              <MessageBar messageBarType={messageType}>{message}</MessageBar>
            )}
            <Stack
              horizontal
              tokens={{ childrenGap: 12 }}
              horizontalAlign="end"
            >
              <DefaultButton
                text="Hold"
                onClick={() => {
                  setApprovalStatus("Hold");
                  handleApprovalSubmit("Hold").catch((error) => {
                    setMessageType(MessageBarType.error);
                    setMessage("Error processing approval.");
                    console.error(error);
                  });
                }}
              />
              <DefaultButton
                text="Approve"
                onClick={() => {
                  setApprovalStatus("Approved");
                  handleApprovalSubmit("Approved").catch((error) => {
                    setMessageType(MessageBarType.error);
                    setMessage("Error processing approval.");
                    console.error(error);
                  });
                }}
              />
              <DefaultButton
                text="Reject"
                onClick={() => {
                  setApprovalStatus("Rejected");
                  handleApprovalSubmit("Rejected").catch((error) => {
                    setMessageType(MessageBarType.error);
                    setMessage("Error processing approval.");
                    console.error(error);
                  });
                }}
              />
            </Stack>
          </Stack>

          <DialogFooter></DialogFooter>
        </Dialog>
      )}
    </div>
  );
};

export default ApproveForm;
