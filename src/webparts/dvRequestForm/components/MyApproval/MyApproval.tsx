import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  DetailsList,
  IColumn,
  IDetailsRowProps,
  IconButton,
  MessageBar,
  MessageBarType,
  SelectionMode,
  Stack,
  Text,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton,
} from "@fluentui/react";
import styles from "../ReviewForm/ReviewRequest.module.scss";

interface IRequestRow {
  key: string;
  Title: string;
  Date: string;
  Department: string;
  Description: string;
  SiteJustification: string;
  PrimaryOwnerTitle: string;
  SecondaryOwnerTitle?: string;
  ApproverTitle: string;
  ApproverEmail: string[];
  Status: string;
  SiteType: string;
  Id: number;
  SiteSlugUrl: string;
  ApprovedBy?: { Title: string };
}

interface IMyApprovalProps {
  props: any;
}

const MyApproval: React.FC<IMyApprovalProps> = ({ props }) => {
  const [approvals, setApprovals] = useState<IRequestRow[]>([]);
  const [message, setMessage] = useState<string | undefined>();
  const [messageType, setMessageType] = useState<MessageBarType>(
    MessageBarType.success,
  );
  const [currentPage, setCurrentPage] = useState(1);
  const [selectedApproval, setSelectedApproval] = useState<IRequestRow | null>(
    null,
  );
  const [isDetailsDialogOpen, setIsDetailsDialogOpen] = useState(false);
  const pageSize = 10;
  console.log(props);

  const currentUserName = props.userDisplayName || "";
  const currentUserEmail = props.userEmail || "";
  const fetchApprovals = async () => {
    try {
      const allRequests: any[] = await props.provider.getReviewData(
        props.RequestFormId,
        "",
        `Approvers/EMail eq '${currentUserEmail}' and Status eq 'InProgress'`,
      );

      const filtered = allRequests.filter((item) => {
        const approverTitle = (item.ApproverTitle || "")
          .toString()
          .toLowerCase();
        const approverEmails = Array.isArray(item.ApproverEmail)
          ? item.ApproverEmail.map((email: string) => email.toLowerCase())
          : [];

        return (
          approverTitle.includes(currentUserName.toLowerCase()) ||
          approverEmails.includes(currentUserEmail.toLowerCase())
        );
      });

      setApprovals(filtered);
      if (filtered.length === 0) {
        setMessageType(MessageBarType.success);
        setMessage("No approval items were found for you.");
      } else {
        setMessage(undefined);
      }
    } catch (error) {
      console.error("Error loading approvals:", error);
      setMessageType(MessageBarType.error);
      setMessage("Unable to load approvals at this time.");
    }
  };
  useEffect(() => {
    fetchApprovals().catch((error) => {});
  }, []);

  const pageCount = Math.max(1, Math.ceil(approvals.length / pageSize));
  const pagedApprovals = React.useMemo(() => {
    const startIndex = (currentPage - 1) * pageSize;
    return approvals.slice(startIndex, startIndex + pageSize);
  }, [approvals, currentPage]);

  React.useEffect(() => {
    if (currentPage > pageCount) {
      setCurrentPage(pageCount);
    }
  }, [currentPage, pageCount]);

  const handleSendRequest = (item: any): void => {
    const currentUrl = window.location.pathname;
    window.history.pushState({}, "", `${currentUrl}?itemid=${item.Id}`);
    window.location.reload();
    // Your existing logic
  };
  const renderTooltipCell =
    (fieldName: keyof IRequestRow) => (item: IRequestRow) => (
      <span title={item[fieldName] != null ? String(item[fieldName]) : ""}>
        {item[fieldName]}
      </span>
    );

  const getDetailText = (value: any) => {
    if (value == null || value === "") {
      return "-";
    }
    if (Array.isArray(value)) {
      return value
        .map((item) =>
          item == null
            ? ""
            : item.Title || item.text || item.name || String(item),
        )
        .filter(Boolean)
        .join(", ");
    }
    return String(value);
  };

  const handleViewDetails = (item: IRequestRow) => {
    setSelectedApproval(item);
    setIsDetailsDialogOpen(true);
  };

  const onRenderRow = (
    props?: IDetailsRowProps,
    defaultRender?: (props?: IDetailsRowProps) => React.ReactElement | null,
  ): React.ReactElement | null => {
    if (!props || !defaultRender) {
      return defaultRender?.(props) ?? null;
    }

    return (
      <div
        onClick={(event) => {
          const target = event.target as HTMLElement;
          if (target.closest("button, a, input, textarea, svg, path")) {
            return;
          }
          handleViewDetails(props.item as IRequestRow);
        }}
        style={{ cursor: "pointer" }}
      >
        {defaultRender(props)}
      </div>
    );
  };
  const columns: IColumn[] = useMemo(
    () => [
      {
        key: "column1",
        name: "Site Title",
        fieldName: "Title",
        minWidth: 150,
        maxWidth: 220,
        isResizable: true,
        onRender: (item: IRequestRow) => {
          const isApprovedWithUrl =
            item.Status === "Approved" && item.SiteSlugUrl;
          return isApprovedWithUrl ? (
            <a
              href={item.SiteSlugUrl}
              target="_blank"
              rel="noopener noreferrer"
              data-interception="off"
              title={item.Title}
              style={{ color: "#0078d4", textDecoration: "none" }}
              onMouseEnter={(e) =>
                (e.currentTarget.style.textDecoration = "underline")
              }
              onMouseLeave={(e) =>
                (e.currentTarget.style.textDecoration = "none")
              }
            >
              {item.Title}
            </a>
          ) : (
            <span title={item.Title}>{item.Title}</span>
          );
        },
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
      },
      {
        key: "column4",
        name: "Approvers",
        fieldName: "ApproverTitle",
        minWidth: 120,
        maxWidth: 160,
        onRender: renderTooltipCell("ApproverTitle"),
      },
      {
        key: "column5",
        name: "Justification",
        fieldName: "SiteJustification",
        minWidth: 160,
        maxWidth: 240,
        onRender: renderTooltipCell("SiteJustification"),
      },
      {
        key: "column6",
        name: "Status",
        fieldName: "Status",
        minWidth: 120,
        maxWidth: 140,
      },
      {
        key: "column7",
        name: "Action",
        minWidth: 110,
        maxWidth: 130,
        isResizable: false,
        onRender: (item: IRequestRow) => (
          <IconButton
            iconProps={{ iconName: "AcceptMedium" }}
            title="Approve/Reject/Hold"
            ariaLabel="Send request"
            disabled={item.Status === "Submitted" || item.Status === "Approved"}
            onClick={() => handleSendRequest(item)}
          />
        ),
      },
    ],
    [],
  );

  return (
    <div className={styles.reviewRequest}>
      <Stack tokens={{ childrenGap: 16 }}>
        {message ? (
          <MessageBar
            messageBarType={messageType}
            // onDismiss={() => setMessage(undefined)}
          >
            {message}
          </MessageBar>
        ) : (
          <>
            <div>
              <div className={styles.tableWrapper}>
                {" "}
                <DetailsList
                  items={pagedApprovals}
                  columns={columns}
                  selectionMode={SelectionMode.none}
                  setKey="set"
                  onRenderRow={onRenderRow}
                />
              </div>

              <Dialog
                hidden={!isDetailsDialogOpen}
                onDismiss={() => setIsDetailsDialogOpen(false)}
                dialogContentProps={{
                  type: DialogType.largeHeader,
                  title: "Approval Details",
                }}
                className="ReviewRequest"
                modalProps={{ isBlocking: false }}
              >
                <Stack tokens={{ childrenGap: 12 }}>
                  <Text>
                    <strong>Site Title:</strong>{" "}
                    {getDetailText(selectedApproval?.Title)}
                  </Text>
                  <Text>
                    <strong>Primary Owner:</strong>{" "}
                    {getDetailText(selectedApproval?.PrimaryOwnerTitle)}
                  </Text>
                  <Text>
                    <strong>Secondary Owner:</strong>{" "}
                    {getDetailText(
                      selectedApproval?.SecondaryOwnerTitle
                        ? selectedApproval?.SecondaryOwnerTitle
                        : "-",
                    )}{" "}
                  </Text>
                  <Text>
                    <strong>Date:</strong>{" "}
                    {getDetailText(selectedApproval?.Date)}
                  </Text>
                  <Text>
                    <strong>Approvers:</strong>{" "}
                    {getDetailText(selectedApproval?.ApproverTitle)}
                  </Text>
                  <Text>
                    <strong>Site Justification:</strong>{" "}
                    {getDetailText(selectedApproval?.SiteJustification)}
                  </Text>
                  <Text>
                    <strong>Status:</strong>{" "}
                    {getDetailText(selectedApproval?.Status)}
                  </Text>
                  <Text>
                    <strong>Department:</strong>{" "}
                    {getDetailText(selectedApproval?.Department)}
                  </Text>{" "}
                  <Text>
                    <strong>Site Type:</strong>{" "}
                    {getDetailText(selectedApproval?.SiteType)}
                  </Text>
                  <Text>
                    <strong>Site Description:</strong>{" "}
                    {getDetailText(selectedApproval?.Description)}
                  </Text>
                </Stack>
                <DialogFooter>
                  <DefaultButton
                    text="Close"
                    onClick={() => setIsDetailsDialogOpen(false)}
                  />
                </DialogFooter>
              </Dialog>
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
              >
                <DefaultButton
                  text="Previous"
                  disabled={currentPage <= 1}
                  onClick={() =>
                    setCurrentPage((page) => Math.max(page - 1, 1))
                  }
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
            </div>
          </>
        )}
      </Stack>
    </div>
  );
};

export default MyApproval;
