import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  DetailsList,
  IColumn,
  IconButton,
  MessageBar,
  MessageBarType,
  SelectionMode,
  Stack,
  Text,
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
  ApproverTitle: string;
  ApproverEmail: string[];
  Status: string;
  Id: number;
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
            <div className={styles.tableWrapper}>
              <DetailsList
                items={pagedApprovals}
                columns={columns}
                selectionMode={SelectionMode.none}
                setKey="set"
              />
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
