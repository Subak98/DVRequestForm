import * as React from "react";
import { useEffect, useState, useMemo } from "react";
import {
  DefaultButton,
  PrimaryButton,
  Stack,
  Text,
  Dialog,
  DialogType,
} from "@fluentui/react";
import styles from "./DvRequestForm.module.scss";
import type { IDvRequestFormProps } from "./IDvRequestFormProps";
import RequestForm from "./RequestForm/RequestForm";
import ReviewRequest from "./ReviewForm/ReviewRequest";
import MyApproval from "./MyApproval/MyApproval";
import ApproveForm from "./ApproveForm/ApproveForm";

const DvRequestForm: React.FC<IDvRequestFormProps> = (props) => {
  const [selectedTab, setSelectedTab] = useState<
    "raise" | "review" | "myapproval"
  >("raise");
  const [showRequestForm, setShowRequestForm] = useState(false);

  // Check for URL id param to determine if showing approval form
  const requestId = useMemo(() => {
    if (typeof window === "undefined") return null;
    const params = new URLSearchParams(window.location.search);

    return params.get("itemid");
  }, []);

  React.useEffect(() => {
    if (typeof window === "undefined" || requestId) return;

    const params = new URLSearchParams(window.location.search);
    const tab = params.get("tab");

    if (tab === "raise" || tab === "review" || tab === "myapproval") {
      setSelectedTab(tab);
    }
  }, [requestId]);

  const navigateToTab = (tab: "raise" | "review" | "myapproval") => {
    setSelectedTab(tab);

    if (typeof window !== "undefined") {
      const url = new URL(window.location.href);
      url.searchParams.delete("itemid");
      url.searchParams.set("tab", tab);
      window.history.pushState({}, "", url.toString());
    }
  };

  // If there's an id in URL, show ApproveForm
  if (requestId) {
    // console.log(requestId);

    return (
      <ApproveForm
        currentWebUrl={props.currentWebUrl}
        spHttpClient={props.spHttpClient}
        context={props.context}
        provider={props.provider}
        requestId={requestId}
        listid={props.RequestFormId}
        useremail={props.userEmail}
      />
    );
  }

  return (
    <div className={styles.dvRequestForm}>
      <Stack tokens={{ childrenGap: 12 }}>
        <div className={styles.tabBar}>
          <DefaultButton
            className={
              selectedTab === "raise" ? styles.activeTab : styles.tabButton
            }
            text="Raise Request"
            onClick={() => {
              navigateToTab("raise");
              setShowRequestForm(false);
            }}
          />
          <DefaultButton
            className={
              selectedTab === "review" ? styles.activeTab : styles.tabButton
            }
            text="My Request"
            onClick={() => {
              navigateToTab("review");
              setShowRequestForm(false);
            }}
          />
          <DefaultButton
            className={
              selectedTab === "myapproval" ? styles.activeTab : styles.tabButton
            }
            text="My Approvals"
            onClick={() => {
              navigateToTab("myapproval");
              setShowRequestForm(false);
            }}
          />
        </div>

        {selectedTab === "raise" ? (
          // <div className={styles.sectionCard}>
          //   <div className={styles.sectionHeader}>
          //     <Text variant="xLarge">Raise Request</Text>
          //     <Text variant="medium">
          //       Start a new site request when you are ready.
          //     </Text>
          //   </div>
          //   <PrimaryButton
          //     text={
          //       showRequestForm ? "Close Request Form" : "Open Request Form"
          //     }
          //     onClick={() => setShowRequestForm((value) => !value)}
          //     style={{ margin: "auto", width: "20%" }}
          //   />
          // </div>
          <RequestForm
            description="Fill in the details below to request a new site."
            currentWebUrl={props.currentWebUrl}
            spHttpClient={props.spHttpClient}
            userDisplayName={props.userDisplayName}
            onCancel={() => setShowRequestForm(false)}
            onNavigateToTab={navigateToTab}
            provider={props.provider}
            props={props}
          />
        ) : selectedTab === "review" ? (
          <ReviewRequest props={props} />
        ) : (
          <MyApproval props={props} />
        )}
      </Stack>

      {/* <Dialog
        hidden={!showRequestForm}
        onDismiss={() => setShowRequestForm(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: "Request a new site",
          // subText: "Complete the details below to raise a new request.",
        }}
        modalProps={{ isBlocking: false }}
        className="requestFormDialog"
      > */}

      {/* </Dialog> */}
    </div>
  );
};

export default DvRequestForm;
