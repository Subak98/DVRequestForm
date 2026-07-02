import * as React from "react";
import { useEffect, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  TextField,
  Dropdown,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Stack,
  Dialog,
  DialogType,
  DialogFooter,
  Text,
} from "@fluentui/react";
import type { IRequestFormProps } from "./IRequestFormProps";
import styles from "./RequestForm.module.scss";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
interface IApproverOption {
  key: string;
  text: string;
  value: string;
}

// const departmentOptions = [
//   { key: "HR", text: "HR", value: "HR" },
//   { key: "Finance", text: "Finance", value: "Finance" },
//   { key: "IT", text: "IT", value: "IT" },
//   { key: "Operations", text: "Operations", value: "Operations" },
// ];

const RequestForm: React.FC<IRequestFormProps> = ({
  description,
  currentWebUrl,
  spHttpClient,
  userDisplayName,
  onCancel,
  onNavigateToTab,
  provider,
  props,
}) => {
  const [siteName, setSiteName] = useState("");
  const [siteDescription, setSiteDescription] = useState("");
  const [sitetype, setSiteType] = useState("Team Site");

  const [department, setDepartment] = useState("");
  const [approver, setApprover] = useState("");
  const [reason, setReason] = useState("");
  const [approverOptions, setApproverOptions] = useState<any[]>([]);
  const [message, setMessage] = useState<string | undefined>();
  const [messageType, setMessageType] = useState<"error" | "success">(
    "success",
  );
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [dialogVisible, setDialogVisible] = useState(false);
  const [draftSavedDialogVisible, setDraftSavedDialogVisible] = useState(false);
  const [submitSuccessDialogVisible, setSubmitSuccessDialogVisible] =
    useState(false);
  const [secondaryOwners, setSecondaryOwners] = React.useState<any[]>([]);
  const [primaryOwners, setPrimaryOwners] = React.useState<any[]>([]);
  const [departmentOptions, setDepartmentOptions] = useState<IApproverOption[]>(
    [],
  );
  const [formDescription, setFormDescription] = useState<any>(null);

  const peoplePickerRef = React.useRef<any>(null);
  const primaryPeoplePickerRef = React.useRef<any>(null);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory,
    spHttpClient: props.context.spHttpClient,
  };
  console.log(props);

  const fetchDepartments = async () => {
    await provider
      .fetchFormDescription(props.formDescription)
      .then((fd: any) => {
        setFormDescription(fd[0]);
      });
    await provider
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
  // console.log(props);
  useEffect(() => {
    fetchDepartments().catch((error) => {
      console.error("Error fetching departments:", error);
    });
  }, [department]);

  useEffect(() => {
    setApprover("");
    setApproverOptions([]);
    if (!department) {
      return;
    }

    const loadApprovers = async () => {
      const filter = `Title eq '${department.replace(/'/g, "''")}'`;
      const requestUrl = `${currentWebUrl}/_api/web/lists/GetByTitle('Departments')/items?$filter=${encodeURIComponent(filter)}&$select=*,Approvers/Title,Approvers/EMail,Approvers/Id&$expand=Approvers&$top=1`;
      try {
        const response = await spHttpClient.get(
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
        console.log(json);

        const items = json.value || [];
        console.log(items);
        const approversarr = items.flatMap(
          (item: any) =>
            item.Approvers?.map((approver: any) => approver.EMail) || [],
        );
        console.log(approversarr);

        setApproverOptions(approversarr);

        if (approversarr.length === 0) {
          setMessageType("error");
          setMessage(`No approvers were found for department ${department}.`);
        } else {
          setMessage(undefined);
        }
      } catch (error) {
        setMessageType("error");
        setMessage("Unable to load approvers for the selected department.");
      }
    };

    void loadApprovers();
  }, [currentWebUrl, department, spHttpClient]);

  const checkSiteExists = async (name: string): Promise<boolean> => {
    const escaped = name.trim().replace(/\s+/g, "");
    // const queryText = `Title:${escaped} AND contentclass:STS_Web`;
    const queryText = `Title:${escaped}`;
    const requestUrl = `${currentWebUrl}/_api/search/query?querytext='${encodeURIComponent(queryText)}'&trimduplicates=false&rowlimit=1&selectproperties='Title,Path'`;

    const response = await spHttpClient.get(
      requestUrl,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: "application/json;odata=nometadata" },
      },
    );

    if (!response.ok) {
      throw new Error("Search request failed.");
    }

    const json = await response.json();
    const rows = json?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    return rows.length > 0;
  };
  const resetFields = async () => {
    setSiteName("");
    setPrimaryOwners([]);
    setSecondaryOwners([]);
    setSiteDescription("");
    setDepartment("");
    setApproverOptions([]);
    setReason("");
    setSiteType("Team Site");
    peoplePickerRef.current?.clearSelectedPersons();
    primaryPeoplePickerRef.current?.clearSelectedPersons();
  };
  // console.log(sitetype);

  const submitRequest = async () => {
    setMessage(undefined);

    if (
      !siteName ||
      primaryOwners.length === 0 ||
      !department ||
      approverOptions.length === 0 ||
      !reason
    ) {
      setMessageType("error");
      setMessage("Please complete all required fields before submitting.");
      return;
    }

    setIsSubmitting(true);
    try {
      const exists = await checkSiteExists(siteName.trim());
      if (exists) {
        setMessageType("error");
        setMessage(
          `The site name "${siteName}" already exists in the tenant. Please choose a different name.`,
        );
      } else {
        await provider
          .saveDraftRequest(props.RequestFormId, {
            siteName,
            sitetype,
            siteDescription,
            primaryOwners,
            secondaryOwners,
            department,
            approver,
            reason,
            approverOptions,
            status: "Submitted",
          })
          .then(async (response: any) => {
            // console.log("Draft saved:", response);
            setDialogVisible(false);
            resetFields().catch((err: any) => {});
            // show confirmation popup then redirect to My Request
            setSubmitSuccessDialogVisible(true);
            onNavigateToTab?.("review");
            // auto-close the popup and form after a short delay
            setTimeout(() => {
              setSubmitSuccessDialogVisible(false);
              onCancel?.();
            }, 3000);
            await props.provider.sendEmail(
              props.context,
              siteName,
              siteDescription,
              primaryOwners.map((user) => user.text).join(", "),
              secondaryOwners.map((user) => user.text).join(", "),
              department,
              approver,
              reason,
              approverOptions,
              response.Id,
              props.PAURL,
            );
          })
          .catch((error) => {
            console.error("Error saving draft:", error);
          });
      }
    } catch (error) {
      setMessageType("error");
      setMessage(
        "Unable to validate the site name at this time. Please try again later.",
      );
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setDialogVisible(true);
  };

  const handleSaveDraft = async () => {
    setMessage("");
    if (!siteName) {
      setMessageType("error");
      setMessage("A site name is required to save this request as a draft.");
      setDialogVisible(false);
      return;
    }
    try {
      const exists = await checkSiteExists(siteName.trim());
      if (exists) {
        setMessageType("error");
        setMessage(
          `The site name "${siteName}" already exists in the tenant. Please choose a different name.`,
        );
        setDialogVisible(false);
      } else {
        await provider
          .saveDraftRequest(props.RequestFormId, {
            siteName,
            sitetype,
            siteDescription,
            primaryOwners,
            secondaryOwners,
            department,
            approver,
            reason,
            approverOptions,
            status: "Draft",
          })
          .then((response) => {
            // console.log("Draft saved:", response);
            setDialogVisible(false);
            resetFields().catch((err: any) => {});
            // onNavigateToTab?.("review");
            // show confirmation popup then close the form
            setDraftSavedDialogVisible(true);
            // auto-close the popup and form after a short delay
            setTimeout(() => {
              onNavigateToTab?.("review");
              setDraftSavedDialogVisible(false);

              onCancel?.();
            }, 2500);
          })
          .catch((error) => {
            console.error("Error saving draft:", error);
          });
      }
    } catch (error) {
      setMessageType("error");
      setMessage(
        "Unable to validate the site name at this time. Please try again later.",
      );
    } finally {
      setIsSubmitting(false);
    }
  };
  // const callFlow = async (): Promise<void> => {
  //   const flowUrl =
  //     "https://defaultb14a476cad7448f8986f8942642cad.2e.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/f645d31e5aa448999fc7a1d80294e8a1/triggers/manual/paths/invoke?api-version=1";

  //   const payload = {
  //     title: siteName,
  //     userEmail: props.context.pageContext.user.email,
  //   };
  //   try {
  //     const response = await fetch(flowUrl, {
  //       method: "POST",
  //       headers: {
  //         "Content-Type": "application/json",
  //       },
  //       body: JSON.stringify(payload),
  //     });
  //     if (response.ok) {
  //       console.log("Flow triggered successfully");
  //     } else {
  //       console.log("Error triggering flow");
  //     }
  //   } catch (error) {
  //     console.error(error);
  //   }
  // };

  const handleConfirmSubmit = async () => {
    // keep dialog open until validation/submission completes
    await submitRequest();
  };

  return (
    <div className={styles.requestForm}>
      <Stack
        tokens={{ childrenGap: 16 }}
        styles={{
          root: {
            background: "#ffffff",
            border: "1px solid #edebe9",
            borderRadius: 8,
            boxShadow: "0 1px 4px rgba(0, 0, 0, 0.08)",
            padding: 24,
            width: "85%",
          },
        }}
      >
        <form onSubmit={handleSubmit} className={styles.form}>
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal tokens={{ childrenGap: 13 }} wrap>
              <Stack.Item grow={1}>
                <TextField
                  label="Site Name"
                  value={siteName}
                  onChange={(_, value) => setSiteName(value || "")}
                  placeholder="Enter the requested site name"
                  required
                />
              </Stack.Item>
              <Stack.Item grow={1}>
                <Dropdown
                  label="Site Type"
                  selectedKey={sitetype}
                  onChange={(_, option) => setSiteType(option?.key as string)}
                  placeholder="Select department"
                  options={[
                    { key: "Communication Site", text: "Communication Site" },
                    { key: "Team Site", text: "Team Site" },
                  ]}
                  required
                />
              </Stack.Item>
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
              <Stack.Item grow={1}>
                <PeoplePicker
                  ref={primaryPeoplePickerRef}
                  context={peoplePickerContext}
                  titleText="Select Primary Owners"
                  personSelectionLimit={10}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required={true}
                  disabled={isSubmitting}
                  searchTextLimit={2}
                  onChange={(items) => setPrimaryOwners(items)}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </Stack.Item>
              <Stack.Item grow={1}>
                <PeoplePicker
                  ref={peoplePickerRef}
                  context={peoplePickerContext}
                  titleText="Select Secondary Owners"
                  personSelectionLimit={10}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  required={false}
                  disabled={isSubmitting}
                  searchTextLimit={2}
                  onChange={(items) => setSecondaryOwners(items)}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </Stack.Item>
            </Stack>

            <TextField
              label="Site Description"
              value={siteDescription}
              onChange={(_, value) => setSiteDescription(value || "")}
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
              label="Reason For Site Creation"
              value={reason}
              onChange={(_, value) => setReason(value || "")}
              placeholder="Describe why the site is needed"
              multiline
              rows={4}
              required
            />

            {message && (
              <MessageBar
                messageBarType={
                  messageType === "error"
                    ? MessageBarType.error
                    : MessageBarType.success
                }
              >
                {message}
              </MessageBar>
            )}

            <Stack
              horizontal
              horizontalAlign="space-between"
              verticalAlign="end"
              tokens={{ childrenGap: 12 }}
              wrap
            >
              <Text variant="smallPlus" styles={{ root: { color: "#605e5c" } }}>
                Requested by {userDisplayName}
              </Text>
              <div style={{ display: "flex", gap: "15px" }}>
                <PrimaryButton
                  type="button"
                  text={isSubmitting ? "Checking..." : "Save"}
                  disabled={isSubmitting}
                  onClick={() => setDialogVisible(true)}
                />
                <DefaultButton
                  type="button"
                  text="Reset"
                  disabled={isSubmitting}
                  onClick={() => {
                    void resetFields();
                  }}
                />
              </div>
            </Stack>
          </Stack>
        </form>
      </Stack>
      <div
        className={styles.formDesc}
        dangerouslySetInnerHTML={{
          __html: formDescription?.Description || "",
        }}
      />
      <Dialog
        hidden={!dialogVisible}
        onDismiss={() => setDialogVisible(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Save request",
          subText:
            "Choose whether to save this form as a draft or submit the request now.",
        }}
      >
        <DialogFooter>
          <DefaultButton text="Save Draft" onClick={handleSaveDraft} />
          <PrimaryButton
            text="Submit Request"
            onClick={handleConfirmSubmit}
            disabled={isSubmitting}
          />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!draftSavedDialogVisible}
        onDismiss={() => setDraftSavedDialogVisible(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Draft saved",
          subText:
            "Draft saved. You can continue editing or submit when ready.",
        }}
        modalProps={{ isBlocking: false }}
      >
        <DialogFooter>
          {/* <PrimaryButton
            text="OK"
            onClick={() => {
              setDraftSavedDialogVisible(false);
              onCancel?.();
            }}
          /> */}
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!submitSuccessDialogVisible}
        onDismiss={() => setSubmitSuccessDialogVisible(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Request submitted",
          subText:
            "Request submitted successfully. Redirecting to My Request...",
        }}
        modalProps={{ isBlocking: false }}
      >
        <DialogFooter></DialogFooter>
      </Dialog>
    </div>
  );
};

export default RequestForm;
