# DV Request Form - Technical Documentation

# Overview

DV Request Form is a SharePoint Framework (SPFx) client-side web part designed for an internal site-request workflow. It allows users to submit requests for new SharePoint sites, save requests as drafts, review their requests, and route approvals through a SharePoint-based approval experience.

The solution uses React, Fluent UI, PnPjs, and the SharePoint REST API to interact with SharePoint lists and user data.

## Solution Purpose

The web part supports the following business flow:

1. A user raises a request for a new site.
2. The request is stored in a SharePoint list with status values such as Draft or Submitted.
3. The request can be reviewed by the requestor and approver.
4. An approver can place the request on Hold, Approve, or Reject it.
5. If approved, the user is redirected to a provisioning flow page with encoded request metadata.

## Technology Stack

- SharePoint Framework v1.22.2
- React 17
- TypeScript 5.8
- Fluent UI React
- PnPjs v4
- Microsoft SPHttpClient
- Heft / RushStack build tooling
- Node.js 22.14.x (recommended)

## Architecture

### Entry Point

The web part entry point is [src/webparts/
/DvRequestFormWebPart.ts](src/webparts/dvRequestForm/DvRequestFormWebPart.ts). It initializes the web part, creates the service provider, and exposes property pane configuration.

### Main UI Components

- [src/webparts/dvRequestForm/components/DvRequestForm.tsx](src/webparts/dvRequestForm/components/DvRequestForm.tsx) – top-level container that switches between Raise Request, My Request, and My Approvals tabs.
- [src/webparts/dvRequestForm/components/RequestForm/RequestForm.tsx](src/webparts/dvRequestForm/components/RequestForm/RequestForm.tsx) – form that captures site name, owners, department, approvers, description, and justification.
- [src/webparts/dvRequestForm/components/ReviewForm/ReviewRequest.tsx](src/webparts/dvRequestForm/components/ReviewForm/ReviewRequest.tsx) – displays existing requests for review.
- [src/webparts/dvRequestForm/components/MyApproval/MyApproval.tsx](src/webparts/dvRequestForm/components/MyApproval/MyApproval.tsx) – lists approval items assigned to the current user.
- [src/webparts/dvRequestForm/components/ApproveForm/ApproveForm.tsx](src/webparts/dvRequestForm/components/ApproveForm/ApproveForm.tsx) – approval dialog used to change request status.

### Service Layer

The SharePoint integration is centralized in [src/webparts/dvRequestForm/service/SharepointServiceProvider.ts](src/webparts/dvRequestForm/service/SharepointServiceProvider.ts). This layer is responsible for:

- retrieving requests and related user data
- saving draft or submitted requests
- updating approval status
- reading departments and form description data
- sending notification payloads to Power Automate

## Key Functionalities

- Create requests for new SharePoint sites
- Save requests as Draft or Submit them for approval
- Load department-specific approvers from SharePoint data
- Support primary and secondary owners through People Picker
- Review and manage requests in a tabbed experience
- Approve, hold, or reject requests
- Redirect approved requests to a provisioning page using encoded configuration data

## Required SharePoint Artifacts

The solution expects the following SharePoint components to be available:

| Artifact                   | Purpose                                                   |
| -------------------------- | --------------------------------------------------------- |
| Request form list          | Stores request records and workflow state                 |
| Departments list           | Provides department names and related department metadata |
| Form description list      | Supplies descriptive content shown in the form UI         |
| Power Automate trigger URL | Used to send notification or provisioning data            |

Typical columns used by the request list include:

- Title
- SiteDescription
- SiteJustification
- Department
- SiteType
- Status
- PrimaryOwner
- SecondaryOwner
- Approvers
- Comments
- ApprovedBy

## Web Part Properties

The web part exposes the following properties through the property pane:

- RequestFormId – the SharePoint list used to store request items
- formDescription – the SharePoint list used to retrieve form description content
- PAURL – the trigger URL for Power Automate integration

## Setup and Development

### Prerequisites

- Node.js 22.14.x
- A Microsoft 365 / SharePoint Online development environment
- An SPFx-compatible tenant and app catalog for deployment

### Install Dependencies

```powershell
npm install
```

### Start Local Development

```powershell
npm run start
```

### Build for Production

```powershell
npm run build
```

## Deployment Notes

1. Build the package using the production script.
2. Package and deploy the SPFx solution to the tenant app catalog.
3. Install the app on the target SharePoint site.
4. Configure the web part properties to point to the correct SharePoint lists and Power Automate endpoint.
5. Ensure the current user and approvers have appropriate permissions to the relevant lists.

## Security and Permissions Considerations

- The solution relies on SharePoint list permissions and user access.
- Person fields and approver lookups depend on valid user identities in the tenant.
- The Power Automate endpoint should be protected and only used by authorized workflows.

## Notes

- The current implementation uses SharePoint lists and REST-based operations rather than a separate backend service.
- The UI is built with Fluent UI components and uses SPFx context objects for SharePoint-aware behavior.
- The provisioning redirect uses URL-encoded JSON configuration for downstream flow integration.

## References

- SharePoint Framework documentation
- PnPjs documentation
- Fluent UI documentation
- SharePoint REST API documentation
