# Support Matrix Portal

An SPFx (SharePoint Framework) React web part that provides a centralized, read-only dashboard for
application support ownership and inventory data, sourced from three SharePoint lists: **Applications**,
**Vendors**, and **Business Units**.

## Summary

- Premium enterprise-style dashboard header
- KPI cards (Total, Active, Decommissioned, Vendor, Cloud, On-Prem, Tier 1, Missing Primary Engineer)
- CSS/HTML mini bar charts (Status, Business Unit, Vendor, Hosting Model, Tier) - no external chart library
- Searchable, filterable, sortable, paginated Application Inventory table
- Application Details side panel (Application Info, Ownership, Support Team) with person pills
- Dynamically computed Data Quality tab (never persisted to SharePoint)
- "My Applications" toggle, matched against `context.pageContext.user.email`
- All list/field internal names centralized in `src/webparts/supportMatrixPortal/constants/ListConstants.ts`
- Data access via PnPjs (`@pnp/sp`) with `$select`/`$expand` for lookups and person fields

This is a **read-only v1**. No create/update/delete operations are implemented.

## Prerequisites

- Node.js 16.13.0 - 18.x (SPFx 1.18.2 requirement)
- SharePoint Online tenant with the three lists created as described below
- Permissions to sideload/deploy SPFx solutions (App Catalog access)

## Folder Structure

```
support-matrix-portal/
├── config/                                  SPFx build configuration
├── src/
│   ├── index.ts
│   └── webparts/
│       └── supportMatrixPortal/
│           ├── SupportMatrixPortalWebPart.ts
│           ├── SupportMatrixPortalWebPart.manifest.json
│           ├── loc/                         Localized strings
│           ├── constants/
│           │   └── ListConstants.ts         <-- list & field internal names (EDIT HERE if needed)
│           ├── models/
│           │   ├── IApplication.ts
│           │   ├── IVendor.ts
│           │   ├── IBusinessUnit.ts
│           │   └── IPersonField.ts
│           ├── services/
│           │   ├── pnpjsConfig.ts           PnPjs SPFI initialization
│           │   └── SupportMatrixService.ts  All SharePoint data access
│           ├── utils/
│           │   └── applicationHelpers.ts    Search, data quality, filter helpers
│           └── components/
│               ├── SupportMatrixPortal.tsx
│               ├── DashboardCards.tsx
│               ├── ChartsSection.tsx
│               ├── FiltersBar.tsx
│               ├── ApplicationsTable.tsx
│               ├── ApplicationDetailsPanel.tsx
│               ├── DataQualityPanel.tsx
│               ├── StatusBadge.tsx
│               ├── PersonPill.tsx
│               └── *.module.scss            Component-scoped styles
├── package.json
├── tsconfig.json
├── gulpfile.js
└── .yo-rc.json
```

## Installation

1. Copy this entire `support-matrix-portal` folder to your machine.
2. Open a terminal in the project root and install dependencies:

   ```bash
   npm install
   ```

3. Trust the local dev certificate (first time only):

   ```bash
   gulp trust-dev-cert
   ```

4. Run locally against the SharePoint Workbench:

   ```bash
   gulp serve
   ```

   Then open your tenant's hosted workbench, e.g. `https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`,
   and add the "Support Matrix Portal" web part. The web part reads from the lists on whichever site the
   workbench page belongs to, so open the workbench on the site that contains the Applications/Vendors/
   Business Units lists.

5. To build and package for deployment:

   ```bash
   gulp bundle --ship
   gulp package-solution --ship
   ```

   The resulting package is at `solution/support-matrix-portal.sppkg`. Upload it to your tenant or site
   App Catalog and add the web part to a modern page.

## Configuring List & Field Names

Everything the app reads from SharePoint is defined in one file:

`src/webparts/supportMatrixPortal/constants/ListConstants.ts`

If your SharePoint lists differ from the internal names assumed here (most commonly because a column's
display name contained spaces and SharePoint encoded them, e.g. `Business_x0020_Unit`), update the
corresponding constant in this file. No other file needs to change.

To find the real internal name of any column:
- Site Settings → List Settings → click the column → look at `Field=` in the URL, or
- Browse to: `<site>/_api/web/lists/getbytitle('Applications')/fields?$select=InternalName,Title`

## SharePoint List Schema Expected

**Applications**
| Display Name | Internal Name (default assumed) | Type |
|---|---|---|
| Application Name | `Title` | Single line of text |
| Business Unit | `BusinessUnit` | Lookup → Business Units |
| Vendor | `Vendor` | Lookup → Vendors |
| Tier | `Tier` | Choice (1, 2, 3) |
| Hosting Model | `HostingModel` | Choice (Cloud, On-Prem, Hybrid) |
| Application Type | `ApplicationType` | Choice |
| Status | `Status` | Choice |
| High Level Details | `HighLevelDetails` | Multiple lines of text |
| Active | `Active` | Yes/No |
| Primary Engineer | `PrimaryEngineer` | Person |
| Secondary Engineer | `SecondaryEngineer` | Person |
| L3 SME | `L3SME` | Person |
| DM | `DM` | Person |
| VP | `VP` | Person |
| Business Owner | `BusinessOwner` | Person |
| L3 Support Engineers | `L3SupportEngineers` | Person (multi) |
| Cloud Operations | `CloudOperations` | Person (multi) |
| App Admin SME | `AppAdminSME` | Person (multi) |

**Vendors**: `Title` (Vendor Name), `VendorType`, `Active`, `Notes`

**Business Units**: `Title` (Business Unit Name), `BusinessUnitCode`, `Active`, `Notes`

## Notes

- The web part loads all applications once (with automatic PnPjs paging for lists larger than the
  configured batch size), then performs search/filter/sort/pagination client-side in React for a fast,
  responsive UI - suitable well beyond the 235+ item scale mentioned in requirements.
- Loading and error states are handled explicitly ("Loading applications...", "Unable to load data",
  "No applications found"). Technical errors are also logged to the browser console.
- This version is strictly read-only; no list writes are performed.
