/**
 * ListConstants.ts
 *
 * Centralized configuration for SharePoint list names and field internal names.
 *
 * IMPORTANT:
 * SharePoint generates an "internal name" for every column at creation time, based on the
 * display name that was used at that moment. If a column is later renamed (e.g. "Title" ->
 * "Application Name"), the INTERNAL NAME does NOT change - only the display name changes.
 * Also, when a column display name contains spaces or special characters, SharePoint may
 * encode them (e.g. "Business Unit" -> "Business_x0020_Unit").
 *
 * If any of the internal names below do not match your actual SharePoint site, update them
 * HERE ONLY. Every service/component in this solution reads field names from this file -
 * nothing is hardcoded elsewhere.
 *
 * To verify the real internal name of a column:
 *   1. Go to List Settings -> click the column -> look at the URL query string "Field=".
 *   2. Or query: <siteUrl>/_api/web/lists/getbytitle('Applications')/fields?$select=InternalName,Title
 */

export const LIST_NAMES = {
  APPLICATIONS: 'Applications',
  VENDORS: 'Vendors',
  BUSINESS_UNITS: 'Business Units'
};

/**
 * Field internal names for the Applications list.
 */
export const APPLICATION_FIELDS = {
  // Title column was renamed to "Application Name" - internal name stays "Title"
  APPLICATION_NAME: 'Title',
  BUSINESS_UNIT: 'BusinessUnit',
  BUSINESS_UNIT_ID: 'BusinessUnitId',
  VENDOR: 'Vendor',
  VENDOR_ID: 'VendorId',
  TIER: 'Tier',
  HOSTING_MODEL: 'HostingModel',
  APPLICATION_TYPE: 'ApplicationType',
  STATUS: 'Status',
  HIGH_LEVEL_DETAILS: 'HighLevelDetails',
  ACTIVE: 'Active',

  // Single person fields
  PRIMARY_ENGINEER: 'PrimaryEngineer',
  SECONDARY_ENGINEER: 'SecondaryEngineer',
  L3_SME: 'L3SME',
  DM: 'DM',
  VP: 'VP',
  BUSINESS_OWNER: 'BusinessOwner',

  // Multi person fields
  L3_SUPPORT_ENGINEERS: 'L3SupportEngineers',
  CLOUD_OPERATIONS: 'CloudOperations',
  APP_ADMIN_SME: 'AppAdminSME',

  ID: 'Id'
};

/**
 * Field internal names for the Vendors list.
 */
export const VENDOR_FIELDS = {
  // Title column was renamed to "Vendor Name" - internal name stays "Title"
  VENDOR_NAME: 'Title',
  VENDOR_TYPE: 'VendorType',
  ACTIVE: 'Active',
  NOTES: 'Notes',
  ID: 'Id'
};

/**
 * Field internal names for the Business Units list.
 */
export const BUSINESS_UNIT_FIELDS = {
  // Title column was renamed to "Business Unit Name" - internal name stays "Title"
  BUSINESS_UNIT_NAME: 'Title',
  BUSINESS_UNIT_CODE: 'BusinessUnitCode',
  ACTIVE: 'Active',
  NOTES: 'Notes',
  ID: 'Id'
};

export const CHOICE_VALUES = {
  TIER: ['1', '2', '3'],
  HOSTING_MODEL: ['Cloud', 'On-Prem', 'Hybrid'],
  APPLICATION_TYPE: ['Operation', 'Vendor Application', 'Decommissioned', 'To Be Transitioned'],
  STATUS: ['Active', 'Under Review', 'To Be Transitioned', 'Decommissioning', 'Decommissioned'],
  VENDOR_TYPE: ['Internal', 'External', 'Partner']
};

/** Single person fields on the Applications list that should be expanded via PnPjs. */
export const SINGLE_PERSON_FIELD_NAMES: string[] = [
  APPLICATION_FIELDS.PRIMARY_ENGINEER,
  APPLICATION_FIELDS.SECONDARY_ENGINEER,
  APPLICATION_FIELDS.L3_SME,
  APPLICATION_FIELDS.DM,
  APPLICATION_FIELDS.VP,
  APPLICATION_FIELDS.BUSINESS_OWNER
];

/** Multi person fields on the Applications list that should be expanded via PnPjs. */
export const MULTI_PERSON_FIELD_NAMES: string[] = [
  APPLICATION_FIELDS.L3_SUPPORT_ENGINEERS,
  APPLICATION_FIELDS.CLOUD_OPERATIONS,
  APPLICATION_FIELDS.APP_ADMIN_SME
];

/** Page size options for the applications table. */
export const PAGE_SIZE_OPTIONS: number[] = [10, 25, 50, 100];

export const DEFAULT_PAGE_SIZE = 25;

/** Maximum number of items PnPjs will retrieve per underlying page request. */
export const LIST_ITEM_BATCH_SIZE = 200;
