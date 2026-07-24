import { IPersonField } from './IPersonField';

export interface IApplication {
  id: number;
  applicationName: string;

  businessUnitId?: number;
  businessUnitName: string;

  vendorId?: number;
  vendorName: string;

  tier: string;
  hostingModel: string;
  applicationType: string;
  status: string;
  highLevelDetails: string;
  active: boolean;

  primaryEngineer?: IPersonField;
  secondaryEngineer?: IPersonField;
  l3SME?: IPersonField;
  dm?: IPersonField;
  vp?: IPersonField;
  businessOwner?: IPersonField;

  l3SupportEngineers: IPersonField[];
  cloudOperations: IPersonField[];
  appAdminSME: IPersonField[];
}

/**
 * Shape of the data collected from the Add/Edit Application form, before being
 * converted into a SharePoint REST payload by SupportMatrixService.
 */
export interface IApplicationFormData {
  applicationName: string;
  businessUnitId?: number;
  vendorId?: number;
  tier: string;
  hostingModel: string;
  applicationType: string;
  status: string;
  highLevelDetails: string;
  active: boolean;

  primaryEngineer?: IPersonField;
  secondaryEngineer?: IPersonField;
  l3SME?: IPersonField;
  dm?: IPersonField;
  vp?: IPersonField;
  businessOwner?: IPersonField;

  l3SupportEngineers: IPersonField[];
  cloudOperations: IPersonField[];
  appAdminSME: IPersonField[];
}

export interface IDataQualityIssues {
  missingPrimaryEngineer: IApplication[];
  missingVendor: IApplication[];
  missingBusinessUnit: IApplication[];
  missingTier: IApplication[];
  missingHostingModel: IApplication[];
  activeButDecommissioned: IApplication[];
  inactiveApplications: IApplication[];
}
