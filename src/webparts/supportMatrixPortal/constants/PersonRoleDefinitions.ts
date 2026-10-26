import { IApplication } from '../models/IApplication';
import { IPersonField } from '../models/IPersonField';
import { APPLICATION_FIELDS } from './ListConstants';

/**
 * Describes one Person/Group role field on the Applications list, sourced directly from
 * the existing APPLICATION_FIELDS internal names and IApplication model - nothing here is
 * a guessed/hardcoded field name.
 */
export interface IPersonRoleDefinition {
  key: string;
  label: string;
  /** SharePoint internal field name (from APPLICATION_FIELDS), used to build update payloads. */
  fieldName: string;
  isMulti: boolean;
  getValue: (app: IApplication) => IPersonField[];
}

const single = (
  key: string,
  label: string,
  fieldName: string,
  getter: (app: IApplication) => IPersonField | undefined
): IPersonRoleDefinition => ({
  key,
  label,
  fieldName,
  isMulti: false,
  getValue: (app: IApplication) => {
    const person = getter(app);
    return person ? [person] : [];
  }
});

const multi = (
  key: string,
  label: string,
  fieldName: string,
  getter: (app: IApplication) => IPersonField[]
): IPersonRoleDefinition => ({
  key,
  label,
  fieldName,
  isMulti: true,
  getValue: getter
});

/**
 * All Person/Group role fields that People Management can search and bulk-remove from.
 * This is the single source of truth - derived from the fields already used by
 * ApplicationFormPanel/SupportMatrixService, so every existing role is covered automatically.
 */
export const PERSON_ROLE_DEFINITIONS: IPersonRoleDefinition[] = [
  single('primaryEngineer', 'Primary Engineer', APPLICATION_FIELDS.PRIMARY_ENGINEER, (a) => a.primaryEngineer),
  single('secondaryEngineer', 'Secondary Engineer', APPLICATION_FIELDS.SECONDARY_ENGINEER, (a) => a.secondaryEngineer),
  single('l3SME', 'L3 SME', APPLICATION_FIELDS.L3_SME, (a) => a.l3SME),
  single('dm', 'DM', APPLICATION_FIELDS.DM, (a) => a.dm),
  single('vp', 'VP', APPLICATION_FIELDS.VP, (a) => a.vp),
  single('businessOwner', 'Business Owner', APPLICATION_FIELDS.BUSINESS_OWNER, (a) => a.businessOwner),
  multi('l3SupportEngineers', 'L3 Support Engineers', APPLICATION_FIELDS.L3_SUPPORT_ENGINEERS, (a) => a.l3SupportEngineers),
  multi('cloudOperations', 'Cloud Operations', APPLICATION_FIELDS.CLOUD_OPERATIONS, (a) => a.cloudOperations),
  multi('appAdminSME', 'App Admin SME', APPLICATION_FIELDS.APP_ADMIN_SME, (a) => a.appAdminSME)
];
