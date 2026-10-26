import { IApplication } from './IApplication';

/**
 * One application where the selected person was found, together with every role
 * (person/group field) they occupy on that application.
 */
export interface IPersonAssignmentMatch {
  application: IApplication;
  roleKeys: string[];
}

/**
 * Outcome of attempting to remove the selected person from a single application's
 * matching role fields.
 */
export interface IBulkRemovalItemResult {
  applicationId: number;
  applicationName: string;
  removedRoleLabels: string[];
  blockedRoleLabels: string[];
  error?: string;
}

export interface IBulkRemovalResult {
  items: IBulkRemovalItemResult[];
}
