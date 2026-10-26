import { IApplication } from '../models/IApplication';
import { IPersonAssignmentMatch } from '../models/IPeopleManagement';
import { PERSON_ROLE_DEFINITIONS, IPersonRoleDefinition } from '../constants/PersonRoleDefinitions';

/**
 * Finds every application where the given person (matched by SharePoint user id) appears
 * in any Person/Group role field, collapsing multiple role matches on the same application
 * into a single entry.
 */
export const findPersonAssignments = (applications: IApplication[], personId: number | undefined): IPersonAssignmentMatch[] => {
  if (!personId) {
    return [];
  }

  const matches: IPersonAssignmentMatch[] = [];
  applications.forEach((app) => {
    const roleKeys: string[] = [];
    PERSON_ROLE_DEFINITIONS.forEach((def) => {
      const values = def.getValue(app);
      if (values.some((p) => p.id === personId)) {
        roleKeys.push(def.key);
      }
    });
    if (roleKeys.length > 0) {
      matches.push({ application: app, roleKeys });
    }
  });

  return matches;
};

export interface IRoleSummaryEntry {
  role: IPersonRoleDefinition;
  count: number;
}

/**
 * Dynamically computes a per-role count across the matched applications. Only roles that
 * actually have at least one match are returned, so this never assumes a fixed role list.
 */
export const computeRoleSummary = (matches: IPersonAssignmentMatch[]): IRoleSummaryEntry[] => {
  return PERSON_ROLE_DEFINITIONS.map((def) => ({
    role: def,
    count: matches.filter((m) => m.roleKeys.indexOf(def.key) !== -1).length
  })).filter((entry) => entry.count > 0);
};

export const getRoleLabels = (roleKeys: string[]): string[] => {
  return roleKeys
    .map((key) => PERSON_ROLE_DEFINITIONS.filter((def) => def.key === key)[0])
    .filter((def): def is IPersonRoleDefinition => !!def)
    .map((def) => def.label);
};
