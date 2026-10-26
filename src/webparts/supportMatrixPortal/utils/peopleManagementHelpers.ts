import { IApplication } from '../models/IApplication';
import { IPersonAssignmentMatch } from '../models/IPeopleManagement';
import { IPersonField } from '../models/IPersonField';
import { PERSON_ROLE_DEFINITIONS, IPersonRoleDefinition } from '../constants/PersonRoleDefinitions';

/**
 * Builds a deduplicated list of every person currently assigned to at least one
 * application, across every Person/Group role field defined in PERSON_ROLE_DEFINITIONS
 * (single- and multi-value alike). This lets People Management find people - such as a
 * former employee - who are still referenced in existing application data even if they
 * can no longer be found through a live directory search.
 *
 * No directory/Graph calls are made here - only the identity data already present in the
 * already-loaded Applications collection is used, so this stays a single pass over data
 * the portal already has in memory.
 */
export const extractAssignedPeople = (applications: IApplication[]): IPersonField[] => {
  const byId = new Map<number, IPersonField>();
  const byEmail = new Map<string, IPersonField>();

  applications.forEach((app) => {
    PERSON_ROLE_DEFINITIONS.forEach((def) => {
      def.getValue(app).forEach((person) => {
        if (!person) {
          return;
        }
        // Prefer the stable SharePoint user id for dedup; fall back to email/login only
        // when an id is not available, since display names are not guaranteed unique.
        if (person.id) {
          if (!byId.has(person.id)) {
            byId.set(person.id, person);
          }
          return;
        }
        const fallbackKey = (person.email || person.loginName || '').toLowerCase();
        if (fallbackKey && !byEmail.has(fallbackKey)) {
          byEmail.set(fallbackKey, person);
        }
      });
    });
  });

  return Array.from(byId.values()).concat(Array.from(byEmail.values()));
};

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
