/**
 * Represents a resolved SharePoint person/group field value.
 */
export interface IPersonField {
  id: number;
  title: string;
  email: string;
  loginName?: string;
}

export type PersonFieldValue = IPersonField | undefined;
export type MultiPersonFieldValue = IPersonField[];

/**
 * A candidate person returned from the SharePoint people picker search, before it has
 * been resolved to a numeric SharePoint user id via ensureUser().
 */
export interface IPersonSuggestion {
  key: string;
  text: string;
  secondaryText: string;
  loginName: string;
}
