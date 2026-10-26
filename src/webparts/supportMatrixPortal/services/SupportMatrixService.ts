import { getSP } from './pnpjsConfig';
import {
  LIST_NAMES,
  APPLICATION_FIELDS,
  VENDOR_FIELDS,
  BUSINESS_UNIT_FIELDS,
  SINGLE_PERSON_FIELD_NAMES,
  MULTI_PERSON_FIELD_NAMES,
  LIST_ITEM_BATCH_SIZE
} from '../constants/ListConstants';
import { IApplication, IApplicationFormData } from '../models/IApplication';
import { IVendor } from '../models/IVendor';
import { IBusinessUnit } from '../models/IBusinessUnit';
import { IPersonField, IPersonSuggestion } from '../models/IPersonField';
import { IBulkRemovalItemResult, IBulkRemovalResult } from '../models/IPeopleManagement';
import { PERSON_ROLE_DEFINITIONS } from '../constants/PersonRoleDefinitions';

/**
 * Raw shape returned by SharePoint REST for a single person field lookup value.
 */
interface IRawPerson {
  Id: number;
  Title: string;
  EMail?: string;
  Name?: string;
}

interface IRawApplicationItem {
  [key: string]: unknown;
}

/**
 * SupportMatrixService
 *
 * Centralizes all PnPjs data access for the Support Matrix Portal.
 * Every SharePoint list/field name used here is sourced from ListConstants.ts.
 */
export class SupportMatrixService {
  /**
   * Builds the $select and $expand strings needed to correctly retrieve lookup and
   * person fields from the Applications list.
   */
  private static buildApplicationsSelectAndExpand(): { select: string[]; expand: string[] } {
    const select: string[] = [
      APPLICATION_FIELDS.ID,
      APPLICATION_FIELDS.APPLICATION_NAME,
      APPLICATION_FIELDS.TIER,
      APPLICATION_FIELDS.HOSTING_MODEL,
      APPLICATION_FIELDS.APPLICATION_TYPE,
      APPLICATION_FIELDS.STATUS,
      APPLICATION_FIELDS.HIGH_LEVEL_DETAILS,
      APPLICATION_FIELDS.ACTIVE,
      `${APPLICATION_FIELDS.BUSINESS_UNIT}/Id`,
      `${APPLICATION_FIELDS.BUSINESS_UNIT}/Title`,
      `${APPLICATION_FIELDS.VENDOR}/Id`,
      `${APPLICATION_FIELDS.VENDOR}/Title`
    ];

    const expand: string[] = [APPLICATION_FIELDS.BUSINESS_UNIT, APPLICATION_FIELDS.VENDOR];

    SINGLE_PERSON_FIELD_NAMES.forEach((fieldName) => {
      select.push(`${fieldName}/Id`, `${fieldName}/Title`, `${fieldName}/EMail`, `${fieldName}/Name`);
      expand.push(fieldName);
    });

    MULTI_PERSON_FIELD_NAMES.forEach((fieldName) => {
      select.push(`${fieldName}/Id`, `${fieldName}/Title`, `${fieldName}/EMail`, `${fieldName}/Name`);
      expand.push(fieldName);
    });

    return { select, expand };
  }

  private static toPersonField(raw: unknown): IPersonField | undefined {
    if (!raw) {
      return undefined;
    }
    const person = raw as IRawPerson;
    if (person.Id === undefined || person.Id === null) {
      return undefined;
    }
    return {
      id: person.Id,
      title: person.Title || '',
      email: person.EMail || '',
      loginName: person.Name || ''
    };
  }

  private static toPersonFieldArray(raw: unknown): IPersonField[] {
    if (!raw || !Array.isArray(raw)) {
      return [];
    }
    return (raw as IRawPerson[])
      .map((p) => SupportMatrixService.toPersonField(p))
      .filter((p): p is IPersonField => p !== undefined);
  }

  private static mapApplication(item: IRawApplicationItem): IApplication {
    const businessUnit = item[APPLICATION_FIELDS.BUSINESS_UNIT] as { Id: number; Title: string } | undefined;
    const vendor = item[APPLICATION_FIELDS.VENDOR] as { Id: number; Title: string } | undefined;

    return {
      id: item[APPLICATION_FIELDS.ID] as number,
      applicationName: (item[APPLICATION_FIELDS.APPLICATION_NAME] as string) || '(No name)',

      businessUnitId: businessUnit ? businessUnit.Id : undefined,
      businessUnitName: businessUnit && businessUnit.Title ? businessUnit.Title : '',

      vendorId: vendor ? vendor.Id : undefined,
      vendorName: vendor && vendor.Title ? vendor.Title : '',

      tier: (item[APPLICATION_FIELDS.TIER] as string) || '',
      hostingModel: (item[APPLICATION_FIELDS.HOSTING_MODEL] as string) || '',
      applicationType: (item[APPLICATION_FIELDS.APPLICATION_TYPE] as string) || '',
      status: (item[APPLICATION_FIELDS.STATUS] as string) || '',
      highLevelDetails: (item[APPLICATION_FIELDS.HIGH_LEVEL_DETAILS] as string) || '',
      active: !!item[APPLICATION_FIELDS.ACTIVE],

      primaryEngineer: SupportMatrixService.toPersonField(item[APPLICATION_FIELDS.PRIMARY_ENGINEER]),
      secondaryEngineer: SupportMatrixService.toPersonField(item[APPLICATION_FIELDS.SECONDARY_ENGINEER]),
      l3SME: SupportMatrixService.toPersonField(item[APPLICATION_FIELDS.L3_SME]),
      dm: SupportMatrixService.toPersonField(item[APPLICATION_FIELDS.DM]),
      vp: SupportMatrixService.toPersonField(item[APPLICATION_FIELDS.VP]),
      businessOwner: SupportMatrixService.toPersonField(item[APPLICATION_FIELDS.BUSINESS_OWNER]),

      l3SupportEngineers: SupportMatrixService.toPersonFieldArray(item[APPLICATION_FIELDS.L3_SUPPORT_ENGINEERS]),
      cloudOperations: SupportMatrixService.toPersonFieldArray(item[APPLICATION_FIELDS.CLOUD_OPERATIONS]),
      appAdminSME: SupportMatrixService.toPersonFieldArray(item[APPLICATION_FIELDS.APP_ADMIN_SME])
    };
  }

  /**
   * Retrieves all items from the Applications list, correctly expanding lookup and
   * person/group fields. Supports large lists via automatic paging.
   */
  public static async getApplications(): Promise<IApplication[]> {
    try {
      const sp = getSP();
      const { select, expand } = SupportMatrixService.buildApplicationsSelectAndExpand();

      const list = sp.web.lists.getByTitle(LIST_NAMES.APPLICATIONS);

      const query = list.items
        .select(...select)
        .expand(...expand)
        .top(LIST_ITEM_BATCH_SIZE);

      const results: IRawApplicationItem[] = [];
      let page = await query.getPaged<IRawApplicationItem[]>();
      results.push(...page.results);

      while (page.hasNext) {
        // eslint-disable-next-line no-await-in-loop
        const nextPage = await page.getNext();
        if (!nextPage) {
          break;
        }
        page = nextPage;
        results.push(...page.results);
      }

      return results.map((item) => SupportMatrixService.mapApplication(item));
    } catch (error) {
      console.error('SupportMatrixService.getApplications failed:', error);
      throw new Error(
        'Unable to load applications. Please check that the "Applications" list name and field internal names in ListConstants.ts match your SharePoint site.'
      );
    }
  }

  /**
   * Retrieves all items from the Vendors list.
   */
  public static async getVendors(): Promise<IVendor[]> {
    try {
      const sp = getSP();
      const items = await sp.web.lists
        .getByTitle(LIST_NAMES.VENDORS)
        .items.select(
          VENDOR_FIELDS.ID,
          VENDOR_FIELDS.VENDOR_NAME,
          VENDOR_FIELDS.VENDOR_TYPE,
          VENDOR_FIELDS.ACTIVE,
          VENDOR_FIELDS.NOTES
        )
        .top(LIST_ITEM_BATCH_SIZE)
        .getAll();

      return (items as IRawApplicationItem[]).map((item) => ({
        id: item[VENDOR_FIELDS.ID] as number,
        vendorName: (item[VENDOR_FIELDS.VENDOR_NAME] as string) || '',
        vendorType: (item[VENDOR_FIELDS.VENDOR_TYPE] as string) || '',
        active: !!item[VENDOR_FIELDS.ACTIVE],
        notes: (item[VENDOR_FIELDS.NOTES] as string) || ''
      }));
    } catch (error) {
      console.error('SupportMatrixService.getVendors failed:', error);
      throw new Error(
        'Unable to load vendors. Please check that the "Vendors" list name and field internal names in ListConstants.ts match your SharePoint site.'
      );
    }
  }

  /**
   * Retrieves all items from the Business Units list.
   */
  public static async getBusinessUnits(): Promise<IBusinessUnit[]> {
    try {
      const sp = getSP();
      const items = await sp.web.lists
        .getByTitle(LIST_NAMES.BUSINESS_UNITS)
        .items.select(
          BUSINESS_UNIT_FIELDS.ID,
          BUSINESS_UNIT_FIELDS.BUSINESS_UNIT_NAME,
          BUSINESS_UNIT_FIELDS.BUSINESS_UNIT_CODE,
          BUSINESS_UNIT_FIELDS.ACTIVE,
          BUSINESS_UNIT_FIELDS.NOTES
        )
        .top(LIST_ITEM_BATCH_SIZE)
        .getAll();

      return (items as IRawApplicationItem[]).map((item) => ({
        id: item[BUSINESS_UNIT_FIELDS.ID] as number,
        businessUnitName: (item[BUSINESS_UNIT_FIELDS.BUSINESS_UNIT_NAME] as string) || '',
        businessUnitCode: (item[BUSINESS_UNIT_FIELDS.BUSINESS_UNIT_CODE] as string) || '',
        active: !!item[BUSINESS_UNIT_FIELDS.ACTIVE],
        notes: (item[BUSINESS_UNIT_FIELDS.NOTES] as string) || ''
      }));
    } catch (error) {
      console.error('SupportMatrixService.getBusinessUnits failed:', error);
      throw new Error(
        'Unable to load business units. Please check that the "Business Units" list name and field internal names in ListConstants.ts match your SharePoint site.'
      );
    }
  }

  /**
   * Searches SharePoint's people picker for users matching the given text.
   */
  public static async searchPeople(query: string): Promise<IPersonSuggestion[]> {
    if (!query || query.trim().length < 2) {
      return [];
    }
    try {
      const sp = getSP();
      const results = await sp.profiles.clientPeoplePickerSearchUser({
        QueryString: query,
        MaximumEntitySuggestions: 10,
        AllowEmailAddresses: true,
        AllowMultipleEntities: false
      });

      return (results || [])
        .filter((entity) => !!entity.Key)
        .map((entity) => ({
          key: entity.Key,
          text: entity.DisplayText || entity.Description || entity.Key,
          secondaryText: (entity.EntityData && entity.EntityData.Email) || entity.Description || '',
          loginName: entity.Key
        }));
    } catch (error) {
      console.error('SupportMatrixService.searchPeople failed:', error);
      return [];
    }
  }

  /**
   * Resolves a people-picker login name/claim into a full IPersonField (with numeric id),
   * ensuring the user exists in the site's user information list.
   */
  public static async resolvePerson(loginName: string, displayText: string, email: string): Promise<IPersonField> {
    const sp = getSP();
    const ensured = await sp.web.ensureUser(loginName);
    return {
      id: ensured.data.Id,
      title: ensured.data.Title || displayText,
      email: ensured.data.Email || email,
      loginName: ensured.data.LoginName || loginName
    };
  }

  /**
   * Converts a list of selected people into a valid numeric SharePoint user id array for a
   * multi-value Person/Group field. The items.add()/items.getById().update() PnPjs endpoints
   * used by this project expect a direct number[] for multi-value Person fields (e.g.
   * FieldNameId: [12, 25]) - NOT the legacy { results: [...] } wrapper, which SharePoint
   * rejects with a 400 "StartObject ... StartArray was expected" error. This mirrors the
   * format already proven working in removePersonFromApplications().
   */
  private static toMultiPersonIdArray(people: IPersonField[]): number[] {
    const seen = new Set<number>();
    people.forEach((p) => {
      const id = Number(p.id);
      if (Number.isInteger(id) && id > 0) {
        seen.add(id);
      }
    });
    return Array.from(seen);
  }

  /**
   * Converts form data into the SharePoint REST field payload used for create/update.
   */
  private static buildApplicationPayload(data: IApplicationFormData): Record<string, unknown> {
    const payload: Record<string, unknown> = {
      [APPLICATION_FIELDS.APPLICATION_NAME]: data.applicationName,
      [APPLICATION_FIELDS.TIER]: data.tier,
      [APPLICATION_FIELDS.HOSTING_MODEL]: data.hostingModel,
      [APPLICATION_FIELDS.APPLICATION_TYPE]: data.applicationType,
      [APPLICATION_FIELDS.STATUS]: data.status,
      [APPLICATION_FIELDS.HIGH_LEVEL_DETAILS]: data.highLevelDetails,
      [APPLICATION_FIELDS.ACTIVE]: data.active,
      [`${APPLICATION_FIELDS.BUSINESS_UNIT_ID}`]: data.businessUnitId || null,
      [`${APPLICATION_FIELDS.VENDOR_ID}`]: data.vendorId || null,
      [`${APPLICATION_FIELDS.PRIMARY_ENGINEER}Id`]: data.primaryEngineer ? data.primaryEngineer.id : null,
      [`${APPLICATION_FIELDS.SECONDARY_ENGINEER}Id`]: data.secondaryEngineer ? data.secondaryEngineer.id : null,
      [`${APPLICATION_FIELDS.L3_SME}Id`]: data.l3SME ? data.l3SME.id : null,
      [`${APPLICATION_FIELDS.DM}Id`]: data.dm ? data.dm.id : null,
      [`${APPLICATION_FIELDS.VP}Id`]: data.vp ? data.vp.id : null,
      [`${APPLICATION_FIELDS.BUSINESS_OWNER}Id`]: data.businessOwner ? data.businessOwner.id : null,
      [`${APPLICATION_FIELDS.L3_SUPPORT_ENGINEERS}Id`]: SupportMatrixService.toMultiPersonIdArray(data.l3SupportEngineers),
      [`${APPLICATION_FIELDS.CLOUD_OPERATIONS}Id`]: SupportMatrixService.toMultiPersonIdArray(data.cloudOperations),
      [`${APPLICATION_FIELDS.APP_ADMIN_SME}Id`]: SupportMatrixService.toMultiPersonIdArray(data.appAdminSME)
    };

    return payload;
  }

  /**
   * Validates that the multi-value Person/Group properties of an application payload are
   * plain number[] (never a { results: [...] } wrapper or an array of person objects) before
   * it is sent to SharePoint.
   */
  private static validateApplicationPayload(payload: Record<string, unknown>): void {
    const multiValueKeys = [
      `${APPLICATION_FIELDS.L3_SUPPORT_ENGINEERS}Id`,
      `${APPLICATION_FIELDS.CLOUD_OPERATIONS}Id`,
      `${APPLICATION_FIELDS.APP_ADMIN_SME}Id`
    ];
    multiValueKeys.forEach((key) => {
      const value = payload[key];
      if (!Array.isArray(value) || !value.every((v) => Number.isInteger(v) && v > 0)) {
        throw new Error(`Invalid update payload for ${key}: expected number[] but received ${JSON.stringify(value)}.`);
      }
    });
  }

  /**
   * Creates a new item in the Applications list. Returns the new item's id.
   */
  public static async createApplication(data: IApplicationFormData): Promise<number> {
    try {
      const sp = getSP();
      const payload = SupportMatrixService.buildApplicationPayload(data);
      SupportMatrixService.validateApplicationPayload(payload);
      console.log('SupportMatrixService.createApplication payload:', JSON.stringify(payload));
      const result = await sp.web.lists.getByTitle(LIST_NAMES.APPLICATIONS).items.add(payload);
      return result.data.Id as number;
    } catch (error) {
      console.error('SupportMatrixService.createApplication failed:', error);
      throw new Error('Unable to create the application. Please check the form values and try again.');
    }
  }

  /**
   * Updates an existing item in the Applications list.
   */
  public static async updateApplication(id: number, data: IApplicationFormData): Promise<void> {
    try {
      const sp = getSP();
      const payload = SupportMatrixService.buildApplicationPayload(data);
      SupportMatrixService.validateApplicationPayload(payload);
      console.log('SupportMatrixService.updateApplication payload:', id, JSON.stringify(payload));
      await sp.web.lists.getByTitle(LIST_NAMES.APPLICATIONS).items.getById(id).update(payload);
    } catch (error) {
      console.error('SupportMatrixService.updateApplication failed:', error);
      throw new Error('Unable to update the application. Please check the form values and try again.');
    }
  }

  /**
   * Deletes an item from the Applications list.
   */
  public static async deleteApplication(id: number): Promise<void> {
    try {
      const sp = getSP();
      await sp.web.lists.getByTitle(LIST_NAMES.APPLICATIONS).items.getById(id).delete();
    } catch (error) {
      console.error('SupportMatrixService.deleteApplication failed:', error);
      throw new Error('Unable to delete the application. Please try again.');
    }
  }

  /**
   * Applies a partial field update to an Applications list item. Unlike updateApplication(),
   * this only touches the fields explicitly present in the payload - used by People
   * Management so unrelated application values are never modified.
   */
  private static async updateApplicationFields(id: number, payload: Record<string, unknown>): Promise<void> {
    const sp = getSP();
    await sp.web.lists.getByTitle(LIST_NAMES.APPLICATIONS).items.getById(id).update(payload);
  }

  /**
   * Returns true if the given update error looks like a SharePoint "value required"
   * validation rejection (i.e. the field cannot legally be left empty), as opposed to a
   * genuine failure (permissions, network, throttling, etc).
   */
  private static isRequiredFieldValidationError(error: unknown): boolean {
    const message = error instanceof Error ? error.message : String(error || '');
    const lower = message.toLowerCase();
    return (
      lower.indexOf('required') !== -1 ||
      lower.indexOf('cannot be blank') !== -1 ||
      lower.indexOf('cannot be empty') !== -1 ||
      lower.indexOf('value does not fall within the expected range') !== -1
    );
  }

  /**
   * Builds a valid numeric SharePoint user id array for a multi-value Person/Group field:
   * dedupes, drops anything that isn't a positive integer, and excludes the selected person.
   */
  private static buildRemainingUserIds(currentPeople: IPersonField[], personIdToRemove: number): number[] {
    const seen = new Set<number>();
    currentPeople.forEach((p) => {
      const id = Number(p.id);
      if (Number.isInteger(id) && id > 0 && id !== personIdToRemove) {
        seen.add(id);
      }
    });
    return Array.from(seen);
  }

  /**
   * Validates that a People Management update payload only contains the shapes SharePoint's
   * item.update() (MERGE) endpoint actually accepts for Person/Group fields: a plain
   * number[] for multi-value fields (never a { results: [...] } wrapper or an array of
   * person objects), and either null or a single positive integer for single-value fields.
   * Throws before the request is sent if an invalid shape is detected.
   */
  private static validatePersonFieldPayload(
    payload: Record<string, unknown>,
    multiRoleDefs: { fieldName: string }[],
    singleRoleDefs: { fieldName: string }[]
  ): void {
    multiRoleDefs.forEach((def) => {
      const key = `${def.fieldName}Id`;
      const value = payload[key];
      if (!(key in payload)) {
        return;
      }
      if (!Array.isArray(value) || !value.every((v) => Number.isInteger(v) && v > 0)) {
        throw new Error(`Invalid payload for field ${key}: expected number[] but received ${JSON.stringify(value)}.`);
      }
    });
    singleRoleDefs.forEach((def) => {
      const key = `${def.fieldName}Id`;
      const value = payload[key];
      if (!(key in payload)) {
        return;
      }
      if (value !== null && !(Number.isInteger(value) && (value as number) > 0)) {
        throw new Error(`Invalid payload for field ${key}: expected null or a positive integer but received ${JSON.stringify(value)}.`);
      }
    });
  }

  /**
   * Removes the given person from every selected application's matching Person/Group role
   * fields. Every matching field (single and multi-value) is updated in a single combined
   * call per application.
   *
   * IMPORTANT: unlike the Add/Edit Application form payload (which posts through items.add()),
   * this goes through items.getById(id).update() (HTTP MERGE). That endpoint expects
   * multi-value Person/Group fields as a direct number[] of SharePoint user ids
   * (e.g. FieldNameId: [12, 25]) - NOT the legacy { results: [...] } wrapper, which SharePoint
   * rejects here with a 400 "StartObject ... StartArray was expected" error.
   *
   * Multi-value fields always keep every other person untouched (only the selected person's
   * id is filtered out of the existing results, deduped, and validated as positive integers).
   * Single-value fields are only ever cleared if SharePoint actually accepts an empty value
   * for that item; if SharePoint rejects the combined update because one of the single-value
   * fields is required, the call is retried with just the multi-value changes so those are not
   * lost, and the required single-value role(s) are reported back as a blocked validation
   * failure instead of being silently dropped or falsely reported as removed.
   */
  public static async removePersonFromApplications(
    personId: number,
    targets: { application: IApplication; roleKeys: string[] }[]
  ): Promise<IBulkRemovalResult> {
    const items: IBulkRemovalItemResult[] = [];

    for (const target of targets) {
      const multiRoleDefs = PERSON_ROLE_DEFINITIONS.filter(
        (def) => def.isMulti && target.roleKeys.indexOf(def.key) !== -1
      );
      const singleRoleDefs = PERSON_ROLE_DEFINITIONS.filter(
        (def) => !def.isMulti && target.roleKeys.indexOf(def.key) !== -1
      );

      const fullPayload: Record<string, unknown> = {};
      multiRoleDefs.forEach((def) => {
        fullPayload[`${def.fieldName}Id`] = SupportMatrixService.buildRemainingUserIds(
          def.getValue(target.application),
          personId
        );
      });
      singleRoleDefs.forEach((def) => {
        fullPayload[`${def.fieldName}Id`] = null;
      });

      const result: IBulkRemovalItemResult = {
        applicationId: target.application.id,
        applicationName: target.application.applicationName,
        removedRoleLabels: [],
        blockedRoleLabels: []
      };

      if (Object.keys(fullPayload).length === 0) {
        items.push(result);
        continue;
      }

      try {
        SupportMatrixService.validatePersonFieldPayload(fullPayload, multiRoleDefs, singleRoleDefs);
      } catch (validationError) {
        console.error(
          `SupportMatrixService.removePersonFromApplications: invalid payload for application ${target.application.id} (${target.application.applicationName}):`,
          validationError
        );
        result.error = validationError instanceof Error ? validationError.message : 'Invalid update payload.';
        items.push(result);
        continue;
      }

      console.log(
        'People Management update payload',
        target.application.id,
        JSON.stringify(fullPayload)
      );

      try {
        // eslint-disable-next-line no-await-in-loop
        await SupportMatrixService.updateApplicationFields(target.application.id, fullPayload);
        result.removedRoleLabels = [...multiRoleDefs, ...singleRoleDefs].map((def) => def.label);
      } catch (fullError) {
        console.error(
          `SupportMatrixService.removePersonFromApplications: combined update failed for application ${target.application.id} (${target.application.applicationName}):`,
          fullError
        );

        if (singleRoleDefs.length > 0 && SupportMatrixService.isRequiredFieldValidationError(fullError)) {
          // Likely cause: one of the single-value fields is a required column and cannot be
          // cleared. Retry with only the multi-value changes so those are not lost, and
          // report the required single-value role(s) as a validation failure.
          if (multiRoleDefs.length > 0) {
            const multiOnlyPayload: Record<string, unknown> = {};
            multiRoleDefs.forEach((def) => {
              multiOnlyPayload[`${def.fieldName}Id`] = fullPayload[`${def.fieldName}Id`];
            });
            try {
              // eslint-disable-next-line no-await-in-loop
              await SupportMatrixService.updateApplicationFields(target.application.id, multiOnlyPayload);
              result.removedRoleLabels = multiRoleDefs.map((def) => def.label);
              result.blockedRoleLabels = singleRoleDefs.map((def) => def.label);
            } catch (multiError) {
              console.error(
                `SupportMatrixService.removePersonFromApplications: multi-value-only retry failed for application ${target.application.id} (${target.application.applicationName}):`,
                multiError
              );
              result.error = multiError instanceof Error ? multiError.message : 'Unable to update this application.';
            }
          } else {
            result.blockedRoleLabels = singleRoleDefs.map((def) => def.label);
          }
        } else {
          result.error = fullError instanceof Error ? fullError.message : 'Unable to update this application.';
        }
      }

      items.push(result);
    }

    return { items };
  }
}