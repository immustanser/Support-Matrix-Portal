import { IApplication, IDataQualityIssues } from '../models/IApplication';

/**
 * Returns true if the given email matches any person field (single or multi) on the application.
 */
export const isUserOnApplication = (app: IApplication, userEmail: string): boolean => {
  if (!userEmail) {
    return false;
  }
  const target = userEmail.toLowerCase();

  const singlePersons = [app.primaryEngineer, app.secondaryEngineer, app.l3SME, app.dm, app.vp, app.businessOwner];
  if (singlePersons.some((p) => !!p && !!p.email && p.email.toLowerCase() === target)) {
    return true;
  }

  const multiPersons = [...app.l3SupportEngineers, ...app.cloudOperations, ...app.appAdminSME];
  return multiPersons.some((p) => !!p.email && p.email.toLowerCase() === target);
};

/**
 * Builds a single lowercase searchable string for an application, covering every field
 * required by the global search box.
 */
export const buildSearchIndex = (app: IApplication): string => {
  const personNames = [
    app.primaryEngineer?.title,
    app.secondaryEngineer?.title,
    app.l3SME?.title,
    app.dm?.title,
    app.vp?.title,
    app.businessOwner?.title,
    ...app.l3SupportEngineers.map((p) => p.title),
    ...app.cloudOperations.map((p) => p.title),
    ...app.appAdminSME.map((p) => p.title)
  ];

  const parts = [
    app.applicationName,
    app.businessUnitName,
    app.vendorName,
    app.status,
    app.applicationType,
    ...personNames
  ];

  return parts.filter(Boolean).join(' | ').toLowerCase();
};

/**
 * Computes Data Quality issues dynamically from the loaded applications. Nothing here is
 * persisted back to SharePoint - it is always calculated on the fly.
 */
export const computeDataQualityIssues = (applications: IApplication[]): IDataQualityIssues => {
  return {
    missingPrimaryEngineer: applications.filter((a) => !a.primaryEngineer || !a.primaryEngineer.title),
    missingVendor: applications.filter((a) => !a.vendorName),
    missingBusinessUnit: applications.filter((a) => !a.businessUnitName),
    missingTier: applications.filter((a) => !a.tier),
    missingHostingModel: applications.filter((a) => !a.hostingModel),
    activeButDecommissioned: applications.filter((a) => a.active === true && a.status === 'Decommissioned'),
    inactiveApplications: applications.filter((a) => a.active === false)
  };
};

export const getUniqueSortedValues = (applications: IApplication[], selector: (app: IApplication) => string): string[] => {
  const values = new Set<string>();
  applications.forEach((app) => {
    const value = selector(app);
    if (value) {
      values.add(value);
    }
  });
  return Array.from(values).sort((a, b) => a.localeCompare(b));
};
