import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Icon } from '@fluentui/react/lib/Icon';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';

import { IApplication } from '../models/IApplication';
import { IVendor } from '../models/IVendor';
import { IBusinessUnit } from '../models/IBusinessUnit';
import { SupportMatrixService } from '../services/SupportMatrixService';
import { DashboardCards } from './DashboardCards';
import { ChartsSection } from './ChartsSection';
import { FiltersBar, IApplicationFilters, EMPTY_FILTERS } from './FiltersBar';
import { ApplicationsTable } from './ApplicationsTable';
import { ApplicationDetailsPanel } from './ApplicationDetailsPanel';
import { ApplicationFormPanel } from './ApplicationFormPanel';
import { DataQualityPanel } from './DataQualityPanel';
import { buildSearchIndex, getUniqueSortedValues, isUserOnApplication } from '../utils/applicationHelpers';
import { CHOICE_VALUES, DEFAULT_PAGE_SIZE } from '../constants/ListConstants';
import styles from './SupportMatrixPortal.module.scss';

export interface ISupportMatrixPortalProps {
  context: WebPartContext;
}

type LoadState = 'loading' | 'loaded' | 'error';

export const SupportMatrixPortal: React.FC<ISupportMatrixPortalProps> = ({ context }) => {
  const [loadState, setLoadState] = React.useState<LoadState>('loading');
  const [errorMessage, setErrorMessage] = React.useState<string>('');
  const [applications, setApplications] = React.useState<IApplication[]>([]);
  const [businessUnits, setBusinessUnits] = React.useState<IBusinessUnit[]>([]);
  const [vendors, setVendors] = React.useState<IVendor[]>([]);

  const [searchText, setSearchText] = React.useState<string>('');
  const [filters, setFilters] = React.useState<IApplicationFilters>({ ...EMPTY_FILTERS });
  const [myApplicationsOnly, setMyApplicationsOnly] = React.useState<boolean>(false);
  const [pageSize, setPageSize] = React.useState<number>(DEFAULT_PAGE_SIZE);

  const [selectedApplication, setSelectedApplication] = React.useState<IApplication | undefined>(undefined);
  const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);

  const [isFormOpen, setIsFormOpen] = React.useState<boolean>(false);
  const [formMode, setFormMode] = React.useState<'create' | 'edit'>('create');
  const [editingApplication, setEditingApplication] = React.useState<IApplication | undefined>(undefined);

  const [applicationToDelete, setApplicationToDelete] = React.useState<IApplication | undefined>(undefined);
  const [isDeleting, setIsDeleting] = React.useState<boolean>(false);
  const [deleteError, setDeleteError] = React.useState<string | undefined>(undefined);

  const currentUserEmail = context.pageContext.user.email || '';

  const loadData = React.useCallback(async (): Promise<void> => {
    setLoadState('loading');
    setErrorMessage('');
    try {
      const [apps, units, vendorList] = await Promise.all([
        SupportMatrixService.getApplications(),
        SupportMatrixService.getBusinessUnits(),
        SupportMatrixService.getVendors()
      ]);
      setApplications(apps);
      setBusinessUnits(units);
      setVendors(vendorList);
      setLoadState('loaded');
    } catch (error) {
      console.error('Failed to load Support Matrix Portal data:', error);
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Unable to load data. Please check list names or field internal names.'
      );
      setLoadState('error');
    }
  }, []);

  React.useEffect(() => {
    loadData().catch((err) => console.error(err));
  }, [loadData]);

  const businessUnitOptions = React.useMemo(() => getUniqueSortedValues(applications, (a) => a.businessUnitName), [applications]);
  const vendorOptions = React.useMemo(() => getUniqueSortedValues(applications, (a) => a.vendorName), [applications]);
  const statusOptions = CHOICE_VALUES.STATUS;
  const tierOptions = CHOICE_VALUES.TIER;
  const hostingModelOptions = CHOICE_VALUES.HOSTING_MODEL;
  const applicationTypeOptions = CHOICE_VALUES.APPLICATION_TYPE;

  const filteredApplications = React.useMemo(() => {
    const search = searchText.trim().toLowerCase();

    return applications.filter((app) => {
      if (myApplicationsOnly && !isUserOnApplication(app, currentUserEmail)) {
        return false;
      }
      if (filters.businessUnit && app.businessUnitName !== filters.businessUnit) {
        return false;
      }
      if (filters.vendor && app.vendorName !== filters.vendor) {
        return false;
      }
      if (filters.status && app.status !== filters.status) {
        return false;
      }
      if (filters.tier && app.tier !== filters.tier) {
        return false;
      }
      if (filters.hostingModel && app.hostingModel !== filters.hostingModel) {
        return false;
      }
      if (filters.applicationType && app.applicationType !== filters.applicationType) {
        return false;
      }
      if (filters.active) {
        const wantActive = filters.active === 'true';
        if (app.active !== wantActive) {
          return false;
        }
      }
      if (search && buildSearchIndex(app).indexOf(search) === -1) {
        return false;
      }
      return true;
    });
  }, [applications, searchText, filters, myApplicationsOnly, currentUserEmail]);

  const handleViewDetails = (application: IApplication): void => {
    setSelectedApplication(application);
    setIsPanelOpen(true);
  };

  const handleAddNew = (): void => {
    setFormMode('create');
    setEditingApplication(undefined);
    setIsFormOpen(true);
  };

  const handleEdit = (application: IApplication): void => {
    setFormMode('edit');
    setEditingApplication(application);
    setIsFormOpen(true);
  };

  const handleFormSaved = (): void => {
    setIsFormOpen(false);
    loadData().catch((err) => console.error(err));
  };

  const handleDeleteRequest = (application: IApplication): void => {
    setDeleteError(undefined);
    setApplicationToDelete(application);
  };

  const handleConfirmDelete = async (): Promise<void> => {
    if (!applicationToDelete) {
      return;
    }
    setIsDeleting(true);
    setDeleteError(undefined);
    try {
      await SupportMatrixService.deleteApplication(applicationToDelete.id);
      setApplicationToDelete(undefined);
      setIsDeleting(false);
      loadData().catch((err) => console.error(err));
    } catch (error) {
      setIsDeleting(false);
      setDeleteError(error instanceof Error ? error.message : 'Unable to delete the application.');
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.headerTextGroup}>
          <h1 className={styles.title}>Support Matrix Portal</h1>
          <p className={styles.subtitle}>Centralized application support ownership and inventory dashboard</p>
        </div>
      </div>

      {loadState === 'loading' && (
        <div className={styles.centeredState}>
          <Spinner size={SpinnerSize.large} />
          <span>Loading applications...</span>
        </div>
      )}

      {loadState === 'error' && (
        <div className={styles.errorState}>
          <p className={styles.errorTitle}>Unable to load data</p>
          <p>{errorMessage}</p>
          <p>Please check list names or field internal names in ListConstants.ts.</p>
          <PrimaryButton text="Retry" onClick={() => loadData()} iconProps={{ iconName: 'Refresh' }} />
        </div>
      )}

      {loadState === 'loaded' && (
        <div className={styles.pivotWrap}>
          <Pivot aria-label="Support Matrix Portal sections">
            <PivotItem headerText="Dashboard" itemIcon="ViewDashboard">
              <div className={styles.tabContent}>
                <DashboardCards applications={applications} />
                <ChartsSection applications={applications} />
              </div>
            </PivotItem>
            <PivotItem headerText="Application Inventory" itemIcon="AppIconDefaultList">
              <div className={styles.tabContent}>
                <FiltersBar
                  searchText={searchText}
                  onSearchTextChange={setSearchText}
                  filters={filters}
                  onFiltersChange={setFilters}
                  businessUnitOptions={businessUnitOptions}
                  vendorOptions={vendorOptions}
                  statusOptions={statusOptions}
                  tierOptions={tierOptions}
                  hostingModelOptions={hostingModelOptions}
                  applicationTypeOptions={applicationTypeOptions}
                  myApplicationsOnly={myApplicationsOnly}
                  onMyApplicationsToggle={setMyApplicationsOnly}
                />
                <div className={styles.resultsSummaryRow}>
                  <p className={styles.resultsSummary}>
                    Showing {filteredApplications.length} of {applications.length} applications
                  </p>
                  <PrimaryButton text="Add Application" iconProps={{ iconName: 'Add' }} onClick={handleAddNew} />
                </div>
                <ApplicationsTable
                  applications={filteredApplications}
                  onViewDetails={handleViewDetails}
                  onEdit={handleEdit}
                  onDelete={handleDeleteRequest}
                  pageSize={pageSize}
                  onPageSizeChange={setPageSize}
                />
              </div>
            </PivotItem>
            <PivotItem headerText="Data Quality" itemIcon="Warning">
              <div className={styles.tabContent}>
                <DataQualityPanel applications={applications} onViewDetails={handleViewDetails} />
              </div>
            </PivotItem>
          </Pivot>
        </div>
      )}

      {loadState === 'loaded' && applications.length === 0 && (
        <div className={styles.centeredState}>
          <Icon iconName="Info" style={{ fontSize: 28 }} />
          <span>No applications found.</span>
        </div>
      )}

      <ApplicationDetailsPanel
        application={selectedApplication}
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
      />

      <ApplicationFormPanel
        isOpen={isFormOpen}
        mode={formMode}
        application={editingApplication}
        businessUnits={businessUnits}
        vendors={vendors}
        onDismiss={() => setIsFormOpen(false)}
        onSaved={handleFormSaved}
      />

      <Dialog
        hidden={!applicationToDelete}
        onDismiss={() => (!isDeleting ? setApplicationToDelete(undefined) : undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Application',
          subText: applicationToDelete
            ? `Are you sure you want to delete "${applicationToDelete.applicationName}"? This action cannot be undone.`
            : ''
        }}
      >
        {deleteError && <p className={styles.errorTitle}>{deleteError}</p>}
        <DialogFooter>
          <PrimaryButton text={isDeleting ? 'Deleting...' : 'Delete'} onClick={() => handleConfirmDelete()} disabled={isDeleting} />
          <DefaultButton text="Cancel" onClick={() => setApplicationToDelete(undefined)} disabled={isDeleting} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default SupportMatrixPortal;
