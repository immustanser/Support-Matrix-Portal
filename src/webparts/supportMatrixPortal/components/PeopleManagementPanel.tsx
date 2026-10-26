import * as React from 'react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { IApplication } from '../models/IApplication';
import { IPersonField } from '../models/IPersonField';
import { IBulkRemovalResult } from '../models/IPeopleManagement';
import { PersonFieldPicker } from './PersonFieldPicker';
import { SupportMatrixService } from '../services/SupportMatrixService';
import { PERSON_ROLE_DEFINITIONS } from '../constants/PersonRoleDefinitions';
import { PAGE_SIZE_OPTIONS, DEFAULT_PAGE_SIZE } from '../constants/ListConstants';
import { findPersonAssignments, computeRoleSummary, getRoleLabels } from '../utils/peopleManagementHelpers';
import { getUniqueSortedValues } from '../utils/applicationHelpers';
import styles from './PeopleManagementPanel.module.scss';

export interface IPeopleManagementPanelProps {
  applications: IApplication[];
  onDataChanged: () => void;
}

export const PeopleManagementPanel: React.FC<IPeopleManagementPanelProps> = ({
  applications,
  onDataChanged
}) => {
  const [selectedPerson, setSelectedPerson] = React.useState<IPersonField | undefined>(undefined);
  const [searchText, setSearchText] = React.useState<string>('');
  const [roleFilter, setRoleFilter] = React.useState<string>('');
  const [buFilter, setBuFilter] = React.useState<string>('');
  const [selectedIds, setSelectedIds] = React.useState<Set<number>>(new Set());
  const [pageSize, setPageSize] = React.useState<number>(DEFAULT_PAGE_SIZE);
  const [currentPage, setCurrentPage] = React.useState<number>(1);

  const [isConfirmOpen, setIsConfirmOpen] = React.useState<boolean>(false);
  const [isRemoving, setIsRemoving] = React.useState<boolean>(false);
  const [removalResult, setRemovalResult] = React.useState<IBulkRemovalResult | undefined>(undefined);
  const [removalResultPersonName, setRemovalResultPersonName] = React.useState<string>('');
  const [removalError, setRemovalError] = React.useState<string | undefined>(undefined);

  const allMatches = React.useMemo(
    () => findPersonAssignments(applications, selectedPerson?.id),
    [applications, selectedPerson]
  );

  const roleSummary = React.useMemo(() => computeRoleSummary(allMatches), [allMatches]);

  const roleFilterOptions = React.useMemo((): IDropdownOption[] => {
    const options: IDropdownOption[] = [{ key: '', text: 'All Roles' }];
    PERSON_ROLE_DEFINITIONS.forEach((def) => {
      options.push({ key: def.key, text: def.label });
    });
    return options;
  }, []);

  const buFilterOptions = React.useMemo((): IDropdownOption[] => {
    const bUs = getUniqueSortedValues(
      allMatches.map((m) => m.application),
      (a) => a.businessUnitName
    );
    const options: IDropdownOption[] = [{ key: '', text: 'All Business Units' }];
    bUs.forEach((bu) => options.push({ key: bu, text: bu }));
    return options;
  }, [allMatches]);

  const filteredMatches = React.useMemo(() => {
    let result = allMatches;
    if (roleFilter) {
      result = result.filter((m) => m.roleKeys.indexOf(roleFilter) !== -1);
    }
    if (buFilter) {
      result = result.filter((m) => m.application.businessUnitName === buFilter);
    }
    if (searchText.trim()) {
      const lower = searchText.trim().toLowerCase();
      result = result.filter(
        (m) =>
          m.application.applicationName.toLowerCase().indexOf(lower) !== -1 ||
          m.application.businessUnitName.toLowerCase().indexOf(lower) !== -1
      );
    }
    return result;
  }, [allMatches, roleFilter, buFilter, searchText]);

  React.useEffect(() => {
    setCurrentPage(1);
  }, [filteredMatches, pageSize]);

  const totalPages = Math.max(1, Math.ceil(filteredMatches.length / pageSize));
  const safePage = Math.min(currentPage, totalPages);
  const pagedMatches = React.useMemo(() => {
    const start = (safePage - 1) * pageSize;
    return filteredMatches.slice(start, start + pageSize);
  }, [filteredMatches, safePage, pageSize]);

  const visibleIds = React.useMemo(
    () => new Set(filteredMatches.map((m) => m.application.id)),
    [filteredMatches]
  );
  const visibleSelectedCount = React.useMemo(() => {
    let count = 0;
    selectedIds.forEach((id) => {
      if (visibleIds.has(id)) {
        count++;
      }
    });
    return count;
  }, [selectedIds, visibleIds]);
  const allVisibleSelected = filteredMatches.length > 0 && visibleSelectedCount === filteredMatches.length;

  const handlePersonChange = (people: IPersonField[]): void => {
    setSelectedPerson(people.length > 0 ? people[0] : undefined);
    setSelectedIds(new Set());
    setSearchText('');
    setRoleFilter('');
    setBuFilter('');
    setRemovalResult(undefined);
    setRemovalResultPersonName('');
    setRemovalError(undefined);
  };

  const toggleSelectAll = (): void => {
    if (allVisibleSelected) {
      const next = new Set(selectedIds);
      filteredMatches.forEach((m) => next.delete(m.application.id));
      setSelectedIds(next);
    } else {
      const next = new Set(selectedIds);
      filteredMatches.forEach((m) => next.add(m.application.id));
      setSelectedIds(next);
    }
  };

  const toggleRow = (id: number): void => {
    const next = new Set(selectedIds);
    if (next.has(id)) {
      next.delete(id);
    } else {
      next.add(id);
    }
    setSelectedIds(next);
  };

  const clearSelection = (): void => {
    setSelectedIds(new Set());
  };

  const selectedTargets = React.useMemo(
    () => allMatches.filter((m) => selectedIds.has(m.application.id)),
    [allMatches, selectedIds]
  );
  const totalRoleAssignments = React.useMemo(
    () => selectedTargets.reduce((sum, t) => sum + t.roleKeys.length, 0),
    [selectedTargets]
  );

  const handleRemoveClick = (): void => {
    setRemovalResult(undefined);
    setRemovalResultPersonName('');
    setRemovalError(undefined);
    setIsConfirmOpen(true);
  };

  const handleConfirmRemoval = async (): Promise<void> => {
    if (!selectedPerson) {
      return;
    }
    const personName = selectedPerson.title;
    setIsRemoving(true);
    setRemovalError(undefined);
    try {
      const targets = selectedTargets.map((m) => ({
        application: m.application,
        roleKeys: m.roleKeys
      }));
      const result = await SupportMatrixService.removePersonFromApplications(
        selectedPerson.id,
        targets
      );
      // Capture the person's name at the time of this operation so the result message
      // remains accurate even though the selected person stays selected afterward.
      setRemovalResultPersonName(personName);
      setRemovalResult(result);
      setIsConfirmOpen(false);
      setIsRemoving(false);
      // Only the processed rows are cleared - the selected person remains selected so the
      // panel can refresh and display their remaining assignments in place.
      setSelectedIds(new Set());
      onDataChanged();
    } catch (error) {
      setIsRemoving(false);
      setRemovalError(error instanceof Error ? error.message : 'An unexpected error occurred.');
    }
  };

  const renderSummary = (): JSX.Element | null => {
    if (!selectedPerson) {
      return null;
    }
    return (
      <div className={styles.summaryCard}>
        <div className={styles.summaryHeader}>
          <div className={styles.summaryPerson}>
            <span className={styles.summaryLabel}>Selected Person:</span>
            <span className={styles.summaryName}>{selectedPerson.title}</span>
            {selectedPerson.email && (
              <span className={styles.summaryEmail}>{selectedPerson.email}</span>
            )}
          </div>
          <div className={styles.summaryTotal}>
            <span className={styles.summaryTotalValue}>{allMatches.length}</span>
            <span className={styles.summaryTotalLabel}>Applications Found</span>
          </div>
        </div>
        {roleSummary.length > 0 && (
          <div className={styles.roleSummaryGrid}>
            {roleSummary.map((entry) => (
              <div className={styles.roleSummaryItem} key={entry.role.key}>
                <span className={styles.roleSummaryCount}>{entry.count}</span>
                <span className={styles.roleSummaryLabel}>{entry.role.label}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  };

  const renderRemovalResult = (): JSX.Element | null => {
    if (!removalResult) {
      return null;
    }
    const totalProcessed = removalResult.items.length;
    // An application counts as successfully updated if SharePoint accepted the update (no
    // top-level error), even if one of its required single-value roles had to be blocked.
    const succeeded = removalResult.items.filter((i) => !i.error);
    const failed = removalResult.items.filter((i) => !!i.error);
    const blocked = removalResult.items.filter((i) => i.blockedRoleLabels.length > 0);
    const personName = removalResultPersonName || 'the selected person';

    let summaryMessageType = MessageBarType.success;
    let summaryMessage = '';

    if (succeeded.length === totalProcessed && totalProcessed > 0) {
      summaryMessageType = MessageBarType.success;
      summaryMessage = `${personName} was successfully removed from ${succeeded.length} application${succeeded.length !== 1 ? 's' : ''}.`;
    } else if (succeeded.length > 0 && failed.length > 0) {
      summaryMessageType = MessageBarType.warning;
      summaryMessage = `Removal completed with some errors. ${personName} was successfully removed from ${succeeded.length} application${succeeded.length !== 1 ? 's' : ''}, but ${failed.length} application${failed.length !== 1 ? 's' : ''} could not be updated.`;
    } else if (succeeded.length === 0 && totalProcessed > 0) {
      summaryMessageType = MessageBarType.error;
      summaryMessage = `Unable to remove ${personName} from the selected applications. No applications were updated.`;
    }

    return (
      <div className={styles.resultSection}>
        {summaryMessage && (
          <MessageBar messageBarType={summaryMessageType} onDismiss={() => setRemovalResult(undefined)}>
            {summaryMessage}
          </MessageBar>
        )}
        {blocked.length > 0 && (
          <MessageBar messageBarType={MessageBarType.warning}>
            {blocked.length} role assignment{blocked.length !== 1 ? 's' : ''} could not be removed because the
            field is required:{' '}
            {blocked
              .map((b) => `${b.applicationName} (${b.blockedRoleLabels.join(', ')})`)
              .join('; ')}
          </MessageBar>
        )}
        {failed.length > 0 && (
          <MessageBar messageBarType={MessageBarType.error}>
            {failed.length} application{failed.length !== 1 ? 's' : ''} failed to update:{' '}
            {failed.map((f) => `${f.applicationName}: ${f.error || 'Unknown error'}`).join('; ')}
          </MessageBar>
        )}
      </div>
    );
  };

  return (
    <div className={styles.container}>
      <h3 className={styles.title}>People Management</h3>
      <p className={styles.subtitle}>
        Manage personnel assignments across multiple applications.
      </p>

      <div className={styles.pickerSection}>
        <PersonFieldPicker
          label="Search for a person"
          itemLimit={1}
          selected={selectedPerson ? [selectedPerson] : []}
          onChange={handlePersonChange}
        />
      </div>

      {renderRemovalResult()}

      {selectedPerson && allMatches.length === 0 && (
        <MessageBar messageBarType={MessageBarType.info}>
          No application assignments found for {selectedPerson.title}.
        </MessageBar>
      )}

      {selectedPerson && allMatches.length > 0 && (
        <>
          {renderSummary()}

          <div className={styles.filtersCard}>
            <div className={styles.filtersRow}>
              <div className={styles.searchBox}>
                <SearchBox
                  placeholder="Search applications..."
                  value={searchText}
                  onChange={(_, value) => setSearchText(value || '')}
                />
              </div>
              <div className={styles.filterDropdown}>
                <Dropdown
                  placeholder="All Roles"
                  options={roleFilterOptions}
                  selectedKey={roleFilter}
                  onChange={(_, option) => setRoleFilter(option ? String(option.key) : '')}
                />
              </div>
              <div className={styles.filterDropdown}>
                <Dropdown
                  placeholder="All Business Units"
                  options={buFilterOptions}
                  selectedKey={buFilter}
                  onChange={(_, option) => setBuFilter(option ? String(option.key) : '')}
                />
              </div>
            </div>
          </div>

          <div className={styles.actionBar}>
            <div className={styles.selectionInfo}>
              {selectedIds.size > 0 && (
                <>
                  <span className={styles.selectedCount}>
                    {selectedIds.size} application{selectedIds.size !== 1 ? 's' : ''} selected
                  </span>
                  <DefaultButton
                    text="Clear Selection"
                    iconProps={{ iconName: 'Cancel' }}
                    onClick={clearSelection}
                    className={styles.clearButton}
                  />
                </>
              )}
            </div>
            <PrimaryButton
              text="Remove From Selected"
              iconProps={{ iconName: 'UserRemove' }}
              disabled={selectedIds.size === 0}
              onClick={handleRemoveClick}
            />
          </div>

          <div className={styles.tableCard}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th className={styles.checkCol}>
                    <Checkbox
                      checked={allVisibleSelected}
                      onChange={toggleSelectAll}
                      ariaLabel="Select all visible applications"
                    />
                  </th>
                  <th className={styles.nameCol}>Application Name</th>
                  <th className={styles.buCol}>Business Unit</th>
                  <th className={styles.rolesCol}>Assigned Role(s)</th>
                </tr>
              </thead>
              <tbody>
                {pagedMatches.map((match) => (
                  <tr key={match.application.id} className={selectedIds.has(match.application.id) ? styles.selectedRow : ''}>
                    <td className={styles.checkCol}>
                      <Checkbox
                        checked={selectedIds.has(match.application.id)}
                        onChange={() => toggleRow(match.application.id)}
                        ariaLabel={`Select ${match.application.applicationName}`}
                      />
                    </td>
                    <td className={styles.nameCol}>
                      <span className={styles.ellipsis} title={match.application.applicationName}>
                        {match.application.applicationName}
                      </span>
                    </td>
                    <td className={styles.buCol}>
                      <span className={styles.ellipsis} title={match.application.businessUnitName}>
                        {match.application.businessUnitName || '-'}
                      </span>
                    </td>
                    <td className={styles.rolesCol}>
                      <span className={styles.roleTags}>
                        {getRoleLabels(match.roleKeys).map((label) => (
                          <span className={styles.roleTag} key={label}>{label}</span>
                        ))}
                      </span>
                    </td>
                  </tr>
                ))}
                {filteredMatches.length === 0 && (
                  <tr>
                    <td colSpan={4} className={styles.emptyState}>
                      No applications match the current filters.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>

            <div className={styles.footerRow}>
              <div className={styles.pageSizeSelector}>
                <span>Rows per page:</span>
                <select
                  value={pageSize}
                  onChange={(e) => setPageSize(Number(e.target.value))}
                >
                  {PAGE_SIZE_OPTIONS.map((size) => (
                    <option key={size} value={size}>{size}</option>
                  ))}
                </select>
              </div>
              <div className={styles.pagination}>
                <IconButton
                  iconProps={{ iconName: 'DoubleChevronLeft' }}
                  disabled={safePage <= 1}
                  onClick={() => setCurrentPage(1)}
                  ariaLabel="First page"
                />
                <IconButton
                  iconProps={{ iconName: 'ChevronLeft' }}
                  disabled={safePage <= 1}
                  onClick={() => setCurrentPage(safePage - 1)}
                  ariaLabel="Previous page"
                />
                <span className={styles.pageInfo}>
                  Page {safePage} of {totalPages} ({filteredMatches.length} items)
                </span>
                <IconButton
                  iconProps={{ iconName: 'ChevronRight' }}
                  disabled={safePage >= totalPages}
                  onClick={() => setCurrentPage(safePage + 1)}
                  ariaLabel="Next page"
                />
                <IconButton
                  iconProps={{ iconName: 'DoubleChevronRight' }}
                  disabled={safePage >= totalPages}
                  onClick={() => setCurrentPage(totalPages)}
                  ariaLabel="Last page"
                />
              </div>
            </div>
          </div>
        </>
      )}

      <Dialog
        hidden={!isConfirmOpen}
        onDismiss={() => (!isRemoving ? setIsConfirmOpen(false) : undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: `Remove ${selectedPerson?.title || ''}?`
        }}
        minWidth={480}
      >
        <div className={styles.confirmBody}>
          <p>
            You are about to remove this person from{' '}
            <strong>{selectedTargets.length}</strong> selected application
            {selectedTargets.length !== 1 ? 's' : ''}.
          </p>
          <p>
            <strong>{totalRoleAssignments}</strong> role assignment
            {totalRoleAssignments !== 1 ? 's' : ''} will be affected.
          </p>
          <p className={styles.confirmNote}>
            The selected person&apos;s matching role assignments will be removed. Other
            personnel assignments will not be changed.
          </p>
          {removalError && (
            <MessageBar messageBarType={MessageBarType.error}>
              {removalError}
            </MessageBar>
          )}
        </div>
        <DialogFooter>
          <PrimaryButton
            text={isRemoving ? 'Removing...' : 'Confirm Removal'}
            onClick={() => handleConfirmRemoval()}
            disabled={isRemoving}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => setIsConfirmOpen(false)}
            disabled={isRemoving}
          />
          {isRemoving && <Spinner size={SpinnerSize.small} />}
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default PeopleManagementPanel;
