import * as React from 'react';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { DefaultButton } from '@fluentui/react/lib/Button';
import styles from './FiltersBar.module.scss';

export interface IApplicationFilters {
  businessUnit: string;
  vendor: string;
  status: string;
  tier: string;
  hostingModel: string;
  applicationType: string;
  active: string;
}

export const EMPTY_FILTERS: IApplicationFilters = {
  businessUnit: '',
  vendor: '',
  status: '',
  tier: '',
  hostingModel: '',
  applicationType: '',
  active: ''
};

export interface IFiltersBarProps {
  searchText: string;
  onSearchTextChange: (value: string) => void;
  filters: IApplicationFilters;
  onFiltersChange: (filters: IApplicationFilters) => void;
  businessUnitOptions: string[];
  vendorOptions: string[];
  statusOptions: string[];
  tierOptions: string[];
  hostingModelOptions: string[];
  applicationTypeOptions: string[];
  myApplicationsOnly: boolean;
  onMyApplicationsToggle: (checked: boolean) => void;
}

const toOptions = (values: string[]): IDropdownOption[] => [
  { key: '', text: 'All' },
  ...values.map((v) => ({ key: v, text: v }))
];

export const FiltersBar: React.FC<IFiltersBarProps> = (props) => {
  const {
    searchText,
    onSearchTextChange,
    filters,
    onFiltersChange,
    businessUnitOptions,
    vendorOptions,
    statusOptions,
    tierOptions,
    hostingModelOptions,
    applicationTypeOptions,
    myApplicationsOnly,
    onMyApplicationsToggle
  } = props;

  const updateFilter = (key: keyof IApplicationFilters) => (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    onFiltersChange({ ...filters, [key]: option ? String(option.key) : '' });
  };

  const hasActiveFilters =
    !!searchText ||
    myApplicationsOnly ||
    Object.keys(filters).some((k) => !!filters[k as keyof IApplicationFilters]);

  const clearAll = (): void => {
    onSearchTextChange('');
    onFiltersChange({ ...EMPTY_FILTERS });
    onMyApplicationsToggle(false);
  };

  return (
    <div className={styles.filtersBar}>
      <div className={styles.topRow}>
        <div className={styles.searchBox}>
          <SearchBox
            placeholder="Search applications, owners, vendors, business units..."
            value={searchText}
            onChange={(_e, newValue) => onSearchTextChange(newValue || '')}
            onClear={() => onSearchTextChange('')}
          />
        </div>
        <div className={styles.myAppsToggle}>
          <Toggle
            label="My Applications"
            inlineLabel
            checked={myApplicationsOnly}
            onChange={(_e, checked) => onMyApplicationsToggle(!!checked)}
          />
        </div>
      </div>
      <div className={styles.filtersRow}>
        <Dropdown
          label="Business Unit"
          selectedKey={filters.businessUnit}
          onChange={updateFilter('businessUnit')}
          options={toOptions(businessUnitOptions)}
        />
        <Dropdown
          label="Vendor"
          selectedKey={filters.vendor}
          onChange={updateFilter('vendor')}
          options={toOptions(vendorOptions)}
        />
        <Dropdown
          label="Status"
          selectedKey={filters.status}
          onChange={updateFilter('status')}
          options={toOptions(statusOptions)}
        />
        <Dropdown
          label="Tier"
          selectedKey={filters.tier}
          onChange={updateFilter('tier')}
          options={toOptions(tierOptions)}
        />
        <Dropdown
          label="Hosting Model"
          selectedKey={filters.hostingModel}
          onChange={updateFilter('hostingModel')}
          options={toOptions(hostingModelOptions)}
        />
        <Dropdown
          label="Application Type"
          selectedKey={filters.applicationType}
          onChange={updateFilter('applicationType')}
          options={toOptions(applicationTypeOptions)}
        />
        <Dropdown
          label="Active"
          selectedKey={filters.active}
          onChange={updateFilter('active')}
          options={[
            { key: '', text: 'All' },
            { key: 'true', text: 'Active' },
            { key: 'false', text: 'Inactive' }
          ]}
        />
      </div>
      {hasActiveFilters && (
        <div className={styles.clearButtonRow}>
          <DefaultButton text="Clear all filters" onClick={clearAll} />
        </div>
      )}
    </div>
  );
};

export default FiltersBar;
