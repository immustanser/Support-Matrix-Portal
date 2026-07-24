import * as React from 'react';
import { DetailsList, IColumn, SelectionMode, IDetailsListStyles } from '@fluentui/react/lib/DetailsList';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { IApplication } from '../models/IApplication';
import { StatusBadge, TierBadge } from './StatusBadge';
import { PAGE_SIZE_OPTIONS } from '../constants/ListConstants';
import styles from './ApplicationsTable.module.scss';

export interface IApplicationsTableProps {
  applications: IApplication[];
  onViewDetails: (application: IApplication) => void;
  onEdit?: (application: IApplication) => void;
  onDelete?: (application: IApplication) => void;
  pageSize: number;
  onPageSizeChange: (size: number) => void;
}

type SortDirection = 'asc' | 'desc';

const detailsListStyles: Partial<IDetailsListStyles> = {
  root: {
    overflowX: 'visible'
  }
};

const getSortValue = (app: IApplication, key: string): string => {
  switch (key) {
    case 'applicationName':
      return app.applicationName || '';
    case 'businessUnitName':
      return app.businessUnitName || '';
    case 'vendorName':
      return app.vendorName || '';
    case 'tier':
      return app.tier || '';
    case 'status':
      return app.status || '';
    case 'hostingModel':
      return app.hostingModel || '';
    default:
      return '';
  }
};

export const ApplicationsTable: React.FC<IApplicationsTableProps> = ({
  applications,
  onViewDetails,
  onEdit,
  onDelete,
  pageSize,
  onPageSizeChange
}) => {
  const [sortKey, setSortKey] = React.useState<string>('applicationName');
  const [sortDirection, setSortDirection] = React.useState<SortDirection>('asc');
  const [currentPage, setCurrentPage] = React.useState<number>(1);

  React.useEffect(() => {
    setCurrentPage(1);
  }, [applications.length, pageSize]);

  const sortedApplications = React.useMemo(() => {
    const copy = [...applications];
    copy.sort((a, b) => {
      const valueA = getSortValue(a, sortKey);
      const valueB = getSortValue(b, sortKey);
      const comparison = valueA.localeCompare(valueB, undefined, { numeric: true, sensitivity: 'base' });
      return sortDirection === 'asc' ? comparison : -comparison;
    });
    return copy;
  }, [applications, sortKey, sortDirection]);

  const totalPages = Math.max(1, Math.ceil(sortedApplications.length / pageSize));
  const safePage = Math.min(currentPage, totalPages);

  const pagedApplications = React.useMemo(() => {
    const start = (safePage - 1) * pageSize;
    return sortedApplications.slice(start, start + pageSize);
  }, [sortedApplications, safePage, pageSize]);

  const onColumnClick = (key: string): void => {
    if (sortKey === key) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortKey(key);
      setSortDirection('asc');
    }
  };

  const renderSortableHeader = (label: string, key: string): JSX.Element => (
    <span>
      {label}
      {sortKey === key && (
        <Icon iconName={sortDirection === 'asc' ? 'SortUp' : 'SortDown'} style={{ marginLeft: 4, fontSize: 12 }} />
      )}
    </span>
  );

  const columns: IColumn[] = [
    {
      key: 'applicationName',
      name: 'Application Name',
      fieldName: 'applicationName',
      minWidth: 160,
      maxWidth: 220,
      isResizable: true,
      onColumnClick: () => onColumnClick('applicationName'),
      onRenderHeader: () => renderSortableHeader('Application Name', 'applicationName'),
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.applicationName}>
          {item.applicationName}
        </span>
      )
    },
    {
      key: 'businessUnitName',
      name: 'Business Unit',
      fieldName: 'businessUnitName',
      minWidth: 120,
      maxWidth: 160,
      isResizable: true,
      onColumnClick: () => onColumnClick('businessUnitName'),
      onRenderHeader: () => renderSortableHeader('Business Unit', 'businessUnitName'),
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.businessUnitName}>
          {item.businessUnitName || '-'}
        </span>
      )
    },
    {
      key: 'vendorName',
      name: 'Vendor',
      fieldName: 'vendorName',
      minWidth: 120,
      maxWidth: 160,
      isResizable: true,
      onColumnClick: () => onColumnClick('vendorName'),
      onRenderHeader: () => renderSortableHeader('Vendor', 'vendorName'),
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.vendorName}>
          {item.vendorName || '-'}
        </span>
      )
    },
    {
      key: 'tier',
      name: 'Tier',
      fieldName: 'tier',
      minWidth: 80,
      maxWidth: 90,
      isResizable: true,
      onColumnClick: () => onColumnClick('tier'),
      onRenderHeader: () => renderSortableHeader('Tier', 'tier'),
      onRender: (item: IApplication) => <TierBadge tier={item.tier} />
    },
    {
      key: 'hostingModel',
      name: 'Hosting Model',
      fieldName: 'hostingModel',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onColumnClick: () => onColumnClick('hostingModel'),
      onRenderHeader: () => renderSortableHeader('Hosting Model', 'hostingModel'),
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.hostingModel}>
          {item.hostingModel || '-'}
        </span>
      )
    },
    {
      key: 'applicationType',
      name: 'Application Type',
      fieldName: 'applicationType',
      minWidth: 130,
      maxWidth: 160,
      isResizable: true,
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.applicationType}>
          {item.applicationType || '-'}
        </span>
      )
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'status',
      minWidth: 120,
      maxWidth: 140,
      isResizable: true,
      onColumnClick: () => onColumnClick('status'),
      onRenderHeader: () => renderSortableHeader('Status', 'status'),
      onRender: (item: IApplication) => <StatusBadge status={item.status} />
    },
    {
      key: 'primaryEngineer',
      name: 'Primary Engineer',
      fieldName: 'primaryEngineer',
      minWidth: 130,
      maxWidth: 160,
      isResizable: true,
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.primaryEngineer?.title || ''}>
          {item.primaryEngineer?.title || 'Not assigned'}
        </span>
      )
    },
    {
      key: 'l3SME',
      name: 'L3 SME',
      fieldName: 'l3SME',
      minWidth: 110,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.l3SME?.title || ''}>
          {item.l3SME?.title || 'Not assigned'}
        </span>
      )
    },
    {
      key: 'vp',
      name: 'VP',
      fieldName: 'vp',
      minWidth: 110,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: IApplication) => (
        <span className={styles.ellipsisCell} title={item.vp?.title || ''}>
          {item.vp?.title || 'Not assigned'}
        </span>
      )
    },
    {
      key: 'active',
      name: 'Active',
      fieldName: 'active',
      minWidth: 70,
      maxWidth: 80,
      isResizable: true,
      onRender: (item: IApplication) => (item.active ? <Icon iconName="CheckMark" style={{ color: '#0e6b0e' }} /> : <Icon iconName="Cancel" style={{ color: '#a4262c' }} />)
    },
    {
      key: 'actions',
      name: 'Actions',
      fieldName: 'actions',
      minWidth: 130,
      maxWidth: 130,
      onRender: (item: IApplication) => (
        <span className={styles.actionsCell}>
          <IconButton
            iconProps={{ iconName: 'View' }}
            title="View Details"
            ariaLabel="View Details"
            onClick={() => onViewDetails(item)}
          />
          {onEdit && (
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              ariaLabel="Edit"
              onClick={() => onEdit(item)}
            />
          )}
          {onDelete && (
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete"
              ariaLabel="Delete"
              onClick={() => onDelete(item)}
            />
          )}
        </span>
      )
    }
  ];

  if (applications.length === 0) {
    return (
      <div className={styles.tableCard}>
        <div className={styles.emptyState}>No applications found. Try adjusting your search or filters.</div>
      </div>
    );
  }

  return (
    <div className={styles.tableCard}>
      <div className={styles.tableScroll}>
        <DetailsList
          items={pagedApplications}
          columns={columns}
          setKey="id"
          selectionMode={SelectionMode.none}
          styles={detailsListStyles}
          compact={false}
          getKey={(item: IApplication) => String(item.id)}
        />
      </div>
      <div className={styles.footerRow}>
        <div className={styles.pageSizeSelector}>
          <span>Rows per page:</span>
          <select
            value={pageSize}
            onChange={(e) => onPageSizeChange(Number(e.target.value))}
          >
            {PAGE_SIZE_OPTIONS.map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
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
            Page {safePage} of {totalPages} ({sortedApplications.length} items)
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
  );
};

export default ApplicationsTable;
