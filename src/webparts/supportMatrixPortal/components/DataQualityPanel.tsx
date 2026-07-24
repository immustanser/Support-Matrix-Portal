import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IApplication } from '../models/IApplication';
import { computeDataQualityIssues } from '../utils/applicationHelpers';
import styles from './DataQualityPanel.module.scss';

export interface IDataQualityPanelProps {
  applications: IApplication[];
  onViewDetails: (application: IApplication) => void;
}

interface IIssueDefinition {
  key: string;
  label: string;
  items: IApplication[];
}

export const DataQualityPanel: React.FC<IDataQualityPanelProps> = ({ applications, onViewDetails }) => {
  const issues = React.useMemo(() => computeDataQualityIssues(applications), [applications]);
  const [expandedKey, setExpandedKey] = React.useState<string>('');

  const issueDefinitions: IIssueDefinition[] = [
    { key: 'missingPrimaryEngineer', label: 'Missing Primary Engineer', items: issues.missingPrimaryEngineer },
    { key: 'missingVendor', label: 'Missing Vendor', items: issues.missingVendor },
    { key: 'missingBusinessUnit', label: 'Missing Business Unit', items: issues.missingBusinessUnit },
    { key: 'missingTier', label: 'Missing Tier', items: issues.missingTier },
    { key: 'missingHostingModel', label: 'Missing Hosting Model', items: issues.missingHostingModel },
    { key: 'activeButDecommissioned', label: 'Active but Status is Decommissioned', items: issues.activeButDecommissioned },
    { key: 'inactiveApplications', label: 'Marked Inactive (Active = No)', items: issues.inactiveApplications }
  ];

  const toggleExpanded = (key: string): void => {
    setExpandedKey(expandedKey === key ? '' : key);
  };

  return (
    <div className={styles.container}>
      <h3 className={styles.title}>Data Quality</h3>
      <p className={styles.subtitle}>
        Automatically calculated from the current application data. Nothing here is stored back to SharePoint.
      </p>
      <div className={styles.issuesGrid}>
        {issueDefinitions.map((issue) => (
          <div className={styles.issueCard} key={issue.key}>
            <div className={styles.issueHeader} onClick={() => toggleExpanded(issue.key)}>
              <span className={styles.issueLabel}>{issue.label}</span>
              <span>
                <span className={`${styles.issueCount} ${issue.items.length === 0 ? styles.issueCountZero : ''}`}>
                  {issue.items.length}
                </span>
                <Icon
                  iconName={expandedKey === issue.key ? 'ChevronUp' : 'ChevronDown'}
                  style={{ marginLeft: 8, fontSize: 12 }}
                />
              </span>
            </div>
            {expandedKey === issue.key && (
              <>
                {issue.items.length === 0 ? (
                  <div className={styles.noIssues}>No applications with this issue.</div>
                ) : (
                  <div className={styles.issueList}>
                    {issue.items.map((app) => (
                      <div
                        className={styles.issueItem}
                        key={app.id}
                        onClick={() => onViewDetails(app)}
                        style={{ cursor: 'pointer' }}
                      >
                        {app.applicationName}
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  );
};

export default DataQualityPanel;
