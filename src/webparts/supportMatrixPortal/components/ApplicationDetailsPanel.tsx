import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { IApplication } from '../models/IApplication';
import { StatusBadge, TierBadge } from './StatusBadge';
import { PersonPill, PersonPillList } from './PersonPill';
import styles from './ApplicationDetailsPanel.module.scss';

export interface IApplicationDetailsPanelProps {
  application?: IApplication;
  isOpen: boolean;
  onDismiss: () => void;
}

const FieldRow: React.FC<{ label: string; children: React.ReactNode }> = ({ label, children }) => (
  <div className={styles.fieldRow}>
    <span className={styles.fieldLabel}>{label}</span>
    <span className={styles.fieldValue}>{children}</span>
  </div>
);

export const ApplicationDetailsPanel: React.FC<IApplicationDetailsPanelProps> = ({ application, isOpen, onDismiss }) => {
  if (!application) {
    return null;
  }

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={application.applicationName}
      closeButtonAriaLabel="Close"
    >
      <div className={styles.headerBadges}>
        <StatusBadge status={application.status} />
        <TierBadge tier={application.tier} />
      </div>

      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Application Information</h3>
        <FieldRow label="Application Name">{application.applicationName}</FieldRow>
        <FieldRow label="Business Unit">{application.businessUnitName || 'Not specified'}</FieldRow>
        <FieldRow label="Vendor">{application.vendorName || 'Not specified'}</FieldRow>
        <FieldRow label="Tier">{application.tier ? `Tier ${application.tier}` : 'Not specified'}</FieldRow>
        <FieldRow label="Hosting Model">{application.hostingModel || 'Not specified'}</FieldRow>
        <FieldRow label="Application Type">{application.applicationType || 'Not specified'}</FieldRow>
        <FieldRow label="Status">{application.status || 'Not specified'}</FieldRow>
        <FieldRow label="Active">
          {application.active ? (
            <span>
              <Icon iconName="CheckMark" style={{ color: '#0e6b0e', marginRight: 6 }} />
              Yes
            </span>
          ) : (
            <span>
              <Icon iconName="Cancel" style={{ color: '#a4262c', marginRight: 6 }} />
              No
            </span>
          )}
        </FieldRow>
        <FieldRow label="High Level Details">
          <span className={styles.detailsText}>{application.highLevelDetails || 'No details provided.'}</span>
        </FieldRow>
      </div>

      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Ownership</h3>
        <FieldRow label="Business Owner">
          <PersonPill person={application.businessOwner} />
        </FieldRow>
        <FieldRow label="VP">
          <PersonPill person={application.vp} />
        </FieldRow>
        <FieldRow label="DM">
          <PersonPill person={application.dm} />
        </FieldRow>
      </div>

      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Support Team</h3>
        <FieldRow label="Primary Engineer">
          <PersonPill person={application.primaryEngineer} />
        </FieldRow>
        <FieldRow label="Secondary Engineer">
          <PersonPill person={application.secondaryEngineer} />
        </FieldRow>
        <FieldRow label="L3 SME">
          <PersonPill person={application.l3SME} />
        </FieldRow>
        <FieldRow label="L3 Support Engineers">
          <PersonPillList people={application.l3SupportEngineers} />
        </FieldRow>
        <FieldRow label="Cloud Operations">
          <PersonPillList people={application.cloudOperations} />
        </FieldRow>
        <FieldRow label="App Admin SME">
          <PersonPillList people={application.appAdminSME} />
        </FieldRow>
      </div>
    </Panel>
  );
};

export default ApplicationDetailsPanel;
