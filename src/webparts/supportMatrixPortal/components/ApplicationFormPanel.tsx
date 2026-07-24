import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { IApplication, IApplicationFormData } from '../models/IApplication';
import { IVendor } from '../models/IVendor';
import { IBusinessUnit } from '../models/IBusinessUnit';
import { CHOICE_VALUES } from '../constants/ListConstants';
import { SupportMatrixService } from '../services/SupportMatrixService';
import { PersonFieldPicker } from './PersonFieldPicker';
import styles from './ApplicationFormPanel.module.scss';

export interface IApplicationFormPanelProps {
  isOpen: boolean;
  mode: 'create' | 'edit';
  application?: IApplication;
  businessUnits: IBusinessUnit[];
  vendors: IVendor[];
  onDismiss: () => void;
  onSaved: () => void;
}

const EMPTY_FORM: IApplicationFormData = {
  applicationName: '',
  businessUnitId: undefined,
  vendorId: undefined,
  tier: '',
  hostingModel: '',
  applicationType: '',
  status: '',
  highLevelDetails: '',
  active: true,
  primaryEngineer: undefined,
  secondaryEngineer: undefined,
  l3SME: undefined,
  dm: undefined,
  vp: undefined,
  businessOwner: undefined,
  l3SupportEngineers: [],
  cloudOperations: [],
  appAdminSME: []
};

const toOptions = (values: string[]): IDropdownOption[] => values.map((v) => ({ key: v, text: v }));

const applicationToFormData = (application: IApplication): IApplicationFormData => ({
  applicationName: application.applicationName,
  businessUnitId: application.businessUnitId,
  vendorId: application.vendorId,
  tier: application.tier,
  hostingModel: application.hostingModel,
  applicationType: application.applicationType,
  status: application.status,
  highLevelDetails: application.highLevelDetails,
  active: application.active,
  primaryEngineer: application.primaryEngineer,
  secondaryEngineer: application.secondaryEngineer,
  l3SME: application.l3SME,
  dm: application.dm,
  vp: application.vp,
  businessOwner: application.businessOwner,
  l3SupportEngineers: application.l3SupportEngineers,
  cloudOperations: application.cloudOperations,
  appAdminSME: application.appAdminSME
});

export const ApplicationFormPanel: React.FC<IApplicationFormPanelProps> = ({
  isOpen,
  mode,
  application,
  businessUnits,
  vendors,
  onDismiss,
  onSaved
}) => {
  const [formData, setFormData] = React.useState<IApplicationFormData>(EMPTY_FORM);
  const [isSaving, setIsSaving] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    if (isOpen) {
      setFormData(mode === 'edit' && application ? applicationToFormData(application) : EMPTY_FORM);
      setError(undefined);
    }
  }, [isOpen, mode, application]);

  const updateField = <K extends keyof IApplicationFormData>(key: K, value: IApplicationFormData[K]): void => {
    setFormData((prev) => ({ ...prev, [key]: value }));
  };

  const businessUnitOptions: IDropdownOption[] = businessUnits.map((bu) => ({ key: bu.id, text: bu.businessUnitName }));
  const vendorOptions: IDropdownOption[] = vendors.map((v) => ({ key: v.id, text: v.vendorName }));

  const isValid = formData.applicationName.trim().length > 0;

  const onSave = async (): Promise<void> => {
    if (!isValid) {
      setError('Application Name is required.');
      return;
    }
    setIsSaving(true);
    setError(undefined);
    try {
      if (mode === 'edit' && application) {
        await SupportMatrixService.updateApplication(application.id, formData);
      } else {
        await SupportMatrixService.createApplication(formData);
      }
      setIsSaving(false);
      onSaved();
    } catch (err) {
      setIsSaving(false);
      setError(err instanceof Error ? err.message : 'An unexpected error occurred while saving.');
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={mode === 'edit' ? 'Edit Application' : 'Add Application'}
      closeButtonAriaLabel="Close"
      isFooterAtBottom
      onRenderFooterContent={() => (
        <div className={styles.footer}>
          <PrimaryButton text={isSaving ? 'Saving...' : 'Save'} onClick={() => onSave()} disabled={isSaving} />
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={isSaving} />
          {isSaving && <Spinner size={SpinnerSize.small} />}
        </div>
      )}
    >
      {error && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.errorBar} onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}

      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Application Information</h3>
        <div className={styles.fieldGroup}>
          <TextField
            label="Application Name"
            required
            value={formData.applicationName}
            onChange={(_, value) => updateField('applicationName', value || '')}
          />
        </div>
        <div className={styles.fieldRow2}>
          <Dropdown
            label="Business Unit"
            options={businessUnitOptions}
            selectedKey={formData.businessUnitId}
            onChange={(_, option) => updateField('businessUnitId', option ? (option.key as number) : undefined)}
          />
          <Dropdown
            label="Vendor"
            options={vendorOptions}
            selectedKey={formData.vendorId}
            onChange={(_, option) => updateField('vendorId', option ? (option.key as number) : undefined)}
          />
        </div>
        <div className={styles.fieldRow2}>
          <Dropdown
            label="Tier"
            options={toOptions(CHOICE_VALUES.TIER)}
            selectedKey={formData.tier}
            onChange={(_, option) => updateField('tier', option ? String(option.key) : '')}
          />
          <Dropdown
            label="Hosting Model"
            options={toOptions(CHOICE_VALUES.HOSTING_MODEL)}
            selectedKey={formData.hostingModel}
            onChange={(_, option) => updateField('hostingModel', option ? String(option.key) : '')}
          />
        </div>
        <div className={styles.fieldRow2}>
          <Dropdown
            label="Application Type"
            options={toOptions(CHOICE_VALUES.APPLICATION_TYPE)}
            selectedKey={formData.applicationType}
            onChange={(_, option) => updateField('applicationType', option ? String(option.key) : '')}
          />
          <Dropdown
            label="Status"
            options={toOptions(CHOICE_VALUES.STATUS)}
            selectedKey={formData.status}
            onChange={(_, option) => updateField('status', option ? String(option.key) : '')}
          />
        </div>
        <div className={styles.fieldGroup}>
          <TextField
            label="High Level Details"
            multiline
            rows={4}
            value={formData.highLevelDetails}
            onChange={(_, value) => updateField('highLevelDetails', value || '')}
          />
        </div>
        <div className={styles.fieldGroup}>
          <Toggle
            label="Active"
            checked={formData.active}
            onChange={(_, checked) => updateField('active', !!checked)}
          />
        </div>
      </div>

      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Ownership</h3>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="Business Owner"
            itemLimit={1}
            selected={formData.businessOwner ? [formData.businessOwner] : []}
            onChange={(people) => updateField('businessOwner', people[0])}
          />
        </div>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="VP"
            itemLimit={1}
            selected={formData.vp ? [formData.vp] : []}
            onChange={(people) => updateField('vp', people[0])}
          />
        </div>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="DM"
            itemLimit={1}
            selected={formData.dm ? [formData.dm] : []}
            onChange={(people) => updateField('dm', people[0])}
          />
        </div>
      </div>

      <div className={styles.section}>
        <h3 className={styles.sectionTitle}>Support Team</h3>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="Primary Engineer"
            itemLimit={1}
            selected={formData.primaryEngineer ? [formData.primaryEngineer] : []}
            onChange={(people) => updateField('primaryEngineer', people[0])}
          />
        </div>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="Secondary Engineer"
            itemLimit={1}
            selected={formData.secondaryEngineer ? [formData.secondaryEngineer] : []}
            onChange={(people) => updateField('secondaryEngineer', people[0])}
          />
        </div>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="L3 SME"
            itemLimit={1}
            selected={formData.l3SME ? [formData.l3SME] : []}
            onChange={(people) => updateField('l3SME', people[0])}
          />
        </div>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="L3 Support Engineers"
            selected={formData.l3SupportEngineers}
            onChange={(people) => updateField('l3SupportEngineers', people)}
          />
        </div>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="Cloud Operations"
            selected={formData.cloudOperations}
            onChange={(people) => updateField('cloudOperations', people)}
          />
        </div>
        <div className={styles.fieldGroup}>
          <PersonFieldPicker
            label="App Admin SME"
            selected={formData.appAdminSME}
            onChange={(people) => updateField('appAdminSME', people)}
          />
        </div>
      </div>
    </Panel>
  );
};

export default ApplicationFormPanel;
