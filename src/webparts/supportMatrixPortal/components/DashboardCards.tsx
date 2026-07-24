import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IApplication } from '../models/IApplication';
import styles from './DashboardCards.module.scss';

export interface IDashboardCardsProps {
  applications: IApplication[];
}

interface ICardDefinition {
  key: string;
  label: string;
  value: number;
  icon: string;
  colorClass: string;
}

export const DashboardCards: React.FC<IDashboardCardsProps> = ({ applications }) => {
  const metrics = React.useMemo(() => {
    const total = applications.length;
    const active = applications.filter((a) => a.active).length;
    const decommissioned = applications.filter((a) => a.status === 'Decommissioned').length;
    const vendorApps = applications.filter((a) => a.applicationType === 'Vendor Application').length;
    const cloudApps = applications.filter((a) => a.hostingModel === 'Cloud').length;
    const onPremApps = applications.filter((a) => a.hostingModel === 'On-Prem').length;
    const tier1Apps = applications.filter((a) => a.tier === '1').length;
    const missingPrimaryEngineer = applications.filter((a) => !a.primaryEngineer || !a.primaryEngineer.title).length;

    return { total, active, decommissioned, vendorApps, cloudApps, onPremApps, tier1Apps, missingPrimaryEngineer };
  }, [applications]);

  const cards: ICardDefinition[] = [
    { key: 'total', label: 'Total Applications', value: metrics.total, icon: 'AppIconDefaultList', colorClass: styles.colorBlue },
    { key: 'active', label: 'Active Applications', value: metrics.active, icon: 'CheckMark', colorClass: styles.colorGreen },
    { key: 'decommissioned', label: 'Decommissioned Applications', value: metrics.decommissioned, icon: 'Delete', colorClass: styles.colorRed },
    { key: 'vendor', label: 'Vendor Applications', value: metrics.vendorApps, icon: 'Handshake', colorClass: styles.colorPurple },
    { key: 'cloud', label: 'Cloud Applications', value: metrics.cloudApps, icon: 'Cloud', colorClass: styles.colorTeal },
    { key: 'onprem', label: 'On-Prem Applications', value: metrics.onPremApps, icon: 'HomeGroup', colorClass: styles.colorGray },
    { key: 'tier1', label: 'Tier 1 Applications', value: metrics.tier1Apps, icon: 'Warning', colorClass: styles.colorOrange },
    { key: 'missingpe', label: 'Missing Primary Engineer', value: metrics.missingPrimaryEngineer, icon: 'ContactCard', colorClass: styles.colorRed }
  ];

  return (
    <div className={styles.cardsGrid}>
      {cards.map((card) => (
        <div className={styles.card} key={card.key}>
          <div className={`${styles.iconWrap} ${card.colorClass}`}>
            <Icon iconName={card.icon} />
          </div>
          <div className={styles.cardBody}>
            <span className={styles.cardValue}>{card.value}</span>
            <span className={styles.cardLabel}>{card.label}</span>
          </div>
        </div>
      ))}
    </div>
  );
};

export default DashboardCards;
