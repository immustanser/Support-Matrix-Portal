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
    const operation = applications.filter((a) => a.applicationType === 'Operation').length;
    const vendorApps = applications.filter((a) => a.applicationType === 'Vendor Application').length;
    const decommissioned = applications.filter((a) => a.applicationType === 'Decommissioned').length;
    const toBeTransitioned = applications.filter((a) => a.applicationType === 'To Be Transitioned').length;

    return { total, operation, vendorApps, decommissioned, toBeTransitioned };
  }, [applications]);

  const cards: ICardDefinition[] = [
    { key: 'all', label: 'All', value: metrics.total, icon: 'AppIconDefaultList', colorClass: styles.colorBlue },
    { key: 'operation', label: 'Operation', value: metrics.operation, icon: 'CheckMark', colorClass: styles.colorGreen },
    { key: 'vendor', label: 'Vendor Applications', value: metrics.vendorApps, icon: 'Handshake', colorClass: styles.colorPurple },
    { key: 'decommissioned', label: 'Decommissioned', value: metrics.decommissioned, icon: 'Delete', colorClass: styles.colorRed },
    { key: 'tobetransitioned', label: 'To be Transitioned', value: metrics.toBeTransitioned, icon: 'Warning', colorClass: styles.colorOrange }
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
