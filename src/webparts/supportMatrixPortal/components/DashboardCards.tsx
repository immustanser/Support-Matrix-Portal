import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IApplication } from '../models/IApplication';
import styles from './DashboardCards.module.scss';

export interface IDashboardCardsProps {
  applications: IApplication[];
  onKpiClick?: (applicationType: string) => void;
}

interface ICardDefinition {
  key: string;
  label: string;
  value: number;
  icon: string;
  colorClass: string;
  applicationType: string;
}

export const DashboardCards: React.FC<IDashboardCardsProps> = ({ applications, onKpiClick }) => {
  const metrics = React.useMemo(() => {
    const total = applications.length;
    const operation = applications.filter((a) => a.applicationType === 'Operation').length;
    const vendorApps = applications.filter((a) => a.applicationType === 'Vendor Application').length;
    const decommissioned = applications.filter((a) => a.applicationType === 'Decommissioned').length;
    const toBeTransitioned = applications.filter((a) => a.applicationType === 'To Be Transitioned').length;

    return { total, operation, vendorApps, decommissioned, toBeTransitioned };
  }, [applications]);

  const cards: ICardDefinition[] = [
    { key: 'all', label: 'All', value: metrics.total, icon: 'AppIconDefaultList', colorClass: styles.colorBlue, applicationType: '' },
    { key: 'operation', label: 'Operation', value: metrics.operation, icon: 'CheckMark', colorClass: styles.colorGreen, applicationType: 'Operation' },
    { key: 'vendor', label: 'Vendor Applications', value: metrics.vendorApps, icon: 'Handshake', colorClass: styles.colorPurple, applicationType: 'Vendor Application' },
    { key: 'decommissioned', label: 'Decommissioned', value: metrics.decommissioned, icon: 'Delete', colorClass: styles.colorRed, applicationType: 'Decommissioned' },
    { key: 'tobetransitioned', label: 'To be Transitioned', value: metrics.toBeTransitioned, icon: 'Warning', colorClass: styles.colorOrange, applicationType: 'To Be Transitioned' }
  ];

  const handleKpiClick = (card: ICardDefinition): void => {
    if (onKpiClick) {
      onKpiClick(card.applicationType);
    }
  };

  const handleKpiKeyDown = (card: ICardDefinition) => (event: React.KeyboardEvent<HTMLDivElement>): void => {
    if (event.key === 'Enter' || event.key === ' ' || event.key === 'Spacebar') {
      event.preventDefault();
      handleKpiClick(card);
    }
  };

  return (
    <div className={styles.cardsGrid}>
      {cards.map((card) => (
        <div
          className={`${styles.card} ${styles.cardClickable}`}
          key={card.key}
          role="button"
          tabIndex={0}
          aria-label={`View ${card.label} applications (${card.value})`}
          onClick={() => handleKpiClick(card)}
          onKeyDown={handleKpiKeyDown(card)}
        >
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
