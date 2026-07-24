import * as React from 'react';
import styles from './StatusBadge.module.scss';

export interface IStatusBadgeProps {
  status: string;
}

export interface ITierBadgeProps {
  tier: string;
}

const STATUS_CLASS_MAP: { [key: string]: string } = {
  Active: 'statusActive',
  'Under Review': 'statusUnderReview',
  'To Be Transitioned': 'statusToBeTransitioned',
  Decommissioning: 'statusDecommissioning',
  Decommissioned: 'statusDecommissioned'
};

export const StatusBadge: React.FC<IStatusBadgeProps> = ({ status }) => {
  if (!status) {
    return <span className={`${styles.badge} ${styles.statusDefault}`}>Unknown</span>;
  }
  const mappedClass = STATUS_CLASS_MAP[status];
  const cssClass = mappedClass ? styles[mappedClass as keyof typeof styles] : styles.statusDefault;
  return <span className={`${styles.badge} ${cssClass}`}>{status}</span>;
};

const TIER_CLASS_MAP: { [key: string]: string } = {
  '1': 'tier1',
  '2': 'tier2',
  '3': 'tier3'
};

export const TierBadge: React.FC<ITierBadgeProps> = ({ tier }) => {
  if (!tier) {
    return <span className={`${styles.badge} ${styles.statusDefault}`}>-</span>;
  }
  const mappedClass = TIER_CLASS_MAP[tier];
  const cssClass = mappedClass ? styles[mappedClass as keyof typeof styles] : styles.statusDefault;
  return <span className={`${styles.badge} ${cssClass}`}>{`Tier ${tier}`}</span>;
};

export default StatusBadge;
