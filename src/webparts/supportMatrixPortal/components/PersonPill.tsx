import * as React from 'react';
import { IPersonField } from '../models/IPersonField';
import styles from './PersonPill.module.scss';

export interface IPersonPillProps {
  person?: IPersonField;
}

export interface IPersonPillListProps {
  people: IPersonField[];
  emptyLabel?: string;
}

const getInitials = (name: string): string => {
  if (!name) {
    return '?';
  }
  const parts = name.trim().split(' ').filter(Boolean);
  if (parts.length === 1) {
    return parts[0].substring(0, 2).toUpperCase();
  }
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
};

export const PersonPill: React.FC<IPersonPillProps> = ({ person }) => {
  if (!person || !person.title) {
    return <span className={styles.empty}>Not assigned</span>;
  }
  return (
    <span className={styles.pill} title={person.email ? `${person.title} (${person.email})` : person.title}>
      <span className={styles.avatar}>{getInitials(person.title)}</span>
      <span className={styles.info}>
        <span className={styles.name}>{person.title}</span>
        {person.email && <span className={styles.email}>{person.email}</span>}
      </span>
    </span>
  );
};

export const PersonPillList: React.FC<IPersonPillListProps> = ({ people, emptyLabel }) => {
  if (!people || people.length === 0) {
    return <span className={styles.empty}>{emptyLabel || 'None assigned'}</span>;
  }
  return (
    <span className={styles.pillContainer}>
      {people.map((person) => (
        <PersonPill key={person.id} person={person} />
      ))}
    </span>
  );
};

export default PersonPill;
