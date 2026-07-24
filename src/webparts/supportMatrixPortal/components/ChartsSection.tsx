import * as React from 'react';
import { IApplication } from '../models/IApplication';
import styles from './ChartsSection.module.scss';

export interface IChartsSectionProps {
  applications: IApplication[];
}

interface IBarDatum {
  label: string;
  value: number;
}

const BAR_COLORS = ['#0078d4', '#8764b8', '#00767e', '#d13438', '#8a6d00', '#6b2fa0', '#0e6b0e', '#605e5c'];

const hexToRgba = (hex: string, alpha: number): string => {
  const parsed = hex.replace('#', '');
  const r = parseInt(parsed.substring(0, 2), 16);
  const g = parseInt(parsed.substring(2, 4), 16);
  const b = parseInt(parsed.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
};

const buildDistribution = (applications: IApplication[], selector: (app: IApplication) => string): IBarDatum[] => {
  const counts = new Map<string, number>();
  applications.forEach((app) => {
    const key = selector(app) || 'Unspecified';
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return Array.from(counts.entries())
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value);
};

const MiniBarChart: React.FC<{ title: string; data: IBarDatum[] }> = ({ title, data }) => {
  const max = data.reduce((m, d) => Math.max(m, d.value), 0);

  return (
    <div className={styles.chartCard}>
      <p className={styles.chartTitle}>{title}</p>
      {data.length === 0 && <div className={styles.emptyState}>No data available</div>}
      {data.map((datum, index) => {
        const color = BAR_COLORS[index % BAR_COLORS.length];
        return (
          <div className={styles.barRow} key={datum.label}>
            <span className={styles.barLabel} title={datum.label}>
              {datum.label}
            </span>
            <span className={styles.barTrack}>
              <span
                className={styles.barFill}
                style={{
                  width: max > 0 ? `${(datum.value / max) * 100}%` : '0%',
                  background: `linear-gradient(90deg, ${hexToRgba(color, 0.85)}, ${color})`
                }}
              />
            </span>
            <span
              className={styles.barValue}
              style={{ color, backgroundColor: hexToRgba(color, 0.12) }}
            >
              {datum.value}
            </span>
          </div>
        );
      })}
    </div>
  );
};

export const ChartsSection: React.FC<IChartsSectionProps> = ({ applications }) => {
  const byBusinessUnit = React.useMemo(() => buildDistribution(applications, (a) => a.businessUnitName), [applications]);
  const byVendor = React.useMemo(() => buildDistribution(applications, (a) => a.vendorName), [applications]);
  const byHostingModel = React.useMemo(() => buildDistribution(applications, (a) => a.hostingModel), [applications]);
  const byTier = React.useMemo(
    () => buildDistribution(applications, (a) => (a.tier ? `Tier ${a.tier}` : '')),
    [applications]
  );

  return (
    <div className={styles.chartsGrid}>
      <MiniBarChart title="Applications by Business Unit" data={byBusinessUnit.slice(0, 10)} />
      <MiniBarChart title="Applications by Vendor" data={byVendor.slice(0, 10)} />
      <MiniBarChart title="Applications by Hosting Model" data={byHostingModel} />
      <MiniBarChart title="Applications by Tier" data={byTier} />
    </div>
  );
};

export default ChartsSection;