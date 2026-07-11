import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './DepartmentsDirectory.module.scss';
import { IDepartmentsDirectoryProps } from './IDepartmentsDirectoryProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { InitialsAvatar } from '../../../common/components/InitialsAvatar';
import { IDepartmentEntry } from '../../../common/services/SettingsService';
import { getCurrentLanguage, pickLocalized, Language } from '../../../common/services/languageService';
import { useSettings } from '../../../common/services/useSettings';
import { linkTarget } from '../../../common/util/format';

/** Parses the optional property-pane JSON override; returns [] when absent/invalid. */
function parseOverride(json: string): IDepartmentEntry[] {
  if (json && json.trim().length > 0) {
    try {
      const parsed: unknown = JSON.parse(json);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed as IDepartmentEntry[];
      }
    } catch {
      /* fall through to central settings */
    }
  }
  return [];
}

const DepartmentsDirectory: React.FunctionComponent<IDepartmentsDirectoryProps> = (props) => {
  // Reactive read: seeds from cache, then fetches + re-renders so the grid never
  // stays empty just because the cache-warming extension hadn't finished yet.
  const settings = useSettings(props.context);
  const language: Language = getCurrentLanguage();
  const override: IDepartmentEntry[] = parseOverride(props.departmentsJson);
  const departments: IDepartmentEntry[] = (override.length > 0 ? override : settings.departments || []).filter(
    (d) => !!d.label
  );

  return (
    <section className={styles.departments}>
      <SectionHeader
        title={props.title || 'Departments'}
        linkText="All Departments"
        linkUrl={props.viewAllUrl}
        showTitle={props.showTitle}
        showLink={props.showViewAll}
      />
      <div className={styles.grid}>
        {departments.map((dept, idx) => {
          const label: string = pickLocalized(dept.label, dept.labelAR, language);
          const description: string = pickLocalized(dept.description, dept.descriptionAR, language);
          return (
            <a className={styles.card} href={dept.url || '#'} key={`${dept.label}-${idx}`} {...linkTarget(settings.openLinksInNewTab)}>
              <div className={styles.icon} style={{ background: dept.accent || '#f5f3f0' }}>
                {dept.icon ? <Icon iconName={dept.icon} /> : <InitialsAvatar name={label} size={40} />}
              </div>
              <div className={styles.body}>
                <div className={styles.name}>{label}</div>
                {description ? <div className={styles.desc}>{description}</div> : null}
              </div>
            </a>
          );
        })}
      </div>
    </section>
  );
};

export default DepartmentsDirectory;
