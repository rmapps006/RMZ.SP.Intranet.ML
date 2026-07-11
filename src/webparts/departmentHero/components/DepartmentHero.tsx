import * as React from 'react';
import styles from './DepartmentHero.module.scss';
import { IDepartmentHeroProps } from './IDepartmentHeroProps';
import { InitialsAvatar } from '../../../common/components/InitialsAvatar';
import { useDepartmentSettings } from '../../../common/services/useDepartmentSettings';
import { getCurrentLanguage, pickLocalized, Language } from '../../../common/services/languageService';

const DepartmentHero: React.FunctionComponent<IDepartmentHeroProps> = (props) => {
  // Property-pane values win; otherwise fall back to the central Department Admin settings.
  const ds = useDepartmentSettings(props.context);
  const language: Language = getCurrentLanguage();
  const eyebrow: string = props.eyebrow || pickLocalized(ds.eyebrow, ds.eyebrowAR, language);
  const departmentName: string = props.departmentName || pickLocalized(ds.departmentName, ds.departmentNameAR, language);
  const description: string = props.description || pickLocalized(ds.description, ds.descriptionAR, language);
  const ownerName: string = props.ownerName || ds.ownerName;
  const ownerRole: string = props.ownerRole || ds.ownerRole;

  return (
    <section className={styles.deptHero}>
      <div className={styles.deptHeroBg} />
      <div className={styles.deptHeroCt}>
        {eyebrow ? <div className={styles.deptEyebrow}>{eyebrow}</div> : null}
        {departmentName ? <h1 className={styles.deptName}>{departmentName}</h1> : null}
        <div className={styles.deptRule} />
        {description ? <p className={styles.deptDesc}>{description}</p> : null}
        {ownerName ? (
          <div className={styles.deptOwnerRow}>
            <InitialsAvatar name={ownerName} size={36} />
            <div>
              <div className={styles.deptOwnerNm}>{ownerName}</div>
              {ownerRole ? <div className={styles.deptOwnerRole}>{ownerRole}</div> : null}
            </div>
          </div>
        ) : null}
      </div>
    </section>
  );
};

export default DepartmentHero;
