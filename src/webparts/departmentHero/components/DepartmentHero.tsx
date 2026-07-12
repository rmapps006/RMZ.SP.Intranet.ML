import * as React from 'react';
import styles from './DepartmentHero.module.scss';
import { IDepartmentHeroProps } from './IDepartmentHeroProps';
import { InitialsAvatar } from '../../../common/components/InitialsAvatar';
import { useDepartmentSettings } from '../../../common/services/useDepartmentSettings';
import { getCurrentLanguage, pickLocalized, isRtl, Language } from '../../../common/services/languageService';

const DepartmentHero: React.FunctionComponent<IDepartmentHeroProps> = (props) => {
  // Property-pane values win; otherwise fall back to the central Department Admin settings.
  const ds = useDepartmentSettings(props.context);
  const language: Language = getCurrentLanguage();
  const eyebrow: string = props.eyebrow || pickLocalized(ds.eyebrow, ds.eyebrowAR, language);
  const departmentName: string = props.departmentName || pickLocalized(ds.departmentName, ds.departmentNameAR, language);
  const description: string = props.description || pickLocalized(ds.description, ds.descriptionAR, language);
  const ownerName: string = props.ownerName || pickLocalized(ds.ownerName, ds.ownerNameAR, language);
  const ownerRole: string = props.ownerRole || pickLocalized(ds.ownerRole, ds.ownerRoleAR, language);
  // Settings resolve async (cache seed, then a live fetch); if both name and
  // description are still empty once that settles, show a minimal placeholder
  // instead of an otherwise-blank hero.
  const hasContent: boolean = !!(departmentName || description);

  return (
    <section className={styles.deptHero} dir={isRtl(language) ? 'rtl' : 'ltr'}>
      <div className={styles.deptHeroBg} />
      <div className={styles.deptHeroCt}>
        {eyebrow ? <div className={styles.deptEyebrow}>{eyebrow}</div> : null}
        {departmentName ? <h1 className={styles.deptName}>{departmentName}</h1> : hasContent ? null : <h1 className={styles.deptName}>Department</h1>}
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
