import * as React from 'react';
import styles from './DepartmentHero.module.scss';
import { IDepartmentHeroProps } from './IDepartmentHeroProps';
import { InitialsAvatar } from '../../../common/components/InitialsAvatar';
import { useDepartmentSettings } from '../../../common/services/useDepartmentSettings';

const DepartmentHero: React.FunctionComponent<IDepartmentHeroProps> = (props) => {
  // Property-pane values win; otherwise fall back to the central Department Admin settings.
  const ds = useDepartmentSettings(props.context);
  const eyebrow: string = props.eyebrow || ds.eyebrow;
  const departmentName: string = props.departmentName || ds.departmentName;
  const description: string = props.description || ds.description;
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
