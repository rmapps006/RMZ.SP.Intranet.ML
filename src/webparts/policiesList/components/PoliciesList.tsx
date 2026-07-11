import * as React from 'react';
import styles from './PoliciesList.module.scss';
import { IPoliciesListProps } from './IPoliciesListProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { formatLongDate, linkTarget } from '../../../common/util/format';
import { IPolicy } from '../../../common/models';
import { getPolicies } from '../services/PoliciesService';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';

const DEPARTMENTS: string[] = ['All Departments', 'HR', 'Finance', 'IT', 'Operations', 'Marketing', 'Legal'];

// Department badge colours, copied from the design inline styles.
const BADGE_STYLES: { [dept: string]: React.CSSProperties } = {
  HR: { background: 'rgba(164,224,230,.2)', color: '#3a8a90' },
  Finance: { background: 'rgba(227,203,168,.35)', color: '#8a6838' },
  IT: { background: 'rgba(100,180,80,.14)', color: '#4e8838' },
  Operations: { background: 'rgba(234,139,110,.16)', color: '#c2664a' },
  Marketing: { background: 'rgba(181,181,181,.22)', color: '#6a6660' },
  Legal: { background: 'rgba(38,38,38,.1)', color: '#3a3632' }
};

function fileTypeClass(t: 'pdf' | 'doc' | 'xls'): string {
  if (t === 'doc') {
    return styles.doc;
  }
  if (t === 'xls') {
    return styles.xls;
  }
  return styles.pdf;
}

const SearchIcon: React.FunctionComponent = () => (
  <svg width="13" height="13" viewBox="0 0 14 14" fill="none" style={{ opacity: 0.38, flexShrink: 0 }} aria-hidden="true">
    <circle cx="6" cy="6" r="4.5" stroke="#8a8682" strokeWidth="1.4" />
    <path d="M9.5 9.5L12 12" stroke="#8a8682" strokeWidth="1.4" strokeLinecap="round" />
  </svg>
);

const PoliciesList: React.FunctionComponent<IPoliciesListProps> = (props) => {
  const [policies, setPolicies] = React.useState<IPolicy[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [search, setSearch] = React.useState<string>('');
  const [dept, setDept] = React.useState<string>('All Departments');
  const settings = useSettings(props.context);
  const navKey: string = useNavKey();
  const newTab: boolean = settings.openLinksInNewTab;
  const detailUrl: string = settings.detailPageUrl;

  // Prefer the shared Detail page (?type=policy) when configured, else the OOTB
  // display form.
  const hrefFor = (p: IPolicy): string | undefined => {
    if (detailUrl && p.id !== undefined) {
      return `${detailUrl}?type=policy&recId=${p.id}`;
    }
    return p.url || undefined;
  };

  React.useEffect(() => {
    let active: boolean = true;
    setLoading(true);
    getPolicies(props.context, props.policiesList)
      .then((items) => {
        if (active) {
          setPolicies(items);
          setLoading(false);
        }
      })
      .catch(() => {
        /* service already falls back internally */
        if (active) {
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, [props.context, props.policiesList, navKey]);

  const filtered: IPolicy[] = policies.filter((p) => {
    const matchesDept: boolean = dept === 'All Departments' || p.department === dept;
    const matchesSearch: boolean =
      search.trim().length === 0 || p.title.toLowerCase().indexOf(search.trim().toLowerCase()) !== -1;
    return matchesDept && matchesSearch;
  });

  return (
    <section className={styles.policies}>
      <SectionHeader title="Policies & Procedures" linkText="All Policies" linkUrl={props.allPoliciesUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      <div className={styles.top}>
        <div className={styles.search}>
          <SearchIcon />
          <input
            type="text"
            placeholder="Search policies…"
            aria-label="Search policies"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
        <div className={styles.chips}>
          {DEPARTMENTS.map((d) => (
            <div
              key={d}
              role="button"
              tabIndex={0}
              aria-pressed={d === dept}
              className={d === dept ? `${styles.chip} ${styles.on}` : styles.chip}
              onClick={() => setDept(d)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  setDept(d);
                }
              }}
            >
              {d}
            </div>
          ))}
        </div>
      </div>

      <div className={styles.table}>
        <div className={styles.hrow}>
          <div className={styles.hc}>Policy Title</div>
          <div className={styles.hc}>Department</div>
          <div className={styles.hc}>Last Updated</div>
          <div className={styles.hc}>Version</div>
        </div>
        {loading ? (
          <div className={styles.empty}>Loading policies…</div>
        ) : filtered.length === 0 ? (
          <div className={styles.empty}>No policies match your search.</div>
        ) : (
          filtered.map((p, i) => {
            const rowInner: JSX.Element = (
              <>
                <div className={styles.title}>
                  <div className={`${styles.dic} ${fileTypeClass(p.fileType)}`}>{p.fileType.toUpperCase()}</div>
                  <div className={styles.nm}>{p.title}</div>
                </div>
                <div className={styles.badge} style={BADGE_STYLES[p.department] || {}}>
                  {p.department}
                </div>
                <div className={styles.date}>{formatLongDate(p.updated)}</div>
                <div className={styles.ver}>{p.version}</div>
              </>
            );
            const key: string = `${p.title}-${i}`;
            const href: string | undefined = hrefFor(p);
            return href ? (
              <a className={`${styles.row} ${styles.rowLink}`} key={key} href={href} {...linkTarget(newTab)} style={{ textDecoration: 'none', color: 'inherit' }}>
                {rowInner}
              </a>
            ) : (
              <div className={styles.row} key={key}>
                {rowInner}
              </div>
            );
          })
        )}
      </div>
    </section>
  );
};

export default PoliciesList;
