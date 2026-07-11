import * as React from 'react';
import styles from './DetailView.module.scss';
import { IDetailViewProps } from './IDetailViewProps';
import { InitialsAvatar } from '../../../common/components/InitialsAvatar';
import { formatLongDate, browserViewUrl, formatClock } from '../../../common/util/format';
import {
  readDetailParams,
  getNewsDetail,
  getEventDetail,
  getPolicyDetail,
  getEmployeeDetail,
  getBenefitDetail,
  IDetailParams,
  INewsDetail,
  IEventDetail,
  IPolicyDetail,
  IEmployeeDetail,
  IBenefitDetail
} from '../services/DetailService';
import { useNavKey } from '../../../common/services/useNavKey';
import { getCurrentLanguage, isRtl } from '../../../common/services/languageService';

type Loaded =
  | { kind: 'news'; data: INewsDetail }
  | { kind: 'event'; data: IEventDetail }
  | { kind: 'policy'; data: IPolicyDetail }
  | { kind: 'employee'; data: IEmployeeDetail }
  | { kind: 'benefit'; data: IBenefitDetail };

const DetailView: React.FunctionComponent<IDetailViewProps> = (props) => {
  const [loading, setLoading] = React.useState<boolean>(true);
  const [result, setResult] = React.useState<Loaded | undefined>(undefined);

  // Re-read the URL params on every client-side navigation (the URL changes
  // without a remount when moving between detail links).
  const navKey: string = useNavKey();
  const params: IDetailParams = React.useMemo(() => readDetailParams(), [navKey]);

  // The page's static dir flips for Arabic, so the hardcoded arrow glyphs need to
  // flip with it rather than always pointing the LTR way.
  const rtl: boolean = isRtl(getCurrentLanguage());
  const backArrow: string = rtl ? '→' : '←';
  const linkArrow: string = rtl ? '←' : '→';

  React.useEffect(() => {
    let active: boolean = true;
    setLoading(true);
    const idNum: number = parseInt(params.id, 10);

    const load = async (): Promise<Loaded | undefined> => {
      if (params.type === 'news' && !isNaN(idNum)) {
        const d = await getNewsDetail(props.context, props.newsList, idNum);
        return d ? { kind: 'news', data: d } : undefined;
      }
      if (params.type === 'event' && !isNaN(idNum)) {
        const d = await getEventDetail(props.context, props.eventsList, idNum);
        return d ? { kind: 'event', data: d } : undefined;
      }
      if (params.type === 'policy' && !isNaN(idNum)) {
        const d = await getPolicyDetail(props.context, props.policiesList, idNum);
        return d ? { kind: 'policy', data: d } : undefined;
      }
      if (params.type === 'benefit' && !isNaN(idNum)) {
        const d = await getBenefitDetail(props.context, props.benefitsList, idNum);
        return d ? { kind: 'benefit', data: d } : undefined;
      }
      if (params.type === 'employee' && params.id) {
        const d = await getEmployeeDetail(props.context, params.id);
        return d ? { kind: 'employee', data: d } : undefined;
      }
      return undefined;
    };

    load()
      .then((r) => {
        if (active) {
          setResult(r);
          setLoading(false);
        }
      })
      .catch(() => {
        if (active) {
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, [props.context, props.newsList, props.eventsList, props.policiesList, props.benefitsList, params]);

  const back: React.ReactNode = props.backUrl ? (
    <a className={styles.back} href={props.backUrl}>
      {backArrow} Back
    </a>
  ) : null;

  if (loading) {
    return (
      <section className={styles.detail}>
        <div className={styles.muted}>Loading…</div>
      </section>
    );
  }

  if (!params.type) {
    return (
      <section className={styles.detail}>
        <div className={styles.muted}>No item specified. Open this page from a News, Event, Policy, Benefit or Directory link.</div>
      </section>
    );
  }

  if (!result) {
    return (
      <section className={styles.detail}>
        {back}
        <div className={styles.muted}>This item could not be found.</div>
      </section>
    );
  }

  if (result.kind === 'news') {
    const n: INewsDetail = result.data;
    return (
      <section className={styles.detail}>
        {back}
        {n.imageUrl ? <div className={styles.hero} style={{ backgroundImage: `url("${n.imageUrl}")` }} /> : null}
        {n.category ? <div className={styles.eyebrow}>{n.category}</div> : null}
        <h1 className={styles.title}>{n.title}</h1>
        <div className={styles.meta}>
          {[formatLongDate(n.date), n.source].filter((v) => !!v).join(' · ')}
        </div>
        {n.body ? <div className={styles.body} dir="auto" dangerouslySetInnerHTML={{ __html: n.body }} /> : null}
        {n.linkUrl ? (
          <a className={styles.cta} href={n.linkUrl} target="_blank" rel="noopener noreferrer">
            Read the full article {linkArrow}
          </a>
        ) : null}
      </section>
    );
  }

  if (result.kind === 'event') {
    const e: IEventDetail = result.data;
    const time: string = [formatClock(e.start), formatClock(e.end)].filter((v) => !!v).join(' – ');
    return (
      <section className={styles.detail}>
        {back}
        {e.category ? <div className={styles.eyebrow}>{e.category}</div> : null}
        <h1 className={styles.title}>{e.title}</h1>
        <div className={styles.factGrid}>
          {e.start ? <div className={styles.fact}><span>Date</span>{formatLongDate(e.start)}</div> : null}
          {time ? <div className={styles.fact}><span>Time</span>{time}</div> : null}
          {e.location ? <div className={styles.fact}><span>Location</span>{e.location}</div> : null}
        </div>
        {e.description ? <div className={styles.body} dir="auto" dangerouslySetInnerHTML={{ __html: e.description }} /> : null}
      </section>
    );
  }

  if (result.kind === 'policy') {
    const p: IPolicyDetail = result.data;
    const viewUrl: string | undefined = browserViewUrl(p.documentUrl);
    return (
      <section className={styles.detail}>
        {back}
        <div className={styles.eyebrow}>Policy</div>
        <h1 className={styles.title}>{p.title}</h1>
        <div className={styles.factGrid}>
          {p.department ? <div className={styles.fact}><span>Department</span>{p.department}</div> : null}
          {p.version ? <div className={styles.fact}><span>Version</span>{p.version}</div> : null}
          {p.fileType ? <div className={styles.fact}><span>Type</span>{p.fileType.toUpperCase()}</div> : null}
          {p.modified ? <div className={styles.fact}><span>Updated</span>{formatLongDate(p.modified)}</div> : null}
        </div>
        {viewUrl ? (
          <a className={styles.cta} href={viewUrl} target="_blank" rel="noopener noreferrer" data-interception="off">
            Open document {linkArrow}
          </a>
        ) : null}
      </section>
    );
  }

  if (result.kind === 'benefit') {
    const b: IBenefitDetail = result.data;
    return (
      <section className={styles.detail}>
        {back}
        <div className={styles.eyebrow}>{b.category || 'Benefit'}</div>
        <h1 className={styles.title}>{b.title}</h1>
        {b.summary ? <div className={styles.meta}>{b.summary}</div> : null}
        <div className={styles.factGrid}>
          {b.eligibility ? <div className={styles.fact}><span>Eligibility</span>{b.eligibility}</div> : null}
          {b.coverage ? <div className={styles.fact}><span>Coverage</span>{b.coverage}</div> : null}
        </div>
        {b.details ? <div className={styles.body} dir="auto" dangerouslySetInnerHTML={{ __html: b.details }} /> : null}
      </section>
    );
  }

  const m: IEmployeeDetail = result.data;
  return (
    <section className={styles.detail}>
      {back}
      <div className={styles.person}>
        <InitialsAvatar name={m.displayName} size={72} />
        <div>
          <h1 className={styles.title}>{m.displayName}</h1>
          <div className={styles.meta}>{[m.jobTitle, m.department].filter((v) => !!v).join(' · ')}</div>
        </div>
      </div>
      <div className={styles.factGrid}>
        {m.email ? <div className={styles.fact}><span>Email</span><a href={`mailto:${m.email}`}>{m.email}</a></div> : null}
        {m.phone ? <div className={styles.fact}><span>Phone</span>{m.phone}</div> : null}
        {m.mobile ? <div className={styles.fact}><span>Mobile</span>{m.mobile}</div> : null}
        {m.office ? <div className={styles.fact}><span>Office</span>{m.office}</div> : null}
      </div>
      {m.email ? (
        <a className={styles.cta} href={`https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(m.email)}`} target="_blank" rel="noopener noreferrer">
          Message on Teams {linkArrow}
        </a>
      ) : null}
    </section>
  );
};

export default DetailView;
