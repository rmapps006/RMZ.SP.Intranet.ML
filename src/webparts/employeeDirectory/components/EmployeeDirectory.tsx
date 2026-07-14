import * as React from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { InitialsAvatar } from '../../../common/components/InitialsAvatar';
import { getPeople, searchPeople, IDirectoryPerson, IDirectoryConfig, FALLBACK_PEOPLE } from '../services/DirectoryService';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget } from '../../../common/util/format';
import { getCurrentLanguage, isRtl, Language } from '../../../common/services/languageService';
import { t, localizeChoice } from '../../../common/services/uiStrings';

const SearchIcon: React.FunctionComponent = () => (
  <svg width="13" height="13" viewBox="0 0 14 14" fill="none" className={styles.searchIcon} aria-hidden="true">
    <circle cx="6" cy="6" r="4.5" stroke="currentColor" strokeWidth="1.4" />
    <path d="M9.5 9.5L12 12" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" />
  </svg>
);

const EnvelopeIcon: React.FunctionComponent = () => (
  <svg width="11" height="11" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth="1.3" aria-hidden="true">
    <rect x="1.5" y="3" width="11" height="8" rx="1" />
    <path d="M2 4l5 3.4L12 4" />
  </svg>
);

const PhoneIcon: React.FunctionComponent = () => (
  <svg width="11" height="11" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinejoin="round" aria-hidden="true">
    <path d="M3 2.4c0 4.2 3.4 7.6 7.6 7.6l1-1.9-2.5-1-1 1C6.7 7.5 5.5 6.3 4.9 4.9l1-1-1-2.5L3 2.4z" />
  </svg>
);

/** A company filter chip: 'All' (empty domain) or a configured company domain. */
interface IChip {
  label: string;
  domain: string;
}

const EmployeeDirectory: React.FunctionComponent<IEmployeeDirectoryProps> = (props) => {
  const [people, setPeople] = React.useState<IDirectoryPerson[]>(FALLBACK_PEOPLE);
  const [searchResults, setSearchResults] = React.useState<IDirectoryPerson[]>([]);
  const [query, setQuery] = React.useState<string>('');
  const [searching, setSearching] = React.useState<boolean>(false);
  const [activeDomain, setActiveDomain] = React.useState<string>('');

  const language: Language = getCurrentLanguage();
  const settings = useSettings(props.context);
  const navKey: string = useNavKey();
  const pageSize: number = props.pageSize && props.pageSize > 0 ? props.pageSize : settings.directoryPageSize;
  const detailUrl: string = settings.detailPageUrl;

  // Build the central directory config (stable across renders via JSON keys).
  const domainsKey: string = JSON.stringify(settings.directoryDomains || []);
  const excludeKey: string = JSON.stringify(settings.directoryExclude || []);
  const config: IDirectoryConfig = React.useMemo(
    () => ({
      domains: (settings.directoryDomains || []).map((d) => (d.domain || '').toLowerCase()).filter((d) => !!d),
      exclude: settings.directoryExclude || [],
      pageSize
    }),
    [domainsKey, excludeKey, pageSize]
  );

  const chips: IChip[] = React.useMemo(() => {
    const list: IChip[] = [{ label: 'All', domain: '' }];
    (settings.directoryDomains || []).forEach((d) => {
      if (d.domain) {
        list.push({ label: d.label || d.domain, domain: d.domain.toLowerCase() });
      }
    });
    return list;
  }, [domainsKey]);

  // Default (no-search) listing.
  React.useEffect(() => {
    let active: boolean = true;
    getPeople(props.context, config)
      .then((result) => {
        if (active) {
          setPeople(result);
        }
      })
      .catch(() => {
        /* keep fallback */
      });
    return () => {
      active = false;
    };
  }, [props.context, config, navKey]);

  // Server-side directory search, debounced.
  React.useEffect(() => {
    const term: string = query.trim();
    if (term.length < 2) {
      setSearchResults([]);
      setSearching(false);
      return;
    }
    setSearching(true);
    let active: boolean = true;
    const handle: number = window.setTimeout(() => {
      searchPeople(props.context, term, config)
        .then((result) => {
          if (active) {
            setSearchResults(result);
            setSearching(false);
          }
        })
        .catch(() => {
          if (active) {
            setSearching(false);
          }
        });
    }, 350);
    return () => {
      active = false;
      window.clearTimeout(handle);
    };
  }, [query, props.context, config]);

  const isSearch: boolean = query.trim().length >= 2;
  const source: IDirectoryPerson[] = isSearch ? searchResults : people;
  const filtered: IDirectoryPerson[] = activeDomain
    ? source.filter((p) => p.domain === activeDomain)
    : source;

  return (
    <section className={styles.directory} dir={isRtl(language) ? 'rtl' : 'ltr'}>
      <SectionHeader title={props.title} linkText={props.linkText} linkUrl={props.fullDirectoryUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      <div className={styles.dirTop}>
        <div className={styles.dirSb}>
          <SearchIcon />
          <input
            type="text"
            placeholder={t('searchDirectory', language)}
            aria-label={t('searchDirectory', language)}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <div className={styles.dirChips}>
          {chips.map((chip) => (
            <div
              key={chip.domain || 'all'}
              role="button"
              tabIndex={0}
              aria-pressed={chip.domain === activeDomain}
              className={chip.domain === activeDomain ? `${styles.chip} ${styles.on}` : styles.chip}
              onClick={() => setActiveDomain(chip.domain)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  setActiveDomain(chip.domain);
                }
              }}
            >
              {localizeChoice(chip.label, language)}
            </div>
          ))}
        </div>
      </div>

      {isSearch ? (
        <div className={styles.dirStatus}>
          {searching
            ? t('searchingDirectory', language)
            : language === 'ar'
              ? `${filtered.length} نتيجة لـ ”${query.trim()}“`
              : `${filtered.length} result${filtered.length === 1 ? '' : 's'} for “${query.trim()}”`}
        </div>
      ) : null}

      <div className={styles.dirG}>
        {filtered.map((person, i) => (
          <div className={styles.dirCd} key={`${person.email || person.displayName}-${i}`}>
            <div className={styles.dirHead}>
              <InitialsAvatar name={person.displayName} size={42} gradient={person.gradient} />
              <div className={styles.dirInfo}>
                {detailUrl && person.email ? (
                  <a
                    className={styles.dirNm}
                    href={`${detailUrl}?type=employee&recId=${encodeURIComponent(person.email)}`}
                    {...linkTarget(settings.openLinksInNewTab)}
                    style={{ textDecoration: 'none', color: 'inherit' }}
                  >
                    {person.displayName}
                  </a>
                ) : (
                  <div className={styles.dirNm}>{person.displayName}</div>
                )}
                {person.jobTitle ? <div className={styles.dirRole}>{person.jobTitle}</div> : null}
                <div className={styles.dirDept}>
                  {[person.department, person.location].filter((v) => !!v).join(' · ')}
                </div>
              </div>
            </div>

            <div className={styles.dirContact}>
              {person.email ? (
                <div className={styles.dirLine}>
                  <EnvelopeIcon />
                  <span>{person.email}</span>
                </div>
              ) : null}
              {person.phone ? (
                <div className={styles.dirLine}>
                  <PhoneIcon />
                  <span>{person.phone}</span>
                </div>
              ) : null}
            </div>

            <div className={styles.dirActs}>
              <a className={styles.dirAct} href={person.phone ? `tel:${person.phone.replace(/[^+\d]/g, '')}` : '#'}>
                <PhoneIcon /> {t('call', language)}
              </a>
              <a className={styles.dirAct} href={person.email ? `mailto:${person.email}` : '#'}>
                <EnvelopeIcon /> {t('email', language)}
              </a>
              <a
                className={styles.dirAct}
                href={person.email ? `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(person.email)}` : '#'}
                target="_blank"
                rel="noreferrer"
              >
                {t('teams', language)}
              </a>
            </div>
          </div>
        ))}
      </div>

      {isSearch && !searching && filtered.length === 0 ? (
        <div className={styles.dirEmpty}>{t('noMatchingPeople', language)}</div>
      ) : null}
    </section>
  );
};

export default EmployeeDirectory;
