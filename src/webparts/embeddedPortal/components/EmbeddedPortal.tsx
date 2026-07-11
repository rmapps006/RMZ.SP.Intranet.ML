import * as React from 'react';
import styles from './EmbeddedPortal.module.scss';
import { IEmbeddedPortalProps } from './IEmbeddedPortalProps';
import { getSP } from '../../../common/services/pnpService';
import { SettingsService, getCachedSettings } from '../../../common/services/SettingsService';

// Default (stricter) sandbox used for ordinary embeds.
const STRICT_SANDBOX: string = 'allow-scripts allow-same-origin allow-forms allow-popups';

// Relaxed sandbox for "app sign-in mode": lets a trusted embedded app run its
// own sign-in — keep its session cookie (allow-same-origin), open the IdP popup
// (allow-popups + allow-popups-to-escape-sandbox), request storage access on
// browsers that block third-party cookies, and allow user-initiated top-nav.
const SIGN_IN_SANDBOX: string = [
  'allow-scripts',
  'allow-same-origin',
  'allow-forms',
  'allow-popups',
  'allow-popups-to-escape-sandbox',
  'allow-storage-access-by-user-activation',
  'allow-top-navigation-by-user-activation'
].join(' ');

/** Returns the origin of a URL, or '' when it can't be parsed. */
function originOf(u: string): string {
  try {
    return new URL(u).origin;
  } catch {
    return '';
  }
}

/** Resolves the configured height to a CSS value: a bare number → px; else used verbatim (e.g. "85vh"). */
function resolveHeight(raw: string): string {
  const v: string = (raw || '').trim();
  if (!v) {
    return '520px';
  }
  return /^\d+$/.test(v) ? `${v}px` : v;
}

const EmbeddedPortal: React.FunctionComponent<IEmbeddedPortalProps> = (props) => {
  const url: string = props.portalUrl || 'https://www.orauae.com';
  const cssHeight: string = resolveHeight(props.frameHeight);
  // Give relative (vh/%) heights a sensible floor so full-screen apps aren't squashed.
  const isRelativeHeight: boolean = /vh|%/i.test(cssHeight);
  const origin: string = originOf(url);
  // In app sign-in mode use the relaxed sandbox and scope device permissions to
  // the embedded app's own origin; otherwise keep the stricter defaults.
  const sandbox: string = props.appSignInMode ? SIGN_IN_SANDBOX : STRICT_SANDBOX;
  const allow: string | undefined = props.appSignInMode
    ? `clipboard-read; clipboard-write;${origin ? ` microphone ${origin}` : ''}`
    : undefined;

  // Hide the URL when the web part toggle asks for it OR when the central
  // setting (Admin screen) hides it for every embed on the intranet.
  const [hideUrl, setHideUrl] = React.useState<boolean>(props.hideUrl === true);

  React.useEffect(() => {
    if (props.hideUrl) {
      setHideUrl(true);
      return;
    }
    let active: boolean = true;
    new SettingsService(getSP(props.context))
      .getSettings()
      .then((s) => {
        if (active && s.embedHideUrl) {
          setHideUrl(true);
        }
      })
      .catch(() => {
        /* defaults to the local toggle */
      });
    return () => {
      active = false;
    };
  }, [props.context, props.hideUrl]);

  return (
    <section className={styles.embeddedPortal}>
      <div className={styles.inner}>
        {/* Custom header so the link can open in a new tab (SectionHeader has no target). */}
        <div className={styles.head}>
          <h2 className={styles.title}>{props.sectionTitle}</h2>
          {props.linkText ? (
            <a
              className={styles.link}
              href={url}
              target="_blank"
              rel="noopener noreferrer"
            >
              {props.linkText}
            </a>
          ) : null}
        </div>

        <div className={styles.frameCard}>
          {/* The fake browser chrome (dots + URL) is hidden entirely when "Hide URL" is on. */}
          {hideUrl ? null : (
            <div className={styles.chrome}>
              <div className={styles.dots}>
                <span className={`${styles.dot} ${styles.dotRed}`} />
                <span className={`${styles.dot} ${styles.dotYellow}`} />
                <span className={`${styles.dot} ${styles.dotGreen}`} />
              </div>
              <div className={styles.pill}>{url}</div>
              <svg
                className={styles.ext}
                width="14"
                height="14"
                viewBox="0 0 14 14"
                fill="none"
                stroke="#b5b5b5"
                strokeWidth="1.3"
              >
                <path d="M9 2h3v3M12 2l-5 5M6 3H3a1 1 0 00-1 1v7a1 1 0 001 1h7a1 1 0 001-1V8" />
              </svg>
            </div>
          )}
          <iframe
            className={styles.frame}
            src={url}
            title={props.sectionTitle || `${getCachedSettings().clientName} Portal`}
            loading="lazy"
            style={{ height: cssHeight, minHeight: isRelativeHeight ? 700 : undefined }}
            sandbox={sandbox}
            allow={allow}
            referrerPolicy="strict-origin-when-cross-origin"
          />
        </div>
      </div>
    </section>
  );
};

export default EmbeddedPortal;
