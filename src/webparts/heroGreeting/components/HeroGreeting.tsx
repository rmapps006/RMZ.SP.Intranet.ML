import * as React from 'react';
import styles from './HeroGreeting.module.scss';
import { IHeroGreetingProps } from './IHeroGreetingProps';
import { getGraphClient } from '../../../common/services/graphService';
import { getCurrentLanguage } from '../../../common/services/languageService';

// Greeting follows the viewer's own local time. Not admin-authored content
// (unlike the rest of this web part's text), so the translation is built in
// rather than a configurable *AR property.
function timeGreeting(): string {
  const language = getCurrentLanguage();
  const h: number = new Date().getHours();
  if (language === 'ar') {
    if (h < 12) {
      return 'صباح الخير';
    }
    if (h < 17) {
      return 'مساء الخير';
    }
    return 'مساء الخير';
  }
  if (h < 12) {
    return 'Good morning';
  }
  if (h < 17) {
    return 'Good afternoon';
  }
  return 'Good evening';
}

const HeroGreeting: React.FunctionComponent<IHeroGreetingProps> = (props) => {
  const [name, setName] = React.useState<string>(
    props.context.pageContext.user.displayName || ''
  );

  React.useEffect(() => {
    let active: boolean = true;
    getGraphClient(props.context)
      .then((client) => client.api('/me').select('displayName').get())
      .then((me: { displayName?: string }) => {
        if (active && me && me.displayName) {
          setName(me.displayName);
        }
      })
      .catch(() => {
        /* keep page-context fallback */
      });
    return () => {
      active = false;
    };
  }, [props.context]);

  return (
    <section className={styles.hero}>
      <div className={styles.bg} aria-hidden="true" />
      <div className={styles.content}>
        {props.showGreeting && name ? (
          <div className={styles.greet}>
            {timeGreeting()}, <span>{name}</span>
          </div>
        ) : null}
        {props.eyebrow ? <div className={styles.eyebrow}>{props.eyebrow}</div> : null}
        <h1
          className={styles.title}
          // Title supports a configurable emphasised second line.
        >
          {props.titleLine1}
          {props.titleEmphasis ? (
            <>
              <br />
              <em>{props.titleEmphasis}</em>
            </>
          ) : null}
        </h1>
        <div className={styles.rule} />
        {props.subtitle ? <p className={styles.sub}>{props.subtitle}</p> : null}
        {props.buttonText ? (
          <a className={styles.btn} href={props.buttonUrl || '#'}>
            {props.buttonText} &nbsp;&rarr;
          </a>
        ) : null}
      </div>
      <div className={styles.caption}>
        {getCurrentLanguage() === 'ar' ? 'صورة الغلاف · تصوير معماري راقٍ' : 'Hero image · Premium Architecture Photography'}
      </div>
    </section>
  );
};

export default HeroGreeting;
