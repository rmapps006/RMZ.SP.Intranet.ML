import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from './pnpService';
import { SettingsService, IIntranetSettings, getCachedSettings } from './SettingsService';
import { useNavKey } from './useNavKey';

/**
 * Returns the central intranet settings. Seeds from the (possibly cold)
 * sessionStorage cache for first paint, then fetches the settings list and
 * re-renders — so web parts don't depend on the header/footer extension having
 * warmed the cache first (which caused intermittent blanks, e.g. the
 * Departments grid rendering empty until a few refreshes warmed the cache).
 */
export function useSettings(context: WebPartContext): IIntranetSettings {
  const [settings, setSettings] = React.useState<IIntranetSettings>(getCachedSettings());
  const navKey: string = useNavKey();

  React.useEffect(() => {
    let active: boolean = true;
    // Re-seed from the (possibly re-warmed) cache immediately on navigation.
    setSettings(getCachedSettings());
    new SettingsService(getSP(context))
      .getSettings()
      .then((s: IIntranetSettings) => {
        if (active) {
          setSettings(s);
        }
      })
      .catch(() => {
        /* keep cached/defaults */
      });
    return () => {
      active = false;
    };
  }, [context, navKey]);

  return settings;
}
