import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from './pnpService';
import {
  DepartmentSettingsService,
  IDepartmentSettings,
  getCachedDepartmentSettings
} from './DepartmentSettingsService';
import { useNavKey } from './useNavKey';

/**
 * Returns the current department settings. Seeds from the (possibly cold)
 * sessionStorage cache for first paint, then fetches the local site's settings
 * list and re-renders — so department web parts don't depend on the header
 * extension having warmed the cache first (which caused intermittent blanks).
 */
export function useDepartmentSettings(context: WebPartContext): IDepartmentSettings {
  const [settings, setSettings] = React.useState<IDepartmentSettings>(getCachedDepartmentSettings());
  const navKey: string = useNavKey();

  React.useEffect(() => {
    let active: boolean = true;
    setSettings(getCachedDepartmentSettings());
    new DepartmentSettingsService(getSP(context))
      .getSettings()
      .then((s: IDepartmentSettings) => {
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
