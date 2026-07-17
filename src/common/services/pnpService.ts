import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const CURRENT = '__current__';
let _cache: { [url: string]: SPFI } = {};
let _ctx: WebPartContext | undefined;

/**
 * Returns a PnP JS v4 instance bound to the current SPFx context, or — when a
 * `siteUrl` is supplied — to that specific site (web) using the same SPFx auth.
 * This lets a web part read a library that lives on a different site than the
 * page it is placed on. Instances are cached per target URL and rebuilt whenever
 * a new context is supplied (SharePoint hands web parts a fresh context on
 * client-side navigation, and reusing the old one can make reads fail). Use for
 * all data access.
 */
export function getSP(context: WebPartContext, siteUrl?: string): SPFI {
  if (_ctx !== context) {
    _cache = {};
    _ctx = context;
  }
  const trimmed: string = (siteUrl || '').trim().replace(/\/+$/, '');
  const key: string = trimmed || CURRENT;
  if (!_cache[key]) {
    _cache[key] = key === CURRENT ? spfi().using(SPFx(context)) : spfi(trimmed).using(SPFx(context));
  }
  return _cache[key];
}
