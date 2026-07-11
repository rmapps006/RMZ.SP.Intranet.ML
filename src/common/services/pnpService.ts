import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import { WebPartContext } from '@microsoft/sp-webpart-base';

let _sp: SPFI | undefined;
let _ctx: WebPartContext | undefined;

/**
 * Returns a PnP JS v4 instance bound to the current SPFx context. Rebuilds when a
 * new context is supplied — SharePoint hands web parts a fresh context on
 * client-side navigation, and reusing the old one can make list reads fail (the
 * cause of widgets showing empty until a full reload). Use for all data access.
 */
export function getSP(context: WebPartContext): SPFI {
  if (!_sp || _ctx !== context) {
    _sp = spfi().using(SPFx(context));
    _ctx = context;
  }
  return _sp;
}
