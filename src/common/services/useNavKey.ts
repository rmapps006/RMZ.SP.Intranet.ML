import * as React from 'react';

/**
 * SharePoint modern pages navigate client-side (SPA): the URL changes without a
 * full reload, and web-part React trees can persist across that navigation
 * without re-running their mount effects — so data fetched on mount goes stale
 * and widgets appear empty until a hard refresh.
 *
 * Detecting that navigation is deliberately belt-and-suspenders because
 * SharePoint's SPA router varies: we (1) patch history.pushState/replaceState,
 * (2) listen to popstate/hashchange, and (3) poll location as a safety net in
 * case the router bypasses our patched history methods. `useNavKey()` returns
 * the current path+search and updates on every client-side navigation — include
 * it in a data-fetch effect's deps to make the widget reload on navigation.
 */

type Listener = () => void;

const listeners: Set<Listener> = new Set<Listener>();
let patched: boolean = false;
let lastKey: string = '';

function currentKey(): string {
  try {
    return window.location.pathname + window.location.search;
  } catch {
    return '';
  }
}

function ensurePatched(): void {
  if (patched || typeof window === 'undefined' || !window.history) {
    return;
  }
  patched = true;
  lastKey = currentKey();

  // Fire only when the URL actually changed, so the poll and the history patch
  // don't double-notify listeners for the same navigation.
  const fire = (): void => {
    const key: string = currentKey();
    if (key === lastKey) {
      return;
    }
    lastKey = key;
    listeners.forEach((l: Listener) => {
      try {
        l();
      } catch {
        /* ignore a single listener error */
      }
    });
  };

  const wrap = (original: History['pushState']): History['pushState'] =>
    function patchedFn(this: History, ...args: Parameters<History['pushState']>): void {
      original.apply(this, args);
      fire();
    };

  try {
    window.history.pushState = wrap(window.history.pushState);
    window.history.replaceState = wrap(window.history.replaceState);
  } catch {
    /* some environments freeze history — fall back to popstate + polling */
  }
  window.addEventListener('popstate', fire);
  window.addEventListener('hashchange', fire);
  // Safety net: catches SPA navigations that don't go through our patched
  // history methods (varies by SharePoint version). Cheap 400ms comparison.
  window.setInterval(fire, 400);
}

/** Returns a key that changes on every client-side navigation. */
export function useNavKey(): string {
  const [key, setKey] = React.useState<string>(currentKey);

  React.useEffect(() => {
    ensurePatched();
    const listener: Listener = () => setKey(currentKey());
    listeners.add(listener);
    // Sync once in case the URL changed between render and effect.
    listener();
    return () => {
      listeners.delete(listener);
    };
  }, []);

  return key;
}
