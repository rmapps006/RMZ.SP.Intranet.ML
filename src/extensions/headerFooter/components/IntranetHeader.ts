import styles from './HeaderFooter.module.scss';
import { ICurrentUser } from '../services/GraphUserService';
import { Language } from '../../../common/services/languageService';

export interface INavItem {
  label: string;
  url: string;
  newTab?: boolean;
}

export interface IIntranetHeaderOptions {
  navItems: INavItem[];
  searchPlaceholder: string;
  searchAriaLabel: string;
  notificationsLabel: string;
  logoUrl: string;
  logoAlt: string;
  subLabel: string;
  homeUrl: string;
  webAbsoluteUrl: string;
  currentPath: string;
  user: ICurrentUser;
  /** Active display language — used to label/drive the language toggle. */
  language: Language;
  /** Invoked when the visitor clicks the language toggle (switches en <-> ar). */
  onLanguageToggle: () => void;
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string
): HTMLElementTagNameMap[K] {
  const node: HTMLElementTagNameMap[K] = document.createElement(tag);
  if (className) {
    node.className = className;
  }
  if (text !== undefined) {
    node.textContent = text;
  }
  return node;
}

/**
 * Determines whether a nav item matches the current page so it can be
 * highlighted with the sand underline (matches `.nav a.on` in the design).
 */
function isActive(itemUrl: string, currentPath: string): boolean {
  const path: string = currentPath.toLowerCase();
  let target: string;
  try {
    // Resolve both absolute and server-relative URLs to a pathname.
    target = new URL(itemUrl, window.location.origin).pathname.toLowerCase();
  } catch {
    target = itemUrl.toLowerCase();
  }

  if (target === '/' || target === '') {
    return path === '/' || path.indexOf('/sitepages/home.aspx') >= 0;
  }
  return path === target || path.indexOf(target) === 0;
}

export class IntranetHeader {
  constructor(private readonly options: IIntranetHeaderOptions) {}

  public render(container: HTMLElement): void {
    const header: HTMLElement = el('div', styles.intranetHeader);
    header.setAttribute('role', 'banner');

    header.appendChild(el('div', styles.accentBar));

    const inner: HTMLElement = el('div', styles.inner);

    inner.appendChild(this.renderLogo());
    inner.appendChild(el('div', styles.logoDivider));
    inner.appendChild(this.renderNav());
    inner.appendChild(this.renderUser());

    header.appendChild(inner);
    container.appendChild(header);
  }

  private renderLogo(): HTMLElement {
    const logo: HTMLAnchorElement = document.createElement('a');
    logo.className = styles.logo;
    logo.href = this.options.homeUrl;
    logo.setAttribute('aria-label', this.options.logoAlt);

    const img: HTMLImageElement = document.createElement('img');
    img.className = styles.logoImg;
    img.src = this.options.logoUrl;
    img.alt = this.options.logoAlt;
    img.height = 34;
    logo.appendChild(img);

    // Small text beside the logo (e.g. "UAE"); hidden when the client clears it.
    if (this.options.subLabel) {
      logo.appendChild(el('span', styles.logoSub, this.options.subLabel));
    }
    return logo;
  }

  private renderNav(): HTMLElement {
    const nav: HTMLElement = el('nav', styles.nav);
    nav.setAttribute('aria-label', 'Primary');

    this.options.navItems.forEach((item) => {
      const active: boolean = isActive(item.url, this.options.currentPath);
      const link: HTMLAnchorElement = document.createElement('a');
      link.className = active ? `${styles.navLink} ${styles.navLinkActive}` : styles.navLink;
      link.href = item.url;
      link.textContent = item.label;
      if (item.newTab) {
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
      }
      if (active) {
        link.setAttribute('aria-current', 'page');
      }
      nav.appendChild(link);
    });

    return nav;
  }

  private renderUser(): HTMLElement {
    const fragment: HTMLElement = el('div', styles.logo);
    fragment.style.flexShrink = '0';
    // Search/bell removed — pin the profile block to the trailing edge (right in
    // LTR, left in RTL — margin-inline-start respects the document's dir).
    fragment.style.marginInlineStart = 'auto';

    fragment.appendChild(this.renderLanguageToggle());

    const avatar: HTMLElement = el('div', styles.avatar, this.options.user.initials);
    avatar.setAttribute('aria-hidden', 'true');
    fragment.appendChild(avatar);

    fragment.appendChild(el('span', styles.profName, this.options.user.displayName));
    return fragment;
  }

  /** English/Arabic switch — labelled with the language it switches *to*. */
  private renderLanguageToggle(): HTMLElement {
    const isArabic: boolean = this.options.language === 'ar';
    const button: HTMLButtonElement = document.createElement('button');
    button.type = 'button';
    button.className = styles.langToggle;
    button.textContent = isArabic ? 'English' : 'العربية';
    button.setAttribute('aria-label', isArabic ? 'Switch to English' : 'التبديل إلى العربية');
    button.addEventListener('click', () => this.options.onLanguageToggle());
    return button;
  }
}
