import styles from './HeaderFooter.module.scss';

export interface IIntranetFooterOptions {
  wordmark: string;
  copyright: string;
  tagline: string;
}

export class IntranetFooter {
  constructor(private readonly options: IIntranetFooterOptions) {}

  public render(container: HTMLElement): void {
    const footer: HTMLElement = document.createElement('div');
    footer.className = styles.intranetFooter;
    footer.setAttribute('role', 'contentinfo');

    const logo: HTMLElement = document.createElement('div');
    logo.className = styles.footerLogo;
    logo.textContent = this.options.wordmark;

    const copy: HTMLElement = document.createElement('div');
    copy.className = styles.footerCopy;
    copy.textContent = this.options.copyright;

    const tag: HTMLElement = document.createElement('div');
    tag.className = styles.footerTag;
    tag.textContent = this.options.tagline;

    footer.appendChild(logo);
    footer.appendChild(copy);
    footer.appendChild(tag);
    container.appendChild(footer);
  }
}
