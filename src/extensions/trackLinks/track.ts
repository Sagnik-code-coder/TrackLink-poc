import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface ITrackLinksApplicationCustomizerProperties {}

interface ILink {
  link: string;
  name: string;
  bookmarked: boolean;
}

export default class TrackLinksApplicationCustomizer extends BaseApplicationCustomizer<ITrackLinksApplicationCustomizerProperties> {
  private lastRecordedUrl: string = '';

  public onInit(): Promise<void> {
    console.log('TrackLinks Application Customizer initialized');

    const normalizeUrl = (url: string): string => {
      try {
        const parsedUrl = new URL(url, window.location.origin);
        return `${parsedUrl.origin}${parsedUrl.pathname}`;
      } catch (error) {
        console.error('Failed to normalize URL:', error);
        return url;
      }
    };

    const addClickedLink = (link: string, name: string): void => {
      const normalizedLink = normalizeUrl(link);
      const links = JSON.parse(sessionStorage.getItem('clickedLinks') || '[]') as ILink[];

      if (!links.some((item) => normalizeUrl(item.link) === normalizedLink)) {
        if (links.length >= 10) links.shift();
        links.push({ link: normalizedLink, name, bookmarked: false });
        sessionStorage.setItem('clickedLinks', JSON.stringify(links));
      }

      console.log('Updated Clicked Links:', links);
    };

    const handleLinkClick = (e: Event): void => {
      e.preventDefault();
      const clickableElement = e.target as HTMLElement;
      let href = '';
      const name = clickableElement.innerText.trim() || 'Unnamed Link';

      if (clickableElement.tagName === 'A') {
        href =
          (clickableElement as HTMLAnchorElement).href ||
          clickableElement.getAttribute('href') ||
          '';
      } else if (
        clickableElement.tagName === 'BUTTON' &&
        clickableElement.getAttribute('role') === 'link'
      ) {
        const parentDiv = clickableElement.closest('div[data-drop-target-key]');
        if (parentDiv) {
          const dropTargetKey = parentDiv.getAttribute('data-drop-target-key');
          if (dropTargetKey) {
            try {
              const parsedKey = JSON.parse(
                dropTargetKey.replace(/&quot;/g, '"')
              );
              const filePath = `${parsedKey[2]}/${clickableElement.textContent?.trim()}`;
              href = `${filePath}`.replace(/ /g, '%20');
            } catch (error) {
              console.error('Error parsing dropTargetKey:', error);
            }
          }
        }
      }

      if (href) {
        const isFile = href.match(/\.(docx|pptx|xlsx|pdf|doc|ppt|xls)$/i);
        const isSharePointLink =
          href.startsWith('/sites/') || href.includes('.sharepoint.com');
        if (isFile || isSharePointLink) addClickedLink(href, name);
        window.location.href = href;
      } else {
        console.warn('No valid link found for the clicked element.');
      }
    };

    const attachEventListeners = (): void => {
      const linkSelector = `a, .ms-ListView-cell a, .ms-DocumentLibraryLink, .ms-Nav-PaneLink, button[role="link"]`;

      document.querySelectorAll(linkSelector).forEach((element) => {
        element.addEventListener('click', handleLinkClick);
      });
    };

    const handleUrlChange = (): void => {
      const currentUrl = window.location.href;
      const links = JSON.parse(sessionStorage.getItem('clickedLinks') || '[]') as ILink[];
      const isUrlAlreadyAdded = links.some((item) =>
        new URL(item.link, window.location.origin).href === new URL(currentUrl, window.location.origin).href
      );

      if (this.lastRecordedUrl !== currentUrl && !isUrlAlreadyAdded) {
        this.lastRecordedUrl = currentUrl;
        const pageName = document.title || 'Untitled Page';
        addClickedLink(currentUrl, pageName);
      }
    };

    const observeUrlChanges = (): void => {
      this.lastRecordedUrl = window.location.href;
      setInterval(() => {
        handleUrlChange();
      }, 1000);
    };

    const observer = new MutationObserver((): void => {
      attachEventListeners();
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
    });

    SPComponentLoader.loadScript(
      'https://mazdausa.sharepoint.com/sites/MCIOneMazdaDev/_layouts/15/SP.RequestExecutor.js'
    )
      .then(() => {
        console.log('SP.RequestExecutor script loaded');
        attachEventListeners();
      })
      .catch((error) => console.error('Failed to load SP.RequestExecutor', error));

    observeUrlChanges();

    return Promise.resolve();
  }
}
