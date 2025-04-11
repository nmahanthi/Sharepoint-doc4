import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as strings from 'DmlCustomUiChangesApplicationCustomizerStrings';
import { spfi } from "@pnp/sp";
import { IHubSiteInfo } from "@pnp/sp/hubsites";
import "@pnp/sp/hubsites";
import { getSP } from '../../pnpjs-config';

const LOG_SOURCE: string = 'DmlCustomUiChangesApplicationCustomizer';

export interface IDmlCustomUiChangesApplicationCustomizerProperties { }

export default class DmlCustomUiChangesApplicationCustomizer
  extends BaseApplicationCustomizer<IDmlCustomUiChangesApplicationCustomizerProperties> {
  private sp: ReturnType<typeof spfi>;
  private previousUrl: string = window.location.href;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.sp = getSP(this.context);
    await this.applyCustomizations();
    this.context.application.navigatedEvent.add(this, this.onNavigated);
    setInterval(() => {
      if (this.previousUrl !== window.location.href) {
        this.previousUrl = window.location.href;
        this.applyCustomizations()
          .catch((error) => {
            window.location.reload();
            console.error("Error applying customizations: ", error);
          });
      }
    }, 1000);
    this.addDraftClickListener();
    return Promise.resolve();
  }

  private async onNavigated(): Promise<void> {
    await this.applyCustomizations();
  }

  private addDraftClickListener(): void {
    const draftLink: HTMLElement | null = document.querySelector('a[href*="Draft"]');
    if (draftLink) {
      draftLink.addEventListener('click', async () => {
        await this.applyCustomizations();
      });
    }
  }

  private async applyCustomizations(): Promise<void> {
    const listabsUrl = this.context.pageContext.list?.serverRelativeUrl.toString();
    const listTitle = listabsUrl?.split('/').pop();

    if (listTitle === "Draft" || listTitle === "Baselines" ||
      listTitle === "ApplicableDocuments" || listTitle === "PreviousVersions") {
      const hubSiteId = this.context.pageContext.legacyPageContext.hubSiteId;

      if (hubSiteId) {
        try {
          const hubsite: IHubSiteInfo = await this.sp.hubSites.getById(hubSiteId)();
          const hubSiteUrl = hubsite.SiteUrl;

          if (hubSiteUrl) {
            const commonCssUrl = `${hubSiteUrl}/SiteAssets/customHideMenuOptions.css`;
            const adminCssUrl = `${hubSiteUrl}/SiteAssets/customAdminStyles.css`;
            this.injectCustomCss(commonCssUrl);
            this.injectCustomCss(adminCssUrl);

          }
        } catch (error) {
          console.error("Error fetching hub site info: ", error);
        }
      }

      if (listTitle === "Draft" && this.isChildFolder()) {
        document.body.classList.add('draft-library');
        this.showNewCommandButton();
      } else {
        document.body.classList.remove('draft-library');
        this.hideNewCommandButton();
      }

      if (listTitle === "Baselines") {
        document.body.classList.add('baselines-library');
        this.showNewCommandButton();
      } else {
        document.body.classList.remove('baselines-library');
        this.hideNewCommandButton();
      }
    } else {
      this.removeCustomCss();
      this.showNewCommandButton();
    }
  }

  private injectCustomCss(cssUrl: string): void {
    if (!document.querySelector(`link[href="${cssUrl}"]`)) {
      const head: HTMLElement = document.getElementsByTagName("head")[0] || document.documentElement;
      const customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssUrl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeend", customStyle);
      console.log(`Custom CSS injected: ${cssUrl}`);
    }
  }

  private removeCustomCss(): void {
    const customStyles: NodeListOf<HTMLLinkElement> = document.querySelectorAll("link[href*='SiteAssets']");
    customStyles.forEach(style => style.remove());
    console.log("All Custom CSS removed.");
  }

  private isChildFolder(): boolean {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get('id') !== null;
  }

  private showNewCommandButton(): void {
    const newCommandButton: HTMLElement | null = document.querySelector('button[data-automationid="newCommand"]');
    if (newCommandButton) {
      newCommandButton.style.display = "inline-block";
    }
  }

  private hideNewCommandButton(): void {
    const newCommandButton: HTMLElement | null = document.querySelector('button[data-automationid="newCommand"]');
    if (newCommandButton) {
      newCommandButton.style.display = "none";
    }
  }
}