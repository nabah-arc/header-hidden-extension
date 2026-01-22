import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
// import { Dialog } from '@microsoft/sp-dialog';

// import * as strings from 'HiddenExtensionApplicationCustomizerStrings';

const LOG_SOURCE: string = 'HiddenExtensionApplicationCustomizer';

export interface IHiddenExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HiddenExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IHiddenExtensionApplicationCustomizerProperties> {

    public onInit(): Promise<void> {
      Log.info(LOG_SOURCE, `Initialized HideSharePointHeader`);
  
      this.hideHeader();
  
      return Promise.resolve();
    }
    private hideHeader(): void {
      const css = `
        div[data-automation-id="Header"],
        div[data-automation-id="SiteHeader"],
        header[class*="ms-SiteHeader"] {
          display: none !important;
        }
    
        div[data-automation-id="CommandBar"],
        div[class*="commandBarWrapper"],
        div[class*="ms-CommandBar"] {
          display: none !important;
        }
    
        div[data-automation-id="PageHeader"],
        div[class*="pageHeader"],
        div[class*="ms-PageHeader"] {
          display: none !important;
        }
    
        #SuiteNavPlaceHolder,
        #O365_NavHeader,
        div[class*="o365cs-base"] {
          display: none !important;
        }
    
        div[data-placeholder="Top"] {
          display: none !important;
        }
    
        /* Uncomment to hide left nav */
        /* nav[role="navigation"] {
          display: none !important;
        } */
    
        div[data-automation-id="CanvasZone"],
        div[class*="CanvasZone"] {
          margin-top: 0 !important;
          padding-top: 0 !important;
        }
      `;
    
      const style = document.createElement("style");
      style.innerHTML = css;
      document.head.appendChild(style);
    }
    
}
