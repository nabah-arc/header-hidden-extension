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
  
      // Check if we should hide headers on this page
      if (!this.shouldHideHeaders()) {
        return Promise.resolve();
      }
  
      // CRITICAL: Inject CSS IMMEDIATELY and SYNCHRONOUSLY to prevent FOUC
      // Do this FIRST before anything else - use multiple methods for maximum coverage
      this.injectCSSImmediately();
      this.injectCSSViaScript();
      
      // Add placeholder overlay to hide headers immediately
      this.addPlaceholderOverlay();
      
      // Also hide existing elements immediately (synchronous)
      this.hideExistingElements();
      
      // Use MutationObserver to hide elements that load later
      this.observeAndHide();
  
      return Promise.resolve();
    }
    
    private injectCSSViaScript(): void {
      // Try to inject CSS using a blocking script approach for immediate effect
      try {
        const css = this.getHeaderHidingCSS();
        const script = document.createElement('script');
        script.textContent = `
          (function() {
            var style = document.createElement('style');
            style.id = 'hideHeaderExtensionCSS-blocking';
            style.type = 'text/css';
            style.innerHTML = ${JSON.stringify(css)};
            if (document.head) {
              document.head.insertBefore(style, document.head.firstChild);
            } else {
              document.documentElement.insertBefore(style, document.documentElement.firstChild);
            }
          })();
        `;
        document.documentElement.appendChild(script);
        document.documentElement.removeChild(script);
      } catch (e) {
        // Fallback if script injection fails
      }
    }
    
    private addPlaceholderOverlay(): void {
      // Add a placeholder overlay that covers the header area immediately
      const overlay = document.createElement('div');
      overlay.id = 'headerHideOverlay';
      overlay.style.cssText = `
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 250px !important;
        background-color: white !important;
        z-index: 999999 !important;
        pointer-events: none !important;
        display: block !important;
      `;
      
      // Insert immediately - try multiple locations
      if (document.body) {
        document.body.insertBefore(overlay, document.body.firstChild);
      } else if (document.documentElement) {
        document.documentElement.appendChild(overlay);
      }
      
      // Also try to insert into head as a style that creates the overlay
      try {
        const overlayStyle = document.createElement('style');
        overlayStyle.id = 'headerHideOverlayStyle';
        overlayStyle.innerHTML = `
          #headerHideOverlay {
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            width: 100% !important;
            height: 250px !important;
            background-color: white !important;
            z-index: 999999 !important;
            pointer-events: none !important;
            display: block !important;
          }
        `;
        if (document.head) {
          document.head.appendChild(overlayStyle);
        }
      } catch (e) {
        // Ignore
      }
      
      // Remove overlay after CSS is confirmed to be applied
      let checkCount = 0;
      const checkInterval = setInterval(() => {
        checkCount++;
        const cssApplied = document.getElementById('hideHeaderExtensionCSS') || 
                          document.getElementById('hideHeaderExtensionCSS-blocking');
        
        // Remove overlay if CSS is applied or after 500ms
        if (cssApplied || checkCount > 10) {
          if (overlay.parentNode) {
            overlay.parentNode.removeChild(overlay);
          }
          const overlayStyle = document.getElementById('headerHideOverlayStyle');
          if (overlayStyle && overlayStyle.parentNode) {
            overlayStyle.parentNode.removeChild(overlayStyle);
          }
          clearInterval(checkInterval);
        }
      }, 50);
    }
    
    private injectCSSImmediately(): void {
      // Check if CSS is already injected
      if (document.getElementById('hideHeaderExtensionCSS')) {
        return;
      }
      
      // Get CSS content
      const css = this.getHeaderHidingCSS();
      
      // Method 1: Inject into head immediately
      const style = document.createElement('style');
      style.id = 'hideHeaderExtensionCSS';
      style.type = 'text/css';
      style.innerHTML = css;
      
      // Insert at the very beginning of head to ensure it loads first
      if (document.head) {
        document.head.insertBefore(style, document.head.firstChild);
      } else {
        // If head doesn't exist yet, inject into documentElement
        if (document.documentElement) {
          document.documentElement.insertBefore(style, document.documentElement.firstChild);
        }
        // Also wait for head and inject there too
        const observer = new MutationObserver((mutations, obs) => {
          if (document.head && !document.getElementById('hideHeaderExtensionCSS')) {
            const style2 = document.createElement('style');
            style2.id = 'hideHeaderExtensionCSS';
            style2.type = 'text/css';
            style2.innerHTML = css;
            document.head.insertBefore(style2, document.head.firstChild);
            obs.disconnect();
          }
        });
        observer.observe(document.documentElement, { childList: true });
      }
      
      // Method 2: Also inject as inline style in documentElement for immediate effect
      try {
        const inlineStyle = document.createElement('style');
        inlineStyle.id = 'hideHeaderExtensionCSS-inline';
        inlineStyle.type = 'text/css';
        inlineStyle.innerHTML = css;
        document.documentElement.appendChild(inlineStyle);
      } catch (e) {
        // Ignore errors
      }
      
      // Also try to inject inline styles on existing elements immediately
      this.applyInlineStylesImmediately();
      
      // Use requestAnimationFrame to apply styles as early as possible in the render cycle
      if (window.requestAnimationFrame) {
        window.requestAnimationFrame(() => {
          this.applyInlineStylesImmediately();
        });
      }
      
      // Also apply immediately after a very short delay to catch late-loading elements
      setTimeout(() => {
        this.applyInlineStylesImmediately();
      }, 0);
    }
    
    private applyInlineStylesImmediately(): void {
      // Apply inline styles directly to elements that might already exist
      const selectors = [
        '[data-automationid="SiteHeader"]',
        '[data-automation-id="SiteHeader"]',
        '[aria-label="SharePoint Site Header"]',
        '[class*="ms-HubNav"]',
        '[class*="hubNavRow"]',
        '#SuiteNavPlaceHolder',
        '#O365_NavHeader'
      ];
      
      selectors.forEach(selector => {
        try {
          const elements = document.querySelectorAll(selector);
          elements.forEach((el: HTMLElement) => {
            const isListHeader = el.closest && el.closest('[class*="ms-List"], [class*="ms-DetailsList"], [class*="ms-Table"]');
            if (!isListHeader) {
              el.style.cssText = 'display: none !important; visibility: hidden !important; height: 0 !important; overflow: hidden !important;';
            }
          });
        } catch (e) {
          // Ignore errors
        }
      });
    }
    
    private shouldHideHeaders(): boolean {
      // Get current page URL
      const pathname = window.location.pathname.toLowerCase();
      
      // Exclude pages where we don't want to hide headers
      const excludePaths = [
        '/_layouts/15/viewlsts.aspx', // Site Contents
        '/_layouts/15/settings.aspx', // Site Settings
        '/_layouts/15/people.aspx', // People
        '/_layouts/15/user.aspx', // User Profile
        '/_layouts/15/appredirect.aspx', // App Redirect
        '/_layouts/15/start.aspx', // Site Contents (alternative)
        '/lists/', // List pages (to preserve list headers)
        '/list/', // List pages (alternative)
        '/_layouts/15/listedit.aspx', // List Settings
        '/_layouts/15/viewlsts.aspx' // View Lists
      ];
      
      // Check if current page should be excluded
      for (const excludePath of excludePaths) {
        if (pathname.indexOf(excludePath) !== -1) {
          return false;
        }
      }
      
      // Only apply on Site Pages or specific pages where webpart is added
      // You can customize this condition based on your specific page URLs
      const includePaths = [
        '/sitepages/', // Site Pages
        '/pages/', // Pages library
        '/site pages/' // Site Pages (alternative)
      ];
      
      // If it's a Site Page, apply the extension
      for (const includePath of includePaths) {
        if (pathname.indexOf(includePath) !== -1) {
          return true;
        }
      }
      
      // Default: apply on pages that are not excluded
      // This ensures it works on the page where webpart is added
      return true;
    }
    
    private hideExistingElements(): void {
      // Hide elements that already exist in DOM
      // Only target site headers, not list headers
      const selectors = [
        '[data-automationid="SiteHeader"]',
        '[data-automation-id="SiteHeader"]',
        '[aria-label="SharePoint Site Header"]',
        '[class*="ms-HubNav"]',
        '[class*="hubNavRow"]',
        '[aria-label*="BMRHUB"]',
        '#SuiteNavPlaceHolder',
        '#O365_NavHeader',
        '[class*="o365cs-base"]'
      ];
      
      selectors.forEach(selector => {
        try {
          const elements = document.querySelectorAll(selector);
          elements.forEach((el: HTMLElement) => {
            // Exclude list headers - check if element is inside a list
            const isListHeader = el.closest('[class*="ms-List"], [class*="ms-DetailsList"], [class*="ms-Table"], [data-automation-id="List"], [data-automation-id="ListView"]');
            if (!isListHeader) {
              el.style.display = 'none';
              el.style.visibility = 'hidden';
              el.style.height = '0';
              el.style.overflow = 'hidden';
            }
          });
        } catch (e) {
          // Ignore selector errors
        }
      });
      
      // Specifically target site headerRow and mainHeader (not list headers)
      try {
        const headerRows = document.querySelectorAll('[class*="headerRow"]');
        headerRows.forEach((el: HTMLElement) => {
          // Only hide if it's a site header, not a list header
          const isListHeader = el.closest('[class*="ms-List"], [class*="ms-DetailsList"], [class*="ms-Table"]');
          const hasSiteHeaderAttr = el.getAttribute('data-automationid') === 'SiteHeader' || 
                                    el.getAttribute('data-automation-id') === 'SiteHeader' ||
                                    el.getAttribute('aria-label') === 'SharePoint Site Header';
          
          if (!isListHeader && (hasSiteHeaderAttr || el.closest('[data-automationid="SiteHeader"]'))) {
            el.style.display = 'none';
            el.style.visibility = 'hidden';
            el.style.height = '0';
            el.style.overflow = 'hidden';
          }
        });
        
        const mainHeaders = document.querySelectorAll('[class*="mainHeader"]');
        mainHeaders.forEach((el: HTMLElement) => {
          // Only hide if it's a site header, not a list header
          const isListHeader = el.closest('[class*="ms-List"], [class*="ms-DetailsList"], [class*="ms-Table"]');
          const hasSiteHeaderAttr = el.getAttribute('data-automationid') === 'SiteHeader' || 
                                    el.getAttribute('data-automation-id') === 'SiteHeader' ||
                                    el.getAttribute('aria-label') === 'SharePoint Site Header';
          
          if (!isListHeader && (hasSiteHeaderAttr || el.closest('[data-automationid="SiteHeader"]'))) {
            el.style.display = 'none';
            el.style.visibility = 'hidden';
            el.style.height = '0';
            el.style.overflow = 'hidden';
          }
        });
      } catch (e) {
        // Ignore selector errors
      }
    }
    
    private hideElementIfHeader(element: HTMLElement): void {
      if (!element || !element.getAttribute) {
        return;
      }
      
      // Check if it's a header element (but not a list header)
      const automationId = element.getAttribute('data-automationid');
      const automationIdHyphen = element.getAttribute('data-automation-id');
      const ariaLabel = element.getAttribute('aria-label');
      const className = element.className ? String(element.className) : '';
      
      // Check if it's inside a list (exclude list headers)
      const isListHeader = element.closest && element.closest('[class*="ms-List"], [class*="ms-DetailsList"], [class*="ms-Table"], [data-automation-id="List"], [data-automation-id="ListView"]');
      
      if (
        !isListHeader && (
          automationId === 'SiteHeader' ||
          automationIdHyphen === 'SiteHeader' ||
          (className && (
            (className.indexOf('ms-HubNav') !== -1 ||
            className.indexOf('hubNavRow') !== -1) ||
            ((className.indexOf('headerRow') !== -1 || className.indexOf('mainHeader') !== -1) &&
             (automationId === 'SiteHeader' || automationIdHyphen === 'SiteHeader' || 
              element.closest && element.closest('[data-automationid="SiteHeader"]')))
          )) ||
          (ariaLabel && (
            ariaLabel.indexOf('SharePoint Site Header') !== -1 ||
            ariaLabel.indexOf('BMRHUB') !== -1
          ))
        )
      ) {
        // Apply inline styles immediately for instant hiding
        element.style.cssText = 'display: none !important; visibility: hidden !important; height: 0 !important; overflow: hidden !important;';
      }
    }
    
    private observeAndHide(): void {
      // Immediately check and hide any existing elements
      this.applyInlineStylesImmediately();
      
      // Create MutationObserver to hide elements as they appear
      const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === 1) { // Element node
              const element = node as HTMLElement;
              
              // Immediately apply inline styles to hide
              this.hideElementIfHeader(element);
              
              // Also check child elements immediately
              const headerElements = element.querySelectorAll ? element.querySelectorAll('[data-automationid="SiteHeader"], [data-automation-id="SiteHeader"], [class*="ms-HubNav"], [class*="hubNavRow"], [class*="headerRow"], [class*="mainHeader"]') : [];
              headerElements.forEach((el: HTMLElement) => {
                this.hideElementIfHeader(el);
              });
            }
          });
        });
      });
      
      // Start observing immediately - use documentElement if body doesn't exist yet
      const targetNode = document.body || document.documentElement;
      observer.observe(targetNode, {
        childList: true,
        subtree: true,
        attributes: false,
        attributeOldValue: false
      });
      
      // If body doesn't exist yet, also observe documentElement
      if (!document.body) {
        const bodyObserver = new MutationObserver((mutations) => {
          if (document.body) {
            observer.observe(document.body, {
              childList: true,
              subtree: true
            });
            bodyObserver.disconnect();
          }
        });
        bodyObserver.observe(document.documentElement, { childList: true });
      }
      
      // Also periodically check for headers (as a backup)
      const intervalId = setInterval(() => {
        this.applyInlineStylesImmediately();
      }, 50); // Check every 50ms
      
      // Stop checking after 5 seconds (headers should be loaded by then)
      setTimeout(() => {
        clearInterval(intervalId);
      }, 5000);
    }
    
    private getHeaderHidingCSS(): string {
      return `
        /* Hide top dark grey header bar */
        #SuiteNavPlaceHolder,
        #O365_NavHeader,
        div[class*="o365cs-base"],
        div[id*="SuiteNav"],
        header[class*="o365cs-base"],
        div[class*="ms-SuiteNav"],
        div[class*="SuiteNav"],
        div[id="O365_NavHeader"],
        div[class*="o365cs-nav"],
        div[class*="o365cs-topbar"],
        div[class*="o365cs-navbar"],
        div[data-sp-feature-name="SuiteNav"],
        div[data-sp-feature-name="O365_NavHeader"] {
          display: none !important;
          visibility: hidden !important;
          height: 0 !important;
          overflow: hidden !important;
        }
    
        /* Hide Hub Navigation bar (dark blue bar with "BMRHUB BI-Projects") */
        div[class*="root"][class*="ms-HubNav"],
        div[class*="ms-HubNav"],
        div[class*="hubNavRow"],
        div[class*="sp-App-hubNav"],
        div[role="navigation"][aria-label*="hub"],
        div[aria-label*="BMRHUB"],
        div[aria-label*="hub site"],
        div[class*="ms-HorizontalNav"],
        div[id="HubNavTitle"],
        a[class*="ms-HubNav-nameLink"],
        div[class*="ms-HubNavItems"] {
          display: none !important;
          visibility: hidden !important;
          height: 0 !important;
          overflow: hidden !important;
        }
    
        /* Hide site header with logo and title (light grey header bar) */
        /* Target the exact element from DevTools - data-automationid (without hyphen) */
        div[data-automationid="SiteHeader"],
        div[data-automation-id="SiteHeader"],
        div[data-automation-id="Header"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]):not([class*="ms-Table"]),
        /* Target parent headerRow and mainHeader classes - but exclude list headers */
        div[class*="headerRow"]:not([class*="ms-List-headerRow"]):not([class*="ms-DetailsList-headerRow"]):not([class*="ms-Table-headerRow"]),
        div[class*="mainHeader"]:not([class*="ms-List-mainHeader"]):not([class*="ms-DetailsList-mainHeader"]),
        div[aria-label="SharePoint Site Header"],
        div[aria-label*="SharePoint Site Header"],
        /* Other header selectors - exclude list headers */
        header[class*="ms-SiteHeader"]:not([class*="ms-List"]),
        div[class*="SiteHeader"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]),
        div[class*="ms-SiteHeader"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]),
        div[class*="od-SiteHeader"]:not([class*="ms-List"]),
        div[class*="od-Header"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]),
        div[class*="sp-siteHeader"]:not([class*="ms-List"]),
        div[class*="sp-header"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]),
        div[role="banner"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]),
        header[role="banner"]:not([class*="ms-List"]),
        div[role="region"][aria-label*="Header"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]),
        div[class*="headerBar"]:not([class*="ms-List"]):not([class*="ms-DetailsList"]),
        div[class*="siteHeaderBar"]:not([class*="ms-List"]) {
          display: none !important;
          visibility: hidden !important;
          height: 0 !important;
          overflow: hidden !important;
        }
        
        /* Exclude list headers - ensure they remain visible */
        div[class*="ms-List-headerRow"],
        div[class*="ms-DetailsList-headerRow"],
        div[class*="ms-Table-headerRow"],
        div[data-automation-id="List"] [class*="headerRow"],
        div[data-automation-id="ListView"] [class*="headerRow"],
        div[class*="ms-List"] [class*="headerRow"],
        div[class*="ms-DetailsList"] [class*="headerRow"] {
          display: table-row !important;
          visibility: visible !important;
          height: auto !important;
          overflow: visible !important;
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
        
        /* Adjust body margin to remove space from hidden header */
        body {
          margin-top: 0 !important;
          padding-top: 0 !important;
        }
        
        /* Hide any top bar with dark grey background */
        div[style*="background-color"][style*="rgb"][style*="grey"],
        div[style*="background-color"][style*="#"][style*="grey"],
        div[class*="topbar"],
        div[class*="navbar"][class*="top"] {
          display: none !important;
        }
      `;
    }
    
}
