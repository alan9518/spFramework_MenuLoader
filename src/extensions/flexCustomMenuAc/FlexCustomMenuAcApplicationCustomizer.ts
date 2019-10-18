import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
// import { Dialog } from '@microsoft/sp-dialog';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

import * as strings from 'FlexCustomMenuAcApplicationCustomizerStrings';

const LOG_SOURCE: string = 'FlexCustomMenuAcApplicationCustomizer';

// Custom Imports

  import * as $ from 'jquery';
  import 'jqueryui';

  require('sp-init');
  require('microsoft-ajax');
  require('sp-runtime');
  require('sharepoint');

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFlexCustomMenuAcApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class FlexCustomMenuAcApplicationCustomizer
  extends BaseApplicationCustomizer<IFlexCustomMenuAcApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    let $relativeURL = window.location.host + this.context.pageContext.web.serverRelativeUrl;
    let $subSiteURL = this.context.pageContext.web.absoluteUrl;
    localStorage.setItem('hostURL', $relativeURL);
    localStorage.setItem('subSiteURL',$subSiteURL);
    console.log('Init', this.context.pageContext.web.serverRelativeUrl);
   
    console.log(this.context.pageContext.web.absoluteUrl);
    console.log( localStorage.getItem('subSiteUrl'));

    this.onRender();
    return Promise.resolve<void>();
  }


  // ------------------------
  // This Method gets Called 
  // When the Page is Rendering
  // Append Custom JS
  // ------------------------
  @override
  public onRender():void
  {
    let $menu = $('#SuiteNavPlaceHolder');     
    let $head: any = document.getElementsByTagName("head")[0] || document.documentElement;
    //  Create Scripts
   

        let $flexMenuScript : any = document.createElement('script');
            $flexMenuScript.type = "text/javascript";
            $flexMenuScript.src = "https://stgflextronics365.sharepoint.com/sites/appcatalog/CDN/FlexMenuProd/scripts/flexMenu.js";

     
          $head.insertBefore($flexMenuScript, $head.firstChild);

        
  }


  
}


// ?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"c699c10a-73af-4212-9d26-00f10ae53d67":{"location":"ClientSideExtension.ApplicationCustomizer"}}
// ?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"c699c10a-73af-4212-9d26-00f10ae53d67":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"testMessage":"Hello as property!"}}}