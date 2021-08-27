import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HideIconsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'HideIconsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHideIconsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  CSSFileLocation: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HideIconsApplicationCustomizer
  extends BaseApplicationCustomizer<IHideIconsApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let fileURL:string=this.context.pageContext.site.serverRelativeUrl + this.properties.CSSFileLocation

    if(fileURL){
 
   const head:any=document.getElementsByName("head")[0] || document.documentElement;
    let customStyle:HTMLLinkElement =document.createElement("link");
    customStyle.href=fileURL;
    customStyle.rel="stylesheet";
    customStyle.type="text/css";
    head.insertAdjacentElement("beforeEnd",customStyle)
   
   }

    

    return Promise.resolve();
  }
}
