import { override } from '@microsoft/decorators';
import { Log, SPEventArgs } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,PlaceholderContent,PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import {SPPermission} from '@microsoft/sp-page-context';
import * as strings from 'HideApplicationCustomizerStrings';
import * as React from "react";
import * as ReactDom from "react-dom";

const LOG_SOURCE: string = 'HideApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHideApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

interface NavigationEventDetails extends Window {
  isNavigatedEventSubscribed: boolean;
  currentPage: string;
  currentHubSiteId: string;
  currentUICultureName: string;
}

declare const window: NavigationEventDetails;

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HideApplicationCustomizer
  extends BaseApplicationCustomizer<IHideApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    //console.log(`LCEVENT:onInit=${window.location.href}`);

    //this.context.placeholderProvider.changedEvent.add(this, this.Hide);

    //if (!(window as any).isNavigatedEventSubscribed) {

      //this.context.application.navigatedEvent.add(this, this.logNavigatedEvent);

      //(window as any).isNavigatedEventSubscribed = true;
    //}
    //this.Hide();
    this.render();

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

    return Promise.resolve();
  }

  onPlaceholdersChanged
  @override
  public onDispose(): Promise<void>{

    //console.log(`LCEVENT:onDispose=${window.location.href}`);

    //this.context.application.navigatedEvent.remove(this, this.logNavigatedEvent);

    //(window as any).isNavigatedEventSubscribed = false;
    //(window as any).currentPage = '';
    
    this.context.application.navigatedEvent.remove(this, this.render);
    
    window.isNavigatedEventSubscribed = false;
    window.currentPage = '';
    window.currentHubSiteId = '';
    window.currentUICultureName = '';

    return Promise.resolve();
  }

  private Hide()
  {
    alert('Hide');
    const canEdit = this.context.pageContext.web.permissions.hasAnyPermissions(SPPermission.manageWeb);
    if(!canEdit)
    {
      let checkExist = setInterval(()=> 
      {
        const setting = document.querySelector("#O365_MainLink_Settings_container").firstChild.firstChild.firstChild.firstChild;
        if(typeof(setting) != "undefined" && setting != null)
        {
          setting.parentElement.parentElement.parentElement.remove();
          console.log("User only has view rights.");
          clearInterval(checkExist);
        }
      
      }, 100);

      let checkExist2 = setInterval(()=> 
      {
        const recycle = document.querySelectorAll("a.ms-Nav-link");
        var str = "Recycle bin";
        var pos = 0;
        var i = 0;
        for(i = 0; i < recycle.length; i++ )
        {
          pos = recycle[i].innerHTML.indexOf(str);
          if(pos > -1)
          {
            if(typeof(recycle) != "undefined" && recycle != null)
            {
              recycle[i].parentElement.remove();
              console.log("User only has view rights.");
              clearInterval(checkExist2);
            }
          }
        }
      
      }, 100);

      let checkExist3 = setInterval(()=> 
      {
        const des = document.querySelectorAll("div.ms-TooltipHost");
        var str = "Description";
        var pos1 = 0;
        var i = 0;
        for(i = 0; i < des.length; i++ )
        {
          pos1 = des[i].innerHTML.indexOf(str);
          if(pos1 > -1)
          {
            if(typeof(des) != "undefined" && des != null)
            {
              const settings= des[i].firstChild;
              //alert(settings.parentElement.parentElement);
              if(typeof(settings) != "undefined" && settings != null){
                settings.parentElement.remove();
              }
              console.log("User only has view rights.");
              clearInterval(checkExist3);
            }
          }
        }
      
      }, 100);

      let checkExist4 = setInterval(()=> 
      {
        const val = document.querySelectorAll("div.ms-DetailsRow-cell");
        var str = "Description";
        var attr = '';
        var i = 0;
        for(i = 0; i < val.length; i++ )
        {
          attr = val[i].getAttribute("data-automation-key");
          if (attr == str)
          {
            //alert(val[i])
            val[i].remove();
            console.log("Description Removed.");
          }
          
          clearInterval(checkExist4);
        }
      
      }, 100);
    }
  }

  private change(eventArgs: any): void
  {
    alert('change');
    this.context.placeholderProvider.changedEvent.add(this, this.Hide.bind(this));
  }
  
  public logNavigatedEvent(args: SPEventArgs): void {

    setTimeout(() => {

      if ((window as any).currentPage !== window.location.href) {

        // REGISTER PAGE VIEW HERE >>>
        console.log(`LCEVENT:navigatedEvent=${window.location.href}`);
        this.Hide();
        (window as any).currentPage = window.location.href;
      }
    }, 3000);
  }

  private render() {
    window.currentPage = window.location.href;
    //window.currentHubSiteId = HubSiteService.getHubSiteId(); // Your custom logic to retrieve the hub site id
    window.currentUICultureName = this.context.pageContext.cultureInfo.currentUICultureName;

    if (!window.isNavigatedEventSubscribed) {
      this.context.application.navigatedEvent.add(this, this.navigationEventHandler);
      window.isNavigatedEventSubscribed = true;
      this.Hide();
    }
  }

  private navigationEventHandler(args: SPEventArgs): void {
    setTimeout(() => {
      if (window.currentHubSiteId !== '') {
        this.onDispose();
        this.onInit();
        return;
      }

      if (window.currentUICultureName !== this.context.pageContext.cultureInfo.currentUICultureName) {
        // Trigger a full page refresh to be sure to have to correct language loaded
        location.reload();
        return;
      }

      // Page URL check
      if (window.currentPage !== window.location.href) {
        window.currentPage = window.location.href;
        this.Hide();
      }
    }, 50);
  }
}
