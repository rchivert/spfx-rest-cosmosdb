import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'Rest2WebPartStrings';
import Rest2 from './components/Rest2';
import { IRest2Props } from './components/IRest2Props';

import { HttpClientResponse, IHttpClientOptions, HttpClient, AadHttpClient  } from '@microsoft/sp-http';
import { setup as pnpSetup } from "@pnp/common";
import * as datefns from 'date-fns';

export interface IRest2WebPartProps {
  description: string;
}

interface IWebApp {
  webapp_uri: string;
  webapp_appid: string;
}

enum enumStorageType {
  cosmos,
  userprofile
}

enum enumWebAppLocation {
  eastus,
  japaneast
}

interface ITiming {
  storageType : enumStorageType;
  webAppLocation : enumWebAppLocation | null;
  duration_get: number;
  duration_post: number;
}



export default class Rest2WebPart extends BaseClientSideWebPart<IRest2WebPartProps> {

private myPromise: Promise<any>;
private webapp : IWebApp;


private async getInitialTimings () : Promise<any>
{
  //  Get the most-performant (closest) Azure Function WebApp
  var startedAt = Date.now();
  this.webapp = await this.getClosestWebApp();
  var endedAt = Date.now();
  var elapsed = datefns.differenceInMilliseconds(endedAt, startedAt);
  console.log ("milliseconds for getClosestWebApp = " + elapsed.toString());

  //  Post data to Cosmos DB, and get data from Cosmos DB
  //
  const requestHeaders: Headers = new Headers();
  requestHeaders.append('Content-type', 'application/json');
  requestHeaders.append('Cache-Control', 'no-cache');

    // create an AadHttpClient
    const aadClient: AadHttpClient = await this.context.aadHttpClientFactory.getClient(this.webapp.webapp_appid);
    console.log("Created aadClient");

    let testdata = {
                    "id":"user2@domain.com",
                    "links":[
                      {
                        "title":"title1x",
                        "url":"url1"
                      },
                      {
                        "title":"title2",
                        "url":"url2"
                      }
                      ],
                    "location": "US-EAST"
                  };

    const requestOptions: IHttpClientOptions =  {
                                                headers: requestHeaders,
                                                body:   JSON.stringify(testdata)
                                                };

    console.log("posting data to cosmos db...");
    startedAt = Date.now();
    let clientPostResponse: HttpClientResponse = await aadClient.post(this.webapp.webapp_uri + '/preferences', AadHttpClient.configurations.v1, requestOptions);
    endedAt = Date.now();
    elapsed = datefns.differenceInMilliseconds(endedAt, startedAt);
    console.log ("milliseconds for posting data to cosmos db = " + elapsed.toString());
    console.log("aadClient post response = " + clientPostResponse.status);

    console.log("getting data from cosmos db...");
    startedAt = Date.now();
    var clientGetResponse : HttpClientResponse = await aadClient.get (this.webapp.webapp_uri + '/preferences/user2@domain.com', AadHttpClient.configurations.v1);
    var txt : string = await clientGetResponse.text();
    var o = txt ? JSON.parse(txt) : [{}];
    endedAt = Date.now();
    elapsed = datefns.differenceInMilliseconds(endedAt, startedAt);
    console.log ("milliseconds for getting data from cosmos db = " + elapsed.toString());

    console.log("response from GET = " + JSON.stringify(o[0]));

    return new Promise<void>(resolve => {
      console.log ("resolving getInitialTimings");
      resolve();
    });
}

private async getClosestWebApp () : Promise<IWebApp>
  {
    //  Get the URL of the most-performant (closest) Azure Function WebApp
    //
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    requestHeaders.append('Cache-Control', 'no-cache');

    const httpClientOptions: IHttpClientOptions = {
      headers: requestHeaders
    };

    let response: HttpClientResponse = await this.context.httpClient.get("https://chiverton365-preferences.trafficmanager.net/hello", HttpClient.configurations.v1, httpClientOptions);
    let txt = await response.text();
    var s = txt ? JSON.parse(txt) : {appid: "", appuri: ""};

    return new Promise<IWebApp> (resolve => {
      console.log("response from traffic mgr 'hello' = " + JSON.stringify(s));
      let wa : IWebApp = {"webapp_appid" : s.appid, "webapp_uri": s.url};
      resolve(wa);
    });
  }

public onInit(): Promise<any> {

  return super.onInit().then(_x => {

    // other init code may be present

    pnpSetup({
      spfxContext: this.context
    });

    //  Use custom myPromise to control when render gets called (https://sharepoint.stackexchange.com/questions/222515/sharepoint-framework-spfx-oninit-promises/222627)
    //
    this.myPromise = this.getInitialTimings();
    });
}

  public render(): void {
    this.myPromise.then (() =>{
      console.log("render");
      const element: React.ReactElement<IRest2Props > = React.createElement(
        Rest2,
        {
          description: this.properties.description,
          ctx: this.context,
          webapp_uri: this.webapp.webapp_uri,
          webapp_appid: this.webapp.webapp_appid
        }
      );

      ReactDom.render(element, this.domElement);

    }).catch(e => {
      console.log(e);
      });
    }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
