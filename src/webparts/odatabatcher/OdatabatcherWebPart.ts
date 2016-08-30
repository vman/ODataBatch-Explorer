import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import styles from './Odatabatcher.module.scss';
import * as strings from 'mystrings';
import { IOdatabatcherWebPartProps } from './IOdatabatcherWebPartProps';

import { ServiceScope, IODataBatchOptions, ODataBatch } from '@microsoft/sp-client-base';

export default class OdatabatcherWebPart extends BaseClientSideWebPart<IOdatabatcherWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.odatabatcher}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
              <a href="https://github.com/SharePoint/sp-dev-docs/wiki" class="ms-Button ${styles.button}">
                <span class="ms-Button-label">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

      this._makeODataBatchRequest();
  }

  private _makeODataBatchRequest(): void {

    // More about Service scopes here: https://sharepoint.github.io/classes/_sp_client_base_.servicescope.html
    const serviceScope: ServiceScope = ServiceScope.startNewRoot();
    serviceScope.finish();

    // Here, 'this' refers to the SPFx BaseClientSideWebPart class. Since I am calling this method from inside the class, I have access to the pageContext.
    const webAbsoluteUrl: string = this.context.pageContext.web.absoluteUrl;

    const batchOpts: IODataBatchOptions = { webUrl: webAbsoluteUrl };

    const odataBatch: ODataBatch = new ODataBatch(serviceScope, batchOpts);

    // Queue a request to get current user's userprofile properties
    const getMyUserProps: Promise<Response> = odataBatch.get(`${webAbsoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`);

    // Queue a request to get all the list titles in the current web
    const getLists: Promise<Response> = odataBatch.get(`${webAbsoluteUrl}/_api/web/lists?$select=Title`);

    // Queue a request to create a list in the current web.
    // POST requests do not work yet due to a bug in the framework : https://github.com/SharePoint/sp-dev-docs/issues/107
    // const postBody: Object = { __metadata: { type: 'SP.List' }, Title: "Developer workbench", BaseTemplate: 100 };
    // const reqHeaders: Headers = new Headers();
    // reqHeaders.append('odata-version', '3.0');
    // const createList: Promise<Response> = odataBatch.post(`${webAbsoluteUrl}/_api/web/lists`,
    // {
    //   body: JSON.stringify(postBody),
    //   headers: reqHeaders
    // });

    // Make the batch request
    odataBatch.execute().then(() => {

      getMyUserProps.then((response: Response) => {
        response.json().then((responseJSON) => {
          console.log(responseJSON);
        });
      });

      getLists.then((response: Response) => {
        response.json().then((responseJSON) => {
          console.log(responseJSON);
        });
      });

      // createList.then((response: Response) => {
      //   response.json().then((responseJSON) => {
      //     console.log(responseJSON);
      //   });
      // });
    });
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
