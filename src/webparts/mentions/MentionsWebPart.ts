import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

// import styles from './Mentions.module.scss';
import * as strings from 'mystrings';
import { IMentionsWebPartProps } from './IMentionsWebPartProps';
import * as jQuery from 'jquery';
import moduleLoader from '@microsoft/sp-module-loader';

const tenantName: string = "kermamoottori";
const clientId: string = "4d1421f5-f3d2-46da-b104-71819ec08191";
const sharepointApi: string = `https://${tenantName}.sharepoint.com`;
const graphApi: string = "https://graph.microsoft.com";

const authConfig: IAuthenticationConfig = {
	tenant: `${tenantName}.onmicrosoft.com`,
	clientId: `${clientId}`,

	/** where to navigate to after AD logs you out */
	postLogoutRedirectUri: window.location.href,

	/** redirect_uri page, this is the page that receives access tokens
	 *  this URL must match, at least, the scheme and origin of at least 1 of
	 *  the Reply URLs entered on your Azure AD Application configuration page
	 */
	redirectUri: `${window.location.href}`,
	endpoints: {}
	// cacheLocation: "localStorage", // enable this for IE, as sessionStorage does not work for localhost.
};
authConfig.endpoints[sharepointApi] = `https://${tenantName}.sharepoint.com/search`;
authConfig.endpoints[graphApi] = "https://graph.microsoft.com";

export default class MentionsWebPart extends BaseClientSideWebPart<IMentionsWebPartProps> {

  private _userdocs: Array<any>;
  private _tribute: any;


  public constructor(context: IWebPartContext) {
    moduleLoader.loadCss("http://zurb.com/playground/uploads/upload/upload/430/tribute.css");
    super(context);
  }

  public render(): void {

    this.domElement.innerHTML = `
    <div id="signinBtn">
    <a href="javascript:;" class="ms-Button"><span class="ms-Button-label">Login to AAD</span></a>
    </div>
    <div id="signoutBtn">
      <a href="javascript:;" class="ms-Button"><span class="ms-Button-label">Logout</span></a>
    </div>
    <!--<a href="javascript:;" id="myDocuments" class="ms-Button"><span class="ms-Button-label">List my documents</span></a>-->
    <ul id="documentList">

    </ul>
    <div class="ms-TextField"  id="tribute-mentions" contenteditable="true">
    </div>
    `;
    this.initTribute();
    this.manageAuthentication();

    var editor = document.getElementById('tribute-mentions');
    this.load(editor);
    editor.onblur = (e) => {
        this.save(editor);
      };
  }

  private save(editor: any): void {
    this.properties.description = editor.innerHTML;
    this.load(editor);
  }

  private load(editor: any): void {
    // if (this.properties.mentionscontent != undefined) {
      editor.innerHTML = this.properties.description;
    // }
  }

  private initTribute(): void {
    moduleLoader.loadScript('http://zurb.com/playground/uploads/upload/upload/435/tribute.js', 'Tribute').then((t: any): void => {
      this._tribute = t;
	    this._userdocs = [];

      var tribute: any = new this._tribute({
        collection: [
          {
            // symbol that starts the lookup
            trigger: '#',

            // function called on select that returns the content to insert
            selectTemplate: (item): string => {
              // return '#' + item.original.value;
              return `#<a href='${item.original.key}' target='_blank'>${item.original.value}</a>`;
            },

            // template for displaying item in menu
            menuItemTemplate: (item): string => {
              return item.string;
            },

            // column to search against in the object (accepts function or string)
            lookup: 'value',

            // column that contains the content to insert by default
            fillAttr: 'value',

            // REQUIRED: array of objects to match
            values: this._userdocs
          }
        ]
      });

      // tribute.attach(document.getElementById('ql-editor-1'));
      // tribute.attach(document.getElementsByClassName('ql-editor'));
      tribute.attach(document.getElementById('tribute-mentions'));
    });
  }

  private getDocumentsForTribute(authContext: AuthenticationContext): void {
    const graphApi: string = "https://graph.microsoft.com";

    moduleLoader.loadScript('https://code.jquery.com/jquery-2.1.1.min.js', 'jQuery').then(($: any): void => {
      const self: any = this;
      // var $documentListElement = jQuery("#documentList");
      const d: any = jQuery.Deferred<any>();

      authContext.acquireToken(graphApi, (error: string, token: string) => {
        if (error || !token) {
          const msg: any = `ADAL error occurred: ${error}`;
          d.rejectWith(this, [msg]);
          return;
        }

      jQuery.ajax({
        type: "GET",
        url: `${graphApi}/v1.0/me/drive/root/children`,
        headers: {
          "Accept": "application/json;odata.metadata=minimal",
          "Authorization": `Bearer ${token}`
        }
        }).done((response: { value: any[] }) => {
          console.log("Successfully fetched documents from O365.");
          d.resolveWith(self);
          var documents: any = response.value;
          for (var document of documents) {
            // var documentLinkElement = `<li><a href='${document.webUrl}' target='_blank'>${document.name}</a></li>`;
            // $documentListElement.append(documentLinkElement);
            this._userdocs.push({key: document.webUrl, value: document.name});
            console.log("document " + document.name + " loaded");
          }

          console.log(response.value);
        }).fail((xhr: JQueryXHR) => {
          const msg: any = `Fetching messages from Office365 failed. ${xhr.status}: ${xhr.statusText}`;
          console.log(msg);
          d.rejectWith(self, [msg]);
        });
      });
    });
  }

  private manageAuthentication(): void {
    moduleLoader.loadScript('https://code.jquery.com/jquery-2.1.1.min.js', 'jQuery').then(($: any): void => {
      const self: any = this;
      const $signInButton: any = jQuery("#signinBtn");
      const $signOutButton: any = jQuery("#signoutBtn");
      // let $myDocumentsButton = jQuery("#myDocuments");
      // $myDocumentsButton.hide();
      moduleLoader.loadScript('//secure.aadcdn.microsoftonline-p.com/lib/1.0.0/js/adal.min.js', 'AuthenticationContext').then((authenticationContext: any): void => {

        const authContext: AuthenticationContext = new AuthenticationContext(authConfig);

        const isCallback: any = authContext.isCallback(window.location.hash);
        if (isCallback) {
          const loginReq: any = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
          console.log(`IS Callback! ${loginReq}`);
          authContext.handleWindowCallback();
          return;
        }
        console.log("Is NOT Callback!");

        /** check login status */
        var user: any = authContext.getCachedUser();

        if (user) {
          console.log(`User is logged-in: ${JSON.stringify(user)}`);
          $signInButton.hide();
          // $myDocumentsButton.show();
          $signOutButton.show();
          self.getDocumentsForTribute(authContext);
          $signOutButton.click(() => {
            authContext.logOut();
          });
        } else {
          console.log("User is NOT logged-in!!");
          $signOutButton.hide();
          $signInButton.show();
          $signInButton.click(() => {
            authContext.login();
          });
        }
      });
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
