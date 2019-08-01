import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import styles from './TwitterWebpartWebPart.module.scss';
import * as strings from 'TwitterWebpartWebPartStrings';

export interface ITwitterWebpartWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  TweetID: string;
  ProfileImageUrl: string;
  Username: string;
}

const logo: any = require('./assets/Twitter.png');

export default class TwitterWebpartWebPart extends BaseClientSideWebPart<ITwitterWebpartWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.twitterWebpart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <span class="${ styles.title }">Tweets</span>
            <div id="spListContainer"/>
          </div>
        </div>
      </div>`;

      this._renderListAsync();
      
  }

  private _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('TweetsVersion1')/Items?$top=5&$orderby=Created`, SPHttpClient.configurations.v1)  
        .then((response: SPHttpClientResponse) => {    
          return response.json();  
        });  
    }  

    private _renderListAsync(): void {  
      this._getListData()
      .then((response) => {
        this._renderList(response.value);
      }); 
  }  

  private _renderList(items: ISPList[]): void {  
    console.log(items);
    let html: string = '<table class="TFtable" border=0 width=100% style="border-collapse: collapse;">';  
    html += ``;  
    items.forEach((item: ISPList) => {  
      html += `  
           <tr>  
              <td>
                  <img class="${ styles.image }" src=${item.ProfileImageUrl} alt="ProfileImage">                
              </td> 
              <td>
                <div><b>${item.Username}</b></div>
                <div>${item.Title}</div>
              </td>  
              <td>
                <a href="https://twitter.com/user/status/${item.TweetID}" target="_blank">
                  <img src=${logo} height="25px" alt="View in Twitter">
                </a>
              </td>  
          </tr>  
          `;  
    });  
    html += `</table>`;  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;  
  }   


     


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return;
  }
}
