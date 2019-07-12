import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SharepointDataWebPart.module.scss';
import * as strings from 'SharepointDataWebPartStrings';
import MockHttpClient from './MockHttpClient';   

export interface ISharepointDataWebPartProps {
  description: string;
}

export interface ISPLists {  
  value: ISPList[];  
}  
export interface ISPList {  
  Title: string;  
}    

import {  
  SPHttpClient  ,SPHttpClientResponse, HttpClient, HttpClientResponse
} from '@microsoft/sp-http';

import {  
  Environment,  
  EnvironmentType  
} from '@microsoft/sp-core-library'; 

export default class SharepointDataWebPart extends BaseClientSideWebPart<ISharepointDataWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.sharepointData}">  
    <div class="${styles.container}">  
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
          <p class="ms-font-l ms-fontColor-white" style="text-align: center">Demo : Retrieve Data from SharePoint List</p>  
        </div>  
      </div>  
      <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}"> 
      <br>  
      <div id="spListContainer" />  
      </div>  
    </div>  
  </div>`;
  //this._renderListAsync(); 
  this._getRestData(); 
  }

  private _getRestData() {
    this.context.httpClient.get("https://jsonplaceholder.typicode.com/todos/", HttpClient.configurations.v1)
    .then((res: HttpClientResponse): Promise<any> => {
      return res.json();
    })
    .then((data: any): void => {
      console.log(data);
      let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';html += `<th>Title</th>`; 

      data.forEach((item) => {
        html += `<tr><td>${item.title}</td></tr>`; 
      });

      html += `</table>`;  
      const listContainer: Element = this.domElement.querySelector('#spListContainer');  
      listContainer.innerHTML = html;
    }, (err: any): void => {
      console.log(err);
      // handle error here
    });
  }

private _getListData(): Promise<ISPLists> {  
  return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('SPFXList')/Items`, SPHttpClient.configurations.v1)  
      .then((response:any) => {   
        debugger;  
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
  let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';  
  html += `<th>Title</th>`;  
  items.forEach((item: ISPList) => {  
    html += `  
         <tr>  
          <td>${item.Title}</td>  
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
