import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ThankGodWebPart.module.scss';
import * as strings from 'ThankGodWebPartStrings';

export interface IThankGodWebPartProps {
  description: string;
}

import {  
  SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http'; 

export interface ISPLists {
  value: ISPList[];
 }
 
 export interface ISPList {
  Title: string;
  Id: string;
  Name: string;
 }
 
export default class ThankGodWebPartWebPart extends BaseClientSideWebPart<IThankGodWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
            
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col" style="width:100%">
              <span><h3>Employee Details</h3></span>
            </div>
            <div class="ms-Grid-col" style="width:100%">
              <span>Title : </span>
            </div>
            <div class="ms-Grid-col" style="width:100%">
              <input type="text" id="Title" style="width:100%">
            </div>
            <div class="ms-Grid-col" style="width:100%">
              <span>Name : </span>
            </div>
            <div class="ms-Grid-col" style="width:100%">
            <input type="text" id="Name" style="width:100%">
            
            </div>
            
            <div class="ms-Grid-col" style="width:100%;text-align:center">
            <br />
              <button type="button" id="btn_add" style="background-color:#009688;color:white;font-weight:bold">Insert Item</button>
              <button type="button" id="btn_update" style="background-color:#009688;color:white;font-weight:bold">Update Item</button>
              <button type="button" id="btn_delete" style="background-color:#009688;color:white;font-weight:bold">Delete Item</button>
            <br />
            </div>
          </div>
        </div>
        <div class="status" style="color:red"></div>
      </div>
      <div >
      <br />
      <br />
      
      <div class="${ styles.container }">
      <div id="spListContainer" />
    </div>

    `;
        
        this._renderListAsync();
        const events: ThankGodWebPartWebPart = this;
        var button = document.querySelector('#btn_add');
        button.addEventListener('click', () => { events.CreateNewItem(); });

        var button = document.querySelector('#btn_update');
        button.addEventListener('click', () => { events.updateItem(); });

        var button = document.querySelector('#btn_delete');
        button.addEventListener('click', () => { events.deleteItem(); });
  }

  private _getListCustomerData(): Promise<ISPLists> {  
    // return this.context.spHttpClient.get(`https://genpactonline.sharepoint.com/sites/ESSDev/_api/web/lists/GetByTitle('DemoList')/items`,SPHttpClient.configurations.v1).
    //    then((responseListCustomer:SPHttpClientResponse) => {  
    //    debugger;
    //    return responseListCustomer.json();
    // });
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl +
       `/_api/web/lists/GetByTitle('DemoList')/items`,SPHttpClient.configurations.v1).
       then((responseListCustomer:SPHttpClientResponse) => {  
       debugger;
       return responseListCustomer.json();
    }); 
}  

private _renderList(items: ISPList[]): void {  
  let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';  
  html += `<th class="ss">EmployeeId</th><th>Title</th><th>EmployeeName</th>`;  
  debugger;
  items.forEach((item: ISPList) => {  
    html += `  
         <tr>  
        <td>${item.Id}</td>  
        <td>${item.Title}</td>
        <td>${item.Name}</td>
         
        </tr>  
        `;  
  });  
  html += `</table>`;  
  const listContainer: Element = this.domElement.querySelector('#spListContainer');  
  listContainer.innerHTML = html;  

}   

private _renderListAsync(): void
{  
   debugger;
    this._getListCustomerData().then((response) => {  
    this._renderList(response.value);  
  });
}  

private CreateNewItem(): void 
{
  this.usermessage('Creating list Item ...');
  let title = (<HTMLInputElement>document.getElementById("Title")).value.trim();
  let name = (<HTMLInputElement>document.getElementById("Name")).value.trim();
  if(title != '' && name != '')
  {
    //Create a array object with all column values
      let requestdata = {};
      requestdata['Title'] = title;
      requestdata['Name'] = name;
      this.usermessage('Creating list Item ...' + requestdata);
      this.addListItem('DemoList',requestdata);
  }
  else
  {
      if(title == '' && name == '')
      {this.usermessage('Please enter title and name');}
      else if(title == '')
      {this.usermessage('Please enter title');}
      else if(name == '')
      {this.usermessage('Please enter name');}
  }
}

private addListItem(listname:string,requestdata:{})
{   
    let requestdatastr = JSON.stringify(requestdata);
    requestdatastr = requestdatastr.substring(1, requestdatastr .length-1);
    console.log(requestdatastr);
    let requestlistItem: string = JSON.stringify({
      '__metadata': {'type': this.getListItemType("DemoList")}
    });
    requestlistItem = requestlistItem.substring(1, requestlistItem .length-1);
    requestlistItem = '{' + requestlistItem + ',' + requestdatastr + '}';
    console.log(requestlistItem);
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listname}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: requestlistItem
        })
        .then((response: SPHttpClientResponse): Promise<ISPList> => {
          console.log('response.json()');
        return response.json();
      })
       .then((item: ISPList): void => {
          console.log('Creation');
          this.usermessage(`List Item created successfully... '(Item Id: ${item.Id})`);
    }, (error: any): void => {
        this.usermessage('List Item Creation Error...');
      });      
      // reload table 
      this._renderListAsync();
}
private usermessage(status: string): void {
  this.domElement.querySelector('.status').innerHTML = status;
}
private getListItemType(name: string) {
  let safeListType = "SP.Data." + name[0].toUpperCase() + name.substring(1) + "ListItem";
  safeListType = safeListType.replace(/_/g,"_x005f_");
  safeListType = safeListType.replace(/ /g,"_x0020_");
  return safeListType;
}

private updateItem(): void 
{  
  let latestItemId: number = undefined;  
  this.updateStatus('Loading latest item...');  
  
  this.getLatestItemId()  
    .then((itemId: number): Promise<SPHttpClientResponse> => {  
      if (itemId === -1) {  
        throw new Error('No items found in the list');  
      }  
  
      latestItemId = itemId;  
      this.updateStatus(`Loading information about item ID: ${itemId}...`);  
        
      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('DemoList')/items(${latestItemId})?$select=Title,Id`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        });  
    })  
    .then((response: SPHttpClientResponse): Promise<ISPList> => {  
      return response.json();  
    })  
    .then((item: ISPList): void => {  
      this.updateStatus(`Item ID1: ${item.Id}, Title: ${item.Title}`);  
      let title = (<HTMLInputElement>document.getElementById("Title")).value.trim();
      let name = (<HTMLInputElement>document.getElementById("Name")).value.trim();
      const body: string = JSON.stringify({  
        // 'Title': `Updated Item ${new Date()}`,
        'Title': `${title}`,
        'Name': `${name}`  
      });  
  
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('DemoList')/items(${item.Id})`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': '',  
            'IF-MATCH': '*',  
            'X-HTTP-Method': 'MERGE'  
          },  
          body: body  
        })  
        .then((response: SPHttpClientResponse): void => {  
          this.updateStatus(`Item with ID: ${latestItemId} successfully updated`); 
           // reload table 
          this._renderListAsync(); 
        }, (error: any): void => {  
          this.updateStatus(`Error updating item: ${error}`);  
        });  
    });  

   
}  
private updateStatus(status: string, items: ISPList[] = []): void 
{  
  this.domElement.querySelector('.status').innerHTML = status;  
  //this.updateItemsHtml(items);  
} 

private deleteItem(): void {  
  if (!window.confirm('Are you sure you want to delete the latest item?')) {  
    return;  
  }  
  
  this.updateStatus('Loading latest items...');  
  let latestItemId: number = undefined;  
  let etag: string = undefined;  
  this.getLatestItemId()  
    .then((itemId: number): Promise<SPHttpClientResponse> => {  
      if (itemId === -1) {  
        throw new Error('No items found in the list');  
      }  
  
      latestItemId = itemId;  
      this.updateStatus(`Loading information about item ID: ${latestItemId}...`);  
      return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('DemoList')/items(${latestItemId})?$select=Id`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        });  
    })  
    .then((response: SPHttpClientResponse): Promise<ISPList> => {  
      etag = response.headers.get('ETag');  
      return response.json();  
    })  
    .then((item: ISPList): Promise<SPHttpClientResponse> => {  
      this.updateStatus(`Deleting item with ID: ${latestItemId}...`);  
      return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('DemoList')/items(${item.Id})`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=verbose',  
            'odata-version': '',  
            'IF-MATCH': etag,  
            'X-HTTP-Method': 'DELETE'  
          }  
        });  
    })  
    .then((response: SPHttpClientResponse): void => {  
      this.updateStatus(`Item with ID: ${latestItemId} successfully deleted`);  
    }, (error: any): void => {  
      this.updateStatus(`Error deleting item: ${error}`);  
    });  

    // reload table 
    this._renderListAsync();
}  


private getLatestItemId(): Promise<number> {  
  return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('DemoList')/items?$orderby=Id desc&$top=1&$select=id`,  
      SPHttpClient.configurations.v1,  
      {  
        headers: {  
          'Accept': 'application/json;odata=nometadata',  
          'odata-version': ''  
        }  
      })  
      .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {  
        return response.json();  
      }, (error: any): void => {  
        reject(error);  
      })  
      .then((response: { value: { Id: number }[] }): void => {  
        if (response.value.length === 0) {  
          resolve(-1);  
        }  
        else {  
          resolve(response.value[0].Id);  
        }  
      });  
  });  
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
