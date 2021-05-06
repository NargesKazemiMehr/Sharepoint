import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp, { Item } from 'sp-pnp-js';

import { Version } from '@microsoft/sp-core-library';

import {

  BaseClientSideWebPart

} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { escape } from '@microsoft/sp-lodash-subset';


import styles from './SpOnlineCrudUsingPnPJsWebPart.module.scss';

import * as strings from 'SpOnlineCrudUsingPnPJsWebPartStrings';
import { Group } from 'sp-pnp-js/lib/graph/groups';

export interface ISpOnlineCrudUsingPnPJsWebPartProps {

  description: string;

}

export interface ISPList {
  ID: string;
  Title:string;
 CustomerName:string;
 OrderNumber: string;
  Destination: string;
  Status: string;
  Owner:string;
}

export default class SpOnlineCrudUsingPnPJsWebPart extends BaseClientSideWebPart<ISpOnlineCrudUsingPnPJsWebPartProps>
{
  private AddEventListeners(): void {
    document.getElementById('AddItemToSPList').addEventListener('click', () => this.AddSPListItem());
    // document.getElementById('Attachment').addEventListener('click', () => this.AttachmentSPItem());
    document.getElementById('UpdateItemInSPList').addEventListener('click', () => this.UpdateSPListItem());

    document.getElementById('DeleteItemFromSPList').addEventListener('click', () => this.DeleteSPListItem());
  }

  private _getSPItems(): Promise<ISPList[]> {
    return pnp.sp.web.lists.getByTitle("Orders").items.get().then((response) => {
      return response;
    });
  }

  private getSPItems(): void {
    this._getSPItems()
      .then((response) => {
        this._renderList(response);
      });
  }

  private _renderList(items: ISPList[]): void {

    let html: string = '<table class="TFtable" border=1 width=style="bordercollapse: collapse;">';

    html += `<th></th><th>ID</th><th>Name</th><th>Order Number</th><th>Customer Name</th><th>Destination</th><th>Owner</th><th>Status</th>`;

    if (items.length > 0) {

      items.forEach((item: ISPList) => {
    pnp.sp.web.lists.getByTitle("Orders").items.getById(+item.ID).select("Status","CustomerName/Title","Owner/Title").expand("CustomerName","Owner").get().then((Myitems: any[]) => {
      console.log(Myitems);
      if( Myitems["CustomerName"]!==undefined  || Myitems["Owner"]!==undefined  ) {
        html += `
        <tr>
        <td>  <input type="radio" id="ID" name="ID" value="${item.ID}"> <br> </td>
        <td>${item.ID}</td>
        <td>${item.Title}</td>
        <td>${item.OrderNumber}</td>
        <td>${Myitems["CustomerName"].Title}</td>
        <td>${item.Destination}</td>
        <td>${Myitems["Owner"].Title}</td>
        <td>${Myitems["Status"].Title}</td>
      </tr>
      `;
      }
      else{
      html += `
      <tr>
      <td>  <input type="radio" id="ID" name="ID" value="${item.ID}"> <br> </td>
      <td>${item.ID}</td>
      <td>${item.Title}</td>
      <td>${item.OrderNumber}</td>
      <td>""</td>
      <td>${item.Destination}</td>
      <td>""</td>
      <td>""</td>
    </tr>
    `;
  }
  });
       });
    }
    else {
      html += "No records...";
    }
    html += `</table>`;
      const listContainer: Element = this.domElement.querySelector('#DivGetItems');
      listContainer.innerHTML = html;
  }
  public render(): void {

    this.domElement.innerHTML = `
    <div class="parentContainer" style="background-color: white">

    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">

      <div class="ms-Grid-col ms-u-lg

  ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
      </div>

    </div>

    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">

      <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:

  x;">Orders Details</div>
    </div>                <button id="UpdateItemInSPList" type="submit" >Update</button>

    <div style="background-color: white" >
        <form onSubmit={this.handleSubmit}>
            <br>
            <div data-role="main" class="ui-content">
              <div >
              <input id="Title"  placeholder="Title"/>
              OrderNumber:<br>
                <input id="OrderNumber" style="width: 100%;"/><br>
                Customer Name:<br>
                <input id="CustomerName" style="width: 100%;"/><br>
                Destination:<br>
                <input id="Destination" style="width: 100%;"/><br>
                Owner:<br>
                <input id="Owner" style="width: 100%;"/><br>
                <button id="AddItemToSPList"  type="submit" >Add</button>
                <input type="file" id="Attachment" name="Attachment"/><br><br>

                <button id="DeleteItemFromSPList"  type="submit" >Delete</button>
                </div>
            </div>
        </form>
      </div>
      <br>
      <div style="background-color: white" id="DivGetItems" />
      </div>
      `;
    this.getSPItems();

    this.AddEventListeners();

  }
  AddSPListItem() {
    pnp.sp.web.lists.getByTitle('Orders').items.add({
      Title: document.getElementById('Title')["value"],
      OrderNumber: document.getElementById('OrderNumber')["value"],
      CustomerName: document.getElementById('CustomerName')["value"],
      Destination: document.getElementById('Destination')["value"],
      Owner: document.getElementById('Owner')["value"],
    });
    this.AttachmentSPItem();
    alert("Record with Orders Name : " + document.getElementById('Title')["value"] + " Added !");
  }

  UpdateSPListItem() {
    var ID = this.domElement.querySelector('input[name = "ID"]:checked')["value"];
    pnp.sp.web.lists.getByTitle("Orders").items.getById(ID).update({
      Title: document.getElementById('Title')["value"],
      OrderNumber: document.getElementById('OrderNumber')["value"]
    });
    alert("Record with Order ID : " + ID + " Updated !");
  }
  DeleteSPListItem() {
    var ID = this.domElement.querySelector('input[name = "empID"]:checked')["value"];
    pnp.sp.web.lists.getByTitle("Orders").items.getById(ID).delete();
    alert("Record with Order ID : " + ID + " Deleted !");
  }
  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }
  private AttachmentSPItem(): void {
    let input = <HTMLInputElement>document.getElementById("Attachment");
    let file = input.files[0];
    pnp.sp.web.getFolderByServerRelativeUrl("/sites/nkSite/OrdersAttachment/").files.add(file.name, file, true);
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

