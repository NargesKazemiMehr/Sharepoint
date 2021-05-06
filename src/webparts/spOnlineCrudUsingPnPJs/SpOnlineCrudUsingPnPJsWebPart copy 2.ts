import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp from 'sp-pnp-js';

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



export interface ISpOnlineCrudUsingPnPJsWebPartProps {

  description: string;

}


export interface ISPList {

  ID: string;

  Title: string;

  OrderNumber: string;

}


export default class SpOnlineCrudUsingPnPJsWebPart extends BaseClientSideWebPart<ISpOnlineCrudUsingPnPJsWebPartProps>

{



  private AddEventListeners(): void {

    document.getElementById('AddItemToSPList').addEventListener('click', () => this.AddSPListItem());

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

    html += `<th></th><th>ID</th><th>Name</th><th>Department</th>`;

    if (items.length > 0) {

      items.forEach((item: ISPList) => {

        html += `

            <tr>

            <td>  <input type="radio" id="empID" name="empID" value="${item.ID}"> <br> </td>



          <td>${item.ID}</td>

          <td>${item.Title}</td>

          <td>${item.OrderNumber}</td>

          </tr>

          `;

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
      </div>

      <div style="background-color: white" >

        <form onSubmit={this.handleSubmit}>

            <br>

            <div data-role="main" class="ui-content">

              <div >





                <input id="OrderNumber"  placeholder="OrderNumber"/>

                <input id="Title"  placeholder="Title"/>

                <button id="AddItemToSPList"  type="submit" >Add</button>

                <button id="UpdateItemInSPList" type="submit" >Update</button>

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
      OrderNumber: document.getElementById('OrderNumber')["value"]
    });
    alert("Record with Orders Name : " + document.getElementById('Orders')["value"] + " Added !");
  }

  UpdateSPListItem() {
    var empID = this.domElement.querySelector('input[name = "empID"]:checked')["value"];
    pnp.sp.web.lists.getByTitle("Orders").items.getById(empID).update({
      Title: document.getElementById('Title')["value"],
      OrderNumber: document.getElementById('OrderNumber')["value"]
    });
    alert("Record with Order ID : " + empID + " Updated !");
  }
  DeleteSPListItem() {
    var empID = this.domElement.querySelector('input[name = "empID"]:checked')["value"];
    pnp.sp.web.lists.getByTitle("Orders").items.getById(empID).delete();
    alert("Record with Order ID : " + empID + " Deleted !");
  }
  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }
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

