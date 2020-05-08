import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'IndexListWebPartStrings';
import { SPComponentLoader } from "@microsoft/sp-loader";

import '../../ExternalReferences/css/style.css';  
import { sp } from "@pnp/sp/presets/all";
import "jquery";
require("bootstrap");
SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.13/css/jquery.dataTables.min.css");
SPComponentLoader.loadScript("https://cdn.datatables.net/1.10.19/js/jquery.dataTables.min.js");

declare var $;

export interface IIndexListWebPartProps {
  ListName: string;
}


//https://daifukuamerica.sharepoint.com/sites/WDH/_layouts/15/workbench.aspx

export default class IndexListWebPart extends BaseClientSideWebPart<IIndexListWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(() => {  
      sp.setup({
         spfxContext: this.context
         });
    });  
  }
  public render(): void {
    this.domElement.innerHTML = `
    <h3 id='enterNotes' style="text-align:center;"></h3>
    <table style="width:100%;display:none" id="projectTable" class="cell-border">
  <thead>
    <th>Team</th>
    <th>Team Code</th>
    <th>Location</th>
    <th>Customer</th>
    <th>Project Number</th>
  </thead>
  <tbody id="tblBody"></tbody>
</table>`;
   
    this.loadIndex();
  }

  public async loadIndex() {
    if(this.properties.ListName) 
    {
      $('#enterNotes').empty();
      $('#projectTable').show();
      await sp.web.lists.getByTitle(this.properties.ListName).items.select("Title","SharePoint_x0020_URL","TeamsCode","Project_x0020_Number", "Location/Title","Customer/Title").expand("Location","Customer").getAll().then((allItems: any[]) => {
        for (var index = 0; index < allItems.length; index++) {
          var element = allItems[index];
          // var created = new Date(element.Created);
          // var date = created.getDate() + "/" + created.getMonth() + 1 + "/" + created.getFullYear();
          $('#tblBody').append('<tr><td><a target="_blank" href="' + element.SharePoint_x0020_URL.Url + '">' + element.Title + '</a></td><td>'+ (element.TeamsCode ? '<a target="_blank" href="' + element.TeamsCode.Url + '">' + element.TeamsCode.Description + '</a>': '') + '</td><td>' + element.Location.Title + '</td><td>' + (element.Customer ? element.Customer.Title : '') +  '</td><td>' + (element.Project_x0020_Number ? element.Project_x0020_Number : '') + '</td></tr>')
        }
      });
      var oTable = $("#projectTable").DataTable({
        columnDefs: [
        {"className": "dt-left", "targets": "_all"}
      ]
    })
    }
    else
    {
      $('#enterNotes').empty();
      $('#projectTable').hide();
      $('#enterNotes').append('Please Enter ListName in Webpart property pane');
    }

  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('ListName', {
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
