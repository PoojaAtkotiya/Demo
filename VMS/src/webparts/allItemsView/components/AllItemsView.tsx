/* tslint:disable:no-unused-variable */
import * as React from 'react';
/* tslint:enable:no-unused-variable */
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import {
  CheckboxVisibility,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,ConstrainMode
  
} from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link'
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import pnp from "sp-pnp-js";
import { Web } from "sp-pnp-js";
import { Item, Items } from 'sp-pnp-js/lib/sharepoint/items';


let _items: {
  key: number,
  name: string,
  value: number,
  Title : string,
  Id :number,
  Category :string,
  Subcategory :string,
  description :string,
  Status :string,
  CreatedBy :string,
  CreatedDate :string,
  // Approvers :string,
  // ApproveRejectedBy :string,
  ApproverComment :string,
  ApproveRejectedDate :string,
   ViewLink :string,
}[] = [];

let _columns = [
 
  {
    key: 'column13',
    
    name: 'View Link',
    fieldName: 'ViewLink', 
    minWidth: 50,
    maxWidth: 60,
    isResizable: true,  
    onRender : item =>(
      <Link data-selection-invoke={ true } href={item.ViewLink }>
     View
      </Link>
    )
  },
  {
    key: 'column2',
    name: 'ID',
    fieldName: 'Id',
    minWidth: 25,
    maxWidth: 40,
    isResizable: true
  },
  {
    key: 'column1',
    name: 'Title',
    fieldName: 'name',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'column5',
    name: 'Description',
    fieldName: 'description',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'column3',
    name: 'Category',
    fieldName: 'Category',
    minWidth: 70,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'column4',
    name: 'Sub Category',
    fieldName: 'Subcategory',
    minWidth: 70,
    maxWidth: 100,
    isResizable: true
  },
 
  {
    key: 'column6',
    name: 'Status',
    fieldName: 'Status',
    minWidth: 70,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'column7',
    name: 'Created By',
    fieldName: 'CreatedBy',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'column8',
    name: 'Created Date',
    fieldName: 'CreatedDate',
    minWidth: 70,
    maxWidth: 100,
    isResizable: true,
  },
  // {
  //   key: 'column9',
  //   name: 'Approvers',
  //   fieldName: 'Approvers',
  //   minWidth: 100,
  //   maxWidth: 200,
  //   isResizable: true
  // },
  // {
  //   key: 'column10',
  //   name: 'Approve/Rejected By',
  //   fieldName: 'ApproveRejectedBy',
  //   minWidth: 100,
  //   maxWidth: 200,
  //   isResizable: true
  // },
  {
    key: 'column11',
    name: 'Approver Comment',
    fieldName: 'ApproverComment',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'column12',
    name: 'Approve/Rejected Date',
    fieldName: 'ApproveRejectedDate',
    minWidth: 70,
    maxWidth: 100,
    isResizable: true
    
  },
  
];

export default class AllItemsView extends React.Component<{},{items: {}[];}> 
{
  private _selection: Selection;
  constructor(props: {}) {
    super(props);  
  
  pnp.sp.web.lists.getByTitle('Connect%20Approval').items.orderBy('ID', false).get().then(
    response => {
      response.map(item =>{
        _items.push({
          key: item.ID,
          name: item.Title,
          value:  item.ID,
          Title : item.Title,
          Id :item.ID,
          Category :item.Category,
          Subcategory :item.Sub_x0020_Category,
          description :item.Feedback_x0020_Description,
          Status :item.Status,
          CreatedBy :item.CreatedByDisplay,
          CreatedDate :(item.Created) ?new Date(item.Created).toLocaleDateString("en-GB") :'',
          // Approvers :item.ApproversDispay,
          // ApproveRejectedBy :item.ApprovedByDisplay,
          ApproverComment :item.Approver_x0020_Comments,
          ApproveRejectedDate : (item.ApproveRejectedDate) ? new Date(item.ApproveRejectedDate).toLocaleDateString("en-GB"): '',
          ViewLink :"https://bajajelect.sharepoint.com/teams/ConnectApp/SitePages/ViewRequestAllFields.aspx?ApproverId=" + item.ID
        })
      })
    }
  )    

  this.state = {
    items: _items
  };
    
  }

  public render() {

    let { items } = this.state;
   
    console.log();
    return (
      <div style = {{wordWrap : 'true'}}>
          <DetailsList
          
            constrainMode = {ConstrainMode.horizontalConstrained}
            items={ items }
            columns={ _columns }
           // setKey='set'
            layoutMode={ DetailsListLayoutMode.fixedColumns }
            //selection={ this._selection }
           // selectionPreservedOnEmptyClick={ true }
            onItemInvoked={ this._onItemInvoked }
            checkboxVisibility ={CheckboxVisibility.hidden}
           
           
            //onRenderItemColumn={ this._onRenderItemColumn }
            // compact={ true }
            
          />
        
      </div>
      
    );
  }

  private _onItemInvoked(item: any): void {
    window.location.href = item.ViewLink;
  }
}