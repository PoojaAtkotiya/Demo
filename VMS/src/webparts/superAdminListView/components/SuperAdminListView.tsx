import * as React from 'react';
import styles from './SuperAdminListView.module.scss';
import { ISuperAdminListViewProps } from './ISuperAdminListViewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from "sp-pnp-js";
import { Web } from "sp-pnp-js";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  CheckboxVisibility
} from 'office-ui-fabric-react/lib/DetailsList';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
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
  Approvers :string, 
  ApproveRejectedBy :string, 
  ApproverComment :string, 
  ApproveRejectedDate :string, 
  ViewLink :string,
}[] = [];


let _Id: {
    UserID: number,
    
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
        Edit/View
        </Link>
      )
    },

    {key: 'column2', 
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
    key: 'column5', 
    name: 'Description', 
    fieldName: 'description', 
    minWidth: 100, 
    maxWidth: 200, 
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
    isResizable: true 
    }, 
    { 
    key: 'column9', 
    name: 'Approvers', 
    fieldName: 'Approvers', 
    minWidth: 100, 
    maxWidth: 200, 
    isResizable: true 
    }, 
    { 
    key: 'column10', 
    name: 'Approve/Rejected By', 
    fieldName: 'ApproveRejectedBy', 
    minWidth: 100, 
    maxWidth: 200, 
    isResizable: true 
    }, 
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
    name: 'Approved Date', 
    fieldName: 'ApproveRejectedDate', 
    minWidth: 70,
    maxWidth: 100,
    isResizable: true 
    },

];

export default class SuperAdminListView extends React.Component<{}, {items: {}[];}> {
  public userId :any;
  //private _selection: Selection;
  
  constructor(props: {}) {
    super(props);
  
     pnp.sp.web.currentUser.get().then(function(res){ 
      var isSuperUser : boolean = false;
      pnp.sp.web.siteGroups.getByName("Super User").users.get().then(function(usersList){

        for (var i = 0; i < usersList.length; i++) 
        {
          console.log("Title: " + usersList[i].Title);
          console.log("ID: " + usersList[i].Id);
          if(usersList[i].Title == res.Title){
              isSuperUser = true;
          }
        }
      if(isSuperUser)
      {
        pnp.sp.web.lists.getByTitle('Connect%20Approval').items.filter("Status eq 'Approve' and SuperUserAcknowledged eq 'Assigned' and NewOrModifiedSuperUser eq 'New'").orderBy("ID",false).get().then(
          response => {
            response.map(item =>{
              _items.push({
                key: item.ID, 
                name: item.Title, 
                value: item.ID, 
                Title : item.Title, 
                Id :item.ID, 
                Category :item.Category, 
                Subcategory :item.Sub_x0020_Category, 
                description :item.Feedback_x0020_Description, 
                Status :item.Status, 
                CreatedBy :item.CreatedByDisplay, 
                CreatedDate :(item.Created) ?new Date(item.Created).toLocaleDateString("en-GB") :'',
                Approvers :item.ApproversDispay, 
                ApproveRejectedBy :item.ApprovedByDisplay, 
                ApproverComment :item.Approver_x0020_Comments, 
                ApproveRejectedDate :(item.ApproveRejectedDate) ? new Date(item.ApproveRejectedDate).toLocaleDateString("en-GB"): '',
                ViewLink :"https://bajajelect.sharepoint.com/teams/ConnectApp/SitePages/SuperUserApproval.aspx?ConnectId=" + item.ID
              })
            })
          }
        )
        return _items;
      }
    })
   })
     this.state = {
       items: _items
    };
  
  }

  public render() {

    this.state = {
      items: _items
    };

    let { items } = this.state;

    return (
      <div>
    
          <DetailsList
            items={items}
         //   onRenderCell={ this._onRenderCell } 
            columns={ _columns }
            layoutMode={ DetailsListLayoutMode.fixedColumns }
            onItemInvoked={ this._onItemInvoked }
            checkboxVisibility = {CheckboxVisibility.hidden}
          
          />
       
      </div>
    );
  }

  // private _onRenderCell(item: any, index: number): JSX.Element {
  //   return(
  //     <div className='ms-ListGhostingExample-itemContent'>
  //         <div className='ms-ListGhostingExample-itemName'>{ item.ID }</div>
  //        {/* <div className='ms-ListGhostingExample-itemIndex'>{ `Item ${index}` }</div> */}
  //       </div>
  //   )
  // }

  
  @autobind
  private _onItemInvoked(item: any): void {
    window.location.href = item.ViewLink;
  }
}




// export default class SuperAdminListView extends React.Component<ISuperAdminListViewProps, {}> {
//   public render(): React.ReactElement<ISuperAdminListViewProps> {
//     return (
//       <div className={ styles.superAdminListView }>
//         <div className={ styles.container }>
//           <div className={ styles.row }>
//             <div className={ styles.column }>
//               <span className={ styles.title }>Welcome to SharePoint!</span>
//               <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
//               <p className={ styles.description }>{escape(this.props.description)}</p>
//               <a href="https://aka.ms/spfx" className={ styles.button }>
//                 <span className={ styles.label }>Learn more</span>
//               </a>
//             </div>
//           </div>
//         </div>
//       </div>
//     );
//   }
// }
