import * as React from 'react';
import styles from './EditConnectApprover.module.scss';
import { IEditConnectApproverProps,ISPList } from './IEditConnectApproverProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, PrimaryButton, DefaultButton, Label } from 'office-ui-fabric-react/lib/';
import pnp from "sp-pnp-js";
import { Web } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';

export default class EditConnectApprover extends React.Component<IEditConnectApproverProps, {}> {
  private listdata: ISPList;
  queryParameters: UrlQueryParameterCollection
  web: any;
  private creatorName :string;


  componentWillMount() {
    
    this.listdata = {
      ID: undefined,
      Title: undefined,
      Category: undefined,
      SubCategory: undefined,
      Description: undefined,
      Status :undefined,
      Approver :undefined,
      ApproverComment :undefined,
      ConnectID:undefined,
      Created:undefined,
      Creator:undefined,
    }
   
    this.web = new Web(this.props.context.pageContext.web.absoluteUrl);
    this.queryParameters = new UrlQueryParameterCollection(window.location.href);
    if (this.queryParameters.getValue("ApproverId")) {
      const Id: number = parseInt(this.queryParameters.getValue("ApproverId"));
      console.log("ApproverId value is : " + Id);
      this.getListData(Id);
      this.creatorName = this.getCreatorName(Id);
      console.log("this.getCreatorName(Id) = "+ this.creatorName);
    }
  }

  private getListData(Id: number): void {
    this._getListData(Id)
      .then((response) => {
        response.map((item: ISPList) => {
          this.listdata = {
            ID: item.ID,
            Title: item.Title,
            Description: item.Description,
            Category: item.Category,
            SubCategory: item.SubCategory,
            Status: item.Status,
            Approver :item.Approver,
            ApproverComment: item.ApproverComment,
            ConnectID :item.ConnectID,
            Created:item.Created,
            Creator:''
            
          };
      
        });
       
       
        this.setState(this.listdata);
      });
      
}

private getCreatorName(Id : number): string
{
 
  this.web.lists.getByTitle('Connect%20Approval').items.getById(Id).fieldValuesAsText.get().then(function(data) {
      //Populate all field values for the List Item
      for (var k in data) 
      {
        console.log(k + " - " + data[k]);  
        if(k.trim() == "creator")  
        {  
          console.log("if : this.CratorName = "+ this.creatorName);    
          this.creatorName = data[k]; 
        
        
        } 

      } 
  });
  this.setState(this.creatorName);
  // console.log("Inside the getCreatotName() : creatorName ==" + this.creatorName );
  return this.creatorName;
  
}

private _getListData(Id): Promise<ISPList[]> {
  return this.web.lists.getByTitle("Connect%20Approval").items.filter("Id eq " + Id).get().then((response) => {
    var data = [];
    response.map((item) => {
      data.push({
        ID: item.ID,
        Title: item.Title,
        Description: item.Feedback_x0020_Description,
        Category: item.Category,
        SubCategory: item.Sub_x0020_Category,
        Status: item.Status,
        Approver: item.Approver,
        ApproverComment :item.Approver_x0020_Comments,
        ConnectID :item.Feedback_x0020_Title,
        Created:item.Created,
       // Creator : item.Creator,
      });
    });
    return data;
  });
}  

  public render(): React.ReactElement<IEditConnectApproverProps> {
   
    const divPadding = {      
      padding:'10px 3px',
      position:'relative',
    };
    const lblHeader ={
      width: '100%',
      padding:'5px 0',
      margin :0,
      fontWeight : 400,
      fontFamily :'"Segoe UI Semibold WestEuropean","Segoe UI Semibold","Segoe UI",Tahoma,Arial,sans-serif',
      fontSize : '14px'
    }
    const lblValue ={

      fontFamily:'"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',
      fontWeight:400,
      fontSize:'14px',
      padding :'0 0 5px',
    }

    return (
      <div className={ styles.editConnectApprover }>

      <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6">
          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Connect ID</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.ConnectID}</label>
            </div>
          </div>

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Category</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.Category}</label>
            </div>
          </div>

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Sub Category</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.SubCategory}</label>
            </div>
          </div>

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding}>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Title</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.Title}</label>
            </div>
          </div>

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Description</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.Description}</label>
            </div>
          </div>  

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Status</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.Status}</label>Dropdown here
            </div>
          </div>  
          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Approver Comment</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.ApproverComment}</label>
              multiline text box here
            </div>
          </div>  
          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Created By</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.Creator}</label>
            </div>
          </div>  
          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
              <label>Created</label>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
              <label>{this.listdata.Created}</label>
            </div>             
          </div> 
       </div>
       <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6 {styles.column}">
       </div> 







              {/* <div className={ styles.container }>
                <div className={ styles.row }>
                  <div className={ styles.column }>
                    <span className={ styles.title }>Welcome to SharePoint!</span>
                    <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
                    <p className={ styles.description }>{escape(this.props.description)}</p>
                    <a href="https://aka.ms/spfx" className={ styles.button }>
                      <span className={ styles.label }>Learn more</span>
                    </a>
                  </div>
                </div>
              </div> */}
      </div>
    );
  }
}
