import * as React from 'react';
import styles from './ViewRequest.module.scss';
import { IViewRequestProps,ISPList } from './IViewRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {PrimaryButton, DefaultButton, Label } from 'office-ui-fabric-react/lib/';
import pnp from "sp-pnp-js";
import { Web } from "sp-pnp-js";
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { Conversation } from 'sp-pnp-js/lib/graph/conversations';


export default class ViewRequest extends React.Component<IViewRequestProps, {}> {
  
  private listdata: ISPList;
  queryParameters: UrlQueryParameterCollection
  web: any;

  componentWillMount() {
    this.listdata = {
      ID: undefined,
      Title: undefined,
      Category: undefined,
      SubCategory: undefined,
      Description: undefined,
      Status: undefined,
      ApproverComment: undefined,
      ApprovedBy : undefined,
      ApprovedRejectedDate :undefined
    }
    
    this.web = new Web(this.props.context.pageContext.web.absoluteUrl);
    this.queryParameters = new UrlQueryParameterCollection(window.location.href);
    if (this.queryParameters.getValue("ConnectId")) {
      const Id: number = parseInt(this.queryParameters.getValue("ConnectId"));
      console.log("Id value is : " + Id);
      this.getListData(Id);
    }
  }

  private getListData(Id: number): void {
    this._getListData(Id)
      .then((response) => {
        //ISPList data;
        response.map((item: ISPList) => {
          this.listdata = {
            ID: item.ID,
            Title: item.Title,
            Description: item.Description,
            Category: item.Category,
            SubCategory: item.SubCategory,
            Status: item.Status,
            ApproverComment: item.ApproverComment,
            ApprovedRejectedDate : item.ApprovedRejectedDate,
            ApprovedBy :item.ApprovedBy
          };
        });
        this.setState(this.listdata);
      });
  
}

private _getListData(Id): Promise<ISPList[]> {
  return this.web.lists.getByTitle("Connect").items.filter("Id eq " + Id).get().then((response) => {
    var data = [];
    response.map((item) => {
      data.push({
        ID: item.ID,
        Title: item.Title,
        Description: item.Feedback_x0020_Description,
        Category: item.Category,
        SubCategory: item.Sub_x0020_Category,
        Status: item.Status,
        ApproverComment: item.ApproverComment,
        ApprovedRejectedDate :item.ApprovedRejectedDate,
        ApprovedBy :item.ApprovedBy
      });
    });
 
    return data;
  });
}  

  public render(): React.ReactElement<IViewRequestProps> { 
    const divPadding = {      
      paddingTop:8,
      paddingBottom :8,
      paddingLeft:3,
      paddingRight:3,
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
    var appRejDate = '';
  
    if(this.listdata.ApprovedRejectedDate){
      appRejDate = new Date(this.listdata.ApprovedRejectedDate).toLocaleDateString('en-GB');
    }
    return (
      <div className={ styles.viewRequest }>
          {/* <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> */}
            <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6">

              <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding}>
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
                  <label>{this.listdata.Status}</label>
                </div>
              </div>  
              <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
                <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
                  <label>Approver Comment</label>
                </div>
                <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
                  <label>{this.listdata.ApproverComment}</label>
                </div>
              </div>  
              <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
                <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
                  <label>Approved By</label>
                </div>
                <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
                  <label>{this.listdata.ApprovedBy}</label>
                </div>
              </div>  
              <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
                <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblHeader}> 
                  <label>Approved/Rejected Date</label>
                </div>
                <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={lblValue}> 
                   <label>{appRejDate}</label>
                </div>             
              </div> 

              <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding} >
              <div className="ms-Grid-col ms-u-sm12 ms-u-md8 ms-u-lg8"> 
              </div>
              <div className="ms-Grid-col ms-u-sm12 ms-u-md4 ms-u-lg4"> 
                <DefaultButton
                  text='Close'
                  href='https://bajajelect.sharepoint.com/teams/ConnectApp/'
                  /> 
              </div>
              </div>         
          </div>
          <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6 {styles.column}">
          </div>         
        {/* </div> */}
      </div>
    );
  }
}