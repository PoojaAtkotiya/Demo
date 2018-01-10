import * as React from 'react';
import styles from './FbForm.module.scss';
import { IFbFormProps, ISPList, IKeyText } from './IFbFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, PrimaryButton, DefaultButton, Label } from 'office-ui-fabric-react/lib/';
import pnp from "sp-pnp-js";
import { Web } from "sp-pnp-js";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import {
  assign,
  autobind
} from 'office-ui-fabric-react/lib/Utilities';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import getFormData, { getFieldData } from 'get-form-data'


export default class FbForm extends React.Component<IFbFormProps, {}>  {
  private category: string;
  private subCategory: string;
  private categorySelectedValue: { key: string | number | undefined, value: string };
  private subCategorySelectedValue: { key: string | number | undefined, value: string };
  private statusSelectedValue: { key: string | number | undefined, value: string };
  private title: string = "";
  private description: string = "";
  private approverComment: string = "";
  private categoryData: IDropdownOption[];
  private subCategoryData: IDropdownOption[];
  private statusOption: IDropdownOption[];
  private listdata: ISPList;
  queryParameters: UrlQueryParameterCollection
  web: any;
  hideDialog: boolean;
  constructor() {
    super();

  }

  componentWillMount() {
    
    this.hideDialog = true
    this.listdata = {
      ID: undefined,
      Title: undefined,
      Category: undefined,
      SubCategory: undefined,
      Description: undefined,

    }
    this.categorySelectedValue = {
      key: "",
      value: ""
    };
    this.subCategorySelectedValue = {
      key: "",
      value: ""
    };
    this.statusSelectedValue = {
      key: "",
      value: ""
    };
    this.web = new Web(this.props.context.pageContext.web.absoluteUrl);
    var data = this.getCategoryData();
    this.queryParameters = new UrlQueryParameterCollection(window.location.href);
    // if (this.queryParameters.getValue("ConnectId")) {
    //   const Id: number = parseInt(this.queryParameters.getValue("ConnectId"));
    //   console.log("Id value is : " + Id);
    //   this.getListData(Id);
    // }
  }

  // private getListData(Id: number): void {

  //   this._getListData(Id)
  //     .then((response) => {
  //       //ISPList data;
  //       response.map((item: ISPList) => {
  //         console.log(item);
  //         this.listdata = {
  //           ID: item.ID,
  //           Title: item.Title,
  //           Description: item.Description,
  //           Category: item.Category,
  //           SubCategory: item.SubCategory
  //         };
  //       });

  //       this.web.lists.getByTitle("CategoryMaster").items.filter("Title eq '" + this.listdata.Category + "'").select('ID,Title').get().then((response) => {


  //         response.map((item) => {

  //           this.listdata.Category = item.ID;
  //           this.categorySelectedValue = {
  //             key: item.ID,
  //             value: item.Title
  //           };
  //           this.setState(this.categorySelectedValue);

  //           var data = { key: item.Id, text: item.Title };
  //           this.getSubCategory(data);
  //           this.web.lists.getByTitle("SubCategoryMaster").items.filter("Title eq '" + this.listdata.SubCategory + "'").select('ID,Title').get().then((response) => {
  //             response.map((itemSubCategory) => {
  //               this.listdata.SubCategory = itemSubCategory.ID;
  //               this.subCategorySelectedValue = {
  //                 key: itemSubCategory.ID,
  //                 value: itemSubCategory.Title
  //               };
  //               this.setState(this.listdata);
  //               this.setState(this.subCategorySelectedValue);
  //             });
  //           });
  //         });
  //       });
  //     });
  // }

  // private _getListData(Id): Promise<ISPList[]> {
  //   return this.web.lists.getByTitle("Connect%20Approval").items.filter("Id eq " + Id).get().then((response) => {
  //     var data = [];
  //     response.map((item) => {
  //       data.push({
  //         ID: item.ID,
  //         Title: item.Title,
  //         Description: item.Feedback_x0020_Description,
  //         Category: item.Category,
  //         SubCategory: item.Sub_x0020_Category,
  //         // Status: item.Status,
  //         // ApproverComment: item.ApproverComment
  //       });
  //     });

  //     console.log(data);
  //     return data;
  //   });
  // }

  private getCategoryData(): IDropdownOption[] {

    var data: IDropdownOption[];
    this._getCategoryData()
      .then((response) => {
        data = this._renderCategoryList(response);
        //  this.categoryData=data;
        this.setState(this.categoryData = data);
      });
    return data;
  }
  private _getCategoryData(): Promise<ISPList[]> {
    return this.web.lists.getByTitle("CategoryMaster").items.get().then((response) => {
      return response;
    });
  }
  /// make is single method
  private _renderCategoryList(items: ISPList[]): IDropdownOption[] {
    return items.map((item: ISPList): IDropdownOption => {
      return {
        key: item.ID,
        text: item.Title,
      };

    });
  }


  @autobind
  public getSubCategory(item: IDropdownOption) {
    this.categorySelectedValue = { key: item.key, value: item.text };
    this.subCategorySelectedValue = { key: "", value: "" };
    var data: IDropdownOption[];
    this._getSubCategoryData(item.key)
      .then((response) => {
        data = this._renderCategoryList(response);
        this.setState(this.subCategoryData = data);
      });
    return data;

  }


  private _getSubCategoryData(categoryTxt: string | number): Promise<ISPList[]> {
    return this.web.lists.getByTitle("SubCategoryMaster").items.filter("Category eq " + categoryTxt).get().then((response) => {
      return response;
    });
  }

  private _renderSubCategoryList(items: ISPList[]): void {
    var data = [];
    items.forEach((item: ISPList) => {
      data.push({
        key: item.ID,
        text: item.Title
      });

    });
    // this.setState({});
  }
  @autobind
  private subCategorySelected(item: IDropdownOption) {
    this.subCategorySelectedValue = { key: item.key, value: item.text };
  }

  @autobind
  private statusSelected(item: IDropdownOption) {
    this.statusSelectedValue = { key: item.key, value: item.text };
  }

  public _saveJMDConnectForm(event) {
    event.preventDefault();
    const _title =   event.target['txtTitle'].value;
    const _description = event.target['txtDescription'].value;
    const _category = document.getElementById("txtCategory-option").textContent;
    const _subCategory = document.getElementById("txtSubCategory-option").textContent;

    var isValid: boolean = false;
    var errorMsg = [];
      if (_category == '' || _category == null || _category == undefined || _category.trim().length <= 0 || _category.toLowerCase().match("select an option")) {
        document.getElementById("txtCategory-option").style.borderColor = "red";
        document.getElementById("txtCategory-option").style.backgroundColor = "lightyellow";
        isValid = false;
        errorMsg.push({
          key: "errorCat",
          value : "Category is required"
        });
      }
      else {
        document.getElementById("txtCategory-option").style.backgroundColor = "white";
        isValid = true;
      }
      if (_subCategory == '' || _subCategory == null || _subCategory == undefined || _subCategory.trim().length <= 0 ||_subCategory.toLowerCase().match("select an option")) {
        document.getElementById("txtSubCategory-option").style.borderColor = "red";
        document.getElementById("txtSubCategory-option").style.backgroundColor = "lightyellow";
        isValid = false;
        errorMsg.push({
          key: "errorSubCat",
          value : "Sub Category is required"
        });
      }
      else {
        document.getElementById("txtSubCategory-option").style.backgroundColor = "white";
        isValid = true;
      }
      if (_title == '' || _title == null ||_title == undefined) {
        document.getElementById("txtTitle").style.borderColor = "red";
        document.getElementById("txtTitle").style.borderWidth = "2";
        document.getElementById("txtTitle").style.backgroundColor = "lightyellow";
        isValid = false;
        errorMsg.push({
          key: "errorTitle",
          value : "Title is required"
        });
      }
      else if( _title.length > 255 ){
        document.getElementById("txtTitle").style.borderColor = "red";
        document.getElementById("txtTitle").style.backgroundColor = "lightyellow";
        isValid = false;
        errorMsg.push({
          key: "errorDesc",
          value : "Title should not be longer than 255 characters."
        });
      }
      else {
        document.getElementById("txtTitle").style.backgroundColor = "white";
        isValid = true;
      }
      if ( _description == '' ||  _description== null|| _description == undefined ) {
  
        document.getElementById("txtDescription").style.borderColor = "red";
        document.getElementById("txtDescription").style.backgroundColor = "lightyellow";
        isValid = false;
        errorMsg.push({
          key: "errorDesc",
          value : "Description is required."
        });
      }
      else if( _description.length > 1000 ){

        document.getElementById("txtDescription").style.borderColor = "red";
        document.getElementById("txtDescription").style.backgroundColor = "lightyellow";
        isValid = false;
        errorMsg.push({
          key: "errorDesc",
          value : "Description should not be longer than 1000 characters."
        });
      }
      else {
        document.getElementById("txtDescription").style.backgroundColor = "white";
        isValid = true;
      }

    console.log("_title = " +_title +",txtTitle ==" + document.getElementById("txtTitle")['value']);
    console.log("event.target.innerText = "+ event.target.innerText);
    console.log("txtSubCategory-option =" + document.getElementById("txtSubCategory-option").textContent);
    console.log("txtCategory-option =" + document.getElementById("txtCategory-option").textContent);

    if(isValid)
    {
      // if (this.queryParameters.getValue("ConnectId")) {
      //       var id = parseInt(this.queryParameters.getValue("ConnectId"));
      //       this.web.lists.getByTitle("Connect").items.getById(id).update({
      //         Category: _category,
      //         Sub_x0020_Category: _subCategory,
      //         Title: _title,
      //         Feedback_x0020_Description: _description,
      //         //Status: _status,
      //        // ApproverComment: _approverComment
      //       }).then((result): void => {
      //         alert("Feedback Updated!");
      //         window.location.reload();
      //       }, (error: any): void => {
      //         console.log(error);
      //         alert("Oops!Something went wrong!!!");
      //       });
      // }
     // else {
            this.web.lists.getByTitle('Connect%20Approval').items.add({
              Category: _category,
              Sub_x0020_Category: _subCategory,
              Title: _title,
              Feedback_x0020_Description: _description
            }).then((result): void => {
              alert("New Feedback has been Added.");
              window.location.href = "https://bajajelect.sharepoint.com/teams/ConnectApp";
            }, (error: any): void => {
              console.log(error);
              alert("Oops!Something went wrong!!!");
            });
        // }
      }
    else{
      console.log(errorMsg);
      var msg = "";
      var newLine = "\r\n"
      errorMsg.forEach(element => {
        msg +=  element.value + newLine;
      });
      alert(msg);  
    }
        
  }


  //save Form Data
  // public _saveJMDConnectForm(): any {
  //   // call valid method ()
  //   this.title = document.getElementById('txtTitle')["value"];
  //   this.description = document.getElementById('txtDescription')["value"];
  //   this.approverComment = document.getElementById('txtApproverComment')["value"];
  //   var isValid: boolean;
  //   isValid = true;

  //   isValid = this.validate();
  //   if (isValid) {
  //     if (this.queryParameters.getValue("ConnectId")) {
  //       var id = parseInt(this.queryParameters.getValue("ConnectId"));
  //       this.web.lists.getByTitle("Connect").items.getById(id).update({
  //         Category: this.categorySelectedValue.value,
  //         Sub_x0020_Category: this.subCategorySelectedValue.value,
  //         Title: this.title,
  //         Feedback_x0020_Description: this.description,
  //         Status: this.statusSelectedValue.value,
  //         ApproverComment: this.approverComment
  //       }).then((result): void => {
  //         alert("Feedback Updated!");
  //         window.location.reload();
  //       }, (error: any): void => {
  //         console.log(error);
  //         alert("Oops!Something went wrong!!!");
  //       });
  //     }
  //     else {
  //       this.web.lists.getByTitle('Connect').items.add({
  //         Category: this.categorySelectedValue.value,
  //         Sub_x0020_Category: this.subCategorySelectedValue.value,
  //         Title: this.title,
  //         Feedback_x0020_Description: this.description
  //       }).then((result): void => {
  //         alert("Feedback Added!");
  //         window.location.reload();
  //       }, (error: any): void => {
  //         console.log(error);
  //         alert("Oops!Something went wrong!!!");
  //       });
  //     }
  //   }
  //   else {
  //     alert('Please fill out this field.');
  //   }
  // }
  @autobind
  private _closeDialog() {
    // this.setState({ hideDialog: true });
    this.hideDialog = true;
  }
  public onsubmit(event) {
    event.preventDefault();
    var form1 = document.querySelector("#frmFeedback");
    var data = getFormData(form1)
    console.log(this.refs.txtCategory);
    console.log(JSON.stringify(data));
  }
  
  public render(): React.ReactElement<IFbFormProps> {
    console.log(this.props.context.pageContext.web.absoluteUrl);
    this.web = new Web(this.props.context.pageContext.web.absoluteUrl);
    this.statusOption = [{ key: 'In Progress', text: 'In Progress' },
    { key: 'Completed', text: 'Completed' },
    { key: 'Rejected', text: 'Rejected' }]
    const errorStyle = {
      color: 'red'
    };
    const divPadding = {      
      padding:'10px 3px',
      position:'relative',
    };
    return (
      <form onSubmit={() => { this._saveJMDConnectForm(event)}}  formNoValidate >
      <div className={styles.fbForm}>
        <div>
          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">

            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"  style={divPadding}>
            <label>
              <span className={styles.Error} style={errorStyle}>*</span>
              Category
              <Dropdown
                placeHolder='Select an Option'
                //label='*Category'
                ref={'txtCategory'}
                id='txtCategory'
                selectedKey={this.listdata.Category}
                ariaLabel='category'
                className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6"
                options={this.categoryData}
                onChanged={this.getSubCategory}
              />
              </label>
            </div>

            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"  style={divPadding}>
            <label>
              <span style={errorStyle}>*</span>
              Sub Category
              <Dropdown
                placeHolder='Select an Option'
                // label='*Sub-Category'
                id='txtSubCategory'
                selectedKey={this.listdata.SubCategory}
                ariaLabel='subCategory'
                className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6"
                options={this.subCategoryData}
                onChanged={this.subCategorySelected}
              />
              </label>
            </div>

            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"  style={divPadding}>
            <label>
              <span style={errorStyle}>*</span>
              Title
            <TextField
                id="txtTitle"
                placeholder="Enter text here"
               // label="*Title"
                className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6"
                value={this.listdata.Title}
            />
              </label>
            </div>
           
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding}>
            <label>
              <span style={errorStyle}>*</span>
              Description
              <TextField
                id="txtDescription"
                multiline
                rows={4}
                placeholder="Enter text here"
               // label="*Description"
                className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6"
                value={this.listdata.Description}

              />
             </label>
            </div>
          
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"  style={divPadding}>
              <PrimaryButton
              type='Submit'
              text='Submit'
                // onClick={() => this._saveJMDConnectForm(event)}
              />
            </div>
           
          </div>
          
        </div>
      </div>
  </form>


    );

  }
}
