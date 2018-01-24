import * as React from 'react';
import styles from './FbForm.module.scss';
import { IFbFormProps, ISPList, IKeyText } from './IFbFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField, PrimaryButton, DefaultButton, Label, Button } from 'office-ui-fabric-react/lib/';
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
import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';


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
  //private statusOption: IDropdownOption[];
  private listdata: ISPList;
  queryParameters: UrlQueryParameterCollection
  web: any;
  hideDialog: string;
  private errorDialog: boolean;
  private confirmDialog: boolean;
  private messegeDialog: boolean;

  constructor() {
    super();

  }
  private errorMsg: string[] = [" "];
  private saveMsg: string = "";
  private saveMsgTitle: string = "Success";

  componentWillMount() {


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
  }



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
    return this.web.lists.getByTitle("CategoryMaster").items.orderBy("Title").get().then((response) => {
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
    return this.web.lists.getByTitle("SubCategoryMaster").items.filter("Category eq " + categoryTxt).orderBy("Title").get().then((response) => {
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

  // private dialog(){
  //   console.log(this.hideDialog);
  //   this.errorDialog=true;
  //   console.log(this.errorDialog);
  // }

  @autobind
  private _showDialog() {
    this.errorMsg = [];
    console.log("_title = " + document.getElementById("txtTitle")["value"]);
    const _title = document.getElementById("txtTitle")["value"].trim();    // event.target['txtTitle'].value.trim();
    const _description = document.getElementById('txtDescription')["value"].trim();
    const _category = document.getElementById("txtCategory-option").textContent;
    const _subCategory = document.getElementById("txtSubCategory-option").textContent;

    var isValid: boolean = true;

    if (_category == '' || _category == null || _category == undefined || _category.trim().length <= 0 || _category.toLowerCase().match("select an option")) {
      document.getElementById("txtCategory-option").style.borderColor = "red";
      document.getElementById("txtCategory-option").style.backgroundColor = "lightyellow";
      isValid = false;
      this.errorMsg.push("Category is required");
    }
    else {
      document.getElementById("txtCategory-option").style.backgroundColor = "white";
    }
    if (_subCategory == '' || _subCategory == null || _subCategory == undefined || _subCategory.trim().length <= 0 || _subCategory.toLowerCase().match("select an option")) {
      document.getElementById("txtSubCategory-option").style.borderColor = "red";
      document.getElementById("txtSubCategory-option").style.backgroundColor = "lightyellow";
      isValid = false;
      this.errorMsg.push("Sub Category is required");
    }
    else {
      document.getElementById("txtSubCategory-option").style.backgroundColor = "white";
    }
    if (_title == '' || _title == null || _title == undefined) {
      //  document.getElementById("txtTitle").parentElement.style.border ="solid thin red";
      document.getElementById("txtTitle").parentElement.parentElement.style.borderColor = "red";
      document.getElementById("txtTitle").style.backgroundColor = "lightyellow";
      isValid = false;
      this.errorMsg.push("Title is required");
    }
    else if (_title.length > 255) {
      //document.getElementById("txtTitle").parentElement.style.border ="solid thin red";
      document.getElementById("txtTitle").parentElement.parentElement.style.borderColor = "red";
      document.getElementById("txtTitle").style.backgroundColor = "lightyellow";
      isValid = false;
      this.errorMsg.push("Title should not be longer than 255 characters.");
    }
    else {
      document.getElementById("txtTitle").parentElement.parentElement.style.borderColor = "";
      // document.getElementById("txtTitle").parentElement.style.border ="";
      document.getElementById("txtTitle").style.backgroundColor = "white";
    }
    if (_description == '' || _description == null || _description == undefined) {

      // document.getElementById("txtDescription").parentElement.style.border ="solid thin red";
      document.getElementById("txtDescription").parentElement.parentElement.style.borderColor = "red";
      document.getElementById("txtDescription").style.backgroundColor = "lightyellow";
      isValid = false;
      this.errorMsg.push("Description is required.");
    }
    else if (_description.length > 1000) {

      // document.getElementById("txtDescription").parentElement.style.border ="solid thin red";
      document.getElementById("txtDescription").parentElement.parentElement.style.borderColor = "red";
      document.getElementById("txtDescription").style.backgroundColor = "lightyellow";
      isValid = false;
      this.errorMsg.push("Description should not be longer than 1000 characters.");
    }
    else {
      //  document.getElementById("txtDescription").parentElement.style.border ="";
      document.getElementById("txtDescription").parentElement.parentElement.style.borderColor = "";
      document.getElementById("txtDescription").style.backgroundColor = "white";
    }


    if (isValid) {
      this.hideDialog == "true";
      this.setState(this.hideDialog);
      this.confirmDialog = true;
    }
    else {
      this.hideDialog == "true";
      this.setState(this.hideDialog);
      this.errorDialog = true;
    }

  }


  @autobind
  private _closeDialog() {
    this.hideDialog == "false";
    this.setState(this.hideDialog);
    this.errorDialog = false;
    this.confirmDialog = false;
  }
  @autobind
  private _closemessegeDialog() {
    this.hideDialog == "false";
    this.setState(this.hideDialog);
    this.messegeDialog = false;
    window.location.href = "https://bajajelect.sharepoint.com/teams/ConnectApp";

  }

  hideConfirmDialog() {
    this.confirmDialog = false;
    this.hideDialog == "false";
    this.setState(this.hideDialog);
  }


  public _saveJMDConnectForm() {
    // console.log(this.refs.btnYes)
    // const btnYes = this.refs.btnYes as PrimaryButton;
    // btnYes.setState({ hidden : true });
  
    const _title = document.getElementById('txtTitle')['value'].trim();
    const _description = document.getElementById('txtDescription')['value'].trim();
    const _category = document.getElementById("txtCategory-option").textContent;
    const _subCategory = document.getElementById("txtSubCategory-option").textContent;

    this.web.lists.getByTitle('Connect%20Approval').items.add({
      Category: _category,
      Sub_x0020_Category: _subCategory,
      Title: _title,
      Feedback_x0020_Description: _description
    }).then((result): void => {

      this.hideConfirmDialog();
      // this.confirmDialog = false;
      // this.hideDialog=="false";
      // this.setState(this.hideDialog);



      this.saveMsgTitle = "Success";
      this.saveMsg = "Connect Call added successfully.";
      this.messegeDialog = true;
      this.hideDialog == "true";
      this.setState(this.hideDialog);



    }, (error: any): void => {
      console.log(error);
      this.saveMsgTitle = "Error";
      this.saveMsg = "Oops!Something went wrong!!!";
      this.hideDialog == "true";
      this.setState(this.hideDialog);
      this.messegeDialog = true;
    });

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

    const lblHeader = {
      width: '100%',
      padding: '5px 0',
      margin: 0,
      fontWeight: 400,
      fontFamily: '"Segoe UI Semibold WestEuropean","Segoe UI Semibold","Segoe UI",Tahoma,Arial,sans-serif',
      fontSize: '14px',
      color: "#002271"
    }

    const errorStyle = {
      color: 'red'
    };
    const divPadding = {
      padding: '3px 3px',
      position: 'relative',
    };

    const divPaddingTitleLabel = {
      padding: '3px 3px 0 3px',
      position: 'relative',
    };

    const divPaddingTitleTB = {
      padding: '0 3px 10px 3px',
      position: 'relative',
    };
    const btnStyle = {
      paddingTop: 10,
      paddingBottom: 7
    }

    var backgroudStyle = {
      backgroundImage: 'url(https://bajajelect.sharepoint.com/teams/ConnectApp/SiteAssets/Images/topography.png)',
      backgroundPosition: 'center',
      backgroundRepeat: 'no-repeat',
      backgroundSize: 'cover',
    }
    return (
      <form formNoValidate className={styles.fbForm} >

        <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={backgroudStyle}>

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding}>
            <label style={lblHeader}>
              <span style={errorStyle}>*</span>
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

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPadding}>
            <label style={lblHeader}>
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
          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPaddingTitleLabel}>
            <label style={lblHeader}>
              <span style={errorStyle}>*</span>
              Title
              </label>
          </div>
          <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6" style={divPaddingTitleTB}>
            <TextField
              id="txtTitle"
              underlined
              placeholder="Enter text here"
              maxLength={255}
              //label="*Title"
              className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6"
              value={this.listdata.Title}
            />
          </div>

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12" style={divPaddingTitleLabel}>
            <label style={lblHeader}>
              <span style={errorStyle}>*</span>
              Description
              </label>
          </div>
          <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6" style={divPaddingTitleTB}>
            <TextField
              id="txtDescription"
              multiline
              underlined
              rows={4}
              maxLength={1000}
              placeholder="Enter text here"
              // label="*Description"
              className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6"
              value={this.listdata.Description}

            />
          </div>

          <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
            <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6">
              <div className="ms-Grid-col ms-u-sm12 ms-u-md8 ms-u-lg8" style={btnStyle}>
                <PrimaryButton
                  // type='Submit'
                  style={{ backgroundColor: '#127316' }}
                  iconProps={{ iconName: 'Add' }}
                  text='Add Connect Call'
                  onClick={this._showDialog}
                />
              </div>
              <div className="ms-Grid-col ms-u-sm12 ms-u-md4 ms-u-lg4" style={btnStyle}>
                <Button
                  style={{ backgroundColor: '#a6a6a6' }}
                  text="Cancel"
                  iconProps={{ iconName: "Cancel" }}
                  href="https://bajajelect.sharepoint.com/teams/ConnectApp/" />
              </div>
            </div>
            <div className="ms-Grid-col ms-u-sm12 ms-u-md6 ms-u-lg6"></div>
          </div>
          <div></div>
        </div>
        <Dialog

          className="dialog"
          isOpen={this.confirmDialog}
          onDismiss={() => this._closeDialog()}
          title='Confirm'
          subText="Are you sure you want to add connect call?" >
          <DialogFooter>
            <PrimaryButton ref="btnYes"
              onClick={() => {
                this._saveJMDConnectForm()
                // return (
                //   <div>
                //     <Spinner size={SpinnerSize.large} label='Please wait, we are loading...' />
                //   </div>
                // )
              }
              } text='Yes' />
            <DefaultButton onClick={() => this._closeDialog()} text='No' />
          </DialogFooter>
        </Dialog>
        <Dialog

          className="dialog"
          isOpen={this.messegeDialog}
          onDismiss={() => this._closemessegeDialog()}
          title={this.saveMsgTitle}
          subText={this.saveMsg} >
        </Dialog>
        <Dialog

          className="dialog"
          isOpen={this.errorDialog}
          onDismiss={() => this._closeDialog()}
          title='Error'>
          <ul>
            {this.errorMsg.filter(x => { return (x !== (undefined || null || '' || ' ')) }).map(element => <li>{element}</li>)}
          </ul>
          <DialogFooter>
            <DefaultButton onClick={() => this._closeDialog()} text='Cancel' />
          </DialogFooter>
        </Dialog>


      </form>


    );

  }
}
