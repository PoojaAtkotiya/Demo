import * as React from 'react';
import styles from './MyPendingRequests.module.scss';
import { IMyPendingRequestsProps } from './IMyPendingRequestsProps';
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
import {
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';

let Count = 0;
let _items: {
  key: number,
  name: string,
  value: number,
  Title: string,
  Id: number,
  Category: string,
  Subcategory: string,
  description: string,
  Status: string,
  CreatedBy: string,
  CreatedDate: string,
  Approvers: string,
  ApproveRejectedBy: string,
  ApproverComment: string,
  ApproveRejectedDate: string,
  EditLink: string,
  Attachments: string
}[] = [];

let _Id: {
  UserID: number,

}[] = [];

let _columns = [
  {
    key: 'column13',
    name: 'Edit Link',
    fieldName: 'EditLink',
    minWidth: 50,
    maxWidth: 60,
    isResizable: true,
    onRender: item => (
      <Link data-selection-invoke={true} href={item.EditLink + item.Id}>
        Edit
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
    name: 'Approve/Rejected Date',
    fieldName: 'ApproveRejectedDate',
    minWidth: 70,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'column14',
    name: 'Attachment',
    fieldName: 'Attachments',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
    onRender: item => ( item.Attachments != null ?
      <Link data-selection-invoke={true} href={item.Attachments} target="_blank">
        Attachment
      </Link>
      : <span>No Attachment</span>
    )
  },
];


export default class MyPendingRequests extends React.Component<{}, { items: {}[]; }>
{
  public userId: any;
  constructor(props: {}) {
    super(props);
    pnp.sp.web.currentUser.get().then(function (res) {
      var isSuperUser: boolean = false;
      var currentUserId = res.Id;
      console.log("currentUserId = " + currentUserId);
      pnp.sp.web.siteGroups.getByName("Super User").users.get().then(function (usersList) {

        for (var i = 0; i < usersList.length; i++) {
          console.log("Title: " + usersList[i].Title);
          console.log("ID: " + usersList[i].Id);
          if (usersList[i].Title == res.Title) {
            isSuperUser = true;
          }
        }
        if (isSuperUser) {
          pnp.sp.web.lists.getByTitle('Connect%20Approval').items.expand("AttachmentFiles").orderBy("ID", false).filter("Status eq 'Approve' and SuperUserAcknowledged eq 'Assigned'").get().then(
            response => {
              response.map(item => {
                _items.push({
                  key: item.ID,
                  name: item.Title,
                  value: item.ID,
                  Title: item.Title,
                  Id: item.ID,
                  Category: item.Category,
                  Subcategory: item.Sub_x0020_Category,
                  description: item.Feedback_x0020_Description,
                  Status: item.Status,
                  CreatedBy: item.CreatedByDisplay,
                  CreatedDate: (item.Created) ? new Date(item.Created).toLocaleDateString("en-GB") : '',
                  Approvers: item.ApproversDispay,
                  ApproveRejectedBy: item.ApprovedByDisplay,
                  ApproverComment: item.Approver_x0020_Comments,
                  ApproveRejectedDate: (item.ApproveRejectedDate) ? new Date(item.ApproveRejectedDate).toLocaleDateString("en-GB") : '',
                  EditLink: "https://bajajelect.sharepoint.com/teams/ConnectApp/SitePages/SuperUserApproval.aspx?ConnectId=",
                  Attachments: (item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : null)
                })
              })
            }
          )
        }
        else {
          pnp.sp.web.lists.getByTitle('Connect%20Approval').items.expand("AttachmentFiles").orderBy("ID", false).filter("Status eq 'In Progress' and Approver eq " + currentUserId).get().then(
            response => {
              console.log(response);
              response.map(item => {
                _items.push({
                  key: item.ID,
                  name: item.Title,
                  value: item.ID,
                  Title: item.Title,
                  Id: item.ID,
                  Category: item.Category,
                  Subcategory: item.Sub_x0020_Category,
                  description: item.Feedback_x0020_Description,
                  Status: item.Status,
                  CreatedBy: item.CreatedByDisplay,
                  CreatedDate: (item.Created) ? new Date(item.Created).toLocaleDateString("en-GB") : '',
                  Approvers: item.ApproversDispay,
                  ApproveRejectedBy: item.ApprovedByDisplay,
                  ApproverComment: item.Approver_x0020_Comments,
                  ApproveRejectedDate: (item.ApproveRejectedDate) ? new Date(item.ApproveRejectedDate).toLocaleDateString("en-GB") : '',
                  EditLink: "https://bajajelect.sharepoint.com/teams/ConnectApp/SitePages/ApprovalForm.aspx?ApproverId=",
                  Attachments: (item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : null)
                })
              })
            }
          )

        }
        return _items;
      })
    })

    this.state = {
      items: _items
    };

  }

  public render() {
    if (_items.length === 0 && Count<7) {
      Count++;
      setTimeout(() => {
        this.setState({ items: _items })
      }, 500);
      if(Count>7)
      {
        return (
          <div>
            <label>No Data Available</label>
          </div>
        )     
      }
      return (
        <div>
          <Spinner size={SpinnerSize.large} label='Please wait, we are loading...' />
        </div>
      )
    }

    let { items } = this.state;
    return (
      <div>
        <DetailsList
          items={items}
          columns={_columns}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          checkboxVisibility={CheckboxVisibility.hidden}
        />

      </div>
    );
  }
}
