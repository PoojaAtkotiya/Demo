import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MyPendingRequestsWebPartStrings';
import MyPendingRequests from './components/MyPendingRequests';
import { IMyPendingRequestsProps } from './components/IMyPendingRequestsProps';
import pnp from "sp-pnp-js";

export interface IMyPendingRequestsWebPartProps {
  description: string;
}

export default class MyPendingRequestsWebPart extends BaseClientSideWebPart<IMyPendingRequestsWebPartProps> {

  public onInit(): Promise<void> {
    
    return super.onInit().then(_ => {
  
      pnp.setup({
        spfxContext: this.context
      });
      
    });
  }

  public render(): void {
    const element: React.ReactElement<{} > = React.createElement(
      MyPendingRequests
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
