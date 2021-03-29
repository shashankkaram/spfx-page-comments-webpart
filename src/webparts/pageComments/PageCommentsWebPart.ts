import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PageCommentsWebPartStrings';
import PageComments from './components/PageComments';
import { IPageCommentsProps } from './components/IPageCommentsProps';

import { PageCommentService } from "../pageComments/services/PageCommentService";
import { sp } from '@pnp/sp';

export interface IPageCommentsWebPartProps {
  description: string;
}

export default class PageCommentsWebPart extends BaseClientSideWebPart<IPageCommentsWebPartProps> {

  private _services: PageCommentService = null;

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
      this._services = new PageCommentService(this.context);
    });
  }

  public render(): void {
    const element: React.ReactElement<IPageCommentsProps> = React.createElement(
      PageComments,
      {
        context:this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
