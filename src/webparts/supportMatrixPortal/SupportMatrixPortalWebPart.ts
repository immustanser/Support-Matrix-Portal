import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SupportMatrixPortalWebPartStrings';
import { SupportMatrixPortal, ISupportMatrixPortalProps } from './components/SupportMatrixPortal';
import { initializePnPjs } from './services/pnpjsConfig';

export interface ISupportMatrixPortalWebPartProps {
  description: string;
}

export default class SupportMatrixPortalWebPart extends BaseClientSideWebPart<ISupportMatrixPortalWebPartProps> {
  protected onInit(): Promise<void> {
    initializePnPjs(this.context);
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ISupportMatrixPortalProps> = React.createElement(SupportMatrixPortal, {
      context: this.context
    });

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
