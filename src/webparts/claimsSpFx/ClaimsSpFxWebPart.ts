import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'ClaimsSpFxWebPartStrings';
import ClaimsSpFx from './components/ClaimsSpFx';
import { IClaimsSpFxProps } from './components/IClaimsSpFxProps';

export interface IClaimsSpFxWebPartProps {
  description: string;
}

export default class ClaimsSpFxWebPart extends BaseClientSideWebPart<IClaimsSpFxWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IClaimsSpFxProps > = React.createElement(
      ClaimsSpFx,
      {
        context: this.context,
        description: this.properties.description,
        siteUrl: this.context.pageContext.web.absoluteUrl
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
