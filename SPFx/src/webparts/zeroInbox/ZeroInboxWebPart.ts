import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "ZeroInboxWebPartStrings";
import ZeroInbox from "./components/ZeroInbox";
import { IZeroInboxProps } from "./components/IZeroInboxProps";
import { Providers, SharePointProvider, TeamsProvider } from "@microsoft/mgt";

export interface IZeroInboxWebPartProps {
  description: string;
}

export default class ZeroInboxWebPart extends BaseClientSideWebPart<IZeroInboxWebPartProps> {
  protected async onInit(): Promise<void> {
    let prov = new SharePointProvider(this.context);
    Providers.globalProvider = prov;
  }

  public render(): void {
    const element: React.ReactElement<IZeroInboxProps> = React.createElement(ZeroInbox, {
      description: this.properties.description,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
