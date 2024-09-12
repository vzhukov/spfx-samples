import * as React from 'react';
import * as ReactDom from 'react-dom';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import DataverseApiWebPart from '../components/DataverseComponent';
import { IDataverseComponentProps } from '../components/IDataverseComponentProps';
import DataverseService, { IDataverseService } from '../services/DataverseService';
import { IDataverseWebPartProps } from './IDataverseWebPartProps';

export default class DataverseWebPart extends BaseClientSideWebPart<IDataverseWebPartProps> {
  private _dataverseService: IDataverseService;

  protected async onInit(): Promise<void> {
    // Consume the service
    this._dataverseService = this.context.serviceScope.consume(DataverseService.serviceKey);

    // Provide Dataverse URL to the service
    this._dataverseService.instanceUrl = this.properties.instanceUrl;
  }

  public render(): void {
    const element: React.ReactElement<IDataverseComponentProps> = React.createElement(
      DataverseApiWebPart,
      {
        dataverseService: this._dataverseService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, _oldValue: any, newValue: any): void {
    // Provide updated Dataverse URL to the service if it's changed
    if (propertyPath === "instanceUrl"
      && this._dataverseService.instanceUrl !== newValue) {
      this._dataverseService.instanceUrl = newValue;
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('instanceUrl', {
                  label: "Instance URL",
                  placeholder: "https://instance.region.dynamics.com"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
