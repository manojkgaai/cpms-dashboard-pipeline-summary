import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DashboardSummaryWebPartStrings';
import DashboardSummary from './components/DashboardSummary';
import { IDashboardSummaryProps } from './components/IDashboardSummaryProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface IDashboardSummaryWebPartProps {
  wptitle: string;
  lists: string;
  webUrl: string;
  onboardfield: string;
  startdtfield: string;
  npmfield: string;
  linkPageUrl: string;
}

export default class DashboardSummaryWebPart extends BaseClientSideWebPart <IDashboardSummaryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDashboardSummaryProps> = React.createElement(
      DashboardSummary,
      {
        wptitle: this.properties.wptitle,
        context:this.context,
        list: this.properties.lists,
        webUrl: this.context.pageContext.web.absoluteUrl,
        onboardfield: this.properties.onboardfield,
        startdtfield: this.properties.startdtfield,
        npmfield: this.properties.npmfield,
        linkPageUrl: this.properties.linkPageUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private listConfigurationChanged(propertyPath: string, oldValue: any, newValue: any) {  
    if (propertyPath === 'lists' && newValue) {  
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);  
      this.context.propertyPane.refresh();  
    }  
    else {  
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);  
    }  
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
                PropertyPaneTextField('wptitle', {
                  label: strings.WebPartTitleLabel
                }),
                PropertyFieldListPicker('lists', {
                  context: this.context,
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Id,
                  onPropertyChange: this.listConfigurationChanged.bind(this),
                  disabled: false,
                  baseTemplate: 100,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneTextField('onboardfield', {
                  label: strings.OnboardFieldLabel
                }),
                PropertyPaneTextField('startdtfield', {
                  label: strings.StartDateFieldLabel
                }),
                PropertyPaneTextField('npmfield', {
                  label: strings.NPMFieldLabel
                }),
                PropertyPaneTextField('linkPageUrl', {
                  label: strings.LinkPageUrlLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
