import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp';

import * as strings from 'EmployeeDirectoryWebPartStrings';
import EmployeeDirectory from './components/EmployeeDirectory';
import { IEmployeeDirectoryProps } from './components/IEmployeeDirectoryProps';
import { IEmployeeDirectoryWebPartProps } from './IEmployeeDirectoryWebPartProps';

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IEmployeeDirectoryProps > = React.createElement(
      EmployeeDirectory,
      {
        title: this.properties.title,
        columns: this.properties.columns,
        exclude: this.properties.exclude,
        sortBy: this.properties.sortBy,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        context: this.context
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
              groupFields: [
                PropertyPaneSlider('columns', {
                  max: 6,
                  min: 1,
                  label: strings.ColumnsFieldLabel
                }),
                PropertyPaneTextField('exclude', {
                  multiline: true,
                  placeholder: strings.ExcludeFieldPlaceholder,
                  label: strings.ExcludeFieldLabel,
                  rows: 8
                }),
                PropertyPaneDropdown('sortBy', {
                  label: strings.SortByFieldLabel,
                  options: [
                    { key: 'Title', text: 'Name' },
                    { key: 'EMail', text: 'Email' },
                    { key: 'Department', text: 'Department' },
                    { key: 'JobTitle', text: 'Job Title' },
                    { key: 'Office', text: 'Office' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
}
