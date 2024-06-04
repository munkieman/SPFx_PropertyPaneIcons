import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PropertyPaneIconsWebPartStrings';
import PropertyPaneIcons from './components/PropertyPaneIcons';
import { IPropertyPaneIconsProps } from './components/IPropertyPaneIconsProps';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';

export interface IPropertyPaneIconsWebPartProps {
  description: string;
  iconPicker1: any;
  iconPicker2: any;
  iconPicker3: any;
}

export default class PropertyPaneIconsWebPart extends BaseClientSideWebPart<IPropertyPaneIconsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IPropertyPaneIconsProps> = React.createElement(
      PropertyPaneIcons,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        iconPicker1: this.properties.iconPicker1,
        iconPicker2: this.properties.iconPicker2,
        iconPicker3: this.properties.iconPicker3
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
                }),
                PropertyFieldIconPicker('iconPicker', {
                  currentIcon: this.properties.iconPicker1,
                  key: "iconPickerId",
                  onSave: (icon: string) => { console.log(icon); this.properties.iconPicker1 = icon; },
                  onChanged:(icon: string) => { console.log(icon);  },
                  buttonLabel: "Icon",
                  renderOption: "panel",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  label: "Icon Picker"              
                }),
                PropertyFieldIconPicker('iconPicker', {
                  currentIcon: this.properties.iconPicker2,
                  key: "iconPickerId",
                  onSave: (icon: string) => { console.log(icon); this.properties.iconPicker2 = icon; },
                  onChanged:(icon: string) => { console.log(icon);  },
                  buttonLabel: "Icon",
                  renderOption: "panel",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  label: "Icon Picker"              
                }),          
                PropertyFieldIconPicker('iconPicker', {
                  currentIcon: this.properties.iconPicker3,
                  key: "iconPickerId",
                  onSave: (icon: string) => { console.log(icon); this.properties.iconPicker3 = icon; },
                  onChanged:(icon: string) => { console.log(icon);  },
                  buttonLabel: "Icon",
                  renderOption: "panel",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  label: "Icon Picker"              
                })          
              ]
            }
          ]
        }
      ]
    };
  }
}
