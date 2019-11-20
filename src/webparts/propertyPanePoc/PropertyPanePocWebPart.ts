import { escape } from "@microsoft/sp-lodash-subset";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";

import * as strings from "PropertyPanePocWebPartStrings";
import { PropertyPaneDateField } from "./customPropertyPaneFields/PropertyPaneDate/PropertyPaneDate";
import { IWebpartConfiguration } from "./models/WebpartConfiguration";

export default class PropertyPanePocWebPart extends BaseClientSideWebPart<IWebpartConfiguration> {

  public render(): void {
    this.domElement.innerHTML = `
    <div>
      <p>${escape(JSON.stringify(this.properties))}</p>
    </div>`
    ;
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
                PropertyPaneTextField("Description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneDateField("WebpartDate", {
                  DayLabel: "Tag",
                  MonthLabel: "Monat",
                  YearLabel: "Jahr",
                  ButtonLabel: "Bestätigen",
                  Value: this.properties.WebpartDate,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
