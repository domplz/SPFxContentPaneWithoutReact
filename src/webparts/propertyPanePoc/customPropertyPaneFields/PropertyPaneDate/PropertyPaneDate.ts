import { IPropertyPaneField, PropertyPaneFieldType } from "@microsoft/sp-webpart-base";
import { IPropertyPaneDateProps, IPropertyPaneDatePropsInternal } from "./PropertyPaneDateProps";

export class PropertyPaneDate implements IPropertyPaneField<IPropertyPaneDatePropsInternal> {
    public type: PropertyPaneFieldType;
    public targetProperty: string;
    public shouldFocus?: boolean;
    public properties: IPropertyPaneDatePropsInternal;

    private value: Date;

    constructor(targetProperty: string, configuration: IPropertyPaneDateProps) {
        this.targetProperty = targetProperty;
        this.shouldFocus = false;
        this.type = PropertyPaneFieldType.Custom;
        this.properties = {
            DayLabel: configuration.DayLabel,
            MonthLabel: configuration.MonthLabel,
            YearLabel: configuration.YearLabel,
            ButtonLabel: configuration.ButtonLabel,
            DefaultValue: configuration.DefaultValue,

            key: targetProperty,
            onRender: this.render.bind(this),
        };

        this.value = this.properties.DefaultValue;
    }

    private render(domElement: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        console.log("render", domElement, context, changeCallback);
        domElement.innerHTML = "";
        domElement.appendChild(this.createHtml(changeCallback));
    }

    private createHtml(changeCallback: (targetProperty?: string, newValue?: any) => void): HTMLElement {

        const yearInputElement = document.createElement("input");
        yearInputElement.type = "number";
        yearInputElement.min = "0";
        yearInputElement.max = "9999";
        yearInputElement.placeholder = this.properties.YearLabel;
        yearInputElement.value = this.value.getFullYear().toString();

        const monthInputElement = document.createElement("input");
        monthInputElement.type = "number";
        monthInputElement.min = "1";
        monthInputElement.max = "12";
        monthInputElement.placeholder = this.properties.MonthLabel;
        monthInputElement.value = (this.value.getMonth() + 1).toString();

        const dayInputElement = document.createElement("input");
        dayInputElement.type = "number";
        dayInputElement.min = "1";
        dayInputElement.max = "31";
        dayInputElement.placeholder = this.properties.DayLabel;
        dayInputElement.value = this.value.getDate().toString();

        const buttonElement = document.createElement("button");
        buttonElement.innerText = this.properties.ButtonLabel;
        buttonElement.onclick = () => {
            const yearValue = yearInputElement.value;
            const monthValue = monthInputElement.value;
            const dayValue = dayInputElement.value;

            const date = this.getDateIfValid(yearValue, monthValue, dayValue);

            if (date) {
                this.value = date;
                changeCallback(this.targetProperty, date);
            } else {
                alert("date invalid");
            }
        };

        const controlElement = document.createElement("div");
        controlElement.className = "custom-date-input";
        controlElement.appendChild(yearInputElement);
        controlElement.appendChild(monthInputElement);
        controlElement.appendChild(dayInputElement);
        controlElement.appendChild(buttonElement);

        return controlElement;
    }

    private getDateIfValid(yearString: string, monthString: string, dayString: string): Date | undefined {
        if (!isNaN(Number(yearString)) && !isNaN(Number(monthString)) && !isNaN(Number(dayString))) {
            const yearNbr = parseInt(yearString, 10);

            if (yearNbr > 0 && yearNbr < 9999) {
                const monthNbr = parseInt(monthString, 10);

                if (monthNbr > 0 && monthNbr < 13) {

                    const dayNbr = parseInt(dayString, 10);
                    if (dayNbr > 0 && dayNbr < 32) {

                        // first convert to UTC (ticks ?) then convert it back to a date object (which automatically adds the correct date offset)
                        return new Date(Date.UTC(yearNbr, monthNbr - 1, dayNbr));
                    }
                }
            }
        }

        return undefined;
    }
}

export function PropertyPaneDateField(targetProperty: string, properties: IPropertyPaneDateProps): IPropertyPaneField<IPropertyPaneDateProps> {

    // Initialize and render the properties
    return new PropertyPaneDate(targetProperty, { ...properties });
}
