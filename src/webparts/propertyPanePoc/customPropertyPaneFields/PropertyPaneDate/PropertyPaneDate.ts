import { throttle } from "@microsoft/sp-lodash-subset";
import { IPropertyPaneField, PropertyPaneFieldType } from "@microsoft/sp-webpart-base";
import { IPropertyPaneDateProps, IPropertyPaneDatePropsInternal } from "./PropertyPaneDateProps";

import styles from "./styles.module.scss";

export class PropertyPaneDate implements IPropertyPaneField<IPropertyPaneDatePropsInternal> {
    public type: PropertyPaneFieldType;
    public targetProperty: string;
    public shouldFocus?: boolean;
    public properties: IPropertyPaneDatePropsInternal;

    private yearInputElement: HTMLInputElement;
    private monthInputElement: HTMLInputElement;
    private dayInputElement: HTMLInputElement;

    private validationElement: HTMLDivElement;

    private changeCallback?: (targetProperty?: string, newValue?: any) => void;

    constructor(targetProperty: string, configuration: IPropertyPaneDateProps) {

        this.targetProperty = targetProperty;
        this.shouldFocus = false;
        this.type = PropertyPaneFieldType.Custom;
        this.properties = {
            Label: configuration.Label,
            DayLabel: configuration.DayLabel,
            MonthLabel: configuration.MonthLabel,
            YearLabel: configuration.YearLabel,
            ButtonLabel: configuration.ButtonLabel,
            Value: configuration.Value ? new Date(configuration.Value.toString()) : new Date(2000, 0, 1),

            key: targetProperty,
            onRender: this.render.bind(this),
        };
    }

    private render(domElement: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        console.log("render", domElement, context, changeCallback);

        this.changeCallback = changeCallback;

        domElement.innerHTML = "";
        domElement.appendChild(this.createHtml());
    }

    private createHtml(): HTMLElement {

        const labelElement = document.createElement("label");
        labelElement.innerText = this.properties.Label;

        const tempValidationElement = document.createElement("div");
        tempValidationElement.className = styles.validationMessageError;

        this.validationElement = tempValidationElement;

        const dateInputElement = document.createElement("div");
        dateInputElement.className = styles.dateInput;

        this.yearInputElement = this.createNumberInputField(this.properties.Value.getFullYear().toString(), this.properties.YearLabel, 0, 9999);
        this.monthInputElement = this.createNumberInputField((this.properties.Value.getMonth() + 1).toString(), this.properties.MonthLabel, 1, 12);
        this.dayInputElement = this.createNumberInputField(this.properties.Value.getDate().toString(), this.properties.DayLabel, 1, 31);

        dateInputElement.appendChild(this.yearInputElement);
        dateInputElement.appendChild(this.monthInputElement);
        dateInputElement.appendChild(this.dayInputElement);

        const controlElement = document.createElement("div");
        controlElement.className = styles.propertyPaneDate;
        controlElement.appendChild(labelElement);
        controlElement.appendChild(dateInputElement);

        controlElement.appendChild(this.validationElement);

        return controlElement;
    }

    private handleInputChange(): void {

        const yearValue = this.yearInputElement.value;
        const monthValue = this.monthInputElement.value;
        const dayValue = this.dayInputElement.value;

        const date = this.getDateIfValid(yearValue, monthValue, dayValue);

        if (date) {

            if (this.changeCallback && typeof this.changeCallback === "function") {

                // this.properties.Value = date;
                this.changeCallback(this.targetProperty, date);
            }
        } else {
            this.validationElement.innerHTML = "Date is invalid";
        }
    }

    private getDateIfValid(yearString: string, monthString: string, dayString: string): Date | undefined {
        if (!isNaN(Number(yearString)) && !isNaN(Number(monthString)) && !isNaN(Number(dayString))) {
            const yearNbr = parseInt(yearString, 10);

            if (yearNbr > 1000 && yearNbr < 9999) {
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

    private createNumberInputField(value: string, placeholder: string, min: number, max: number): HTMLInputElement {

        const inputField = document.createElement("input");
        inputField.type = "number";
        inputField.min = min.toString();
        inputField.max = max.toString();
        inputField.maxLength = max.toString().length;
        inputField.placeholder = placeholder;
        inputField.value = value;
        inputField.oninput = throttle( () => this.handleInputChange(), 1000, {trailing: true, leading: false });

        return inputField;
    }
}

export function PropertyPaneDateField(targetProperty: string, properties: IPropertyPaneDateProps): IPropertyPaneField<IPropertyPaneDateProps> {

    // Initialize and render the properties
    return new PropertyPaneDate(targetProperty, { ...properties });
}
