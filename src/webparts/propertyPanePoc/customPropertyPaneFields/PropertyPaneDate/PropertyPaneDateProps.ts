import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-webpart-base";

export interface IPropertyPaneDateProps {

    Value: Date;

    Label: string;
    DayLabel: string;
    MonthLabel: string;
    YearLabel: string;
    ButtonLabel: string;
}

export interface IPropertyPaneDatePropsInternal extends IPropertyPaneDateProps, IPropertyPaneCustomFieldProps {
}
