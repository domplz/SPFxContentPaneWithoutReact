import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-webpart-base";

export interface IPropertyPaneDateProps {
    DayLabel: string;
    MonthLabel: string;
    YearLabel: string;
    ButtonLabel: string;
    DefaultValue: Date;
}

export interface IPropertyPaneDatePropsInternal extends IPropertyPaneDateProps, IPropertyPaneCustomFieldProps {
}
