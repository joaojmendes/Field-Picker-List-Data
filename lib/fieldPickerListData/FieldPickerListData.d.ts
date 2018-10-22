/// <reference types="react" />
import * as React from "react";
import { IFieldPickerListDataProps } from "./IFieldPickerListDataProps";
import { IFieldPickerListDataState } from "./IFieldPickerListDataState";
export declare class FieldPickerListData extends React.Component<IFieldPickerListDataProps, IFieldPickerListDataState> {
    private _value;
    private _spservice;
    constructor(props: IFieldPickerListDataProps);
    render(): React.ReactElement<IFieldPickerListDataProps>;
    private getTextFromItem(item);
    private onItemChanged(selectedItems);
    private onFilterChanged(filterText, tagList);
    private loadListItems(filterText);
}
