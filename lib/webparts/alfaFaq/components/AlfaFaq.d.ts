import * as React from "react";
import { IAlfaFaqProps } from "./IAlfaFaqProps";
import "./alfaFaq.css";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
interface IAccordionItem {
    Id: string;
    Title: string;
    Content: string;
    Category: string;
    helpful?: boolean;
}
export interface IAlfaFaqState {
    items: Array<IAccordionItem>;
    choices: Array<string>;
    allowMultipleExpanded: boolean;
    allowZeroExpanded: boolean;
    searchText: string;
    userEmail: string;
    expandedItems: string[];
}
export default class AlfaFaq extends React.Component<IAlfaFaqProps, IAlfaFaqState> {
    private _sp;
    constructor(props: IAlfaFaqProps);
    private getUserEmail;
    private onSearchTextChange;
    private updateExpandedItems;
    private highlightText;
    private getListItems;
    private logItemOpen;
    componentDidUpdate(prevProps: IAlfaFaqProps): void;
    render(): React.ReactElement<IAlfaFaqProps>;
}
export {};
