import { DisplayMode } from "@microsoft/sp-core-library";
export interface IAlfaFaqProps {
    listId: string;
    accordionTitle: string;
    columnTitle: string;
    accordianTitleColumn: string;
    accordianContentColumn: string;
    accordianSortColumn: string;
    isSortDescending: boolean;
    allowZeroExpanded: boolean;
    allowMultipleExpanded: boolean;
    displayMode: DisplayMode;
    updateProperty: (value: string) => void;
    onConfigure: () => void;
    webhookUrl: string;
    enableLogging: boolean;
}
