import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
export interface IAlfaFaqWebPartProps {
    listId: string;
    accordionTitle: string;
    columnTitle: string;
    allowZeroExpanded: boolean;
    allowMultipleExpanded: boolean;
    accordianTitleColumn: string;
    accordianContentColumn: string;
    accordianSortColumn: string;
    isSortDescending: false;
    webhookUrl: string;
    enableLogging: boolean;
}
export default class AlfaFaqWebPart extends BaseClientSideWebPart<IAlfaFaqWebPartProps> {
    private _sp;
    private listColumns;
    spfxContext: any;
    private allListColumns;
    private columnChoices;
    private columnsDropdownDisabled;
    private choicesDropdownDisabled;
    protected onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected get disableReactivePropertyChanges(): boolean;
    protected get dataVersion(): Version;
    private loadColumns;
    private loadAllColumns;
    private loadCateogryChoices;
    protected onPropertyPaneConfigurationStart(): void;
    protected onPropertyPaneFieldChanged(): void;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
