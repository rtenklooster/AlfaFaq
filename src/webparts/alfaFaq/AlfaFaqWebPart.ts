import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";

import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";

import * as strings from "AlfaFaqWebPartStrings";
import AlfaFaq from "./components/AlfaFaq";
import { IAlfaFaqProps } from "./components/IAlfaFaqProps";
import { getSP } from "../../utils/pnpjs-config"
import { SPFI } from "@pnp/sp";
import { IField, IFieldInfo } from "@pnp/sp/fields";


export interface IAlfaFaqWebPartProps {
  listId: string;
  accordionTitle: string;
  columnTitle: string;
  //selectedChoice: string;
  allowZeroExpanded: boolean;
  allowMultipleExpanded: boolean;
  accordianTitleColumn: string;
  accordianContentColumn: string;
  accordianSortColumn: string;
  isSortDescending: false;
  webhookUrl: string;
  enableLogging: boolean;
}

export default class AlfaFaqWebPart extends BaseClientSideWebPart<
  IAlfaFaqWebPartProps
> {

  private _sp: SPFI;
  
  private listColumns: IPropertyPaneDropdownOption[];spfxContext
  private allListColumns: IPropertyPaneDropdownOption[];
  private columnChoices: IPropertyPaneDropdownOption[];

  private columnsDropdownDisabled = true;
  private choicesDropdownDisabled = true;

  protected async onInit(): Promise<void> {
    super.onInit();
  
    //Initialize our _sp object that we can then use in other packages without having to pass around the context.
    //  Check out pnpjsConfig.ts for an example of a project setup file.
    this._sp = getSP(this.context);
  }

  public render(): void {
    const element: React.ReactElement<IAlfaFaqProps> = React.createElement(
      AlfaFaq,
      {
        listId: this.properties.listId,
        columnTitle: this.properties.columnTitle,
        //selectedChoice: this.properties.selectedChoice,
        accordionTitle: this.properties.accordionTitle,
        accordianTitleColumn: this.properties.accordianTitleColumn,
        accordianContentColumn: this.properties.accordianContentColumn,
        accordianSortColumn: this.properties.accordianSortColumn,
        isSortDescending: this.properties.isSortDescending,
        allowZeroExpanded: this.properties.allowZeroExpanded,
        allowMultipleExpanded: this.properties.allowMultipleExpanded,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.accordionTitle = value;
        },
        onConfigure: () => {
          this.context.propertyPane.open();
        },
        webhookUrl: this.properties.webhookUrl,
        enableLogging: this.properties.enableLogging,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  private loadColumns(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error) => void
      ) => {
        if (!this.properties.listId) {
          console.log("Geen lijst geselecteerd");
          return null;
        }

        const spListColumns = this._sp.web.lists
          .getById(this.properties.listId)
          .fields.filter(
            "ReadOnlyField eq false and Hidden eq false and (TypeAsString eq 'Choice' or TypeAsString eq 'MultiChoice')"
          )
          ();
        spListColumns.then((columnResult) => {
          const listColumns = [];
          columnResult.forEach((column) => {
            listColumns.push({
              key: column.Title,
              text: column.Title + (column.TypeAsString === 'MultiChoice' ? ' (Multi-select)' : ''),
            });
          });
          resolve(listColumns);
        }).catch((error) => {
          reject(error);
        });
      }
    );
  }

  private loadAllColumns(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error) => void
      ) => {
        if (!this.properties.listId) {
          console.log("Geen lijst geselecteerd");
          return null;
        }

        const spListColumns = this._sp.web.lists
          .getById(this.properties.listId)
          .fields.filter("ReadOnlyField eq false and Hidden eq false")
          ();
        spListColumns.then((columnResult) => {
          const listColumns = [];
          columnResult.forEach((column) => {
            listColumns.push({
              key: column.InternalName,
              text: column.Title + " - [" + column.InternalName + "]",
            });
          });
          resolve(listColumns);
        }).catch((error) => {
          reject(error);
        });
      }
    );
  }

  private loadCateogryChoices(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error) => void
      ) => {
        if (!this.properties.columnTitle) {
          console.log("No Columns Selected");
          return null;
        }

        const categoryField: IField = this._sp.web.lists
          .getById(this.properties.listId)
          .fields.getByInternalNameOrTitle(this.properties.columnTitle);
          
        const choices: Promise<IFieldInfo> = categoryField.select("Choices")();
        choices.then((result) => {
          // console.clear();
          // console.log(result.Choices);
          const columnChoices = [];
          result.Choices.forEach((choice) => {
            columnChoices.push({
              key: choice,
              text: choice,
            });
          });
          resolve(columnChoices);
        }).catch((error) => {
          reject(error);
        });
      }
    );
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.columnsDropdownDisabled = !this.properties.listId;
    this.choicesDropdownDisabled = !this.properties.columnTitle;

    //if (this.lists) {
    //  return;
    //}

    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "lists, column and choices"
    );
    if (this.properties.listId) {
      this.loadColumns().then(
        (columnOptions: IPropertyPaneDropdownOption[]): void => {
          this.listColumns = columnOptions;
          this.columnsDropdownDisabled = !this.properties.listId;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        }
      );
      this.loadAllColumns().then(
        (allcolumnOptions: IPropertyPaneDropdownOption[]): void => {
          this.allListColumns = allcolumnOptions;
          this.columnsDropdownDisabled = !this.properties.listId;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        }
      );
    }
    if (this.properties.columnTitle) {
      this.loadCateogryChoices().then(
        (choiceOptions: IPropertyPaneDropdownOption[]): void => {
          this.columnChoices = choiceOptions;
          this.choicesDropdownDisabled = !this.properties.columnTitle;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        }
      );
    }
  }

  protected onPropertyPaneFieldChanged(): void {
    if (this.properties.listId) {
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Columns"
      );

      this.loadColumns().then(
        (columnOptions: IPropertyPaneDropdownOption[]): void => {
          // store items
          this.listColumns = columnOptions;
          // enable item selector
          this.columnsDropdownDisabled = false;
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        }
      );
      this.loadAllColumns().then(
        (allcolumnOptions: IPropertyPaneDropdownOption[]): void => {
          this.allListColumns = allcolumnOptions;
          this.columnsDropdownDisabled = !this.properties.listId;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        }
      );
    }

    if (this.properties.columnTitle) {
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Choices"
      );
      this.loadCateogryChoices().then(
        (choiceOption: IPropertyPaneDropdownOption[]): void => {
          // store items
          this.columnChoices = choiceOption;
          // enable item selector
          this.choicesDropdownDisabled = false;
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        }
      );
    }

    if (this.properties.listId) {
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Data"
      );
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
      this.context.propertyPane.refresh();
    }
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
                PropertyFieldListPicker("listId", {
                  label: "Selecteer een lijst",
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyPaneDropdown("columnTitle", {
                  label: "Kies de keuze kolom voor de categorieën",
                  options: this.listColumns,
                  disabled: this.columnsDropdownDisabled,
                }),
                PropertyPaneDropdown("accordianTitleColumn", {
                  label: "Kies de kolom voor de vraag",
                  options: this.allListColumns,
                  disabled: this.choicesDropdownDisabled,
                }),
                PropertyPaneDropdown("accordianContentColumn", {
                  label: "Kies de kolom voor het antwoord",
                  options: this.allListColumns,
                  disabled: this.choicesDropdownDisabled,
                }),
                PropertyPaneDropdown("accordianSortColumn", {
                  label: "Kies de kolom waarop moet worden gesorteerd",
                  options: this.allListColumns,
                  disabled: this.choicesDropdownDisabled,
                }),
                PropertyPaneToggle("isSortDescending", {
                  label: "Sorteer oplopend of aflopend",
                  onText: "Oplopend", 
                  offText: "Aflopend",
                  disabled: !this.properties.accordianSortColumn
                }),
                PropertyPaneToggle("allowZeroExpanded", {
                  label: "Sta geen uitgeklapte items toe",
                  checked: this.properties.allowZeroExpanded,
                  key: "allowZeroExpanded",
                }),
                PropertyPaneToggle("allowMultipleExpanded", {
                  label: "Sta meerdere uitgeklapte items toe",
                  checked: this.properties.allowMultipleExpanded,
                  key: "allowMultipleExpanded",
                }),
                PropertyPaneTextField("webhookUrl", {
                  label: "Webhook URL",
                  value: this.properties.webhookUrl,
                }),
                PropertyPaneToggle("enableLogging", {
                  label: "Automatisch loggen inschakelen",
                  onText: "Ingeschakeld",
                  offText: "Uitgeschakeld",
                  checked: this.properties.enableLogging,
                })
              ],
            },
          ],
        },
      ],
    };
  }
}
