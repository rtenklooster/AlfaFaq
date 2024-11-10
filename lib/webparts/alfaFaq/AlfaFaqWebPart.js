var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { PropertyPaneToggle, PropertyPaneDropdown, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy, } from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import * as strings from "AlfaFaqWebPartStrings";
import AlfaFaq from "./components/AlfaFaq";
import { getSP } from "../../utils/pnpjs-config";
var AlfaFaqWebPart = /** @class */ (function (_super) {
    __extends(AlfaFaqWebPart, _super);
    function AlfaFaqWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.columnsDropdownDisabled = true;
        _this.choicesDropdownDisabled = true;
        return _this;
    }
    AlfaFaqWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                _super.prototype.onInit.call(this);
                //Initialize our _sp object that we can then use in other packages without having to pass around the context.
                //  Check out pnpjsConfig.ts for an example of a project setup file.
                this._sp = getSP(this.context);
                return [2 /*return*/];
            });
        });
    };
    AlfaFaqWebPart.prototype.render = function () {
        var _this = this;
        var element = React.createElement(AlfaFaq, {
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
            updateProperty: function (value) {
                _this.properties.accordionTitle = value;
            },
            onConfigure: function () {
                _this.context.propertyPane.open();
            },
            webhookUrl: this.properties.webhookUrl,
            enableLogging: this.properties.enableLogging,
        });
        ReactDom.render(element, this.domElement);
    };
    AlfaFaqWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(AlfaFaqWebPart.prototype, "disableReactivePropertyChanges", {
        get: function () {
            return true;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(AlfaFaqWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse("1.0");
        },
        enumerable: false,
        configurable: true
    });
    AlfaFaqWebPart.prototype.loadColumns = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (!_this.properties.listId) {
                console.log("Geen lijst geselecteerd");
                return null;
            }
            var spListColumns = _this._sp.web.lists
                .getById(_this.properties.listId)
                .fields.filter("ReadOnlyField eq false and Hidden eq false and TypeAsString eq 'Choice'")();
            spListColumns.then(function (columnResult) {
                var listColumns = [];
                columnResult.forEach(function (column) {
                    listColumns.push({
                        key: column.Title,
                        text: column.Title,
                    });
                });
                resolve(listColumns);
            }).catch(function (error) {
                reject(error);
            });
        });
    };
    AlfaFaqWebPart.prototype.loadAllColumns = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (!_this.properties.listId) {
                console.log("Geen lijst geselecteerd");
                return null;
            }
            var spListColumns = _this._sp.web.lists
                .getById(_this.properties.listId)
                .fields.filter("ReadOnlyField eq false and Hidden eq false")();
            spListColumns.then(function (columnResult) {
                var listColumns = [];
                columnResult.forEach(function (column) {
                    listColumns.push({
                        key: column.InternalName,
                        text: column.Title + " - [" + column.InternalName + "]",
                    });
                });
                resolve(listColumns);
            }).catch(function (error) {
                reject(error);
            });
        });
    };
    AlfaFaqWebPart.prototype.loadCateogryChoices = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (!_this.properties.columnTitle) {
                console.log("No Columns Selected");
                return null;
            }
            var categoryField = _this._sp.web.lists
                .getById(_this.properties.listId)
                .fields.getByInternalNameOrTitle(_this.properties.columnTitle);
            var choices = categoryField.select("Choices")();
            choices.then(function (result) {
                // console.clear();
                // console.log(result.Choices);
                var columnChoices = [];
                result.Choices.forEach(function (choice) {
                    columnChoices.push({
                        key: choice,
                        text: choice,
                    });
                });
                resolve(columnChoices);
            }).catch(function (error) {
                reject(error);
            });
        });
    };
    AlfaFaqWebPart.prototype.onPropertyPaneConfigurationStart = function () {
        var _this = this;
        this.columnsDropdownDisabled = !this.properties.listId;
        this.choicesDropdownDisabled = !this.properties.columnTitle;
        //if (this.lists) {
        //  return;
        //}
        this.context.statusRenderer.displayLoadingIndicator(this.domElement, "lists, column and choices");
        if (this.properties.listId) {
            this.loadColumns().then(function (columnOptions) {
                _this.listColumns = columnOptions;
                _this.columnsDropdownDisabled = !_this.properties.listId;
                _this.context.propertyPane.refresh();
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                _this.render();
            });
            this.loadAllColumns().then(function (allcolumnOptions) {
                _this.allListColumns = allcolumnOptions;
                _this.columnsDropdownDisabled = !_this.properties.listId;
                _this.context.propertyPane.refresh();
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                _this.render();
            });
        }
        if (this.properties.columnTitle) {
            this.loadCateogryChoices().then(function (choiceOptions) {
                _this.columnChoices = choiceOptions;
                _this.choicesDropdownDisabled = !_this.properties.columnTitle;
                _this.context.propertyPane.refresh();
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                _this.render();
            });
        }
    };
    AlfaFaqWebPart.prototype.onPropertyPaneFieldChanged = function () {
        var _this = this;
        if (this.properties.listId) {
            this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Columns");
            this.loadColumns().then(function (columnOptions) {
                // store items
                _this.listColumns = columnOptions;
                // enable item selector
                _this.columnsDropdownDisabled = false;
                // clear status indicator
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                // re-render the web part as clearing the loading indicator removes the web part body
                _this.render();
                // refresh the item selector control by repainting the property pane
                _this.context.propertyPane.refresh();
            });
            this.loadAllColumns().then(function (allcolumnOptions) {
                _this.allListColumns = allcolumnOptions;
                _this.columnsDropdownDisabled = !_this.properties.listId;
                _this.context.propertyPane.refresh();
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                _this.render();
            });
        }
        if (this.properties.columnTitle) {
            this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Choices");
            this.loadCateogryChoices().then(function (choiceOption) {
                // store items
                _this.columnChoices = choiceOption;
                // enable item selector
                _this.choicesDropdownDisabled = false;
                // clear status indicator
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                // re-render the web part as clearing the loading indicator removes the web part body
                _this.render();
                // refresh the item selector control by repainting the property pane
                _this.context.propertyPane.refresh();
            });
        }
        if (this.properties.listId) {
            this.context.statusRenderer.displayLoadingIndicator(this.domElement, "Data");
            this.context.statusRenderer.clearLoadingIndicator(this.domElement);
            this.render();
            this.context.propertyPane.refresh();
        }
    };
    AlfaFaqWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return AlfaFaqWebPart;
}(BaseClientSideWebPart));
export default AlfaFaqWebPart;
//# sourceMappingURL=AlfaFaqWebPart.js.map