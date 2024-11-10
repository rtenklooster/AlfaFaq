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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
import * as React from "react";
import styles from "./AlfaFaq.module.scss";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import "./alfaFaq.css";
import { Pivot, PivotItem, TextField } from '@fluentui/react';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { Accordion, AccordionItem, AccordionItemHeading, AccordionItemButton, AccordionItemPanel, } from "react-accessible-accordion";
import { getSP } from "../../../utils/pnpjs-config";
var AlfaFaq = /** @class */ (function (_super) {
    __extends(AlfaFaq, _super);
    function AlfaFaq(props) {
        var _this = _super.call(this, props) || this;
        _this.onSearchTextChange = function (event, newValue) {
            _this.setState({ searchText: newValue || "" }, _this.updateExpandedItems);
            console.log("Zoektekst: " + newValue);
        };
        _this.updateExpandedItems = function () {
            var _a = _this.state, searchText = _a.searchText, items = _a.items;
            if (searchText) {
                var expandedItems = items
                    .filter(function (item) { return item[_this.props.accordianTitleColumn].toLowerCase().includes(searchText.toLowerCase()) ||
                    item[_this.props.accordianContentColumn].toLowerCase().includes(searchText.toLowerCase()); })
                    .map(function (item) { return item['ID']; });
                _this.setState({ expandedItems: expandedItems });
            }
            else {
                _this.setState({ expandedItems: [] });
            }
        };
        _this.highlightText = function (text, highlight) {
            if (!highlight)
                return text;
            var parts = text.split(new RegExp("(".concat(highlight, ")"), 'gi'));
            return parts.map(function (part, index) {
                return part.toLowerCase() === highlight.toLowerCase() ? "<mark>".concat(part, "</mark>") : part;
            }).join('');
        };
        _this.state = {
            items: [],
            choices: [],
            allowMultipleExpanded: _this.props.allowMultipleExpanded,
            allowZeroExpanded: _this.props.allowZeroExpanded,
            searchText: "",
            userEmail: "",
            expandedItems: []
        };
        _this._sp = getSP();
        _this.getListItems();
        _this.getUserEmail(); // Gebruikers-e-mail ophalen
        return _this;
    }
    AlfaFaq.prototype.getUserEmail = function () {
        return __awaiter(this, void 0, void 0, function () {
            var user, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._sp.web.currentUser()];
                    case 1:
                        user = _a.sent();
                        this.setState({ userEmail: user.Email }); // Gebruikers-e-mail opslaan in de state
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        console.error("Failed to fetch user email", error_1);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    AlfaFaq.prototype.getListItems = function () {
        var _this = this;
        if (this.props.listId && this.props.columnTitle) {
            var theAccordianList = this._sp.web.lists.getById(this.props.listId);
            theAccordianList.fields.getByInternalNameOrTitle(this.props.columnTitle).select("Choices")().then(function (field) {
                _this.setState({ choices: __spreadArray(["Alle"], field.Choices, true) });
            });
            var orderByQuery = '';
            if (this.props.accordianSortColumn) {
                orderByQuery = "<OrderBy>\n          <FieldRef Name='".concat(this.props.accordianSortColumn, "' ").concat(this.props.isSortDescending ? 'Ascending="False"' : '', " />\n        </OrderBy>");
            }
            var query = "<View>\n        <Query>\n          ".concat(orderByQuery, "\n        </Query>\n      </View>");
            theAccordianList.getItemsByCAMLQuery({ ViewXml: query }).then(function (results) {
                _this.setState({ items: results });
                console.dir(results);
            }).catch(function (error) {
                console.log("Failed to get list items!");
                console.log(error);
            });
        }
    };
    AlfaFaq.prototype.logItemOpen = function (itemId) {
        return __awaiter(this, void 0, void 0, function () {
            var userEmail, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log("LogItemOpen called for ID: " + itemId);
                        if (!(this.props.enableLogging && this.props.webhookUrl)) return [3 /*break*/, 7];
                        userEmail = this.state.userEmail;
                        if (!(itemId && userEmail)) return [3 /*break*/, 5];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        console.log("Posting webhook for id " + itemId);
                        return [4 /*yield*/, fetch(this.props.webhookUrl, {
                                method: "POST",
                                headers: {
                                    "Content-Type": "application/json",
                                },
                                body: JSON.stringify({
                                    id: itemId,
                                    email: userEmail, // E-mail toevoegen aan de payload
                                }),
                                mode: "no-cors", // Optioneel: gebruik dit als CORS een probleem is
                            })];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        error_2 = _a.sent();
                        console.error("Failed to log item open:", error_2);
                        return [3 /*break*/, 4];
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        // Log een bericht als een van de waarden ontbreekt
                        console.log("Skipping POST: Either itemId or userEmail is missing.");
                        _a.label = 6;
                    case 6: return [3 /*break*/, 8];
                    case 7:
                        console.log("Posting skipped for id " + itemId);
                        _a.label = 8;
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    AlfaFaq.prototype.componentDidUpdate = function (prevProps) {
        if (prevProps.listId !== this.props.listId) {
            this.getListItems();
        }
        if (prevProps.allowMultipleExpanded !== this.props.allowMultipleExpanded ||
            prevProps.allowZeroExpanded !== this.props.allowZeroExpanded) {
            this.setState({
                allowMultipleExpanded: this.props.allowMultipleExpanded,
                allowZeroExpanded: this.props.allowZeroExpanded,
            });
        }
    };
    AlfaFaq.prototype.render = function () {
        var _this = this;
        var listSelected = typeof this.props.listId !== "undefined" && this.props.listId.length > 0;
        var _a = this.state, allowMultipleExpanded = _a.allowMultipleExpanded, allowZeroExpanded = _a.allowZeroExpanded, searchText = _a.searchText, expandedItems = _a.expandedItems;
        return (React.createElement("div", { className: styles.alfaFaq },
            !listSelected && (React.createElement(Placeholder, { iconName: "ExpandAll", iconText: "Stel je wepbart in", description: "Kies een lijst met vragen en antwoorden om weer te geven.", buttonLabel: "Kies hier je lijst", onConfigure: this.props.onConfigure })),
            listSelected && (React.createElement("div", null,
                React.createElement(WebPartTitle, { displayMode: this.props.displayMode, title: "Kies hieronder het gewenste onderwerp.", updateProperty: this.props.updateProperty }),
                React.createElement(TextField, { placeholder: "Zoek...", onChange: this.onSearchTextChange, value: this.state.searchText }),
                React.createElement(Pivot, null, this.state.choices.map(function (category, index) { return (React.createElement(PivotItem, { headerText: category, key: index },
                    React.createElement(Accordion, { allowZeroExpanded: allowZeroExpanded, allowMultipleExpanded: allowMultipleExpanded, preExpanded: expandedItems, onChange: function (uuid) {
                            console.log("Accordionkey: " + uuid);
                            _this.logItemOpen(uuid.toString());
                        } },
                        _this.state.items
                            .filter(function (item) { return (category === "Alle" || item[_this.props.columnTitle] === category) &&
                            (!searchText || item[_this.props.accordianTitleColumn].toLowerCase().includes(searchText.toLowerCase()) ||
                                item[_this.props.accordianContentColumn].toLowerCase().includes(searchText.toLowerCase())); })
                            .map(function (item) { return (React.createElement(AccordionItem, { uuid: item['ID'], key: item.Id },
                            React.createElement(AccordionItemHeading, null,
                                React.createElement(AccordionItemButton, { title: item[_this.props.accordianTitleColumn], onClick: function () { console.log("Accordion item clicked!"); _this.logItemOpen(item.Id); } }, item[_this.props.accordianTitleColumn])),
                            React.createElement(AccordionItemPanel, null,
                                React.createElement("p", { dangerouslySetInnerHTML: { __html: _this.highlightText(item[_this.props.accordianContentColumn], searchText) } })))); }),
                        _this.state.items.filter(function (item) { return (category === "Alle" || item[_this.props.columnTitle] === category) &&
                            (!searchText || item[_this.props.accordianTitleColumn].toLowerCase().includes(searchText.toLowerCase()) ||
                                item[_this.props.accordianContentColumn].toLowerCase().includes(searchText.toLowerCase())); }).length === 0 && (React.createElement("p", null, "Deze categorie bevat geen vragen, of je zoekopdracht heeft geen resultaten."))))); }))))));
    };
    return AlfaFaq;
}(React.Component));
export default AlfaFaq;
//# sourceMappingURL=AlfaFaq.js.map