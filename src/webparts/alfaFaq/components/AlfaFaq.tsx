import * as React from "react";
import styles from "./AlfaFaq.module.scss";
import { IAlfaFaqProps } from "./IAlfaFaqProps";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import "./alfaFaq.css";
import { Pivot, PivotItem, TextField } from '@fluentui/react';
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from "react-accessible-accordion";
import { getSP } from "../../../utils/pnpjs-config";

interface IAccordionItem {
  Id: string;
  Title: string;
  Content: string;
  Category: string | string[];
  helpful?: boolean;
}

export interface IAlfaFaqState {
  items: Array<IAccordionItem>;
  choices: Array<string>;
  allowMultipleExpanded: boolean;
  allowZeroExpanded: boolean;
  searchText: string; // Nieuwe state voor zoektekst
  userEmail: string;
  expandedItems: string[];
}

export default class AlfaFaq extends React.Component<
  IAlfaFaqProps,
  IAlfaFaqState
> {

  private _sp: SPFI;

  constructor(props: IAlfaFaqProps) {
    super(props);

    this.state = {
      items: [],
      choices: [],
      allowMultipleExpanded: this.props.allowMultipleExpanded,
      allowZeroExpanded: this.props.allowZeroExpanded,
      searchText: "",
      userEmail: "",  // Toevoegen van state voor userEmail
      expandedItems: []
    };

    this._sp = getSP();
    this.getListItems();

    this.getUserEmail(); // Gebruikers-e-mail ophalen
  }

  private async getUserEmail(): Promise<void> {
    try {
      const user = await this._sp.web.currentUser(); // Huidige gebruiker ophalen
      this.setState({ userEmail: user.Email }); // Gebruikers-e-mail opslaan in de state
    } catch (error) {
      console.error("Failed to fetch user email", error);
    }
  }
  private onSearchTextChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ searchText: newValue || "" }, this.updateExpandedItems);
    console.log("Zoektekst: "+newValue);
  }

  private updateExpandedItems = () => {
    const { searchText, items } = this.state;
    if (searchText) {
      const expandedItems = items
        .filter(item => item[this.props.accordianTitleColumn].toLowerCase().includes(searchText.toLowerCase()) ||
                        item[this.props.accordianContentColumn].toLowerCase().includes(searchText.toLowerCase()))
        .map(item => item['ID']);
      this.setState({ expandedItems });
    } else {
      this.setState({ expandedItems: [] });
    }
  }

  private highlightText = (text: string, highlight: string): string => {
    if (!highlight) return text;
    const parts = text.split(new RegExp(`(${highlight})`, 'gi'));
    return parts.map((part, index) => 
      part.toLowerCase() === highlight.toLowerCase() ? `<mark>${part}</mark>` : part
    ).join('');
  }

  private getListItems(): void {
    if (this.props.listId && this.props.columnTitle) {
      const theAccordianList = this._sp.web.lists.getById(this.props.listId);
      theAccordianList.fields.getByInternalNameOrTitle(this.props.columnTitle).select("Choices")().then(field => {
        this.setState({ choices: ["Alle", ...field.Choices] });
      });

      let orderByQuery = '';
      if (this.props.accordianSortColumn) {
        orderByQuery = `<OrderBy>
          <FieldRef Name='${this.props.accordianSortColumn}' ${this.props.isSortDescending ? 'Ascending="False"' : ''} />
        </OrderBy>`;
      }

      const query = `<View>
        <Query>
          ${orderByQuery}
        </Query>
      </View>`;

      theAccordianList.getItemsByCAMLQuery({ ViewXml: query }).then((results: Array<IAccordionItem>) => {
        this.setState({ items: results });
        console.dir(results);
      }).catch(error => {
        console.log("Failed to get list items!");
        console.log(error);
      });
    }
  }
  private async logItemOpen(itemId: string): Promise<void> {
    console.log("LogItemOpen called for ID: " + itemId);
    
    // Check of logging is ingeschakeld en een webhook URL is ingesteld
    if (this.props.enableLogging && this.props.webhookUrl) {
      const { userEmail } = this.state;  // Gebruikers-e-mail ophalen uit de state
  
      // Check of itemId en userEmail beiden bestaan
      if (itemId && userEmail) {
        try {
          console.log("Posting webhook for id " + itemId);
  
          await fetch(this.props.webhookUrl, {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              id: itemId,
              email: userEmail,  // E-mail toevoegen aan de payload
            }),
            mode: "no-cors",  // Optioneel: gebruik dit als CORS een probleem is
          });
        } catch (error) {
          console.error("Failed to log item open:", error);
        }
      } else {
        // Log een bericht als een van de waarden ontbreekt
        console.log("Skipping POST: Either itemId or userEmail is missing.");
      }
    } else {
      console.log("Posting skipped for id " + itemId);
    }
  }

  /**
   * Controleert of een item in een bepaalde categorie valt
   * Werkt zowel met single-select als multi-select velden
   */
  private isItemInCategory(item: IAccordionItem, category: string): boolean {
    if (category === "Alle") return true;
    
    // Debug informatie printen om te zien wat er in het item zit
    console.log("Item kategorie check:", {
      item: item,
      category: category,
      columnTitle: this.props.columnTitle,
      itemKeys: Object.keys(item),
      itemValue: item[this.props.columnTitle],
      hasMultiChoice: item["Meerderekeuzea"] !== undefined
    });
    
    // Probeer eerst de geconfigureerde kolom
    let itemCategory = item[this.props.columnTitle];
    
    // Als dat niet werkt, probeer dan ook "Meerderekeuzea" als fallback voor multi-select
    if ((itemCategory === null || itemCategory === undefined) && item["Meerderekeuzea"] !== undefined) {
      console.log("Fallback naar Meerderekeuzea veld");
      itemCategory = item["Meerderekeuzea"];
    }
    
    // Als het nog steeds null of undefined is
    if (itemCategory === null || itemCategory === undefined) {
      console.log("Geen categorie gevonden voor dit item");
      return false;
    }
    
    // Als het een array is (multi-select)
    if (Array.isArray(itemCategory)) {
      const result = itemCategory.some(value => value === category);
      console.log(`Multi-select check: ${itemCategory.join(',')} bevat ${category}? ${result}`);
      return result;
    }
    
    // Als het een string is (single-select)
    console.log(`Single-select check: ${itemCategory} == ${category}? ${itemCategory === category}`);
    return itemCategory === category;
  }

  public componentDidUpdate(prevProps: IAlfaFaqProps): void {
    if (prevProps.listId !== this.props.listId) {
      this.getListItems();
    }

    if (
      prevProps.allowMultipleExpanded !== this.props.allowMultipleExpanded ||
      prevProps.allowZeroExpanded !== this.props.allowZeroExpanded
    ) {
      this.setState({
        allowMultipleExpanded: this.props.allowMultipleExpanded,
        allowZeroExpanded: this.props.allowZeroExpanded,
      });
    }
  }

  public render(): React.ReactElement<IAlfaFaqProps> {
    const listSelected: boolean = typeof this.props.listId !== "undefined" && this.props.listId.length > 0;
    const { allowMultipleExpanded, allowZeroExpanded, searchText, expandedItems } = this.state;
    return (
      <div className={styles.alfaFaq}>
        {!listSelected && (
          <Placeholder
            iconName="ExpandAll"
            iconText="Stel je wepbart in"
            description="Kies een lijst met vragen en antwoorden om weer te geven."
            buttonLabel="Kies hier je lijst"
            onConfigure={this.props.onConfigure}
          />
        )}
        {listSelected && (
          <div>
            <WebPartTitle
              displayMode={this.props.displayMode}
              title="Kies hieronder het gewenste onderwerp."
              updateProperty={this.props.updateProperty}
            />
            <TextField
              placeholder="Zoek..."
              onChange={this.onSearchTextChange}
              value={this.state.searchText}
            />
            <Pivot>
              {this.state.choices.map((category, index) => (
                <PivotItem headerText={category} key={index}>
                  <Accordion
                    allowZeroExpanded={allowZeroExpanded}
                    allowMultipleExpanded={allowMultipleExpanded}
                    preExpanded={expandedItems}
                    onChange={ (uuid) => { 
                      console.log("Accordionkey: "+uuid);
                      this.logItemOpen(uuid.toString()); 
                    }} 
                  >
                    {this.state.items
                      .filter(item => this.isItemInCategory(item, category) && 
                                      (!searchText || item[this.props.accordianTitleColumn].toLowerCase().includes(searchText.toLowerCase()) || 
                                      item[this.props.accordianContentColumn].toLowerCase().includes(searchText.toLowerCase())))
                      .map((item: IAccordionItem) => (
                        <AccordionItem uuid={item['ID']} key={item.Id}>
                          <AccordionItemHeading>
                            <AccordionItemButton 
                              title={item[this.props.accordianTitleColumn]} 
                              onClick={() => { console.log("Accordion item clicked!"); this.logItemOpen(item.Id); }}
                            >
                              {item[this.props.accordianTitleColumn]}
                            </AccordionItemButton>
                          </AccordionItemHeading>
                          <AccordionItemPanel>
                            <p dangerouslySetInnerHTML={{ __html: this.highlightText(item[this.props.accordianContentColumn], searchText) }} />
                          </AccordionItemPanel>
                        </AccordionItem>
                      ))}
                      {this.state.items.filter(item => this.isItemInCategory(item, category) && 
                                      (!searchText || item[this.props.accordianTitleColumn].toLowerCase().includes(searchText.toLowerCase()) || 
                                      item[this.props.accordianContentColumn].toLowerCase().includes(searchText.toLowerCase()))).length === 0 && (
                      <p>Deze categorie bevat geen vragen, of je zoekopdracht heeft geen resultaten.</p>
                    )}
                  </Accordion>
                </PivotItem>
              ))}
            </Pivot>
          </div>
        )}
      </div>
    );
  }
}
