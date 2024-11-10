import { DisplayMode } from "@microsoft/sp-core-library";

export interface IAlfaFaqProps {
  listId: string;
  accordionTitle: string;
  columnTitle: string;
  //selectedChoice: string;
  accordianTitleColumn: string;
  accordianContentColumn: string;
  accordianSortColumn: string;
  isSortDescending: boolean;
  allowZeroExpanded: boolean;
  allowMultipleExpanded: boolean;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  onConfigure: () => void;
  webhookUrl: string; // Nieuwe property voor webhook URL
  enableLogging: boolean; // Nieuwe property voor logging toggle
}
