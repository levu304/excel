import type { Color, PageSetup } from ":/types";
import type { Workbook } from "./workbook";

export type WorksheetProperties = {
  /**
   * Color of the tab
   */
  tabColor: Partial<Color>;

  /**
   * The worksheet column outline level (default: 0)
   */
  outlineLevelCol: number;

  /**
   * The worksheet row outline level (default: 0)
   */
  outlineLevelRow: number;

  /**
   * The outline properties which controls how it will summarize rows and columns
   */
  outlineProperties: {
    summaryBelow: boolean;
    summaryRight: boolean;
  };
  /**
   * Default row height (default: 15)
   */
  defaultRowHeight: number;

  /**
   * Default column width (optional)
   */
  defaultColWidth?: number;

  /**
   * default: 55
   */
  dyDescent: number;
  showGridLines: boolean;
};

export type WorksheetState = "visible" | "hidden" | "veryHidden";

export type WorksheetOptions = {
  workbook: Workbook;
  id: string;
  orderNo: number;
  name: string;
  state: WorksheetState;
  properties: WorksheetProperties;
  pageSetup: PageSetup;
};

export class Worksheet {
  constructor(options: Partial<WorksheetOptions>) {}
}
