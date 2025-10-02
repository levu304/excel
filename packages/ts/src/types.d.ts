export type PageSetup = {
  /**
   * Whitespace on the borders of the page. Units are inches.
   */
  margins: Margins;

  /**
   * Orientation of the page - i.e. taller (`'portrait'`) or wider (`'landscape'`).
   *
   * `'portrait'` by default
   */
  orientation: "portrait" | "landscape";

  /**
   * Horizontal Dots per Inch. Default value is 4294967295
   */
  horizontalDpi: number;

  /**
   * Vertical Dots per Inch. Default value is 4294967295
   */
  verticalDpi: number;

  /**
   * Whether to use fitToWidth and fitToHeight or scale settings.
   *
   * Default is based on presence of these settings in the pageSetup object - if both are present,
   * scale wins (i.e. default will be false)
   */
  fitToPage: boolean;

  /**
   * How many pages wide the sheet should print on to. Active when fitToPage is true
   *
   * Default is 1
   */
  fitToWidth: number;

  /**
   * How many pages high the sheet should print on to. Active when fitToPage is true
   *
   * Default is 1
   */
  fitToHeight: number;

  /**
   * Percentage value to increase or reduce the size of the print. Active when fitToPage is false
   *
   * Default is 100
   */
  scale: number;

  /**
   * Which order to print the pages.
   *
   * Default is `downThenOver`
   */
  pageOrder: "downThenOver" | "overThenDown";

  /**
   * Print without colour
   *
   * false by default
   */
  blackAndWhite: boolean;

  /**
   * Print with less quality (and ink)
   *
   * false by default
   */
  draft: boolean;

  /**
   * Where to place comments
   *
   * Default is `None`
   */
  cellComments: "atEnd" | "asDisplayed" | "None";

  /**
   * Where to show errors
   *
   * Default is `displayed`
   */
  errors: "dash" | "blank" | "NA" | "displayed";

  /**
   * 	What paper size to use (see below)
   *
   * | Name                          | Value       |
   * | ----------------------------- | ---------   |
   * | Letter                        | `undefined` |
   * | Legal                         |  `5`        |
   * | Executive                     |  `7`        |
   * | A4                            |  `9`        |
   * | A5                            |  `11`       |
   * | B5 (JIS)                      |  `13`       |
   * | Envelope #10                  |  `20`       |
   * | Envelope DL                   |  `27`       |
   * | Envelope C5                   |  `28`       |
   * | Envelope B5                   |  `34`       |
   * | Envelope Monarch              |  `37`       |
   * | Double Japan Postcard Rotated |  `82`       |
   * | 16K 197x273 mm                |  `119`      |
   */
  paperSize: PaperSize;

  /**
   * Whether to show the row numbers and column letters, `false` by default
   */
  showRowColHeaders: boolean;

  /**
   * Whether to show grid lines, `false` by default
   */
  showGridLines: boolean;

  /**
   * Which number to use for the first page
   */
  firstPageNumber: number;

  /**
   * 	Whether to center the sheet data horizontally, `false` by default
   */
  horizontalCentered: boolean;

  /**
   * 	Whether to center the sheet data vertically, `false` by default
   */
  verticalCentered: boolean;

  /**
   * Set Print Area for a sheet, e.g. `'A1:G20'`
   */
  printArea: string;

  /**
   * Repeat specific rows on every printed page, e.g. `'1:3'`
   */
  printTitlesRow: string;

  /**
   * Repeat specific columns on every printed page, e.g. `'A:C'`
   */
  printTitlesColumn: string;
};

export type Color = {
  /**
   * Hex string for alpha-red-green-blue e.g. FF00FF00
   */
  argb: string;

  /**
   * Choose a theme by index
   */
  theme: number;
};

export type WorksheetViewCommon = {
  /**
   * Sets the worksheet view's orientation to right-to-left, `false` by default
   */
  rightToLeft: boolean;

  /**
   * The currently selected cell
   */
  activeCell: string;

  /**
   * Shows or hides the ruler in Page Layout, `true` by default
   */
  showRuler: boolean;

  /**
   * Shows or hides the row and column headers (e.g. A1, B1 at the top and 1,2,3 on the left,
   * `true` by default
   */
  showRowColHeaders: boolean;

  /**
   * Shows or hides the gridlines (shown for cells where borders have not been defined),
   * `true` by default
   */
  showGridLines: boolean;

  /**
   * 	Percentage zoom to use for the view, `100` by default
   */
  zoomScale: number;

  /**
   * 	Normal zoom for the view, `100` by default
   */
  zoomScaleNormal: number;
};

export type WorksheetViewNormal = {
  /**
   * Controls the view state
   */
  state: "normal";

  /**
   * Presentation style
   */
  style: "pageBreakPreview" | "pageLayout";
};

export type WorksheetViewFrozen = {
  /**
   * Where a number of rows and columns to the top and left are frozen in place.
   * Only the bottom left section will scroll
   */
  state: "frozen";

  /**
   * Presentation style
   */
  style?: "pageBreakPreview";

  /**
   * How many columns to freeze. To freeze rows only, set this to 0 or undefined
   */
  xSplit?: number;

  /**
   * How many rows to freeze. To freeze columns only, set this to 0 or undefined
   */
  ySplit?: number;

  /**
   * Which cell will be top-left in the bottom-right pane. Note: cannot be a frozen cell.
   * Defaults to first unfrozen cell
   */
  topLeftCell?: string;
};

export type WorksheetViewSplit = {
  /**
   * Where the view is split into 4 sections, each semi-independently scrollable.
   */
  state: "split";

  /**
   * Presentation style
   */
  style?: "pageBreakPreview" | "pageLayout";

  /**
   * How many points from the left to place the splitter.
   * To split vertically, set this to 0 or undefined
   */
  xSplit?: number;

  /**
   * How many points from the top to place the splitter.
   * To split horizontally, set this to 0 or undefined
   */
  ySplit?: number;

  /**
   * Which cell will be top-left in the bottom-right pane
   */
  topLeftCell?: string;

  /**
   * Which pane will be active
   */
  activePane?: "topLeft" | "topRight" | "bottomLeft" | "bottomRight";
};

export type WorksheetView = WorksheetViewCommon &
  (WorksheetViewNormal | WorksheetViewFrozen | WorksheetViewSplit);

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
  properties: WorksheetProperties;
  pageSetup: PageSetup;
  headerFooter: Partial<HeaderFooter>;
  views?: Partial<WorksheetView>[];
  state?: WorksheetState;
  autoFilter?: AutoFilter;
};

export type HeaderFooter = {
  /**
   * Set the value of differentFirst as true, which indicates that headers/footers for first page are different from the other pages, `false` by default
   */
  differentFirst: boolean;
  /**
   * Set the value of differentOddEven as true, which indicates that headers/footers for odd and even pages are different, `false` by default
   */
  differentOddEven: boolean;
  /**
   * Set header string for odd pages, could format the string and `null` by default
   */
  oddHeader: string;
  /**
   * Set footer string for odd pages, could format the string and `null` by default
   */
  oddFooter: string;
  /**
   * Set header string for even pages, could format the string and `null` by default
   */
  evenHeader: string;
  /**
   * Set footer string for even pages, could format the string and `null` by default
   */
  evenFooter: string;
  /**
   * Set header string for the first page, could format the string and `null` by default
   */
  firstHeader: string;
  /**
   * Set footer string for the first page, could format the string and `null` by default
   */
  firstFooter: string;
};

export type AutoFilter =
  | string
  | {
      from: string | { row: number; column: number };
      to: string | { row: number; column: number };
    };

export type CfvoTypes =
  | "percentile"
  | "percent"
  | "num"
  | "min"
  | "max"
  | "formula"
  | "autoMin"
  | "autoMax";

export interface Cvfo {
  type: CfvoTypes;
  value?: number;
}

export type Font = {
	name: string;
	size: number;
	family: number;
	scheme: 'minor' | 'major' | 'none';
	charset: number;
	color: Partial<Color>;
	bold: boolean;
	italic: boolean;
	underline: boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
	vertAlign: 'superscript' | 'subscript';
	strike: boolean;
	outline: boolean;
}

export type Style = {
	numFmt: string;
	font: Partial<Font>;
	alignment: Partial<Alignment>;
	protection: Partial<Protection>;
	border: Partial<Borders>;
	fill: Fill;
}

export interface ConditionalFormattingBaseRule {
  priority: number;
  style?: Partial<Style>;
}
export interface ExpressionRuleType extends ConditionalFormattingBaseRule {
  type: "expression";
  formulae?: any[];
}

export interface CellIsRuleType extends ConditionalFormattingBaseRule {
  type: "cellIs";
  formulae?: any[];
  operator?: CellIsOperators;
}

export interface Top10RuleType extends ConditionalFormattingBaseRule {
  type: "top10";
  rank: number;
  percent: boolean;
  bottom: boolean;
}

export interface AboveAverageRuleType extends ConditionalFormattingBaseRule {
  type: "aboveAverage";
  aboveAverage: boolean;
}

export interface ColorScaleRuleType extends ConditionalFormattingBaseRule {
  type: "colorScale";
  cfvo?: Cvfo[];
  color?: Partial<Color>[];
}

export interface IconSetRuleType extends ConditionalFormattingBaseRule {
  type: "iconSet";
  showValue?: boolean;
  reverse?: boolean;
  custom?: boolean;
  iconSet?: IconSetTypes;
  cfvo?: Cvfo[];
}

export interface ContainsTextRuleType extends ConditionalFormattingBaseRule {
  type: "containsText";
  operator?: ContainsTextOperators;
  text?: string;
}

export interface TimePeriodRuleType extends ConditionalFormattingBaseRule {
  type: "timePeriod";
  timePeriod?: TimePeriodTypes;
}

export type DataBarRuleType = ConditionalFormattingBaseRule & {
  type: "dataBar";
  gradient?: boolean;
  minLength?: number;
  maxLength?: number;
  showValue?: boolean;
  border?: boolean;
  negativeBarColorSameAsPositive?: boolean;
  negativeBarBorderColorSameAsPositive?: boolean;
  axisPosition?: "auto" | "middle" | "none";
  direction?: "context" | "leftToRight" | "rightToLeft";
  cfvo?: Cvfo[];
};

export type ConditionalFormattingRule =
  | ExpressionRuleType
  | CellIsRuleType
  | Top10RuleType
  | AboveAverageRuleType
  | ColorScaleRuleType
  | IconSetRuleType
  | ContainsTextRuleType
  | TimePeriodRuleType
  | DataBarRuleType;

export type ConditionalFormattingOptions = {
  ref: string;
  rules: ConditionalFormattingRule[];
};

export type WorksheetProtection = {
	objects: boolean;
	scenarios: boolean;
	selectLockedCells: boolean;
	selectUnlockedCells: boolean;
	formatCells: boolean;
	formatColumns: boolean;
	formatRows: boolean;
	insertColumns: boolean;
	insertRows: boolean;
	insertHyperlinks: boolean;
	deleteColumns: boolean;
	deleteRows: boolean;
	sort: boolean;
	autoFilter: boolean;
	pivotTables: boolean;
	spinCount: number;
}

export type Image = {
	extension: 'jpeg' | 'png' | 'gif';
	base64?: string;
	filename?: string;
	buffer?: Buffer;
}
