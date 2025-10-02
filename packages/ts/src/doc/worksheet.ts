import type {
  AutoFilter,
  ConditionalFormattingOptions,
  HeaderFooter,
  Image,
  PageSetup,
  WorksheetOptions,
  WorksheetProperties,
  WorksheetProtection,
  WorksheetState,
  WorksheetView,
} from ":/types";
import type { Workbook } from "./workbook";
import type { Row } from "./row";
import { Column } from "./column";
import { DataValidations } from "./data-validations";

export class Worksheet {
  private _workbook: Workbook;
  private id: string;
  private orderNo: number;
  private _name: string;
  private state: WorksheetState;
  private _rows: Row[];
  private _columns: Column[];
  private _keys: Record<string, Column[]>;
  private _merges: Record<string, unknown>;
  private rowBreaks: unknown[];
  private properties: WorksheetProperties;
  private pageSetup: Partial<PageSetup>;
  private headerFooter: Partial<HeaderFooter>;
  private views: Partial<WorksheetView>[];
  private autoFilter: AutoFilter | null;
  private dataValidations: DataValidations;
  private conditionalFormattings: ConditionalFormattingOptions[];
  private pivotTables: unknown[];
  private tables: Record<string, unknown>;
  private sheetProtection: WorksheetProtection | null;
  private _media: Image[];

  constructor(options: WorksheetOptions) {
    this._workbook = options.workbook;
    this.id = options.id;
    this.orderNo = options.orderNo;

    // and a name
    this._name = options.name; 

    // add a state
    this.state = options.state || "visible";

    // rows allows access organised by row. Sparse array of arrays indexed by row-1, col
    // Note: _rows is zero based. Must subtract 1 to go from cell.row to index
    this._rows = [];

    // column definitions
    this._columns = [];

    // column keys (addRow convenience): key ==> this._collumns index
    this._keys = {};

    // keep record of all merges
    this._merges = {};

    // record of all row and column pageBreaks
    this.rowBreaks = [];

    // for tabColor, default row height, outline levels, etc
    this.properties = Object.assign(
      {},
      {
        defaultRowHeight: 15,
        dyDescent: 55,
        outlineLevelCol: 0,
        outlineLevelRow: 0,
      },
      options.properties
    );

    // for all things printing
    this.pageSetup = Object.assign(
      {},
      {
        margins: {
          left: 0.7,
          right: 0.7,
          top: 0.75,
          bottom: 0.75,
          header: 0.3,
          footer: 0.3,
        },
        orientation: "portrait",
        horizontalDpi: 4294967295,
        verticalDpi: 4294967295,
        fitToPage: !!(
          options.pageSetup &&
          (options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) &&
          !options.pageSetup.scale
        ),
        pageOrder: "downThenOver",
        blackAndWhite: false,
        draft: false,
        cellComments: "None",
        errors: "displayed",
        scale: 100,
        fitToWidth: 1,
        fitToHeight: 1,
        paperSize: undefined,
        showRowColHeaders: false,
        showGridLines: false,
        firstPageNumber: undefined,
        horizontalCentered: false,
        verticalCentered: false,
        rowBreaks: null,
        colBreaks: null,
      },
      options.pageSetup
    );

    this.headerFooter = Object.assign(
      {},
      {
        differentFirst: false,
        differentOddEven: false,
        oddHeader: null,
        oddFooter: null,
        evenHeader: null,
        evenFooter: null,
        firstHeader: null,
        firstFooter: null,
      },
      options.headerFooter
    );

    this.dataValidations = new DataValidations();

    // for freezepanes, split, zoom, gridlines, etc
    this.views = options?.views || [];

    this.autoFilter = options.autoFilter || null;

    // for images, etc
    this._media = [];

    // worksheet protection
    this.sheetProtection = null;

    // for tables
    this.tables = {};

    this.pivotTables = [];

    this.conditionalFormattings = [];
  }

  get name() {
    return this._name;
  }

  set name(name) {
    if (name === undefined) {
      name = `sheet${this.id}`;
    }

    if (this._name === name) return;

    if (typeof name !== 'string') {
      throw new Error('The name has to be a string.');
    }

    if (name === '') {
      throw new Error('The name can\'t be empty.');
    }

    if (name === 'History') {
      throw new Error('The name "History" is protected. Please use a different name.');
    }

    // Illegal character in worksheet name: asterisk (*), question mark (?),
    // colon (:), forward slash (/ \), or bracket ([])
    if (/[*?:/\\[\]]/.test(name)) {
      throw new Error(`Worksheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`);
    }

    if (/(^')|('$)/.test(name)) {
      throw new Error(`The first or last character of worksheet name cannot be a single quotation mark: ${name}`);
    }

    if (name && name.length > 31) {
      // eslint-disable-next-line no-console
      console.warn(`Worksheet name ${name} exceeds 31 chars. This will be truncated`);
      name = name.substring(0, 31);
    }

    if (this._workbook._worksheets.find(ws => ws && ws.name.toLowerCase() === name.toLowerCase())) {
      throw new Error(`Worksheet name already exists: ${name}`);
    }

    this._name = name;
  }

  get workbook() {
    return this._workbook;
  }

  get columns() {
    return this._columns;
  }

  // set the columns from an array of column definitions.
  // Note: any headers defined will overwrite existing values.
  set columns(value: Column[]) {
    // calculate max header row count
    this._headerRowCount = value.reduce((pv, cv) => {
      const headerCount = (cv.header && 1) || (cv.headers && cv.headers.length) || 0;
      return Math.max(pv, headerCount);
    }, 0);

    // construct Column objects
    let count = 1;
    const columns = (this._columns = []);
    value.forEach(defn => {
      const column = new Column(this, count++, false);
      columns.push(column);
      column.defn = defn;
    });
  }
}