import type { Worksheet } from "./worksheet";

export class Workbook {
  public _worksheets: Worksheet[];

  constructor() {
    this._worksheets = [];
  }
}
