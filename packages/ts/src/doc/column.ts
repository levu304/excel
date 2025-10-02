import type { Worksheet } from "./worksheet";

export class Column {
  private _header: string | null;

  constructor(worksheet: Worksheet, number: number, defn: () => {} | boolean) {
    if (defn !== false) {
        this.defn = defn;
    }
  }

  get defn() {
    return {
      header: this._header,
      key: this.key,
      width: this.width,
      style: this.style,
      hidden: this.hidden,
      outlineLevel: this.outlineLevel,
    };
  }
}
