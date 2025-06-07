declare namespace ExcelScript {
  export type Range = any;
  export enum ClearApplyTo {
    contents = "contents"
  }
  export enum VerticalAlignment {
    center = "center"
  }
  export interface Worksheet {
    getRange(address: string): Range;
    getName(): string;
  }
  export interface Workbook {
    getWorksheet(name: string): Worksheet;
    getActiveWorksheet(): Worksheet;
  }
}