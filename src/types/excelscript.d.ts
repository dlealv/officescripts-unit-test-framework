declare namespace ExcelScript {
  enum ClearApplyTo {
    contents = "contents"
  }
  enum VerticalAlignment {
    center = "center"
  }
  interface Range {
    clear(what: ClearApplyTo): void;
    getFormat(): {
      setVerticalAlignment(v: VerticalAlignment): void;
      getFont(): { setColor(color: string): void };
    };
    setValue(val: any): void;
    getValue(): any;
    getAddress(): string;
    getCellCount(): number;
  }
  interface Worksheet {
    getRange(address: string): Range;
    getName(): string;
  }
  interface Workbook {
    getWorksheet(name: string): Worksheet;
    getActiveWorksheet(): Worksheet;
  }
}