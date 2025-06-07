export namespace ExcelScript {
  export enum ClearApplyTo {
    contents = "contents"
  }
  export enum VerticalAlignment {
    center = "center"
  }
  export class Range {
    private value: any = "";
    private address: string;
    constructor(address: string = "A1") {
      this.address = address;
    }
    clear(_what: ClearApplyTo) {
      this.value = "";
    }
    getFormat() {
      return {
        setVerticalAlignment: (_v: VerticalAlignment) => {},
        getFont: () => ({
          setColor: (_color: string) => {}
        })
      };
    }
    setValue(val: any) {
      this.value = val;
    }
    getValue() {
      return this.value;
    }
    getAddress() {
      return this.address;
    }
    getCellCount() {
      return 1;
    }
  }

  export class Worksheet {
    private name: string;
    private ranges: { [addr: string]: Range } = {};
    constructor(name: string = "Sheet1") {
      this.name = name;
    }
    getRange(address: string) {
      if (!this.ranges[address]) {
        this.ranges[address] = new Range(address);
      }
      return this.ranges[address];
    }
    getName() {
      return this.name;
    }
  }

  export class Workbook {
    private sheets: { [name: string]: Worksheet } = {};
    constructor(sheetNames: string[] = ["Sheet1"]) {
      for (const name of sheetNames) {
        this.sheets[name] = new Worksheet(name);
      }
    }
    getWorksheet(name: string) {
      if (!this.sheets[name]) {
        this.sheets[name] = new Worksheet(name);
      }
      return this.sheets[name];
    }
    getActiveWorksheet() {
      // Return the first worksheet by default
      const sheetNames = Object.keys(this.sheets);
      return this.sheets[sheetNames[0]];
    }
  }
}