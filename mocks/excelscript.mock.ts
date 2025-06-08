// excelscript.mock.ts
//
// - This file provides a minimal mock implementation of the ExcelScript namespace for local testing of Office Scripts.
// - The mocks simulate enough ExcelScript behavior to allow isolated unit tests and logger validation without the Office Scripts runtime.
// - For more realism, expand methods or add features as needed for your test scenarios.
//
// Notes on implementation:
// - Enum values (ClearApplyTo, VerticalAlignment) match those in the real ExcelScript API.
// - The Range class holds a value and address, with methods to clear and format the cell, and to get/set values.
// - Worksheet manages named ranges and provides access by address, creating them as needed.
// - Workbook manages named worksheets and can retrieve them by name or return the "active" worksheet (first sheet).
// - At the end, the namespace is attached to globalThis for compatibility with test environments (Node, ts-node).

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

// Make ExcelScript available globally for Node/ts-node test environments
if (typeof globalThis !== "undefined" && typeof ExcelScript !== "undefined") {
  // @ts-ignore
  globalThis.ExcelScript = ExcelScript;
}