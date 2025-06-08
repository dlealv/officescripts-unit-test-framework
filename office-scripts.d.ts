// office-scripts.d.ts
//
// - Declares a minimal subset of the Office Scripts ExcelScript namespace for use in local development and testing.
// - This file provides enough typing to write and test scripts that interact with Excel ranges, worksheets, and workbooks.
// - Only selected types, enums, and interfaces are included for simplicity and to avoid overcomplicating local mocks.
// - Replace 'any' with more precise types as needed for stricter type checking.
//
// Field documentation and rationale:
// - Range: Declared as 'any' to maximize flexibility for mocking and to avoid implementation detail here.
// - ClearApplyTo and VerticalAlignment: Minimal enums for supported operations.
// - Worksheet and Workbook: Only essential methods are included for getting ranges and worksheet names.

declare namespace ExcelScript {
  export type Range = any; // Represents a cell range in Excel. Use 'any' for flexibility in mocks.

  export enum ClearApplyTo {
    contents = "contents" // Indicates clearing only the contents of a range.
  }

  export enum VerticalAlignment {
    center = "center" // Indicates vertical alignment is centered.
  }

  export interface Worksheet {
    getRange(address: string): Range; // Returns a Range object for the specified address (e.g., "A1:B2").
    getName(): string; // Returns the worksheet's name.
  }

  export interface Workbook {
    getWorksheet(name: string): Worksheet; // Returns the worksheet with the specified name.
    getActiveWorksheet(): Worksheet; // Returns the currently active worksheet.
  }
}