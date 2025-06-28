// File: types/office-scripts/index.d.ts
// - Declares a minimal subset of the Office Scripts ExcelScript namespace for use in local development and testing.
// - This file provides enough typing to write and test scripts that interact with Excel ranges, worksheets, and workbooks.
// - Only selected types, enums, and interfaces are included for simplicity and to avoid overcomplicating local mocks.
//
// Field documentation and rationale:
// - Workbook: Only the class declaration is included for basic compatibility.
// - Expand this file with more classes, interfaces, or enums as your test and implementation needs grow.

declare namespace ExcelScript {
  class Workbook {
  }
}