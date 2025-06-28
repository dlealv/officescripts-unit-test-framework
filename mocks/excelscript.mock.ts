//
// - This file provides a minimal mock implementation of the ExcelScript namespace for local testing of Office Scripts.
// - The mocks simulate enough ExcelScript behavior to allow isolated unit tests and logger validation without the Office Scripts runtime.
// - For more realism, expand methods or add features as needed for your test scenarios.
//
// Usage:
// - This file is intended for use in local Node.js/TypeScript test environments and CI pipelines.
// - It should be loaded or referenced by main-wrapper.ts before running any Office Script-based tests.
// - Example: let workbook = new ExcelScript.Workbook()
// - Extend this mock as your tests require additional ExcelScript features.
//
// Notes:
// - No import/export keywords are used for Office Script compatibility.
// - Attach to globalThis if needed for Node.js/ts-node environments.

export namespace ExcelScript {
  export class Workbook {
  }
}